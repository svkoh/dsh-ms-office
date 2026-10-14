import { test } from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import { readFileSync } from 'node:fs';

const source = (name) => readFileSync(new URL(`../server/actions/${name}.js`, import.meta.url), 'utf8');
const plain = (value) => JSON.parse(JSON.stringify(value));
const loadable = (value = {}) => Object.assign(value, { load() { return this; } });

// Load host fragments directly: registry metadata is integrated independently.
function hostHarness(host, context, { util = {}, supported = () => true, globals = {} } = {}) {
  const utilities = { okResult: (result) => ({ ok: true, result }), errResult: (code, error) => ({ ok: false, code, error }), ...util };
  const sandbox = { window: { __UTIL__: utilities }, Office: { context: { requirements: { isSetSupported: supported } } }, [host]: { run: async (callback) => callback(context) }, ...globals };
  vm.createContext(sandbox);
  vm.runInContext(source('runtime'), sandbox);
  vm.runInContext(source(host === 'Excel' ? 'excel' : 'ppt'), sandbox);
  return sandbox.window.__OFFICE_ACTIONS__.hosts[host];
}

function tracked(value, writes, prefix) {
  return new Proxy(loadable(value), { set(target, key, next) { writes.push([`${prefix}.${String(key)}`, next]); target[key] = next; return true; } });
}

function excelHarness(options = {}) {
  const writes = [], calls = [], loads = [];
  const format = tracked({ columnWidth: 80, rowHeight: 20, wrapText: false, horizontalAlignment: 'General', verticalAlignment: 'Bottom', font: tracked({ name: 'Calibri', size: 11, bold: false }, writes, 'font'), fill: tracked({ color: '#FFFFFF' }, writes, 'fill'), autofitColumns() { calls.push(['autofitColumns']); }, autofitRows() { calls.push(['autofitRows']); } }, writes, 'format');
  format.load = (fields) => { loads.push(fields); return format; };
  writes.length = 0;
  const range = tracked({ address: 'Report!A1:B2', rowCount: 2, columnCount: 2, values: [[1, 2], [3, 4]], formulas: [['', ''], ['', '']], numberFormat: [['General']], format }, writes, 'range');
  let frozenAddress = options.frozenAddress ?? null;
  const freezePanes = {
    freezeRows(count) { calls.push(['freezeRows', count]); frozenAddress = `Report!1:${count}`; },
    freezeColumns(count) { calls.push(['freezeColumns', count]); frozenAddress = `Report!A:${String.fromCharCode(64 + count)}`; },
    freezeAt(frozenRange) { calls.push(['freezeAt', frozenRange]); frozenAddress = frozenRange.address; },
    unfreeze() { calls.push(['unfreeze']); frozenAddress = null; },
    getLocationOrNullObject() { calls.push(['getLocationOrNullObject']); return loadable({ isNullObject: frozenAddress === null, address: frozenAddress }); },
  };
  const sheet = loadable({ name: 'Report', freezePanes, getRange(address) { calls.push(['getRange', address]); return loadable({ address: `Report!${address}` }); } });
  const context = { sync: async () => {}, workbook: { worksheets: { getItem(name) { if (name !== 'Report') throw new Error('missing sheet'); return sheet; } }, tables: { getItem() { throw new Error('missing table'); } } } };
  const host = hostHarness('Excel', context, { util: { excelRange: () => range }, supported: options.supported });
  return { host, range, format, writes, calls, loads };
}

function pptHarness(types = ['TextBox', 'GeometricShape', 'Image'], options = {}) {
  const writes = [], reads = [];
  const shapes = types.map((type, index) => {
    const font = tracked({ name: 'Aptos', size: 22, color: '#112233', bold: true, italic: false }, writes, `s${index}.font`);
    const textRange = tracked({ text: `text-${index}`, font }, writes, `s${index}.textRange`);
    textRange.load = (fields) => { reads.push([index, 'text', fields]); if (options.failStyleRead) throw new Error('style unavailable'); return textRange; };
    const fill = tracked({ type: 'Solid', foregroundColor: '#FFFFFF', transparency: 0.3, setSolidColor(color) { this.foregroundColor = color; this.type = 'Solid'; this.transparency = 0; } }, writes, `s${index}.fill`);
    fill.load = (fields) => { reads.push([index, 'fill', fields]); return fill; };
    const shape = tracked({ id: `s${index}`, name: `shape-${index}`, type, left: index * 100, top: 20, width: 90, height: 40, zOrderPosition: index, fill }, writes, `s${index}`);
    Object.defineProperty(shape, 'textFrame', { get() { if (!['TextBox', 'GeometricShape'].includes(type)) throw new Error('InvalidArgument: no text frame'); return { textRange }; } });
    return shape;
  });
  writes.length = 0;
  const slide = { id: 'slide-1', shapes: loadable({ items: shapes }) };
  const context = { sync: async () => {}, presentation: { slides: loadable({ items: [slide] }) } };
  const host = hostHarness('PowerPoint', context, options);
  return { host, shapes, writes, reads, context };
}

test('freezeAt receives the supplied frozen range, not a derived anchor or combined row/column calls', async () => {
  const h = excelHarness();
  const result = await h.host.actions.excel_set_freeze_panes({ sheet: 'Report', mode: 'range', range: 'A1:B2' });
  assert.equal(result.ok, true);
  assert.equal(result.result.frozenRange, 'Report!A1:B2');
  assert.deepEqual(h.calls.filter(([name]) => name.startsWith('freeze')).map(([name]) => name), ['freezeAt']);
  assert.equal(h.calls.find(([name]) => name === 'freezeAt')[1].address, 'Report!A1:B2');
});

test('freeze rows, columns, none dispatch exactly one mutation and preview reads current location only', async () => {
  for (const [mode, method] of [['rows', 'freezeRows'], ['columns', 'freezeColumns'], ['none', 'unfreeze']]) {
    const h = excelHarness({ frozenAddress: 'Report!A1:B2' });
    const args = { sheet: 'Report', mode, ...(mode === 'none' ? {} : { count: 2 }) };
    const preview = await h.host.previewReaders.excel_set_freeze_panes(args);
    assert.equal(preview.result.frozenRange, 'Report!A1:B2');
    assert.equal(preview.result.previewOnly, true);
    assert.equal(h.calls.some(([name]) => ['freezeRows', 'freezeColumns', 'freezeAt', 'unfreeze'].includes(name)), false);
    const result = await h.host.actions.excel_set_freeze_panes(args);
    assert.equal(result.ok, true);
    assert.equal(h.calls.filter(([name]) => name === method).length, 1);
    if (mode === 'none') assert.equal(result.result.frozenRange, null);
  }
});

test('freeze validates mode, dimensions, local A1 syntax and sheet before any mutation', async () => {
  const invalid = [{ mode: 'rows', count: '2' }, { mode: 'rows', count: 0 }, { mode: 'columns', count: 16385 }, { mode: 'range', range: 'Other!A1:B2' }, { mode: 'range', range: 'A:A' }, { mode: 'range', range: 'XFE1' }, { mode: 'range', range: 'A1048577' }, { mode: 'none', count: 1 }, { mode: 'none', range: 'A1' }, { mode: 'rows', count: 2, range: 'A1' }, { mode: 'bogus' }];
  for (const args of invalid) {
    const h = excelHarness();
    assert.equal((await h.host.actions.excel_set_freeze_panes({ sheet: 'Report', ...args })).code, 'bad_args');
    assert.equal(h.calls.length, 0);
  }
  const h = excelHarness();
  assert.equal((await h.host.actions.excel_set_freeze_panes({ sheet: 'missing', mode: 'rows', count: 1 })).code, 'execution');
  assert.equal(h.calls.length, 0);
});

test('ExcelApi 1.7 is required for freezing and 1.2 only for dimension/autofit fields', async () => {
  const h = excelHarness({ supported: (_set, version) => version === '1.1' });
  assert.equal((await h.host.actions.excel_set_freeze_panes({ sheet: 'Report', mode: 'none' })).code, 'requirement');
  assert.equal((await h.host.actions.format_range({ address: 'A1', columnWidth: 40 })).code, 'requirement');
  assert.equal(h.writes.length, 0);
  const result = await h.host.actions.format_range({ address: 'A1', wrapText: true });
  assert.equal(result.ok, true);
  assert.equal('columnWidth' in result.result.verification.layout, false);
});

test('format layout writes finite point sizes, canonical alignment and actual readback, preserving numberFormat contract', async () => {
  const h = excelHarness();
  const result = await h.host.actions.format_range({ address: 'A1:B2', columnWidth: 120, rowHeight: 28, wrapText: true, horizontalAlignment: 'left', verticalAlignment: 'center', numberFormat: '0.00' });
  assert.equal(result.ok, true);
  assert.deepEqual(plain(result.result.verification.layout), { columnWidth: 120, rowHeight: 28, wrapText: true, horizontalAlignment: 'Left', verticalAlignment: 'Center' });
  assert.deepEqual(plain(h.range.numberFormat), [['0.00']]);
  assert.ok(h.loads.some((fields) => fields.includes('columnWidth,rowHeight')));
  const preview = await h.host.previewReaders.format_range({ address: 'A1:B2' });
  assert.deepEqual(plain(preview.result.layout), plain(result.result.verification.layout));
});

test('autofit uses the actual lowercase Office.js method names', async () => {
  const h = excelHarness();
  assert.equal((await h.host.actions.format_range({ address: 'A1:B2', autoFit: true, autoFitRows: true })).ok, true);
  assert.deepEqual(h.calls, [['autofitColumns'], ['autofitRows']]);
});

test('invalid layout, conflicting dimensions and absent table cannot partially format a range', async () => {
  const invalid = [{ columnWidth: Infinity }, { rowHeight: NaN }, { columnWidth: '80' }, { rowHeight: 0 }, { wrapText: 'false' }, { autoFitRows: 1 }, { autoFit: 'yes' }, { horizontalAlignment: 'centre' }, { verticalAlignment: 'Middle' }, { autoFit: true, columnWidth: 30 }, { autoFitRows: true, rowHeight: 20 }, { tableStyle: 'Medium2' }];
  for (const patch of invalid) {
    const h = excelHarness();
    assert.equal((await h.host.actions.format_range({ address: 'A1', bold: true, ...patch })).code, 'bad_args');
    assert.equal(h.writes.length, 0);
  }
  const h = excelHarness();
  assert.equal((await h.host.actions.format_range({ address: 'A1', bold: true, tableStyle: 'Medium2', tableName: 'missing' })).code, 'execution');
  assert.equal(h.writes.length, 0);
});

test('Excel range images require 1.7, preflight full dimensions and read ClientResult only after sync', async () => {
  function rendering(rows = 2, cols = 2, supported = true) {
    let dimensionsLoaded = false, imageLoaded = false, imageCalls = 0, imageQueued = false;
    const range = {
      load(fields) { assert.match(fields, /rowCount/); assert.match(fields, /columnCount/); assert.match(fields, /address/); },
      get rowCount() { assert.equal(dimensionsLoaded, true); return rows; },
      get columnCount() { assert.equal(dimensionsLoaded, true); return cols; },
      get address() { assert.equal(dimensionsLoaded, true); return 'Report!A1:B2'; },
      getImage() { assert.equal(dimensionsLoaded, true); imageCalls++; imageQueued = true; return { get value() { assert.equal(imageLoaded, true); return 'iVBORw0KGgo='; } }; },
    };
    const context = { sync: async () => { dimensionsLoaded = true; if (imageQueued) imageLoaded = true; } };
    const host = hostHarness('Excel', context, { util: { excelRange: () => range }, supported: (set, version) => { assert.equal(set, 'ExcelApi'); assert.equal(version, '1.7'); return supported; } });
    return { render: host.actions.excel_render_range, get calls() { return imageCalls; } };
  }
  const unsupported = rendering(2, 2, false);
  assert.equal((await unsupported.render({ address: 'Report!A1:B2' })).code, 'requirement');
  assert.equal(unsupported.calls, 0);
  const oversized = rendering(100, 51);
  assert.equal((await oversized.render({ address: 'A1:AY100' })).code, 'bad_args');
  assert.equal(oversized.calls, 0);
  for (const maxCells of [0, 10001, 1.5, '4', Infinity]) {
    const h = rendering();
    assert.equal((await h.render({ address: 'A1:B2', maxCells })).code, 'bad_args');
    assert.equal(h.calls, 0);
  }
  const h = rendering();
  assert.equal((await h.render({ address: 'A1:B2', maxCells: 3 })).code, 'bad_args');
  assert.equal(h.calls, 0);
  const result = await h.render({ address: 'A1:B2', maxCells: 4 });
  assert.equal(result.ok, true);
  assert.deepEqual(plain(result.result), { host: 'Excel', address: 'Report!A1:B2', rows: 2, cols: 2, mimeType: 'image/png', base64: 'iVBORw0KGgo=' });
  assert.equal(h.calls, 1);
  assert.equal((await rendering().render({ address: '' })).code, 'bad_args');
});

test('PPT new textboxes apply explicit fill/transparency/outline and retain unspecified default shape styles', async () => {
  async function add(patch) {
    const calls = [];
    const fill = tracked({ foregroundColor: '#101010', transparency: 0.6, setSolidColor(color) { calls.push(['setSolidColor', color]); this.foregroundColor = color; }, clear() { throw new Error('textbox default fill must not be cleared'); } }, calls, 'fill');
    const lineFormat = tracked({ color: '#999999', weight: 3, visible: false }, calls, 'line');
    const shape = loadable({ id: 'new-textbox', fill, lineFormat, textFrame: { textRange: { font: {}, paragraphFormat: {} } } });
    const context = { sync: async () => {}, presentation: { slides: { getItem(id) { assert.equal(id, 'slide-1'); return { shapes: { addTextBox(text) { assert.equal(text, 'Readable on dark slides'); return shape; } } }; } } } };
    const host = hostHarness('PowerPoint', context);
    const result = await host.actions.ppt_add_elements({ slideId: 'slide-1', coordinateSpace: { width: 960, height: 540 }, targetSize: { width: 720, height: 540 }, elements: [{ type: 'text', text: 'Readable on dark slides', x: 10, y: 10, w: 500, h: 40, ...patch }] });
    assert.equal(result.ok, true, result.error);
    return { calls, fill, lineFormat };
  }
  const explicit = await add({ fill: '#ffffff', fillTransparency: 0.25, lineColor: '#123456', lineWidth: 2 });
  assert.deepEqual(explicit.calls, [['setSolidColor', '#FFFFFF'], ['fill.foregroundColor', '#FFFFFF'], ['fill.transparency', 0.25], ['line.visible', true], ['line.color', '#123456'], ['line.weight', 2]]);
  assert.equal(explicit.fill.transparency, 0.25);
  const opaque = await add({ fill: '#FFFFFF' });
  assert.equal(opaque.fill.transparency, 0);
  assert.equal(opaque.calls.some(([key]) => key.startsWith('line.')), false);
  const defaults = await add({});
  assert.deepEqual(defaults.calls, []);
  assert.equal(defaults.fill.foregroundColor, '#101010'); assert.equal(defaults.fill.transparency, 0.6);
  assert.deepEqual({ color: defaults.lineFormat.color, weight: defaults.lineFormat.weight, visible: defaults.lineFormat.visible }, { color: '#999999', weight: 3, visible: false });
  const widthOnly = await add({ lineWidth: 4 });
  assert.deepEqual(widthOnly.calls, [['line.weight', 4]]);
  assert.equal(widthOnly.lineFormat.visible, false);
  const hidden = await add({ lineWidth: 0 });
  assert.deepEqual(hidden.calls, [['line.visible', false]]);
  const transparencyOnly = await add({ fillTransparency: 0.1 });
  assert.deepEqual(transparencyOnly.calls, [['fill.transparency', 0.1]]);
});

test('PPT patches only supplied fields and preserves fill transparency on a color-only update', async () => {
  const h = pptHarness();
  const result = await h.host.actions.ppt_update_shapes({ slideId: 'slide-1', updates: [{ shapeId: 's0', text: 'revised', font: { bold: false }, fill: { color: '#abcdef' }, left: 42 }] });
  assert.equal(result.ok, true);
  const shape = result.result.shapes[0];
  assert.equal(shape.text, 'revised');
  assert.equal(shape.left, 42);
  assert.deepEqual(plain(shape.font), { name: 'Aptos', size: 22, color: '#112233', bold: false, italic: false });
  assert.equal(shape.fill.color, '#ABCDEF');
  assert.equal(shape.fill.transparency, 0.3);
  assert.equal(shape.width, 90);
  assert.equal(h.writes.some(([key]) => ['s0.width', 's0.height', 's0.top', 's0.font.name', 's0.font.size', 's1.left'].includes(key)), false);
});

test('PPT prevalidates every update, missing target, duplicate target and unsupported type before any setter', async () => {
  const cases = [[{ shapeId: 's0', left: 10 }, { shapeId: 's1', width: -1 }], [{ shapeId: 's0', text: 'x' }, { shapeId: 'missing', text: 'y' }], [{ shapeId: 's0', left: 10 }, { shapeId: 's0', top: 10 }], [{ shapeId: 's0', left: 10 }, { shapeId: 's2', text: 'image text' }], [{ shapeId: 's0', font: { bold: 'false' } }], [{ shapeId: 's0', font: { size: Infinity } }], [{ shapeId: 's0', fill: { color: '#FFFFFF', transparency: 2 } }], [{ shapeId: 's0', runs: [] }], [{ shapeId: 's0' }]];
  for (const updates of cases) {
    const h = pptHarness();
    assert.equal((await h.host.actions.ppt_update_shapes({ slideId: 'slide-1', updates })).ok, false);
    assert.equal(h.writes.length, 0);
  }
  const h = pptHarness();
  assert.equal((await h.host.actions.ppt_update_shapes({ slideId: 'missing', updates: [{ shapeId: 's0', left: 1 }] })).code, 'not_found');
  assert.equal(h.writes.length, 0);
});

test('PPT property-resolution failure happens before writes and preview is selected shapes only', async () => {
  const failing = pptHarness(['TextBox'], { failStyleRead: true });
  assert.equal((await failing.host.actions.ppt_update_shapes({ slideId: 'slide-1', updates: [{ shapeId: 's0', left: 99 }] })).code, 'execution');
  assert.equal(failing.writes.length, 0);
  const h = pptHarness();
  const preview = await h.host.previewReaders.ppt_update_shapes({ slideId: 'slide-1', updates: [{ shapeId: 's1', text: 'new' }] });
  assert.equal(preview.result.previewOnly, true);
  assert.equal(preview.result.shapes.length, 1);
  assert.equal(preview.result.shapes[0].id, 's1');
  assert.equal(preview.result.shapes[0].text, 'text-1');
  assert.equal(preview.result.shapes[0].font.size, 22);
  assert.equal(h.writes.length, 0);
});

test('PPT inspect defaults stay geometric; requested styles explicitly mark unreadable types', async () => {
  const h = pptHarness();
  const geometry = await h.host.actions.ppt_inspect_slides({});
  assert.equal(geometry.ok, true);
  assert.equal(h.reads.length, 0);
  assert.equal('stylesRead' in geometry.result.slides[0].shapes[0], false);
  const styled = await h.host.actions.ppt_inspect_slides({ withStyles: true });
  assert.equal(styled.ok, true);
  assert.equal(styled.result.slides[0].shapes[0].stylesRead, true);
  assert.equal(styled.result.slides[0].shapes[2].stylesRead, false);
  assert.match(styled.result.slides[0].shapes[2].stylesReadReason, /Image/);
  assert.equal('text' in styled.result.slides[0].shapes[2], false);
});

test('PPT geometry-only patch supports image shapes and never requests their textFrame', async () => {
  const h = pptHarness();
  const result = await h.host.actions.ppt_update_shapes({ slideId: 'slide-1', updates: [{ shapeId: 's2', top: 0, width: 75 }] });
  assert.equal(result.ok, true);
  assert.equal(result.result.shapes[0].stylesRead, false);
  assert.deepEqual(h.writes, [['s2.top', 0], ['s2.width', 75]]);
});

test('PPT explicit empty text and false font flags work without touching unrequested fill or geometry', async () => {
  const h = pptHarness();
  const result = await h.host.actions.ppt_update_shapes({ slideId: 'slide-1', updates: [{ shapeId: 's0', text: '', font: { bold: false, italic: false } }] });
  assert.equal(result.ok, true);
  assert.equal(result.result.shapes[0].text, '');
  assert.deepEqual(h.writes.map(([key]) => key), ['s0.textRange.text', 's0.font.bold', 's0.font.italic']);
});

test('PPT 1.4 gate rejects mutation before accessing shapes', async () => {
  const h = pptHarness(['TextBox'], { supported: () => false });
  const result = await h.host.actions.ppt_update_shapes({ slideId: 'slide-1', updates: [{ shapeId: 's0', text: 'x' }] });
  assert.equal(result.code, 'requirement');
  assert.equal(h.reads.length, 0);
  assert.equal(h.writes.length, 0);
});

test('PPT style snapshots do not collapse equal shape IDs across slides or invent non-solid fill colors', async () => {
  const h = pptHarness(['TextBox']);
  const other = pptHarness(['TextBox']);
  other.shapes[0].textFrame.textRange.text = 'another slide';
  other.shapes[0].fill.type = 'Gradient';
  other.shapes[0].fill.transparency = null;
  h.context.presentation.slides.items.push({ id: 'slide-2', shapes: loadable({ items: other.shapes }) });
  const result = await h.host.actions.ppt_inspect_slides({ withStyles: true });
  assert.equal(result.ok, true);
  assert.equal(result.result.slides[0].shapes[0].text, 'text-0');
  assert.equal(result.result.slides[1].shapes[0].text, 'another slide');
  assert.equal('color' in result.result.slides[1].shapes[0].fill, false);
  assert.equal(result.result.slides[1].shapes[0].fill.transparency, null);
  assert.equal(other.reads.some(([, kind, fields]) => kind === 'fill' && fields === 'foregroundColor'), false);
});

test('PPT rendering filters in document order, keeps original indexes and preflights every requested ID', async () => {
  const h = pptHarness();
  const rendered = [];
  h.context.presentation.slides.items = ['old-page', 'new-1', 'new-2'].map((id) => ({ id, getImageAsBase64(options) { rendered.push({ id, options }); return { value: `image-${id}` }; } }));
  const result = await h.host.actions.ppt_render_slides({ slideIds: ['new-2', 'new-1'], width: 800, height: 600 });
  assert.equal(result.ok, true);
  assert.deepEqual(plain(result.result.slides), [{ index: 2, id: 'new-1', base64: 'image-new-1' }, { index: 3, id: 'new-2', base64: 'image-new-2' }]);
  assert.deepEqual(rendered.map(({ id }) => id), ['new-1', 'new-2']);
  rendered.length = 0;
  for (const [slideIds, code] of [[['new-1', 'missing'], 'not_found'], [[], 'bad_args'], [['new-1', 'new-1'], 'bad_args'], [[1], 'bad_args']]) {
    assert.equal((await h.host.actions.ppt_render_slides({ slideIds })).code, code);
    assert.equal(rendered.length, 0);
  }
  assert.equal((await h.host.actions.ppt_render_slides({})).ok, true);
  assert.deepEqual(rendered.map(({ id }) => id), ['old-page', 'new-1', 'new-2']);
});

// Tiny DOM fixture adapter: it models only XML members used by pptXml and
// orderedPptSlides; no Office objects or native slide-size API are introduced.
function xmlFixtureParser() {
  const node = (xml) => ({ getAttribute(name) { return xml.match(new RegExp(`\\b${name}="([^"]*)"`))?.[1] ?? null; }, getAttributeNS(_ns, name) { return xml.match(new RegExp(`\\br:${name}="([^"]*)"`))?.[1] ?? null; } });
  return class {
    parseFromString(xml) {
      return { getElementsByTagName: () => [], getElementsByTagNameNS(_ns, name) {
        if (name === 'sldSz') return [...xml.matchAll(/<p:sldSz\b[^>]*\/>/g)].map(([match]) => node(match));
        if (name === 'Relationship') return [...xml.matchAll(/<Relationship\b[^>]*\/>/g)].map(([match]) => node(match));
        if (name === 'sldIdLst') return xml.includes('<p:sldIdLst>') ? [{ getElementsByTagNameNS() { return [...xml.matchAll(/<p:sldId\b[^>]*\/>/g)].map(([match]) => node(match)); } }] : [];
        return [];
      } };
    }
  };
}

function presentationInfoHarness(size) {
  const encoder = new TextEncoder();
  const entries = {
    'ppt/presentation.xml': `<p:presentation><p:sldIdLst><p:sldId id="259" r:id="rId2"/><p:sldId id="256" r:id="rId1"/></p:sldIdLst>${size}</p:presentation>`,
    'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>',
    'ppt/slides/slide1.xml': '<slide/>', 'ppt/slides/slide2.xml': '<slide/>',
  };
  return hostHarness('PowerPoint', {}, { util: { readPptFile: (callback) => callback([new Uint8Array([1])]), unzip: () => entries, inflateEntry: (entry) => encoder.encode(entry) }, globals: { DOMParser: xmlFixtureParser(), TextDecoder, Uint8Array }, supported: () => false });
}

test('presentation info uses XML EMUs, converts to points and preserves relationship slide order', async () => {
  const host = presentationInfoHarness('<p:sldSz cx="9144000" cy="6858000"/>');
  const result = await host.actions.ppt_read_presentation_info();
  assert.equal(result.ok, true);
  assert.deepEqual(plain(result.result.slideSize), { width: 720, height: 540, unit: 'pt' });
  assert.equal(result.result.aspectRatio, 4 / 3);
  assert.deepEqual(plain(result.result.slideIds), ['259', '256']);
});

test('presentation info does not guess dimensions for missing or invalid sldSz', async () => {
  for (const size of ['', '<p:sldSz cx="0" cy="6858000"/>', '<p:sldSz cx="oops" cy="6858000"/>']) {
    const result = await presentationInfoHarness(size).actions.ppt_read_presentation_info();
    assert.equal(result.code, 'execution');
    assert.match(result.error, /slide size/);
  }
});
