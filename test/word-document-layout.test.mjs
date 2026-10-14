import { test } from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import fs from 'node:fs';
const source = fs.readFileSync(new URL('../server/actions/word.js', import.meta.url), 'utf8');
const pane = fs.readFileSync(new URL('../server/taskpane.js', import.meta.url), 'utf8');
const plain = (value) => JSON.parse(JSON.stringify(value));

function fixture({ supported = true, requirements = { WordApi: '1.3', WordApiDesktop: '1.3' }, texts = ['Summary', 'Details'], listed = false, failBorderRead = false, fixedBorderWidth = false } = {}) {
  let writes = 0, runs = 0, listIdQueued = false, listIdLoaded = false, pendingBorderFailure = false;
  const gates = [], events = [], pendingBorders = new Set();
  const tracked = (data, prefix = '') => new Proxy(data, { set(object, key, value) { writes++; events.push(`${prefix}.${String(key)}`); object[key] = value; return true; } });
  const list = {
    get id() { assert.equal(listIdLoaded, true, 'List.id must be loaded and synced before reading'); events.push('list.id.read'); return 42; },
    load(fields) { assert.equal(fields, 'id'); listIdQueued = true; events.push('list.id.load'); },
    setLevelNumbering(level, type) { writes++; assert.equal(level, 0); assert.equal(type, 'Arabic'); },
    setLevelBullet(level, type) { writes++; assert.equal(level, 0); assert.equal(type, 'Solid'); },
  };
  const paragraphs = texts.map((text, index) => {
    const data = { text, styleBuiltIn: 'Normal', alignment: 'Left', spaceBefore: 0, spaceAfter: 0, lineSpacing: 12, leftIndent: 0, rightIndent: 0, firstLineIndent: 0, isListItem: listed,
      load() {}, font: tracked({ name: 'Arial', size: 12, color: '#000000', bold: false, italic: false, load() {} }, `p${index}.font`),
      startNewList() { writes++; listIdLoaded = false; return list; }, detachFromList() { writes++; }, attachToList(id, level) { assert.equal(id, 42); assert.equal(level, 0); writes++; events.push('paragraph.attachToList'); } };
    return new Proxy(data, { set(object, key, value) {
      writes++; events.push(`p${index}.${String(key)}`); object[key] = value;
      // A semantic style changes inherited font values; explicit font setters
      // must execute afterwards. This deliberately exposes reversed ordering.
      if (key === 'styleBuiltIn' && value === 'Heading1') Object.assign(object.font, { name: 'Theme Heading', size: 24, color: '#336699', bold: false, italic: false });
      return true;
    } });
  });
  const paras = { items: paragraphs, load() {} };
  const page = tracked({ topMargin: 72, bottomMargin: 72, leftMargin: 72, rightMargin: 72, headerDistance: 36, footerDistance: 36, load() {} }, 'page');
  const borders = new Map();
  const table = tracked({ getBorder(location) {
    if (!borders.has(location)) {
      const data = { color: '#000000', width: 1, type: fixedBorderWidth ? 'Wave' : 'Single' };
      let loaded = false;
      const border = new Proxy({ load(fields) { assert.equal(fields, 'color,type,width'); events.push(`border.${location}.load`); pendingBorders.add(() => { loaded = true; }); if (failBorderRead) pendingBorderFailure = true; } }, {
        get(object, key) { if (['color', 'width', 'type'].includes(key)) { assert.equal(loaded, true, 'border values must be loaded and synced'); return data[key]; } return object[key]; },
        set(_object, key, value) { writes++; events.push(`border.${location}.${String(key)}`); if (!(key === 'width' && fixedBorderWidth)) data[key] = value; loaded = false; return true; },
      });
      borders.set(location, border);
    }
    return borders.get(location);
  } }, 'table');
  const ctx = { document: { body: { paragraphs: paras, tables: { items: [table], load() {} } }, getSelection: () => ({ paragraphs: paras }), sections: { items: [{ pageSetup: page }], load() {} } }, async sync() {
    events.push('sync');
    if (pendingBorderFailure) throw new Error('border read failed');
    for (const apply of pendingBorders) apply(); pendingBorders.clear();
    if (listIdQueued) { listIdLoaded = true; listIdQueued = false; }
  } };
  const Word = { run: async (fn) => { runs++; return fn(ctx); }, BuiltInStyleName: { title: 'Title', subtitle: 'Subtitle', normal: 'Normal', heading1: 'Heading1', heading2: 'Heading2', heading9: 'Heading9', caption: 'Caption' }, Alignment: { left: 'Left', centered: 'Centered', right: 'Right', justified: 'Justified' }, ListNumbering: { arabic: 'Arabic' }, ListBullet: { solid: 'Solid' } };
  const window = { __OFFICE_ACTIONS__: { util: { okResult: (result) => ({ ok: true, result }), errResult: (code, error) => ({ ok: false, code, error }) }, hosts: {}, isSetSupported(name, version) {
    gates.push([name, version]);
    if (!supported || !requirements[name]) return false;
    const current = requirements[name].split('.').map(Number), requested = version.split('.').map(Number);
    return current[0] > requested[0] || (current[0] === requested[0] && current[1] >= requested[1]);
  } } };
  const context = vm.createContext({ window, Word });
  vm.runInContext(pane.slice(pane.indexOf('function mapBuiltInStyle('), pane.indexOf('// Excel：')), context);
  window.__OFFICE_ACTIONS__.util.mapBuiltInStyle = vm.runInContext('mapBuiltInStyle', context);
  window.__OFFICE_ACTIONS__.util.mapAlignment = vm.runInContext('mapAlignment', context);
  vm.runInContext(source, context);
  return { host: window.__OFFICE_ACTIONS__.hosts.Word, paragraphs, page, borders, gates, events, writes: () => writes, runs: () => runs, style: (name) => window.__OFFICE_ACTIONS__.util.mapBuiltInStyle(name) };
}

test('semantic style inherits font first and explicit font/paragraph patches override it afterwards', async () => {
  const f = fixture();
  const out = await f.host.actions.word_format_paragraphs({ targets: [{ index: 0, expectedText: 'Summary', style: 'Heading 1', font: { color: '#245B8A', bold: true }, spaceAfter: 8, lineSpacing: 18 }] });
  assert.equal(out.ok, true); assert.equal(out.result.formatted, 1);
  const p = out.result.paragraphs[0];
  assert.equal(p.styleBuiltIn, 'Heading1'); assert.equal(p.font.name, 'Theme Heading'); assert.equal(p.font.size, 24); assert.equal(p.font.color, '#245B8A'); assert.equal(p.font.bold, true); assert.equal(p.spaceAfter, 8); assert.equal(p.lineSpacing, 18); assert.equal(f.paragraphs[1].font.color, '#000000');
  assert.ok(f.events.indexOf('p0.styleBuiltIn') < f.events.lastIndexOf('p0.font.color'));
  assert.deepEqual(f.gates, [['WordApi', '1.3']]);
});

test('patches without style preserve absent font, spacing and other paragraph fields', async () => {
  const f = fixture();
  const out = await f.host.actions.word_format_paragraphs({ targets: [{ index: 0, expectedText: 'Summary', font: { color: '#245B8A' }, spaceAfter: 8 }] });
  assert.equal(out.ok, true);
  const p = out.result.paragraphs[0];
  assert.deepEqual(plain(p.font), { name: 'Arial', size: 12, color: '#245B8A', bold: false, italic: false });
  assert.equal(p.styleBuiltIn, 'Normal'); assert.equal(p.spaceBefore, 0); assert.equal(p.lineSpacing, 12); assert.equal(p.leftIndent, 0);
  assert.equal(f.events.includes('p0.styleBuiltIn'), false);
  assert.equal(f.events.some((event) => event.startsWith('p1.')), false);
});

test('all paragraph text targets preflight before any writes', async () => {
  for (const bad of [{ index: 1, expectedText: 'stale', spaceAfter: 9 }, { index: 20, expectedText: 'missing', spaceAfter: 9 }]) {
    const f = fixture(); const out = await f.host.actions.word_format_paragraphs({ targets: [{ index: 0, expectedText: 'Summary', spaceAfter: 8 }, bad] }); assert.equal(out.code, 'bad_args'); assert.equal(f.writes(), 0);
  }
});

test('invalid paragraph fields fail before entering Office', async () => {
  for (const extra of [{ font: { size: NaN } }, { font: { color: 'red' } }, { font: { bold: 'false' } }, { spaceAfter: -1 }, { lineSpacing: 0 }, { bogus: 1 }, { style: 'FakeHeading' }]) {
    const f = fixture(); const out = await f.host.actions.word_format_paragraphs({ targets: [{ index: 0, expectedText: 'Summary', ...extra }] }); assert.equal(out.code, 'bad_args'); assert.equal(f.writes(), 0); assert.equal(f.runs(), 0);
  }
  const f = fixture(); assert.equal((await f.host.actions.word_format_paragraphs({ targets: [{ index: 0, expectedText: 'Summary', spaceAfter: 1 }, { index: 0, expectedText: 'Summary', spaceAfter: 2 }] })).code, 'bad_args'); assert.equal(f.writes(), 0);
});

test('paragraph preview reads target style and font without mutation', async () => {
  const f = fixture(); const out = await f.host.previewReaders.word_format_paragraphs({ targets: [{ index: 1, expectedText: 'Details', font: { size: 14 } }] }); assert.equal(out.ok, true); assert.equal(out.result.paragraphs[0].font.size, 12); assert.equal(f.writes(), 0);
});

test('built-in style mapping includes Heading1..9 and localized caption', () => {
  const f = fixture(); assert.equal(f.style('Heading-9'), 'heading9'); assert.equal(f.style('标题4'), 'heading4'); assert.equal(f.style('题注'), 'caption'); assert.equal(f.style('Heading10'), null);
});

test('legacy base paragraph properties work on WordApi 1.1 without fictitious range.paragraphFormat', async () => {
  const f = fixture({ requirements: { WordApi: '1.1' } }); const out = await f.host.actions.set_paragraph_format({ alignment: 'center', indent: 2.54, lineSpacing: 18 }); assert.equal(out.ok, true); assert.equal(out.result.formatted, 2);
  for (const p of f.paragraphs) { assert.equal(p.alignment, 'Centered'); assert.equal(p.leftIndent, 72); assert.equal(p.lineSpacing, 18); }
  assert.equal(out.result.paragraphs[0].styleRead, false);
});

test('both semantic style entry points require WordApi 1.3 before any writes', async () => {
  const f = fixture({ requirements: { WordApi: '1.1', WordApiDesktop: '1.3' } });
  assert.equal((await f.host.actions.word_format_paragraphs({ targets: [{ index: 0, expectedText: 'Summary', style: 'Heading1' }] })).code, 'requirement');
  assert.equal((await f.host.actions.set_paragraph_format({ style: 'Heading1', alignment: 'left' })).code, 'requirement');
  assert.equal(f.runs(), 0); assert.equal(f.writes(), 0);
  assert.deepEqual(f.gates, [['WordApi', '1.3'], ['WordApi', '1.3']]);
  const supported = fixture();
  const result = await supported.host.actions.set_paragraph_format({ style: 'Heading1', font: { bold: true } });
  assert.equal(result.ok, true); assert.equal(result.result.paragraphs[0].font.size, 24); assert.equal(result.result.paragraphs[0].font.bold, true);
});

test('list formatting loads and syncs list ID before attaching with real list enums', async () => {
  for (const listType of ['numbered', 'bulleted']) {
    const f = fixture(); assert.equal((await f.host.actions.set_paragraph_format({ listType })).ok, true); assert.ok(f.writes() >= 3);
    const load = f.events.indexOf('list.id.load'), sync = f.events.indexOf('sync', load + 1), read = f.events.indexOf('list.id.read'), attach = f.events.indexOf('paragraph.attachToList');
    assert.ok(load >= 0 && load < sync && sync < read && read < attach);
    assert.ok(f.gates.some(([name, version]) => name === 'WordApi' && version === '1.3'));
  }
  const f = fixture({ requirements: { WordApi: '1.1', WordApiDesktop: '1.3' } }); assert.equal((await f.host.actions.set_paragraph_format({ listType: 'numbered', alignment: 'center' })).code, 'requirement'); assert.equal(f.writes(), 0);
  const existing = fixture({ listed: true }); assert.equal((await existing.host.actions.set_paragraph_format({ listType: 'bulleted', alignment: 'center' })).code, 'bad_args'); assert.equal(existing.writes(), 0);
});

test('page setup requires the Desktop API specifically and previews actual page values', async () => {
  const missing = fixture({ requirements: { WordApi: '1.3', WordApiDesktop: '1.2' } }); assert.equal((await missing.host.actions.word_set_page_setup({ leftMargin: 60 })).code, 'requirement'); assert.equal(missing.runs(), 0); assert.deepEqual(missing.gates, [['WordApiDesktop', '1.3']]);
  const f = fixture(); const pre = await f.host.previewReaders.word_set_page_setup({ leftMargin: 60 }); assert.equal(pre.result.sections[0].values.leftMargin, 72); assert.equal(f.writes(), 0);
  const out = await f.host.actions.word_set_page_setup({ leftMargin: 60 }); assert.equal(out.ok, true); assert.equal(out.result.pageSetup[0].values.leftMargin, 60); assert.equal(out.result.pageSetup[0].values.rightMargin, 72);
  assert.deepEqual(f.gates, [['WordApiDesktop', '1.3'], ['WordApiDesktop', '1.3']]);
});

test('table border width is patched only explicitly and actual color/type/width are loaded and returned', async () => {
  const f = fixture();
  const initial = await f.host.actions.set_table_borders({ color: '#D9DEE5', locations: ['Top'] });
  assert.equal(initial.ok, true); assert.equal(f.borders.get('Top').width, 1); assert.equal(f.events.includes('border.Top.width'), false);
  const out = await f.host.actions.set_table_borders({ color: '#D9DEE5', width: 0.5, locations: ['Top'] });
  assert.equal(out.ok, true); assert.equal(f.borders.get('Top').width, 0.5); assert.equal(f.borders.get('Top').type, 'Single'); assert.equal(out.result.width, 0.5);
  assert.deepEqual(plain(out.result.borders), [{ table: 1, location: 'Top', color: '#D9DEE5', type: 'Single', width: 0.5 }]);
  assert.match(out.result.warning, /fixed widths/);
});

test('table border validation rejects any invalid or missing target before table style or border writes', async () => {
  const invalid = [{ locations: ['Top', 'bogus'] }, { locations: ['Top', 'Top'] }, { locations: [] }, { locations: 'Top' }, { tableIndexes: [1, 2] }, { tableIndexes: [1, 0] }, { tableIndexes: [1, 1] }, { tableIndexes: ['1'] }, { tableIndexes: [] }, { width: '0.5' }, { width: Infinity }, { width: 0 }, { enableGrid: 'true' }];
  for (const patch of invalid) {
    const f = fixture(); const out = await f.host.actions.set_table_borders({ color: '#D9DEE5', enableGrid: true, ...patch }); assert.equal(out.ok, false); assert.equal(f.writes(), 0);
  }
  const unsupported = fixture({ requirements: { WordApi: '1.1' } }); assert.equal((await unsupported.host.actions.set_table_borders({ color: '#D9DEE5' })).code, 'requirement'); assert.equal(unsupported.runs(), 0);
});

test('border load failure precedes even enableGrid writes, and fixed-width readback is not the request', async () => {
  const failing = fixture({ failBorderRead: true });
  assert.equal((await failing.host.actions.set_table_borders({ color: '#D9DEE5', width: 0.5, enableGrid: true })).code, 'execution'); assert.equal(failing.writes(), 0);
  const fixed = fixture({ fixedBorderWidth: true });
  const result = await fixed.host.actions.set_table_borders({ color: '#D9DEE5', width: 0.5, locations: ['Top'] });
  assert.equal(result.ok, true); assert.equal(result.result.width, 0.5); assert.equal(result.result.borders[0].width, 1); assert.match(result.result.warning, /fixed widths/);
});
