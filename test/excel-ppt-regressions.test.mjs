import { test } from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import { readFileSync } from 'node:fs';
import { loadActions, implOf } from './helpers/actions.mjs';

function excelHarness(rows = 10000, cols = 20) {
  const events = [], applied = [];
  let phase = 0;
  function range(r, c, label) {
    const object = { rowCount: r, columnCount: c, address: 'Sheet1!A1',
      load: (props) => events.push({ label, props, phase, cells: r * c }),
      getCell: () => ({ getResizedRange: (dr, dc) => range(dr + 1, dc + 1, 'bounded') }),
      format: { load: (props) => events.push({ label: 'format', props, phase, cells: r * c }), fill: { color: '#fff' }, font: { name: 'Arial', size: 12, bold: false } },
    };
    for (const key of ['values', 'text', 'formulas', 'numberFormat']) {
      let assigned;
      Object.defineProperty(object, key, { get: () => assigned || Array.from({ length: r }, () => Array(c).fill(key === 'numberFormat' ? '0.00' : 'x')), set: (v) => { assigned = v; events.push({ assigned: key, value: v }); } });
    }
    Object.defineProperty(object.format, 'numberFormat', { get() { throw new Error('Invalid format.numberFormat read'); }, set() { throw new Error('Invalid format.numberFormat write'); } });
    return object;
  }
  const original = range(rows, cols, 'original');
  const ws = { name: 'Sheet1', load() {}, getUsedRange: () => original, autoFilter: { apply: (...args) => applied.push(args) } };
  const tables = { items: [{ id: 't1', name: 'T1', getRange: () => original }], load() {} };
  const util = { okResult: (result) => ({ ok: true, result }), errResult: (code, error) => ({ ok: false, code, error }), excelRange: () => original };
  const { A, sandbox } = loadActions({ host: 'Excel', util, supported: true });
  sandbox.Excel.run = (callback) => callback({ workbook: { getSelectedRange: () => ({ rowCount: 1, columnCount: 1, load() {}, getCell: () => original }), tables, worksheets: { getActiveWorksheet: () => ws, getItem: () => ws } }, sync: async () => { phase++; } });
  return { action: (name, args = {}) => implOf(A, 'Excel', name)(args), preview: (name, args) => A.hosts.Excel.previewReaders[name](args), events, applied, original };
}

for (const [name, args, max] of [['read_range', { address: 'A1:XFD100000', limit: 7 }, 7], ['read_document', {}, 5000]]) {
  test(`${name} loads dimensions before bounded data, including very wide ranges`, async () => {
    const h = excelHarness(100000, 16384);
    const result = await h.action(name, args);
    assert.equal(result.ok, true, result.error);
    assert.equal(result.result.rows, 100000);
    assert.equal(result.result.cols, 16384);
    assert.equal(result.result.truncated, true);
    const payload = h.events.filter((e) => /values|text|formulas|numberFormat/.test(e.props || ''));
    assert.equal(payload.length, 1);
    assert.equal(payload[0].label, 'bounded');
    assert.ok(payload[0].phase >= 1);
    assert.ok(payload[0].cells <= max);
  });
}

test('read_range returns all three bounded matrices and rejects invalid limits', async () => {
  const h = excelHarness(20, 3);
  const r = await h.action('read_range', { address: 'A1:C20', limit: 7 });
  for (const key of ['values', 'formulas', 'numberFormat']) assert.equal(r.result[key].length * r.result[key][0].length, 6);
  for (const limit of [0, -1, 1.5, 'x', Infinity]) assert.equal((await h.action('read_range', { address: 'A1', limit })).code, 'bad_args');
});

for (const [includeValues, rows, expected] of [[false, 2, false], [true, 2, true], [true, 1000, false]]) {
  test(`list_tables loads values only when requested and within budget (${includeValues}, ${rows})`, async () => {
    const h = excelHarness(rows, 2);
    const r = await h.action('list_tables', { includeValues, limit: 10 });
    assert.equal(r.ok, true, r.error);
    assert.equal(h.events.some((e) => (e.props || '').includes('values')), expected);
    assert.equal(r.result.tables[0].values !== undefined, expected);
  });
}

test('format_range writes a matrix to Range.numberFormat', async () => {
  const h = excelHarness(1, 1);
  const r = await h.action('format_range', { address: 'A1', numberFormat: '0.00' });
  assert.equal(r.ok, true, r.error);
  assert.deepEqual(JSON.parse(JSON.stringify(h.events.find((e) => e.assigned).value)), [['0.00']]);
});

for (const name of ['read_styles', 'read_selection']) {
  test(`${name} loads and reads Range.numberFormat`, async () => {
    const h = excelHarness(1, 1);
    const r = await h.action(name, { withStyles: true });
    assert.equal(r.ok, true, r.error);
    assert.equal(r.result.cells[0].numberFormat[0][0], '0.00');
    assert.ok(h.events.some((e) => e.props.includes('numberFormat') && !e.props.includes('format/numberFormat')));
  });
}

test('apply_filter uses documented criteria and a separate 3-argument call per 1-based column', async () => {
  const h = excelHarness(10, 4);
  const r = await h.action('apply_filter', { address: 'Sheet1!A1:D10', columns: [{ column: 1, criteria: ['a'] }, { column: 4, filterType: 'Custom', criteria: { criterion1: '>2' } }] });
  assert.equal(r.ok, true, r.error);
  assert.equal(h.applied.length, 2);
  assert.equal(h.applied[0].length, 3);
  assert.equal(h.applied[0][0], h.original);
  assert.equal(h.applied[0][1], 0);
  assert.equal(h.applied[1][1], 3);
  assert.deepEqual(JSON.parse(JSON.stringify(h.applied[0][2])), { values: ['a'], filterOn: 'Values' });
  assert.equal(h.applied[1][2].criterion1, '>2');
});

test('invalid filter input never queues a partial filter', async () => {
  for (const columns of [{}, [null], [{ column: 0, criteria: ['a'] }], [{ column: 1.5, criteria: ['a'] }], [{ column: 5, criteria: ['a'] }], [{ column: 1, criteria1: 'ignored' }], [{ column: 1, filterType: 'Bogus', criteria: {} }], [{ column: 1, criteria: ['a'] }, { column: 2, criteria: [] }]]) {
    const h = excelHarness(10, 4);
    assert.equal((await h.action('apply_filter', { address: 'A1:D10', columns })).code, 'bad_args');
    assert.equal(h.applied.length, 0);
  }
});

test('empty filter list enables the filter range with one argument', async () => {
  const h = excelHarness(10, 4);
  assert.equal((await h.action('apply_filter', { address: 'A1:D10' })).ok, true);
  assert.equal(h.applied[0].length, 1);
});

test('replace_all preview reads bounded current state without invoking a write or simulating matching', async () => {
  const h = excelHarness(100000, 20);
  const r = await h.preview('replace_all', { search: '*x', replace: 'y' });
  assert.equal(r.ok, true, r.error);
  assert.equal(r.result.previewOnly, true);
  assert.equal(r.result.truncated, true);
  assert.equal(r.result.search, '*x');
  assert.equal(r.result.replace, 'y');
  assert.equal(r.result.replaced, undefined);
  assert.match(r.result.text, /仅预览当前状态.*已截断/);
  assert.equal((await h.action('replace_all', { search: 'x', replace: 'y', dryRun: true })).code, 'unsupported');
  assert.ok(h.events.filter((e) => (e.props || '').includes('text')).every((e) => e.cells <= 5000));
  assert.equal(h.events.some((e) => e.assigned), false);
});

// Minimal namespace-aware XML DOM mock for controlled OOXML fixtures. No browser or Office.
class FixtureDOMParser {
  parseFromString(xml) {
    const root = { children: [] }, stack = [root];
    const descendants = (node, ns, local) => node.children.flatMap((child) => [...((ns === '*' || child.ns === ns) && child.local === local ? [child] : []), ...descendants(child, ns, local)]);
    for (const token of xml.match(/<[^>]+>|[^<]+/g) || []) {
      if (token.startsWith('</')) { stack.pop(); continue; }
      if (token.startsWith('<?')) continue;
      if (!token.startsWith('<')) { stack.at(-1).textContent = (stack.at(-1).textContent || '') + token; continue; }
      const name = token.match(/^<([^\s/>]+)/)[1], attrs = {};
      for (const m of token.matchAll(/([^\s=]+)=["']([^"']*)["']/g)) attrs[m[1]] = m[2];
      const namespaces = { ...(stack.at(-1).namespaces || {}) };
      for (const [key, value] of Object.entries(attrs)) if (key === 'xmlns' || key.startsWith('xmlns:')) namespaces[key === 'xmlns' ? '' : key.slice(6)] = value;
      const [prefix, local] = name.includes(':') ? name.split(':') : ['', name];
      const node = { children: [], namespaces, ns: namespaces[prefix], local, textContent: '', getAttribute: (key) => attrs[key] || '', getAttributeNS: (ns, key) => Object.entries(attrs).find(([attr]) => attr.includes(':') && namespaces[attr.split(':')[0]] === ns && attr.split(':')[1] === key)?.[1] || '', getElementsByTagNameNS(ns, key) { return descendants(this, ns, key); }, getElementsByTagName(key) { return descendants(this, '*', key); } };
      stack.at(-1).children.push(node);
      if (!token.endsWith('/>')) stack.push(node);
    }
    root.getElementsByTagNameNS = (ns, key) => descendants(root, ns, key);
    root.getElementsByTagName = (key) => descendants(root, '*', key);
    return root;
  }
}

function pptHarness(overrides = {}) {
  const p = 'http://schemas.openxmlformats.org/presentationml/2006/main', r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const rel = (id, type, target, mode = '') => `<Relationship Id="${id}" Type="${r}/${type}" Target="${target}" ${mode ? `TargetMode="${mode}"` : ''}/>`;
  const text = (value) => `<z:sld xmlns:z="${p}" xmlns:d="${a}"><d:t>${value}</d:t></z:sld>`;
  const entries = {
    'ppt/presentation.xml': `<x:presentation xmlns:x="${p}" xmlns:q="${r}"><x:sldIdLst><x:sldId id="300" q:id="r2"/><x:sldId id="256" q:id="r1"/><x:sldId id="401" q:id="r3"/></x:sldIdLst></x:presentation>`,
    'ppt/_rels/presentation.xml.rels': `<Relationships>${rel('r1', 'slide', 'slides/slide1.xml')}${rel('r3', 'slide', '/ppt/slides/custom.xml')}${rel('r2', 'slide', './slides/slide9.xml')}</Relationships>`,
    'ppt/slides/slide1.xml': text('Second'), 'ppt/slides/slide9.xml': text('First'), 'ppt/slides/custom.xml': text('Third'), 'ppt/slides/slide2.xml': text('ORPHAN'),
    'ppt/slides/_rels/slide9.xml.rels': `<Relationships>${rel('n1', 'notesSlide', '../notesSlides/notes7.xml')}</Relationships>`,
    'ppt/slides/_rels/custom.xml.rels': `<Relationships>${rel('n3', 'notesSlide', '/ppt/notesSlides/empty.xml')}</Relationships>`,
    'ppt/notesSlides/notes7.xml': text('First notes'), 'ppt/notesSlides/empty.xml': text(''),
    ...overrides,
  };
  // Exercise the existing ZIP parser against actual stored local-file headers.
  const chunks = Object.entries(entries).filter(([, value]) => value !== null).map(([name, value]) => {
    const filename = Buffer.from(name), body = Buffer.from(value), header = Buffer.alloc(30);
    header.writeUInt32LE(0x04034b50); header.writeUInt32LE(body.length, 18); header.writeUInt16LE(filename.length, 26);
    return Buffer.concat([header, filename, body]);
  });
  const taskpane = readFileSync(new URL('../server/taskpane.js', import.meta.url), 'utf8');
  const zipContext = { TextDecoder };
  vm.createContext(zipContext);
  vm.runInContext(taskpane.slice(taskpane.indexOf('function unzip('), taskpane.indexOf('window.__UTIL__ =')), zipContext);
  const util = { okResult: (result) => ({ ok: true, result }), errResult: (code, error) => ({ ok: false, code, error }), unzip: zipContext.unzip, inflateEntry: zipContext.inflateEntry, readPptFile: (extractor) => { try { return extractor(chunks); } catch (e) { return { ok: false, code: 'execution', error: e.message }; } } };
  const { A, sandbox } = loadActions({ host: 'PowerPoint', util });
  sandbox.TextDecoder = TextDecoder; sandbox.DOMParser = FixtureDOMParser;
  return (name) => implOf(A, 'PowerPoint', name)({});
}

test('PPT text follows sldIdLst and relationships, ignores orphan files, returns stable IDs', async () => {
  const r = await pptHarness()('read_document');
  assert.equal(r.ok, true, r.error);
  assert.equal(r.result.slides, 3);
  assert.equal(r.result.text, '--- 幻灯片 1 ---\nFirst\n\n--- 幻灯片 2 ---\nSecond\n\n--- 幻灯片 3 ---\nThird');
  assert.deepEqual(Array.from(r.result.slideIds), ['300', '256', '401']);
});

test('PPT notes keep display indices across missing notes and resolve relative/absolute targets', () => {
  const r = pptHarness()('ppt_read_notes');
  assert.equal(r.ok, true, r.error);
  assert.equal(r.result.slidesWithNotes, 2);
  assert.equal(r.result.notes, '--- 幻灯片 1 备注 ---\nFirst notes\n\n--- 幻灯片 3 备注 ---\n(无备注文本)');
  assert.deepEqual(Array.from(r.result.slideIds), ['300', '401']);
});

for (const override of [{ 'ppt/presentation.xml': null }, { 'ppt/_rels/presentation.xml.rels': null }, { 'ppt/slides/slide9.xml': null }]) {
  test('PPT missing required metadata/slide reports execution error instead of filename-order fallback', async () => {
    const r = await pptHarness(override)('read_document');
    assert.equal(r.code, 'execution');
    assert.match(r.error, /Missing/);
  });
}
