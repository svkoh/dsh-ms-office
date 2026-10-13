import { test } from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import { readFileSync } from 'node:fs';

const source = (path) => readFileSync(new URL('../' + path, import.meta.url), 'utf8');
const registrySource = source('server/actions/registry.js');
const paneSource = source('server/taskpane.js');
const wordSource = source('server/actions/word.js');
const ok = (result) => ({ ok: true, result });
const err = (code, error) => ({ ok: false, code, error });
const tick = () => new Promise(setImmediate);

function registryFixture({ mode = 'ask', preview = () => ok({ text: 'before' }), read = () => ok({ text: 'document' }) } = {}) {
  let writes = 0, approvedFactory;
  const window = {
    __CONFIRM_MODE__: mode,
    __REGISTER_APPROVAL_EXECUTOR__: (factory) => { approvedFactory = factory; },
    __ACTIONS_META__: { write: { hosts: ['Word'], destructive: true }, read_document: { hosts: ['Word'], destructive: false } },
    __OFFICE_ACTIONS__: { util: { errResult: err }, hosts: { Word: {
      actions: { write: async (args) => { writes++; assert.equal(args.confirm, undefined); return ok({ written: true }); }, read_document: read },
      previewReaders: { write: preview },
    } } },
  };
  const context = vm.createContext({ window, Office: { context: { host: 'Word', document: {} } } });
  vm.runInContext(registrySource, context);
  return { window, writes: () => writes, approve: (cmd) => approvedFactory(cmd) };
}

test('registry ignores external confirm and approval-shaped args in ask mode', async () => {
  const f = registryFixture();
  const out = await f.window.__EXECUTE__('write', { confirm: true, commandId: 'forged', approvedBy: 'user' });
  assert.equal(out.code, 'confirm_required');
  assert.equal(f.writes(), 0);
});

test('private approval snapshots the command, is single use, and awaits after-state verification', async () => {
  const f = registryFixture();
  const cmd = { commandId: 'one', action: 'write', args: { confirm: true } };
  const run = f.approve(cmd);
  cmd.commandId = 'two'; cmd.action = 'missing';
  const out = await run();
  assert.equal(out.ok, true);
  assert.equal(out.result.afterState.text, 'before');
  assert.equal(out.result.document, 'document');
  assert.equal((await run()).code, 'rejected');
  assert.equal(f.writes(), 1);
});

test('successful write survives thrown and returned post-verification failures', async () => {
  for (const failure of [() => { throw Error('read broken'); }, () => Promise.reject(Error('async read broken')), () => err('execution', 'read failed')]) {
    const f = registryFixture({ preview: failure, read: failure });
    const out = await f.approve({ commandId: 'one', action: 'write', args: {} })();
    assert.equal(out.ok, true);
    assert.equal(out.result.written, true);
    assert.equal(out.result.verificationWarnings.length, 2);
    assert.equal(f.writes(), 1);
  }
});

test('auto mode awaits writes and refuses failed pre-write preview', async () => {
  const good = registryFixture({ mode: 'auto' });
  const out = await good.window.__EXECUTE__('write', { confirm: true });
  assert.equal(out.result.approvedBy, 'auto');
  assert.equal(out.result.previousState.text, 'before');
  assert.equal(out.result.afterState.text, 'before');
  const bad = registryFixture({ mode: 'auto', preview: () => { throw Error('preview failed'); } });
  assert.equal((await bad.window.__EXECUTE__('write', {})).ok, false);
  assert.equal(bad.writes(), 0);
});

function paneFixture({ preview = () => ok({ text: 'before' }), claim = true, resultStatus = 200, pollCommand = null } = {}) {
  const nodes = new Map();
  const node = (id) => {
    if (!nodes.has(id)) nodes.set(id, { textContent: '', className: '', hidden: false, setAttribute() {}, querySelectorAll: () => [], addEventListener() {} });
    return nodes.get(id);
  };
  let timerId = 0, writes = 0;
  const timers = new Map(), requests = [];
  const window = {};
  const context = vm.createContext({ window, document: { getElementById: node, documentElement: {} },
    Office: { context: { host: 'Word', document: {} }, onReady() {} },
    setTimeout: (callback, delay) => { const id = ++timerId; timers.set(id, { callback, delay }); return id; },
    clearTimeout: (id) => timers.delete(id),
    fetch: async (url, options) => {
      const body = options && options.body ? JSON.parse(options.body) : undefined;
      requests.push({ url, body });
      if (url.includes('/office/poll?')) return { ok: true, json: async () => pollCommand };
      if (url.endsWith('/office/result')) return { ok: resultStatus < 400, status: resultStatus, json: async () => ({ ok: true }) };
      if (url.endsWith('/start')) return { ok: claim, status: 409, json: async () => ({ ok: true, commandId: decodeURIComponent(url.split('/').at(-2)), status: 'executing' }) };
      return { ok: true, json: async () => ({ ok: true }) };
    },
  });
  vm.runInContext(paneSource, context);
  window.__ACTIONS__ = { write: { destructive: true, preview } };
  window.__REGISTER_APPROVAL_EXECUTOR__((cmd) => {
    const id = cmd.commandId;
    return async () => { writes++; return ok({ commandId: id }); };
  });
  const command = (id = 'one') => ({ commandId: id, action: 'write', args: { confirm: true }, expiresAt: Date.now() + 180000 });
  return { context, window, requests, timers, node, command, writes: () => writes,
    show: (cmd = command()) => { context.cmd = cmd; return vm.runInContext('showApproval(cmd)', context); },
    decide: (decision) => { context.decision = decision; vm.runInContext('approvalPending.resolve(decision)', context); },
  };
}

test('pane refuses both thrown and failed previews without claiming or writing', async () => {
  for (const preview of [() => { throw Error('broken'); }, () => err('execution', 'broken')]) {
    const f = paneFixture({ preview });
    await f.show();
    assert.equal(f.writes(), 0);
    assert.equal(f.requests.some((r) => r.url.endsWith('/start')), false);
    assert.equal(f.requests.at(-1).body.code, 'execution');
    assert.equal(f.timers.size, 0);
  }
});

test('pane blocks early approval, claims after preview and user approval, and includes instance', async () => {
  let finish;
  const f = paneFixture({ preview: () => new Promise((resolve) => { finish = resolve; }) });
  const done = f.show();
  await tick();
  f.decide('confirm');
  assert.equal(f.node('approval-confirm').disabled, true);
  assert.equal(f.requests.length, 0);
  finish(ok({ text: 'before' }));
  await tick();
  assert.equal(f.node('approval-confirm').disabled, false);
  f.decide('confirm');
  await done;
  assert.equal(f.writes(), 1);
  assert.ok(f.requests[0].url.endsWith('/commands/one/start'));
  assert.ok(f.requests[0].body.instance.startsWith('inst-'));
  assert.equal(f.requests[1].body.instance, f.requests[0].body.instance);
  assert.equal(f.requests[1].body.result.approvedBy, 'user');
  assert.equal(f.timers.size, 0);
});

test('pane refuses expired command and failed atomic claim after approval', async () => {
  for (const expired of [true, false]) {
    const f = paneFixture({ claim: false });
    const cmd = f.command();
    const done = f.show(cmd);
    await tick();
    if (expired) cmd.expiresAt = Date.now() - 1;
    f.decide('confirm'); const out = await done;
    assert.equal(f.writes(), 0);
    assert.equal(out.ok, false);
    assert.equal(f.requests.some((r) => r.url.endsWith('/office/result')), false);
    assert.equal(f.requests.some((r) => r.url.endsWith('/start')), !expired);
  }
});

test('old approval timeout cannot resolve the next approval; hanging preview can time out', async () => {
  const f = paneFixture();
  const first = f.show();
  const oldCallback = [...f.timers.values()][0].callback;
  await tick(); f.decide('reject'); await first;
  const second = f.show(f.command('two'));
  await tick(); oldCallback(); await tick();
  assert.equal(f.requests.length, 1);
  f.decide('confirm'); await second;
  assert.equal(f.writes(), 1);
  const hung = paneFixture({ preview: () => new Promise(() => {}) });
  const done = hung.show();
  [...hung.timers.values()][0].callback();
  await done;
  assert.equal(hung.requests.at(-1).body.code, 'rejected');
  assert.equal(hung.writes(), 0);
});

test('pane poll cannot bypass approval with confirm:true and all direct executions claim first', async () => {
  const cmd = { commandId: 'polled', action: 'write', args: { confirm: true }, expiresAt: Date.now() + 180000 };
  const ask = paneFixture({ pollCommand: cmd });
  vm.runInContext('ensureActions = async () => {}', ask.context);
  const waiting = vm.runInContext('poll()', ask.context);
  await tick();
  assert.equal(ask.node('approval').hidden, false);
  assert.equal(ask.requests.some((r) => r.url.endsWith('/start')), false);
  ask.decide('reject'); await waiting;
  assert.equal(ask.writes(), 0);
  assert.equal(cmd.args.confirm, undefined);
  for (const destructive of [false, true]) {
    const direct = paneFixture({ pollCommand: { ...cmd, args: {} } });
    direct.window.__ACTIONS__.write.destructive = destructive;
    direct.window.__EXECUTE__ = async () => {
      assert.ok(direct.requests.at(-1).url.endsWith('/start'));
      return ok({ written: true });
    };
    vm.runInContext("ensureActions = async () => {}; confirmMode = 'auto'", direct.context);
    await vm.runInContext('poll()', direct.context);
    assert.equal(direct.requests.at(-1).body.ok, true);
  }
});

test('expired claims and stale result responses do not mark the bridge disconnected', async () => {
  const expired = paneFixture({ pollCommand: { commandId: 'old', action: 'read', args: {}, expiresAt: Date.now() - 1 } });
  vm.runInContext('ensureActions = async () => {}', expired.context);
  await vm.runInContext('poll()', expired.context);
  assert.equal(vm.runInContext('pendingStatus.key', expired.context), 'connected');
  const stale = paneFixture({ resultStatus: 404 });
  const done = stale.show(); await tick(); stale.decide('reject'); await done;
  assert.equal(stale.timers.size, 0);
});

test('registry hot reload refreshes the private pane executor without exposing approval globally', async () => {
  const f = paneFixture();
  let writes = 0;
  f.window.__ACTIONS_META__ = { write: { hosts: ['Word'], destructive: true } };
  f.window.__OFFICE_ACTIONS__ = { util: { errResult: err }, hosts: { Word: {
    actions: { write: async () => ok({ version: ++writes }) },
    previewReaders: { write: () => ok({ text: 'before' }) },
  } } };
  vm.runInContext(registrySource, f.context);
  const first = f.show(); await tick(); f.decide('confirm'); await first;
  vm.runInContext(registrySource, f.context);
  const second = f.show(f.command('two')); await tick(); f.decide('confirm'); await second;
  assert.equal(writes, 2);
  assert.equal(f.window.__EXECUTE_APPROVED__, undefined);
  const denied = await f.window.__EXECUTE__('write', { confirm: true });
  assert.equal(denied.code, 'confirm_required');
});

function wordFixture() {
  const calls = [];
  let selected = 'original';
  const range = (id) => ({ text: id, font: {}, load() {}, select() { selected = id; },
    getRange: () => range(id), getOoxml: () => ({ value: '' }),
    insertText() {}, insertField() {},
    insertParagraph: (_text, location) => { calls.push({ type: 'paragraph', target: id, location }); return { getRange: () => range('inserted') }; },
    insertBreak: (_type, location) => calls.push({ type: 'break', target: id, location }),
    insertOoxml: (_xml, location) => calls.push({ type: 'equation', target: id, location }),
    insertTable: (_rows, _cols, location, values) => { calls.push({ type: 'table', target: id, location }); return { rowCount: 1, values, load() {}, getRange: () => range('inserted') }; },
    paragraphs: { getFirst: () => ({ insertParagraph: (_text, location) => { calls.push({ type: 'paragraph', target: id + '-paragraph', location }); return { getRange: () => range('inserted') }; } }) },
  });
  const body = range('body');
  body.paragraphs = { load() {}, items: ['first', 'last'].map((id) => ({ getRange: (edge) => range(id + '-' + edge) })) };
  const ctx = { document: { body, getSelection: () => range(selected), sections: { items: [{}], load() {} } }, sync: async () => {} };
  const A = { hosts: {}, util: { okResult: ok, errResult: err, mapBuiltInStyle: () => null }, isSetSupported: () => true };
  const Word = { run: (fn) => fn(ctx), InsertLocation: Object.fromEntries(['start', 'end', 'before', 'after', 'replace'].map((v) => [v, v])) };
  vm.runInNewContext(wordSource, { window: { __OFFICE_ACTIONS__: A }, Word, Office: { context: {} } });
  return { ...A.hosts.Word, calls, selected: () => selected };
}

test('Word selection-relative insertions retain original selection through preview and execution', async () => {
  const cases = [
    ['word_insert_caption', { label: 'Figure', title: 'title' }, 'original'],
    ['word_insert_caption', { label: 'Figure', title: 'title', position: 'beforeSelection' }, 'original'],
    ['word_insert_break', { type: 'page' }, 'original'],
    ['word_insert_break', { type: 'page', location: 'beforeSelection' }, 'original'],
    ['word_insert_equation', { expression: 'x' }, 'original'],
    ['word_insert_equation', { expression: 'x', target: { kind: 'selection' }, mode: 'replace' }, 'original'],
    ['insert_table', { rows: [['x']], location: 'beforeSelection' }, 'original'],
    ['insert_table', { rows: [['x']], location: 'afterSelection' }, 'original'],
    ['insert_paragraph', { text: 'x', location: 'afterSelectionParagraph' }, 'original-paragraph'],
    ['insert_paragraph', { text: 'x', location: 'afterSelection' }, 'original'],
  ];
  for (const [action, args, target] of cases) {
    const f = wordFixture();
    await f.selectTargets[action](args);
    assert.equal(f.selected(), 'original', action);
    const out = await f.actions[action](args);
    assert.equal(out.ok, true, out.error);
    assert.equal(f.calls[0].target, target, action);
  }
});

test('Word body start/end and equation search targets follow their own semantics', async () => {
  for (const [action, args, edge] of [
    ['insert_paragraph', { location: 'start' }, 'first-Start'],
    ['insert_paragraph', {}, 'last-End'],
    ['insert_table', { location: 'start' }, 'first-Start'],
    ['word_insert_break', { location: 'end' }, 'last-End'],
    ['word_insert_equation', { target: { kind: 'body' }, mode: 'prepend' }, 'first-Start'],
  ]) {
    const f = wordFixture(); await f.selectTargets[action](args); assert.equal(f.selected(), edge);
  }
  const f = wordFixture();
  await f.selectTargets.word_insert_equation({ target: { kind: 'search', query: 'x' } });
  assert.equal(f.selected(), 'original');
});
