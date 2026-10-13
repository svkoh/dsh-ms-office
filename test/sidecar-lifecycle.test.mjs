import { test } from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import { readFileSync } from 'node:fs';
import { EventEmitter } from 'node:events';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const source = readFileSync(new URL('../lib/sidecar-host.mjs', import.meta.url), 'utf8');
const deferred = () => {
  let resolve;
  const promise = new Promise((done) => { resolve = done; });
  return { promise, resolve };
};
const flush = async () => { for (let i = 0; i < 30; i++) await Promise.resolve(); };

function harness({ health, config = {} } = {}) {
  let settings = { version: 1, port: null, autostart: true, ...config }, clock = 0;
  const children = [], writes = [], requests = [], timers = new Map(), routes = new Map();
  const fs = {
    readFileSync: () => JSON.stringify(settings),
    mkdirSync() {},
    writeFileSync(_file, body) { settings = JSON.parse(body); writes.push({ ...settings }); },
    readdirSync() { throw new Error('Lifecycle tests must not publish skills'); },
  };
  const net = { createServer() {
    const server = new EventEmitter();
    server.listen = () => { queueMicrotask(() => server.emit('listening')); return server; };
    server.close = (callback) => callback();
    return server;
  } };
  function spawn(_command, _args, options) {
    const child = new EventEmitter();
    child.pid = 700 + children.length;
    child.exitCode = null;
    child.stdout = new EventEmitter(); child.stderr = new EventEmitter();
    child.signals = [];
    child.kill = (signal) => { child.signals.push(signal); child.exitCode = 0; child.emit('exit', 0); return true; };
    child.options = options;
    children.push(child);
    return child;
  }
  const sandbox = {
    fs, net, spawn, path, fileURLToPath, os: { homedir: () => '/mock/home' },
    process: { env: { DSH_HOME: '/mock/dsh' }, platform: 'linux', execPath: '/mock/node' },
    console: { log() {} }, Buffer, TextDecoder, Error,
    AbortSignal: { timeout: () => ({ mocked: true }) },
    Date: { now: () => clock },
    setTimeout: (fn, delay) => { const id = {}; timers.set(id, { fn, delay }); return id; },
    clearTimeout: (id) => timers.delete(id),
    fetch: async (url) => {
      requests.push(url);
      return health ? health(children.at(-1), requests.length) : { ok: true, json: async () => ({ plugin: 'test-office', pid: children.at(-1).pid }) };
    },
  };
  vm.createContext(sandbox);
  const executable = source.replace(/^import .*;\r?\n/gm, '').replace('export function runSidecarHost', 'function runSidecarHost').replaceAll('import.meta.url', JSON.stringify('file:///mock/plugin/lib/sidecar-host.mjs'));
  vm.runInContext(executable + '\nglobalThis.lifecycle = { startServer, applySettings, killChild, waitHealth, registerRoutes, statusView };', sandbox, { filename: 'sidecar-host.mjs' });
  const state = { pluginName: 'test-office', pluginDir: '/mock/plugin', entry: 'server/server.js', defaultPort: 37531, healthPath: '/health', bridge: '/test-office', child: null, pid: null, disposed: false, running: false, phase: 'stopped', restarts: [], starting: false, startPromise: null, settingsPromise: Promise.resolve(), restartTimer: null, requestedPort: null, currentPort: 0, lastError: null, logs: '' };
  const ctx = { webServer: { register(route) { routes.set(route.path, route.handler); return () => routes.delete(route.path); } } };
  const tick = async (delay, now = clock) => {
    clock = now;
    const due = [...timers].filter(([, timer]) => timer.delay === delay);
    for (const [id, timer] of due) { timers.delete(id); timer.fn(); }
    await flush();
  };
  const save = (configPatch) => {
    const req = { async *[Symbol.asyncIterator]() { yield Buffer.from(JSON.stringify({ config: configPatch })); } };
    const response = { status: null, body: null, writeHead(status) { this.status = status; }, end(body) { this.body = JSON.parse(body); } };
    const promise = routes.get('/test-office/save')(req, response);
    return { promise, response };
  };
  return { ...sandbox.lifecycle, state, children, writes, requests, timers, routes, ctx, tick, save, setConfig: (patch) => { settings = { ...settings, ...patch }; }, getConfig: () => settings };
}

test('startup remains starting until health confirms both plugin and child PID, and concurrent starts share a launch', async () => {
  const gate = deferred();
  const h = harness({ health: () => gate.promise });
  const pending = h.startServer(h.state);
  assert.equal(h.startServer(h.state), pending);
  await flush();
  assert.equal(h.children.length, 1);
  assert.equal(h.state.running, false);
  assert.equal(h.state.starting, true);
  assert.equal(h.statusView(h.state).phase, 'starting');
  assert.equal(h.children[0].options.env.PORT, '37531');
  gate.resolve({ ok: true, json: async () => ({ plugin: 'test-office', pid: h.children[0].pid }) });
  await pending;
  assert.equal(h.state.running, true);
  assert.equal(h.state.phase, 'healthy');
  assert.equal(h.state.starting, false);
  assert.equal(h.state.startPromise, null);
});

for (const mismatch of ['plugin', 'pid']) {
  test(`health ${mismatch} mismatch cannot mark a different service healthy`, async () => {
    const h = harness({ health: (child) => ({ ok: true, json: async () => ({ plugin: mismatch === 'plugin' ? 'unrelated-service' : 'test-office', pid: mismatch === 'pid' ? child.pid + 1 : child.pid }) }) });
    const pending = h.startServer(h.state);
    await flush();
    assert.equal(h.state.phase, 'starting');
    assert.equal(h.state.running, false);
    await h.tick(250, 10001);
    await pending;
    assert.equal(h.state.phase, 'unhealthy');
    assert.equal(h.state.running, false);
    assert.match(h.state.lastError, /health timeout/);
    assert.equal(h.requests.length, 2);
  });
}

test('late exit and error from an old child cannot clear a healthy replacement or schedule a restart', async () => {
  const h = harness();
  await h.startServer(h.state);
  const previous = h.children[0];
  h.setConfig({ port: 38000 });
  await h.applySettings(h.state);
  const replacement = h.children[1];
  assert.equal(previous.signals[0], 'SIGTERM');
  assert.equal(h.state.child, replacement);
  previous.emit('exit', 99);
  previous.emit('error', new Error('late old-child error'));
  assert.equal(h.state.child, replacement);
  assert.equal(h.state.pid, replacement.pid);
  assert.equal(h.state.running, true);
  assert.equal(h.state.phase, 'healthy');
  assert.equal(h.state.lastError, null);
  assert.equal(h.state.restarts.length, 0);
  assert.equal(h.state.restartTimer, null);
  assert.equal(h.timers.size, 0);
});

test('disabling autostart stops the child and cancels scheduled restart state', async () => {
  const h = harness();
  await h.startServer(h.state);
  const child = h.state.child;
  child.exitCode = 2; child.emit('exit', 2);
  assert.equal(h.state.phase, 'unhealthy');
  assert.ok(h.state.restartTimer);
  h.setConfig({ autostart: false });
  await h.applySettings(h.state);
  assert.equal(h.state.child, null);
  assert.equal(h.state.pid, null);
  assert.equal(h.state.running, false);
  assert.equal(h.state.phase, 'stopped');
  assert.equal(h.state.restartTimer, null);
  assert.equal(h.timers.size, 0);
  await h.tick(1000);
  assert.equal(h.children.length, 1);
});

test('disabling autostart while health is pending kills the child and prevents late health from reviving it', async () => {
  const gate = deferred();
  const h = harness({ health: () => gate.promise });
  const pending = h.startServer(h.state);
  await flush();
  const child = h.state.child;
  h.setConfig({ autostart: false });
  const stop = h.applySettings(h.state);
  await flush();
  assert.equal(child.signals[0], 'SIGTERM');
  assert.equal(h.state.child, null);
  gate.resolve({ ok: true, json: async () => ({ plugin: 'test-office', pid: child.pid }) });
  await flush();
  await h.tick(250);
  await Promise.all([pending, stop]);
  assert.equal(h.state.phase, 'stopped');
  assert.equal(h.state.running, false);
  assert.equal(h.state.starting, false);
  assert.equal(h.state.pid, null);
  assert.equal(h.timers.size, 0);
});

test('concurrent save requests serialize persistence and lifecycle changes, preserving previous patch fields', async () => {
  const gate = deferred();
  const h = harness({ health: () => gate.promise, config: { autostart: false } });
  const dispose = h.registerRoutes(h.state, h.ctx);
  const first = h.save({ autostart: true, port: 39000 });
  await flush();
  assert.equal(h.children.length, 1);
  const second = h.save({ autostart: false });
  await flush();
  assert.equal(h.writes.length, 1, 'second save must not persist while the first is still waiting for health');
  assert.equal(second.response.body, null);
  assert.equal(h.state.phase, 'starting');
  gate.resolve({ ok: true, json: async () => ({ plugin: 'test-office', pid: h.children[0].pid }) });
  await Promise.all([first.promise, second.promise]);
  assert.equal(h.writes.length, 2);
  assert.equal(h.writes[1].port, 39000);
  assert.equal(h.writes[1].autostart, false);
  assert.equal(first.response.status, 200);
  assert.equal(second.response.status, 200);
  assert.equal(second.response.body.status.phase, 'stopped');
  assert.equal(h.state.child, null);
  assert.equal(h.children[0].signals[0], 'SIGTERM');
  assert.equal(h.timers.size, 0);
  dispose();
  assert.equal(h.routes.size, 0);
});

test('a rejected save does not prevent a later valid save from applying', async () => {
  const h = harness({ config: { autostart: false } });
  h.registerRoutes(h.state, h.ctx);
  const bad = h.save({ port: 80 });
  const good = h.save({ port: 39001, autostart: false });
  await Promise.all([bad.promise, good.promise]);
  assert.equal(bad.response.status, 500);
  assert.match(bad.response.body.error, /port/);
  assert.equal(good.response.status, 200);
  assert.equal(h.writes.length, 1);
  assert.equal(h.getConfig().port, 39001);
  assert.equal(h.children.length, 0);
});
