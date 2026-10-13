import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { startBridge, cleanupStateDir } from './helpers/server.mjs';

let bridge;
before(async () => { bridge = await startBridge({ env: { COMMAND_TIMEOUT_MS: '800' } }); });
after(async () => { if (bridge) { await bridge.stop(); cleanupStateDir(bridge.stateDir); } });
const post = (route, body) => bridge.req(route, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
async function register(instance, host = 'Word') {
  await post('/office/heartbeat', { host, instanceId: instance });
}
async function waitCommand(instance, host = 'Word') {
  for (let i = 0; i < 100; i++) {
    const { body } = await bridge.req(`/office/poll?instance=${instance}&host=${host}`);
    if (body.commandId) return body;
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
  throw new Error('command did not become available');
}
const submit = (instance, extra = {}) => post('/office/command', { action: 'read_selection', instance, args: {}, ...extra });
const claim = (command, instance) => post(`/office/commands/${command.commandId}/start`, { instance });
const complete = (command, instance, result = { text: 'done' }) => post('/office/result', { commandId: command.commandId, instance, ok: true, result });

test('non-object JSON is rejected on every POST endpoint and the server stays alive', async () => {
  for (const route of ['/office/command', '/office/result', '/office/heartbeat', '/office/hello', '/office/config']) {
    for (const body of [null, [], 7, 'text', true]) {
      const response = await post(route, body);
      assert.equal(response.status, 400, route);
      assert.equal(response.body.code, 'bad_args', route);
    }
  }
  assert.equal((await bridge.req('/office/status')).status, 200);
});

test('invalid command fields fail before routing', async () => {
  for (const patch of [{ args: null }, { args: [] }, { args: 'text' }, { host: 5 }, { instance: {} }, { idempotencyKey: [] }]) {
    const response = await submit('invalid-fields', patch);
    assert.equal(response.status, 400);
    assert.equal(response.body.code, 'bad_args');
  }
});

test('external confirm is stripped and a command can only be claimed once by its target', async () => {
  const instance = 'confirm-target';
  await register(instance);
  const waiting = submit(instance, { args: { confirm: true, text: 'safe' } });
  const command = await waitCommand(instance);
  assert.equal(command.args.confirm, undefined);
  assert.equal(command.args.text, 'safe');
  assert.ok(command.expiresAt > Date.now());
  assert.equal((await claim(command, 'wrong-instance')).status, 400);
  assert.equal((await complete(command, instance)).status, 409);
  assert.equal((await claim(command, instance)).body.status, 'executing');
  assert.equal((await claim(command, instance)).status, 409);
  assert.equal((await bridge.req(`/office/poll?instance=${instance}&host=Word`)).body.commandId, null);
  await complete(command, instance);
  assert.equal((await waiting).body.ok, true);
  const status = await bridge.req(`/office/commands/${command.commandId}`);
  assert.equal(status.body.status, 'completed');
  assert.equal(status.body.response.result.text, 'done');
});

test('documents run independently while the same instance is serialized', async () => {
  await register('parallel-word'); await register('parallel-excel', 'Excel');
  const first = submit('parallel-word');
  const word = await waitCommand('parallel-word');
  const second = submit('parallel-excel');
  const excel = await waitCommand('parallel-excel', 'Excel');
  assert.equal((await submit('parallel-word')).body.code, 'busy');
  assert.equal((await bridge.req('/office/status')).body.pendingCommands.length, 2);
  await claim(word, 'parallel-word'); await claim(excel, 'parallel-excel');
  await complete(excel, 'parallel-excel'); await complete(word, 'parallel-word');
  assert.equal((await first).body.ok, true); assert.equal((await second).body.ok, true);
});

test('host must match the named document instance', async () => {
  await register('host-mismatch', 'Excel');
  assert.equal((await submit('host-mismatch', { host: 'Word' })).status, 400);
});

test('unstarted expired commands cannot be claimed or succeed later', async () => {
  await register('expires-before-execution');
  const waiting = submit('expires-before-execution');
  const command = await waitCommand('expires-before-execution');
  const response = await waiting;
  assert.equal(response.body.status, 'expired');
  assert.equal(response.body.executionStarted, false);
  assert.equal((await claim(command, 'expires-before-execution')).body.code, 'timeout');
  assert.equal((await complete(command, 'expires-before-execution')).status, 404);
});

test('execution timeout preserves the slot and a late result can still be queried', async () => {
  const instance = 'expires-during-execution';
  await register(instance);
  const waiting = submit(instance);
  const command = await waitCommand(instance);
  await claim(command, instance);
  const response = await waiting;
  assert.equal(response.body.status, 'execution_unknown');
  assert.equal(response.body.executionStarted, true);
  assert.equal((await submit(instance)).body.code, 'busy');
  assert.equal((await complete(command, 'wrong-instance')).status, 400);
  assert.equal((await complete(command, instance, { text: 'late' })).body.ok, true);
  const status = await bridge.req(`/office/commands/${command.commandId}`);
  assert.equal(status.body.status, 'completed');
  assert.equal(status.body.response.result.text, 'late');
});

test('idempotent retries share the command and completed result; changed arguments are rejected', async () => {
  const instance = 'idempotency-target';
  await register(instance);
  const extra = { idempotencyKey: 'operation-1' };
  const first = submit(instance, extra);
  const command = await waitCommand(instance);
  const retry = submit(instance, extra);
  assert.equal((await submit(instance, { ...extra, args: { changed: true } })).status, 409);
  await claim(command, instance); await complete(command, instance);
  const a = await first, b = await retry;
  assert.equal(a.body.commandId, b.body.commandId);
  const replay = await submit(instance, extra);
  assert.equal(replay.body.commandId, a.body.commandId);
  assert.deepEqual(replay.body.result, a.body.result);
  assert.equal((await bridge.req(`/office/poll?instance=${instance}&host=Word`)).body.commandId, null);
});

test('test startup writes an isolated manifest without Office registration', () => {
  assert.ok(fs.existsSync(path.join(bridge.stateDir, 'manifest.xml')));
  assert.doesNotMatch(bridge.logs(), /add-in registered|加载项已注册/);
});
