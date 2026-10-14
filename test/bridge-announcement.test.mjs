// 端口公告契约：端口可变（3000 被占会落到备用口），而调用方（agent）看不到插件设置页。
// 所以服务必须把实际端口写到一个机器可读的文件里，技能文档指向它。
// 这个文件一旦消失，「不要写死端口」就变成一句无法执行的空话 —— 技能会退化成让 agent 猜。
import { test, describe, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { cleanupStateDir, startBridge } from './helpers/server.mjs';

let srv;
before(async () => { srv = await startBridge(); });
after(async () => { if (srv) { await srv.stop(); cleanupStateDir(srv.stateDir); } });

describe('端口公告（bridge.json）', () => {
  const file = () => path.join(srv.stateDir, 'bridge.json');
  const read = () => JSON.parse(fs.readFileSync(file(), 'utf8'));

  test('服务启动后写出 bridge.json', () => {
    assert.ok(fs.existsSync(file()), `没有写出 ${file()}`);
  });

  test('公告里的端口 == 实际监听的端口', () => {
    const info = read();
    assert.equal(info.port, srv.port);
    assert.equal(info.url, `http://127.0.0.1:${srv.port}`);
  });

  test('公告里的 url 真的能连上（不是自说自话）', async () => {
    const info = read();
    const r = await fetch(`${info.url}/office/status`);
    assert.equal(r.status, 200);
  });

  test('公告是可解析 JSON，且带 plugin / pid / startedAt', () => {
    const info = read();
    assert.equal(info.plugin, 'dsh-ms-office');
    assert.equal(typeof info.pid, 'number');
    assert.ok(!Number.isNaN(Date.parse(info.startedAt)), `startedAt 不是 ISO 时间：${info.startedAt}`);
  });
});

test('default state namespace publishes only under the current package name', async () => {
  const stateHome = fs.mkdtempSync(path.join(path.dirname(srv.stateDir), 'dsh-ms-office-default-test-'));
  let isolated;
  try {
    isolated = await startBridge({ env: { DSH_HOME: stateHome, DSH_OFFICE_STATE_DIR: undefined } });
    assert.deepEqual(fs.readdirSync(stateHome), ['dsh-ms-office']);
    const announcement = JSON.parse(fs.readFileSync(path.join(stateHome, 'dsh-ms-office', 'bridge.json'), 'utf8'));
    assert.equal(announcement.plugin, 'dsh-ms-office');
    assert.equal(announcement.url, isolated.base);
    const status = await isolated.req('/office/status');
    assert.equal(status.body.plugin, 'dsh-ms-office');
  } finally {
    if (isolated) { await isolated.stop(); cleanupStateDir(isolated.stateDir); }
    cleanupStateDir(stateHome);
  }
});
