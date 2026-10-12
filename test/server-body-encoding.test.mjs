// 请求体编码：桥接服务必须按 UTF-8 严格解码。
//
// 这条曾经真的坏过：客户端（PowerShell 的 ConvertTo-Json）按本地代码页发中文，
// Buffer.toString('utf8') 把非法字节静默替换成 U+FFFD，服务端回 200，文档标题变成乱码。
// 静默写坏用户文档比报错更糟，所以这里要求：非法 UTF-8 明确拒绝，合法 UTF-8 原样保留。
import { test, describe, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { cleanupStateDir, startBridge } from './helpers/server.mjs';

let srv;
before(async () => { srv = await startBridge(); });
after(async () => { if (srv) { await srv.stop(); cleanupStateDir(srv.stateDir); } });

function post(pathname, body) {
  return srv.req(pathname, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json; charset=utf-8' },
    body,
  });
}

describe('请求体编码', () => {
  test('非法 UTF-8 字节 → 400 bad_encoding（不再静默替换成 U+FFFD）', async () => {
    // {"action":"<0xC3>"} —— 0xC3 后面跟的是 ASCII 引号，不是续接字节
    const bad = Buffer.from([0x7b, 0x22, 0x61, 0x63, 0x74, 0x69, 0x6f, 0x6e, 0x22, 0x3a, 0x22, 0xc3, 0x22, 0x7d]);
    const r = await post('/office/command', bad);
    assert.equal(r.status, 400);
    assert.equal(r.body.code, 'bad_encoding');
  });

  test('合法 UTF-8 的中文原样保留：写入什么，读回就是什么', async () => {
    const title = 'DSH 架构详解 · 测试文档';
    const r = await post('/office/heartbeat', JSON.stringify({
      host: 'PowerPoint', instanceId: 'inst-encoding-test', docTitle: title,
    }));
    assert.equal(r.status, 200);
    const status = await srv.req('/office/status');
    const inst = status.body.instances.find((i) => i.instanceId === 'inst-encoding-test');
    assert.ok(inst, 'status 里没有这条实例（心跳没登记）');
    assert.equal(inst.docTitle, title, '中文在解码环节被改坏了');
  });

  test('带 BOM 的 JSON 也能解析（BOM 会让 JSON.parse 直接失败）', async () => {
    const r = await post('/office/command', '\uFEFF' + JSON.stringify({ action: 'read_selection', host: 'Word' }));
    assert.equal(r.status, 200); // 能走到「缺 instance」这一步，说明 BOM 已被剥掉、JSON 解析成功
    assert.equal(r.body.code, 'instance_required');
  });

  test('空 body 不算编码错误（保持原有约定）', async () => {
    const r = await post('/office/command', '');
    assert.equal(r.status, 400);
    assert.equal(r.body.code, 'bad_args');
  });

  test('编码错误后服务仍然可用（没有把连接状态搞坏）', async () => {
    await post('/office/command', Buffer.from([0x7b, 0xff, 0x22, 0x7d]));
    const r = await srv.req('/office/status');
    assert.equal(r.status, 200);
  });
});
