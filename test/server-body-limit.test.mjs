// 请求体超限：必须回 413 body_too_large，而不是把连接直接掐掉。
//
// 这条曾经真的坏过：超限时调用 req.destroy()，socket 一起没了，413 还没写出去，
// 客户端只看到 "socket hang up"。**光看代码会以为已经修好了**，所以这里用真发一次超限请求来测。
import { test, describe, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { cleanupStateDir, startBridge } from './helpers/server.mjs';

const LIMIT = 2048;
let srv;

before(async () => { srv = await startBridge({ env: { MAX_BODY_BYTES: String(LIMIT) } }); });
after(async () => { if (srv) { await srv.stop(); cleanupStateDir(srv.stateDir); } });

function bigBody(bytes, validJson) {
  const filler = 'x'.repeat(Math.max(0, bytes - 40));
  return validJson ? JSON.stringify({ action: 'read_selection', pad: filler }) : `{"pad":"${filler}"`;
}

describe(`请求体上限（MAX_BODY_BYTES=${LIMIT}）`, () => {
  test('超限的**合法 JSON** → 413 body_too_large', async () => {
    const r = await srv.req('/office/command', {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: bigBody(LIMIT * 4, true),
    });
    assert.equal(r.status, 413);
    assert.equal(r.body.code, 'body_too_large');
  });

  test('超限的**非法 JSON** 也回 413（不能因为解析失败就抢先报 400）', async () => {
    const r = await srv.req('/office/command', {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: bigBody(LIMIT * 4, false),
    });
    assert.equal(r.status, 413);
    assert.equal(r.body.code, 'body_too_large');
  });

  test('未超限的非法 JSON → 400 bad_json（区分开，不是一律 413）', async () => {
    const r = await srv.req('/office/command', {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{not json',
    });
    assert.equal(r.status, 400);
    assert.equal(r.body.code, 'bad_json');
  });

  test('未超限的合法请求正常处理（上限没被设得过紧）', async () => {
    const r = await srv.req('/office/command', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'read_selection', host: 'Word' }),
    });
    assert.equal(r.status, 200);
    assert.equal(r.body.code, 'instance_required');
  });

  test('超限响应后服务仍然可用（没有把连接状态搞坏）', async () => {
    await srv.req('/office/command', {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: bigBody(LIMIT * 8, true),
    });
    const r = await srv.req('/office/status');
    assert.equal(r.status, 200);
  });
});
