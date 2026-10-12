// 桥接服务的 HTTP 冒烟测试：来源校验、错误码、能力表、版本注入、静态文件白名单。
// 服务是真实 spawn 出来的（随机端口 + 临时 state 目录），测的是用户机器上真正跑的那条路径。
import { test, describe, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { readCapabilities } from './helpers/paths.mjs';
import { cleanupStateDir, startBridge } from './helpers/server.mjs';

const capabilities = readCapabilities();
let srv;

before(async () => { srv = await startBridge(); });
after(async () => { if (srv) { await srv.stop(); cleanupStateDir(srv.stateDir); } });

describe('来源校验（/office/* 无鉴权的老问题的正解）', () => {
  test('无 Origin 的本机调用放行', async () => {
    const r = await srv.req('/office/status');
    assert.equal(r.status, 200);
  });

  test('Origin = 自身来源放行', async () => {
    const r = await srv.req('/office/status', { headers: { Origin: srv.base } });
    assert.equal(r.status, 200);
  });

  test('Origin = 外部站点 → 403 forbidden_origin', async () => {
    const r = await srv.req('/office/status', { headers: { Origin: 'http://evil.example' } });
    assert.equal(r.status, 403);
    assert.equal(r.body.code, 'forbidden_origin');
  });

  test('Host = 外部域名（DNS 重绑定）→ 403', async () => {
    const r = await srv.rawGet('/office/status', { Host: 'evil.example' });
    assert.equal(r.status, 403);
  });

  test('Host = 127.0.0.1:port 放行', async () => {
    const r = await srv.rawGet('/office/status', { Host: `127.0.0.1:${srv.port}` });
    assert.equal(r.status, 200);
  });

  test('不返回 Access-Control-Allow-Origin: *', async () => {
    const r = await srv.req('/office/status', { headers: { Origin: 'http://evil.example' } });
    assert.equal(r.headers.get('access-control-allow-origin'), null);
  });

  test('跨来源 POST /office/config → 403（历史漏洞：任意网页可改配置）', async () => {
    const r = await srv.req('/office/config', {
      method: 'POST',
      headers: { Origin: 'http://evil.example', 'Content-Type': 'application/json' },
      body: JSON.stringify({ confirmMode: 'auto' }),
    });
    assert.equal(r.status, 403);
  });
});

describe('请求体错误码区分', () => {
  test('非法 JSON → 400 bad_json', async () => {
    const r = await srv.req('/office/command', {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{not json',
    });
    assert.equal(r.status, 400);
    assert.equal(r.body.code, 'bad_json');
  });

  test('缺 action → 400 bad_args', async () => {
    const r = await srv.req('/office/command', {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({}),
    });
    assert.equal(r.status, 400);
    assert.equal(r.body.code, 'bad_args');
  });

  test('非法 host → 400 bad_args', async () => {
    const r = await srv.req('/office/command', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'read_selection', host: 'Notepad' }),
    });
    assert.equal(r.status, 400);
  });

  test('缺 instance → instance_required', async () => {
    const r = await srv.req('/office/command', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'read_selection', host: 'Word' }),
    });
    assert.equal(r.body.code, 'instance_required');
  });

  test('未注册实例 → addin_offline（快速失败，不傻等 90s）', async () => {
    const r = await srv.req('/office/command', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'read_selection', host: 'Word', instance: 'inst-nope' }),
    });
    assert.equal(r.body.code, 'addin_offline');
  });
});

describe('确认模式默认值', () => {
  test('/office/config 默认 confirmMode = ask（与窗格初值一致）', async () => {
    const r = await srv.req('/office/config');
    assert.equal(r.body.confirmMode, 'ask');
  });

  test('非法 confirmMode → 400', async () => {
    const r = await srv.req('/office/config', {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ confirmMode: 'yolo' }),
    });
    assert.equal(r.status, 400);
  });
});

describe('能力表与错误码清单', () => {
  test('/office/capabilities 返回全部 action', async () => {
    const r = await srv.req('/office/capabilities');
    assert.equal(r.status, 200);
    assert.equal(r.body.actions.length, capabilities.length);
  });

  test('/office/errors 含 body_too_large 与 confirm_required', async () => {
    const r = await srv.req('/office/errors');
    const codes = r.body.codes.map((c) => c.code);
    assert.ok(codes.includes('body_too_large'));
    assert.ok(codes.includes('confirm_required'));
  });
});

describe('/actions.js 拼接与元数据注入', () => {
  test('返回 200 且是分片拼接产物', async () => {
    const r = await srv.req('/actions.js');
    assert.equal(r.status, 200);
    for (const part of ['runtime.js', 'word.js', 'excel.js', 'ppt.js', 'common.js', 'registry.js']) {
      assert.ok(r.text.includes(`======== actions/${part} ========`), `拼接结果里缺少 ${part} 的分段标记`);
    }
  });

  test('注入 __ACTIONS_VERSION__ 与 __ACTIONS_META__', async () => {
    const r = await srv.req('/actions.js');
    assert.match(r.text, /window\.__ACTIONS_VERSION__ = ".*?";/);
    assert.match(r.text, /window\.__ACTIONS_META__ = \{.*?\};/s);
  });

  test('__ACTIONS_META__ 每条 hosts/destructive 与 capabilities.json 一致', async () => {
    const r = await srv.req('/actions.js');
    const meta = JSON.parse(r.text.match(/window\.__ACTIONS_META__ = (\{.*?\});/s)[1]);
    assert.deepEqual(Object.keys(meta).sort(), capabilities.map((c) => c.name).sort());
    for (const c of capabilities) {
      assert.deepEqual(meta[c.name].hosts, c.hosts, `${c.name} hosts 不一致`);
      assert.equal(meta[c.name].destructive, c.destructive === true, `${c.name} destructive 不一致`);
    }
  });

  test('/office/actions-version 等于各分片 mtime 的最大值', async () => {
    const fs = await import('node:fs');
    const path = await import('node:path');
    const { ACTIONS, readParts } = await import('./helpers/paths.mjs');
    const expected = String(Math.max(...readParts().map((f) => fs.statSync(path.join(ACTIONS, f)).mtimeMs)));
    const r = await srv.req('/office/actions-version');
    assert.equal(r.body.version, expected);
  });
});

describe('外壳代码版本注入', () => {
  test('/taskpane.js 注入 __CODE_VERSION__，心跳回传同一个值', async () => {
    const tp = await srv.req('/taskpane.js');
    assert.equal(tp.status, 200);
    const injected = JSON.parse(tp.text.match(/window\.__CODE_VERSION__ = (".*?");/)[1]);

    const hb = await srv.req('/office/heartbeat', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ host: 'Word', instanceId: 'inst-test-1', docUrl: 'file:///x.docx', docTitle: 'x' }),
    });
    assert.ok(hb.body.codeVersion);
    assert.equal(hb.body.codeVersion, injected, '窗格靠这个相等关系判断要不要自动 reload');
  });

  test('/office/status 列出刚注册的实例并带 codeVersion', async () => {
    const r = await srv.req('/office/status');
    assert.ok(r.body.instances.some((i) => i.instanceId === 'inst-test-1'));
    assert.ok(r.body.codeVersion);
  });
});

describe('静态文件白名单', () => {
  const blocked = ['/package.json', '/manifest.xml', '/sideload.ps1', '/capabilities.json', '/actions/word.js',
    '/actions/registry.js', '/actions/parts.json', '/%2e%2e/package.json', '/%2e%2e%2fpackage.json',
    '/..%5C..%5Cpackage.json', '/office/../package.json'];
  for (const p of blocked) {
    test(`非白名单路径被拒 ${p}`, async () => {
      const r = await srv.req(p);
      assert.equal(r.status, 404);
    });
  }

  test('/ 返回窗格页面', async () => {
    const r = await srv.req('/');
    assert.equal(r.status, 200);
    assert.ok(r.text.includes('DSH Office 执行器'));
  });

  test('/pako.min.js 可加载（窗格解压 OOXML 依赖）', async () => {
    const r = await srv.req('/pako.min.js');
    assert.equal(r.status, 200);
  });

  test('user-actions.js 不存在时兜底返回空对象', async () => {
    const r = await srv.req('/user-actions.js');
    assert.equal(r.status, 200);
    assert.ok(r.text.includes('__USER_ACTIONS__'));
  });
});
