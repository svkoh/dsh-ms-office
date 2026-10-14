// 起一个隔离的桥接服务用于测试：随机端口 + 临时 state 目录，绝不碰用户真实的 DSH_HOME。
import { spawn } from 'node:child_process';
import net from 'node:net';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { SERVER } from './paths.mjs';

// WHATWG Fetch forbids these ports even when the loopback listener is healthy.
const unsafeHttpPorts = new Set([1,7,9,11,13,15,17,19,20,21,22,23,25,37,42,43,53,69,77,79,87,95,101,102,103,104,109,110,111,113,115,117,119,123,135,137,139,143,161,179,389,427,465,512,513,514,515,526,530,531,532,540,548,554,556,563,587,601,636,989,990,993,995,1719,1720,1723,2049,3659,4045,5060,5061,6000,6566,6665,6666,6667,6668,6669,6697,10080]);

export function freePort() {
  return new Promise((resolve, reject) => {
    const srv = net.createServer();
    srv.on('error', reject);
    srv.listen(0, '127.0.0.1', () => {
      const { port } = srv.address();
      if (unsafeHttpPorts.has(port)) {
        // Keep this socket reserved until the next candidate is acquired, so an
        // OS allocator cannot hand us the same forbidden port repeatedly.
        freePort().then(safe => srv.close(() => resolve(safe)), error => srv.close(() => reject(error)));
      } else srv.close(() => resolve(port));
    });
  });
}

/**
 * @param {{env?: Record<string,string>, readyTimeoutMs?: number}} opts
 * @returns {Promise<{port:number, base:string, stateDir:string, req:Function, rawGet:Function, stop:Function, logs:Function}>}
 */
export async function startBridge(opts = {}) {
  const port = await freePort();
  const stateDir = fs.mkdtempSync(path.join(os.tmpdir(), 'dsh-ms-office-test-'));
  const child = spawn(process.execPath, [path.join(SERVER, 'server.js')], {
    cwd: SERVER,
    env: {
      ...process.env,
      PORT: String(port),
      DSH_OFFICE_STATE_DIR: stateDir,
      ...(opts.env || {}),
      OFFICE_SKIP_SIDELOAD: '1',
    },
    stdio: ['ignore', 'pipe', 'pipe'],
  });

  let out = '';
  child.stdout.on('data', (d) => { out += d.toString(); });
  child.stderr.on('data', (d) => { out += d.toString(); });

  const base = `http://127.0.0.1:${port}`;
  const deadline = Date.now() + (opts.readyTimeoutMs ?? 15000);
  let up = false;
  while (Date.now() < deadline) {
    if (child.exitCode !== null) {
      throw new Error(`桥接服务提前退出（code=${child.exitCode}）：\n${out}`);
    }
    try {
      const r = await fetch(`${base}/office/status`);
      if (r.ok) { up = true; break; }
    } catch { /* 还没起来 */ }
    await new Promise((r) => setTimeout(r, 100));
  }
  if (!up) {
    child.kill();
    throw new Error(`桥接服务在超时内未就绪：\n${out}`);
  }

  async function req(pathname, init = {}) {
    const r = await fetch(base + pathname, init);
    const text = await r.text();
    let body = null;
    try { body = JSON.parse(text); } catch { body = text; }
    return { status: r.status, body, text, headers: r.headers };
  }

  // fetch() 不允许覆盖 Host 头（forbidden header），DNS 重绑定那条只能用裸 http 请求测
  function rawGet(pathname, headers) {
    return new Promise((resolve) => {
      const r = net.connect(port, '127.0.0.1', () => {
        r.write(`GET ${pathname} HTTP/1.1\r\nHost: ${headers.Host}\r\nConnection: close\r\n\r\n`);
      });
      let data = '';
      r.on('data', (c) => { data += c; });
      r.on('end', () => resolve({ status: Number((data.match(/^HTTP\/1\.\d (\d+)/) || [])[1] || 0), text: data }));
      r.on('error', (e) => resolve({ status: 0, text: String(e && e.message) }));
    });
  }

  return {
    port,
    base,
    stateDir,
    req,
    rawGet,
    logs: () => out,
    stop: () => new Promise((resolve) => {
      if (child.exitCode !== null) return resolve();
      child.once('exit', () => resolve());
      child.kill();
      setTimeout(() => { try { child.kill('SIGKILL'); } catch { /* 已退出 */ } resolve(); }, 3000);
    }),
  };
}

/** 供 after() 清理 */
export function cleanupStateDir(stateDir) {
  try { fs.rmSync(stateDir, { recursive: true, force: true }); } catch { /* 临时目录，删不掉也无所谓 */ }
}
