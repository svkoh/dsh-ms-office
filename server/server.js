// server.js — DSH ↔ Office 加载项桥接服务
//
// 职责：
//   1) 托管加载项静态页面（taskpane.html/js），供 Office 通过 http://127.0.0.1:<port> 加载
//   2) /office/command —— DSH 侧（agent 经 pwsh）POST 一条 Office 操作指令，同步等待加载项执行结果（最长 waitMs）
//   3) /office/poll   —— 加载项轮询取指令（带 ?host= + ?instance= 精确路由，避免多文档抢指令）
//   4) /office/result —— 加载项回传执行结果
//   5) /office/heartbeat / /office/status —— 心跳（按 host/实例记录）与状态（含在线文档列表）
//
// 多文档模型：Word / Excel / PowerPoint 各自运行一个加载项实例，指令可指定 host/instance 精确路由。
//
// 安全模型：本服务只服务两方——(a) 从本服务自己加载的 Office 加载项页面，
// (b) 本机 agent（pwsh/curl，不带 Origin 头）。浏览器发起的跨来源请求一律拒绝：
// 不返回任何 Access-Control-Allow-* 头，并且凡带 Origin 头的请求必须等于本服务自身来源；
// 同时校验 Host 头，防止 DNS 重绑定把攻击者域名指向 127.0.0.1。
//
// 无第三方依赖（Node >= 18 自带 fetch）。启动：node server.js（端口可用环境变量 PORT 覆盖）

const http = require('node:http');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const crypto = require('node:crypto');

const PORT = Number(process.env.PORT || 3000);
const ROOT = __dirname;
const COMMAND_TIMEOUT_MS = Number(process.env.COMMAND_TIMEOUT_MS || 90000);
const MAX_BODY_BYTES = Number(process.env.MAX_BODY_BYTES || 64 * 1024 * 1024);
const VALID_HOSTS = ['Word', 'Excel', 'PowerPoint'];
// 覆盖类操作确认模式：auto=自动 re-read 后执行（结果附 previousState，不打扰用户）；ask=窗格显示审批面板，用户点确认才执行
// 默认 ask（审批模式）：所有覆盖类操作都经用户确认，最安全；可用环境变量 OFFICE_CONFIRM_MODE 覆盖
let CONFIRM_MODE = process.env.OFFICE_CONFIRM_MODE || 'ask';

// ---- 能力注册表 ----
// 唯一真源是 capabilities.json：这里只负责读它并对外提供（/office/capabilities 给 AI 发现能力，
// 以及注入进 /actions.js 的 __ACTIONS_META__ 给加载项判断 hosts/destructive）。
// 改能力表只改 capabilities.json；actions 分片里只放实现，不再重复 hosts/destructive。
const CAPABILITIES = require('./capabilities.json');

// ---- 错误码注册表（供 AI 层 GET /office/errors 发现：code/含义/应如何处理/是否可重试）----
// layer: server=桥接服务入口校验；addin=加载项执行。retry: true=修好后可重发；false=不要无脑重试。
const ERROR_CODES = [
  { code: 'instance_required', layer: 'server', meaning: '指令缺少 instance 参数（多文档必须指定目标实例）', aiAction: 'GET /office/status 查看 instances（按 docUrl 识别目标文档），带 instance 重发', retry: true },
  { code: 'addin_offline', layer: 'server', meaning: '目标加载项窗格未开启或已离线', aiAction: '提醒用户打开对应文档的「DSH Office 执行器」窗格；不要重试', retry: false },
  { code: 'busy', layer: 'server', meaning: '上一条指令仍在执行中（队列占用）', aiAction: '稍等片刻后重发同一条指令', retry: true },
  { code: 'bad_json', layer: 'server', meaning: '请求体不是合法 JSON', aiAction: '检查 POST body 格式后重发', retry: true },
  { code: 'body_too_large', layer: 'server', meaning: '请求体超过服务端上限（默认 64 MiB）', aiAction: '缩小请求体（如分批写入、改用更小的图片/HTML）后重发', retry: true },
  { code: 'bad_args', layer: 'server|addin', meaning: '参数缺失或非法（如缺 action、host 非法、缺必填参数）', aiAction: '按 error 文本修正参数后重发', retry: true },
  { code: 'timeout', layer: 'server', meaning: '指令发出后超时未返回（默认 90s）', aiAction: 'GET /office/status 确认窗格在线；在线可重试一次，仍超时提醒用户重开窗格', retry: true },
  { code: 'unknown_action', layer: 'addin', meaning: 'action 名称不存在', aiAction: 'GET /office/capabilities 查看支持的 action 名，用正确名称重发', retry: true },
  { code: 'unsupported_host', layer: 'addin', meaning: '该 action 不支持当前应用（如 search 仅 Word）', aiAction: '按 error 文本换用支持的 action 或换应用', retry: true },
  { code: 'confirm_required', layer: 'addin', meaning: 'ask 模式下覆盖类操作需用户审批', aiAction: '把 result.preview 展示给用户；等待用户在 Office 窗格批准（窗格会附加 confirm:true）。AI 不得预先发送 confirm:true', retry: true },
  { code: 'rejected', layer: 'addin', meaning: '用户在窗格拒绝了该操作（或审批超时视为拒绝）', aiAction: '停止该操作、尊重用户决定，可询问替代方案；不要重发', retry: false },
  { code: 'not_found', layer: 'addin', meaning: '定位/查找目标不存在（书签、文本、命名区域、区域地址等）', aiAction: '把「未找到」如实告知用户，不要重试同一查找；可建议换关键词或换定位方式', retry: false },
  { code: 'readonly', layer: 'addin', meaning: '文档处于只读/查看模式，覆盖类操作被拒绝', aiAction: '告知用户文档为只读，请另存为可编辑副本或切换编辑模式后重试；不要重试', retry: false },
  { code: 'requirement', layer: 'addin', meaning: '当前 Office 环境缺少所需 API（如批注 API 不可用）', aiAction: '如实告知用户该功能在当前环境不可用，绝不假装成功', retry: false },
  { code: 'unsupported', layer: 'addin', meaning: '该参数组合在当前环境不支持（如 Excel replace_all 的 dryRun）', aiAction: '按 error 文本换用支持的调用方式', retry: true },
  { code: 'execution', layer: 'addin', meaning: '执行期异常', aiAction: '把 error 原样报告给用户；若为 Office.js 环境限制，按 error 文本判断', retry: false },
];

// ---- 结果语义约定（"无匹配"不是错误：只读/查找类空结果 = ok:true + count:0/空数组）----
const RESULT_SEMANTICS = [
  { scope: 'search / replace_all(dryRun)', behavior: '无命中返回 ok:true + count:0（非错误）', aiAction: '把「未找到」告知用户，不要重试' },
  { scope: 'read_* / list_* / read_range', behavior: '空内容返回 ok:true + 空数组/空字符串（非错误）', aiAction: '正常处理空结果即可' },
  { scope: 'read_selection（Word 无选区）', behavior: '返回光标上下文 "左 | 右"（| = 光标位置），非错误', aiAction: '把它当作光标位置信息使用' },
  { scope: '覆盖类写入（confirm 后执行）', behavior: '成功结果附 previousState/afterState/document', aiAction: '用 afterState/document 核验写入结果，不盲目信任' },
];

// ---- 指令队列（一次只处理一条，先到先得）----
let pending = null;        // { commandId, action, args, host, instance, resolve, timer }  host/instance: null = 任意
const heartbeats = {};     // host -> 最近心跳时间戳（按 host 摘要）
const instances = {};      // instanceId -> { host, docUrl, docTitle, lastBeat }
const hellos = {};         // host -> { ts, version } 窗格启动上报（诊断用）

const OFFLINE_MS = 15000;      // 心跳每 5s 一次，15s 无心跳视为窗格离线
const INSTANCE_TTL_MS = 60 * 60 * 1000; // 实例记录上限：避免长会话里不断开文档把 instances 撑大

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
};

// 静态文件白名单：窗格加载所需的全部文件。
// 不做目录级静态服务——server/ 里还有 manifest.xml / *.ps1 / capabilities.json / package.json，
// 它们没有任何理由出现在 HTTP 面上。白名单同时让「路径穿越」从"靠 URL 规范化兜住"变成"不可能命中"。
// 注意 actions.js 不在磁盘上：它是下面 ACTIONS_PARTS 的拼接产物（见 readActionsBundle）。
const STATIC_FILES = new Set(['taskpane.html', 'taskpane.js', 'actions.js', 'pako.min.js', 'user-actions.js']);

// actions 分片（顺序 = 依赖顺序，清单在 actions/parts.json，测试也读同一份）。
// 拆分的理由：action 实现按宿主分文件，Word/Excel/PowerPoint 的代码不再互相穿插；
// 拼接的理由：浏览器侧的加载契约（一个 <script src="/actions.js">）和热更新机制保持不变。
const ACTIONS_PARTS = require('./actions/parts.json').parts.map((f) => `actions/${f}`);

function sendJson(res, code, obj, extraHeaders) {
  const body = JSON.stringify(obj);
  res.writeHead(code, {
    'Content-Type': 'application/json; charset=utf-8',
    'Content-Length': Buffer.byteLength(body),
    ...(extraHeaders || {}),
  });
  res.end(body);
}

// ---- 本机来源校验 ----
// 浏览器一定会在跨来源请求（含表单 POST / fetch POST）上带 Origin 头，且无法伪造；
// 因此「带 Origin 且不等于自身来源」= 一定是别的网页在打本服务 → 拒绝。
// Host 校验挡 DNS 重绑定（攻击者域名解析到 127.0.0.1 时 Host 会是攻击者域名）。
const SELF_ORIGINS = new Set([`http://127.0.0.1:${PORT}`, `http://localhost:${PORT}`]);
const SELF_HOSTS = new Set([`127.0.0.1:${PORT}`, `localhost:${PORT}`, '127.0.0.1', 'localhost']);

function isLocalRequest(req) {
  const origin = req.headers.origin;
  if (typeof origin === 'string' && origin && !SELF_ORIGINS.has(origin)) return false;
  const host = req.headers.host;
  if (typeof host === 'string' && host && !SELF_HOSTS.has(host)) return false;
  return true;
}

// ---- 运行期目录（settings.json 与按端口改写过的 manifest 都写这里，不碰源码树）----
function stateDir() {
  const env = process.env.DSH_OFFICE_STATE_DIR;
  if (typeof env === 'string' && env.trim()) return env.trim();
  const home = process.env.DSH_HOME && process.env.DSH_HOME.trim() ? process.env.DSH_HOME.trim() : path.join(os.homedir(), '.dsh');
  return path.join(home, 'dsh-office-for-mso');
}

// 代码版本：taskpane.js 的 mtime（毫秒）。窗格心跳比对发现不一致 → 自动 reload，无需手动重开窗格。
function getCodeVersion() {
  try {
    const st = fs.statSync(path.join(ROOT, 'taskpane.js'));
    return String(st.mtimeMs);
  } catch (e) { return '0'; }
}

// actions 版本：**所有分片** mtime 的最大值（毫秒），与 /office/actions-version 一致。
// 取最大值而不是只取入口，是因为改了任意一个宿主分片都必须触发窗格热更新。
function getActionsVersion() {
  let max = 0;
  for (const rel of ACTIONS_PARTS) {
    try { max = Math.max(max, fs.statSync(path.join(ROOT, rel)).mtimeMs); } catch (e) { /* 缺分片由 readActionsBundle 报错 */ }
  }
  return String(max);
}

// 把分片按序拼成一个脚本。缺分片时抛错——静默少拼一段会让窗格缺 action 且难以察觉。
function readActionsBundle() {
  const parts = ACTIONS_PARTS.map((rel) => {
    const full = path.join(ROOT, rel);
    if (!fs.existsSync(full)) throw new Error(`actions 分片缺失: ${rel}`);
    return `// ======== ${rel} ========\n${fs.readFileSync(full, 'utf8').trimEnd()}`;
  });
  return parts.join('\n\n') + '\n';
}

// 注入进 /actions.js 的运行期元数据：只带加载项真正要用的字段（hosts / destructive）。
// 让能力表只有一个真源（capabilities.json），actions 分片不必再抄一遍。
function actionsMeta() {
  const meta = {};
  for (const entry of CAPABILITIES) meta[entry.name] = { hosts: entry.hosts, destructive: entry.destructive === true };
  return meta;
}

// ---- 启动时按 PORT 生成加载项清单，并注册 WEF（Windows）----
// Office 窗格只能从清单里的 URL 加载；3000 被占时会落到备用口，必须同步改写。
// 生成物写到 stateDir()，源码树里的 manifest.xml 永远保持模板原样（端口 3000）。
function writeRuntimeManifest(port) {
  const template = fs.readFileSync(path.join(ROOT, 'manifest.xml'), 'utf8');
  const xml = template.replace(
    /http:\/\/(?:localhost|127\.0\.0\.1):\d+\/taskpane\.html/g,
    `http://127.0.0.1:${port}/taskpane.html`
  );
  const file = path.join(stateDir(), 'manifest.xml');
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, xml, 'utf8');
  return file;
}

function ensureSideloaded() {
  let manifest;
  try {
    manifest = writeRuntimeManifest(PORT);
  } catch (e) {
    console.log('[dsh-office-for-mso] 无法生成运行期清单，跳过自动注册:', e.message);
    return;
  }
  if (process.platform !== 'win32') {
    console.log(`[dsh-office-for-mso] 非 Windows 平台，跳过加载项自动注册；手动加载清单：${manifest}`);
    return;
  }
  const { spawnSync } = require('node:child_process');
  const reg = (args) => spawnSync('reg', args, { stdio: 'ignore', windowsHide: true }).status === 0;
  try {
    const xml = fs.readFileSync(manifest, 'utf8');
    const m = xml.match(/<Id>\s*([0-9a-fA-F-]{36})\s*<\/Id>/);
    if (!m) { console.log('[dsh-office-for-mso] manifest.xml 缺少 <Id>，跳过自动注册'); return; }
    const guid = m[1];
    const key = 'HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer';
    // Always refresh the registry value so a port change rewrites the path Office reads.
    if (reg(['add', key, '/v', guid, '/t', 'REG_SZ', '/d', manifest, '/f'])) {
      console.log(`[dsh-office-for-mso] 加载项已注册 (WEF: ${guid}) → http://127.0.0.1:${PORT}/taskpane.html`);
      console.log('[dsh-office-for-mso] 若端口刚变更，请关闭并重新打开 Office 文档以刷新窗格');
    } else {
      console.log(`[dsh-office-for-mso] 加载项自动注册失败（可手动运行 sideload.ps1）；清单在 ${manifest}`);
    }
  } catch (e) {
    console.log('[dsh-office-for-mso] 加载项自动注册跳过:', e.message);
  }
}

// 请求体读取：分块累积后一次性解码（避免大 body 上反复拼接字符串）。
// 超限抛 code='body_too_large'，非法 JSON 抛 code='bad_json'——两者都不该被当成同一种错误。
//
// 超限后**不能** req.destroy()：那会把 socket 一起销毁，413 还没写出去连接就断了，
// 调用方只会看到 "socket hang up" 而不是错误码。正确做法是停止缓冲、继续把剩余数据读完
// （下面 settled 后直接丢弃），让响应正常发出。
function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    let settled = false;
    const fail = (error) => {
      if (settled) return;
      settled = true;
      reject(error);
    };
    req.on('data', (chunk) => {
      if (settled) return; // 已判定失败：剩余分块继续读并丢弃，只为让响应能写出去
      size += chunk.length;
      if (size > MAX_BODY_BYTES) {
        const error = new Error(`body too large (max ${MAX_BODY_BYTES} bytes)`);
        error.code = 'body_too_large';
        return fail(error);
      }
      chunks.push(chunk);
    });
    req.on('end', () => {
      if (settled) return;
      settled = true;
      const text = Buffer.concat(chunks).toString('utf8');
      if (!text) return resolve({});
      try {
        resolve(JSON.parse(text));
      } catch (e) {
        const error = new Error('bad json');
        error.code = 'bad_json';
        reject(error);
      }
    });
    req.on('error', fail);
  });
}

// 统一的 body 读取失败应答：区分「太大」和「不是 JSON」
function sendBodyError(res, error) {
  if (error && error.code === 'body_too_large') {
    // 请求体没读完就回包，这条连接不能再复用（否则剩余字节会被当成下一个请求）
    sendJson(res, 413, { ok: false, code: 'body_too_large', error: error.message }, { connection: 'close' });
  } else {
    sendJson(res, 400, { ok: false, code: 'bad_json', error: 'bad json' });
  }
}

/** 丢弃长时间没有心跳的实例记录（窗格关掉后 instances 不会自己消失）。 */
function pruneInstances(now) {
  for (const [id, value] of Object.entries(instances)) {
    if (now - value.lastBeat > INSTANCE_TTL_MS) delete instances[id];
  }
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host || '127.0.0.1'}`);
  const p = url.pathname;

  // ---- 来源校验：任何请求都必须来自本机自身来源（或本机无 Origin 的 agent 调用）----
  if (!isLocalRequest(req)) {
    sendJson(res, 403, { ok: false, code: 'forbidden_origin', error: 'cross-origin request rejected: this bridge only serves its own add-in page and local callers' });
    return;
  }

  // ---- 静态文件（白名单内；API 路径 /office/* 已单独处理）----
  if (req.method === 'GET' && !p.startsWith('/office/')) {
    let file = '';
    try { file = p === '/' ? 'taskpane.html' : decodeURIComponent(p.slice(1)); } catch { /* 非法百分号编码 → 走 404 */ }
    if (!STATIC_FILES.has(file)) { sendJson(res, 404, { error: 'not found' }); return; }
    // user-actions.js：用户定制文件（gitignore）；不存在时返回空对象，保证窗格总能加载
    if (file === 'user-actions.js') {
      const ua = path.join(ROOT, 'user-actions.js');
      const body = fs.existsSync(ua) ? fs.readFileSync(ua, 'utf8') : 'window.__USER_ACTIONS__ = {};\n';
      res.writeHead(200, { 'Content-Type': MIME['.js'], 'Cache-Control': 'no-store' });
      res.end(body);
      return;
    }
    // actions.js：磁盘上没有这个文件，它是 actions/ 下分片的拼接产物。
    // 前缀注入 actions 版本 + 能力元数据：
    //   版本 → 窗格热更新后可验证新脚本确实执行（防止加载成功但 IIFE 抛错/旧值残留）
    //   元数据 → hosts/destructive 的唯一真源（capabilities.json），分片里不再自带一份
    if (file === 'actions.js') {
      let src;
      try { src = readActionsBundle(); }
      catch (e) {
        // 拼不出来就说清楚是哪一段缺了，而不是给窗格一个空脚本（那会表现成「所有 action 都不认识」）
        res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8', 'Cache-Control': 'no-store' });
        res.end(`actions bundle unavailable: ${String(e && e.message || e)}\n`);
        return;
      }
      res.writeHead(200, { 'Content-Type': MIME['.js'], 'Cache-Control': 'no-store' });
      res.end(
        `;window.__ACTIONS_VERSION__ = ${JSON.stringify(getActionsVersion())};\n`
        + `;window.__ACTIONS_META__ = ${JSON.stringify(actionsMeta())};\n`
        + src
      );
      return;
    }
    const full = path.join(ROOT, file);
    if (fs.existsSync(full) && fs.statSync(full).isFile()) {
      const ext = path.extname(full).toLowerCase();
      res.writeHead(200, { 'Content-Type': MIME[ext] || 'application/octet-stream', 'Cache-Control': 'no-store' });
      if (file === 'taskpane.js') {
        // 前缀注入代码版本标记（须在文件头部，taskpane.js 顶部读取 window.__CODE_VERSION__）
        const src = fs.readFileSync(full, 'utf8');
        res.end(`;window.__CODE_VERSION__ = ${JSON.stringify(getCodeVersion())};\n` + src);
        return;
      }
      fs.createReadStream(full).pipe(res);
      return;
    }
    sendJson(res, 404, { error: 'not found' });
    return;
  }

  if (req.method === 'OPTIONS') { res.writeHead(204); res.end(); return; }

  // POST /office/command —— DSH 侧发指令，同步等待结果；body.host 可选（Word/Excel/PowerPoint），缺省=任意文档
  if (req.method === 'POST' && p === '/office/command') {
    let body;
    try { body = await readBody(req); } catch (e) { sendBodyError(res, e); return; }
    if (!body.action || typeof body.action !== 'string') { sendJson(res, 400, { ok: false, code: 'bad_args', error: 'action required' }); return; }
    const host = body.host && typeof body.host === 'string' ? body.host : null;
    if (host && !VALID_HOSTS.includes(host)) { sendJson(res, 400, { ok: false, code: 'bad_args', error: `invalid host: ${host} (valid: ${VALID_HOSTS.join('/')})` }); return; }
    const instance = body.instance && typeof body.instance === 'string' ? body.instance : null; // 必填：指定加载项实例（多文档精确路由，防误执行）
    if (!instance) {
      sendJson(res, 200, { ok: false, code: 'instance_required', error: '必须指定加载项实例（instanceId）：先 GET /office/status 查看 instances（按 docUrl 识别目标文档），带 instance 重发；多文档时未指定实例会导致指令被错误文档执行' });
      return;
    }
    if (pending) { sendJson(res, 409, { ok: false, code: 'busy', error: 'a command is already pending (previous command still executing); retry shortly' }); return; }
    // 快速失败：目标 host/instance 窗格不在线时立即返回明确错误（而不是傻等 90s 超时）
    const now = Date.now();
    pruneInstances(now);
    if (!instances[instance]) {
      sendJson(res, 200, { ok: false, code: 'addin_offline', error: `指定的加载项实例不存在（instanceId=${instance}）：请确认目标文档的窗格已开启，或重新查 /office/status 获取当前实例` });
      return;
    }
    if (now - instances[instance].lastBeat > OFFLINE_MS) {
      sendJson(res, 200, { ok: false, code: 'addin_offline', error: '指定的加载项实例已离线（窗格关闭）：请重新打开该文档的窗格' });
      return;
    }
    if (host && (!heartbeats[host] || now - heartbeats[host] > OFFLINE_MS)) {
      sendJson(res, 200, { ok: false, code: 'addin_offline', error: `「${host}」加载项窗格未开启：请打开该文档，并在 开始/开发人员 → 加载项 → 开发人员加载项 中打开「DSH Office 执行器」窗格并保持开启，然后重试` });
      return;
    }
    const commandId = crypto.randomUUID();
    const result = await new Promise((resolve) => {
      const timer = setTimeout(() => resolve({ ok: false, code: 'timeout', error: 'timeout: no add-in heartbeat or execution', timeout: true }), COMMAND_TIMEOUT_MS);
      pending = { commandId, action: body.action, args: body.args || {}, host, instance, resolve: (r) => { clearTimeout(timer); resolve(r); } };
    });
    if (pending && pending.commandId === commandId) pending = null; // 清理（超时路径）
    sendJson(res, 200, { ok: result.ok !== false, commandId, host: host || 'any', instance: instance || 'any', ...result });
    return;
  }

  // GET /office/poll —— 加载项取指令；?host= + ?instance= 精确匹配（多文档时只有目标实例能取走）
  if (req.method === 'GET' && p === '/office/poll') {
    const host = url.searchParams.get('host') || '';
    const instance = url.searchParams.get('instance') || '';
    if (pending && (!pending.host || pending.host === host) && (!pending.instance || pending.instance === instance)) {
      sendJson(res, 200, { commandId: pending.commandId, action: pending.action, args: pending.args });
    } else {
      sendJson(res, 200, { commandId: null });
    }
    return;
  }

  // POST /office/result —— 加载项回传结果
  if (req.method === 'POST' && p === '/office/result') {
    let body;
    try { body = await readBody(req); } catch (e) { sendBodyError(res, e); return; }
    if (pending && pending.commandId === body.commandId) {
      const r = pending.resolve;
      pending = null;
      // 错误也转发 code/result（如 confirm_required 的 result.preview），否则 AI 拿不到预览
      r(body.ok === false
        ? { ok: false, error: body.error || 'add-in error', code: body.code, result: body.result }
        : { ok: true, result: body.result });
      sendJson(res, 200, { ok: true });
    } else {
      sendJson(res, 404, { ok: false, error: 'unknown or expired commandId' });
    }
    return;
  }

  // POST /office/heartbeat —— body: { host, instanceId, docUrl, docTitle }；实例注册 + host 摘要；响应带 codeVersion
  if (req.method === 'POST' && p === '/office/heartbeat') {
    let body;
    try { body = await readBody(req); } catch { body = {}; }
    const host = body.host && typeof body.host === 'string' && VALID_HOSTS.includes(body.host) ? body.host : null;
    const now = Date.now();
    if (host) heartbeats[host] = now;
    if (host && body.instanceId) {
      instances[String(body.instanceId)] = { host, docUrl: body.docUrl ? String(body.docUrl) : '', docTitle: body.docTitle ? String(body.docTitle) : '', lastBeat: now };
      pruneInstances(now);
    }
    sendJson(res, 200, { ok: true, host, codeVersion: getCodeVersion() });
    return;
  }

  // POST /office/hello —— 窗格启动上报（每次 JS 加载执行一次；用于验证自动刷新/reload 是否生效）
  if (req.method === 'POST' && p === '/office/hello') {
    let body;
    try { body = await readBody(req); } catch { body = {}; }
    const host = body.host && typeof body.host === 'string' && VALID_HOSTS.includes(body.host) ? body.host : null;
    if (host) hellos[host] = { ts: Date.now(), version: body.version || '' };
    sendJson(res, 200, { ok: true });
    return;
  }

  // GET /office/status —— 状态：pending 指令 + 各在线文档（host 摘要）+ 实例明细（含 docUrl 识别文档）+ 窗格启动记录
  if (req.method === 'GET' && p === '/office/status') {
    const now = Date.now();
    pruneInstances(now);
    const hosts = {};
    for (const h of VALID_HOSTS) if (heartbeats[h]) hosts[h] = heartbeats[h];
    const instanceList = Object.entries(instances).map(([id, v]) => ({
      instanceId: id,
      host: v.host,
      docUrl: v.docUrl,
      docTitle: v.docTitle,
      online: now - v.lastBeat <= OFFLINE_MS,
      lastBeat: v.lastBeat,
    }));
    sendJson(res, 200, {
      pending: pending ? { commandId: pending.commandId, action: pending.action, host: pending.host || 'any', instance: pending.instance || 'any' } : null,
      hosts,
      instances: instanceList,
      hellos,
      codeVersion: getCodeVersion(),
      lastHeartbeat: Object.keys(heartbeats).length ? Math.max(...Object.values(heartbeats)) : 0,
    });
    return;
  }

  // GET /office/actions-version —— actions 分片版本（全部分片 mtime 的最大值），窗格据此热更新实现
  if (req.method === 'GET' && p === '/office/actions-version') {
    sendJson(res, 200, { version: getActionsVersion() });
    return;
  }

  // GET/POST /office/config —— 运行时配置（confirmMode: auto|ask，窗格开关同步用）
  if (req.method === 'GET' && p === '/office/config') {
    sendJson(res, 200, { confirmMode: CONFIRM_MODE });
    return;
  }
  if (req.method === 'POST' && p === '/office/config') {
    let body;
    try { body = await readBody(req); } catch { body = {}; }
    if (body.confirmMode === 'auto' || body.confirmMode === 'ask') {
      CONFIRM_MODE = body.confirmMode;
      sendJson(res, 200, { ok: true, confirmMode: CONFIRM_MODE });
    } else {
      sendJson(res, 400, { ok: false, code: 'bad_args', error: 'confirmMode must be "auto" or "ask"' });
    }
    return;
  }

  // GET /office/capabilities —— 能力发现：返回 action 注册表（名称/平台/是否破坏性/参数说明）
  if (req.method === 'GET' && p === '/office/capabilities') {
    sendJson(res, 200, { version: '2.0.0', actions: CAPABILITIES });
    return;
  }

  // GET /office/errors —— 错误码发现：返回错误码枚举 + 结果语义约定（"无匹配"≠错误）
  if (req.method === 'GET' && p === '/office/errors') {
    sendJson(res, 200, { version: '1.1.0', codes: ERROR_CODES, semantics: RESULT_SEMANTICS });
    return;
  }

  sendJson(res, 404, { error: 'unknown endpoint' });
});

server.listen(PORT, '127.0.0.1', () => {
  console.log(`[dsh-office-for-mso] listening on http://127.0.0.1:${PORT}`);
  console.log(`  add-in taskpane: http://127.0.0.1:${PORT}/taskpane.html`);
  console.log(`  command endpoint: POST /office/command  (timeout ${COMMAND_TIMEOUT_MS}ms, host/instance routed)`);
  console.log(`  capabilities: ${CAPABILITIES.length} actions`);
  ensureSideloaded();
});
