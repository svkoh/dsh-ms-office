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

// ---- 文案（i18n）----
// 结构与文案分离：capabilities.json 只留代码真正依赖的元数据（name/hosts/destructive），
// 给人看、给 AI 看的文字全在 i18n/<lang>.json。这样改一句参数说明不会动到决定
// 「平台校验」和「审批拦截」的数据，而且「加了 action 却忘了写说明」变成一个可测的缺口。
const I18N_DIR = path.join(__dirname, '..', 'i18n');
const LANGS = ['zh', 'en'];
const I18N = {};
for (const lang of LANGS) I18N[lang] = JSON.parse(fs.readFileSync(path.join(I18N_DIR, `${lang}.json`), 'utf8'));

// 语言探测优先级：?lang= → Accept-Language → 环境变量 OFFICE_LANG → 系统区域 → zh。
// 判定规则与项目其余部分一致：en* 用英文，其余一律中文（不做「不认识就回退英文」）。
function normalizeLang(value) {
  if (typeof value !== 'string') return null;
  const tag = value.trim().toLowerCase();
  if (tag.startsWith('en')) return 'en';
  if (tag.startsWith('zh')) return 'zh';
  return null;
}
const DEFAULT_LANG =
  normalizeLang(process.env.OFFICE_LANG)
  || normalizeLang(Intl.DateTimeFormat().resolvedOptions().locale)
  || 'zh';

function langOf(url, req) {
  return normalizeLang(url.searchParams.get('lang'))
    || normalizeLang(req.headers['accept-language'])
    || DEFAULT_LANG;
}

/** 取服务端文案；{name} 占位符用 vars 替换。 */
function msg(lang, key, vars) {
  let text = (I18N[lang] && I18N[lang].server.messages[key]) || key;
  if (vars) for (const k of Object.keys(vars)) text = text.split('{' + k + '}').join(String(vars[k]));
  return text;
}

// ---- 错误码注册表（供 AI 层 GET /office/errors 发现：code/含义/应如何处理/是否可重试）----
// 这里只放结构（code/layer/retry）；含义与处理建议是文案，按请求语言从 i18n 取。
// layer: server=桥接服务入口校验；addin=加载项执行。retry: true=修好后可重发；false=不要无脑重试。
const ERROR_CODES = [
  { code: 'instance_required', layer: 'server', retry: true },
  { code: 'addin_offline', layer: 'server', retry: false },
  { code: 'busy', layer: 'server', retry: true },
  { code: 'bad_json', layer: 'server', retry: true },
  { code: 'body_too_large', layer: 'server', retry: true },
  { code: 'bad_args', layer: 'server|addin', retry: true },
  { code: 'timeout', layer: 'server', retry: true },
  { code: 'unknown_action', layer: 'addin', retry: true },
  { code: 'unsupported_host', layer: 'addin', retry: true },
  { code: 'confirm_required', layer: 'addin', retry: true },
  { code: 'rejected', layer: 'addin', retry: false },
  { code: 'not_found', layer: 'addin', retry: false },
  { code: 'readonly', layer: 'addin', retry: false },
  { code: 'requirement', layer: 'addin', retry: false },
  { code: 'unsupported', layer: 'addin', retry: true },
  { code: 'execution', layer: 'addin', retry: false },
];

/** 错误码 + 该语言的含义/处理建议。结果语义约定（"无匹配"≠错误）整条都是文案，直接从 i18n 取。 */
function errorCatalog(lang) {
  const dict = I18N[lang].server.errors;
  return ERROR_CODES.map((e) => ({ ...e, meaning: dict[e.code].meaning, aiAction: dict[e.code].action }));
}

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

// 代码版本：taskpane.js 与 i18n 字典的 mtime 最大值（毫秒）。
// 窗格心跳比对发现不一致 → 自动 reload，无需手动重开窗格。
// 把 i18n 一起算进来，是因为改文案（i18n/*.json）后窗格必须重新取字典，
// 否则会出现「改了文案但窗格没反应」——这与改 taskpane.js 是同一类问题，所以用同一个机制。
const CODE_VERSION_FILES = [
  path.join(ROOT, 'taskpane.js'),
  ...LANGS.map((l) => path.join(I18N_DIR, `${l}.json`)),
];
function getCodeVersion() {
  let max = 0;
  for (const file of CODE_VERSION_FILES) {
    try { max = Math.max(max, fs.statSync(file).mtimeMs); } catch (e) { /* 缺文件时忽略 */ }
  }
  return String(max);
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

// ---- 把实际端口写成机器可读的公告 ----
// 端口是可变的（3000 被占会落到备用口），而调用方（agent）看不到插件设置页。
// 没有这个文件，「不要写死端口」就只是一句无法执行的空话——所以它是那条规则的落点。
// 与 manifest.xml 一样写在 stateDir()，源码树保持干净。
function writeBridgeAnnouncement(port) {
  const file = path.join(stateDir(), 'bridge.json');
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const payload = {
    plugin: 'dsh-office-for-mso',
    url: `http://127.0.0.1:${port}`,
    port,
    pid: process.pid,
    startedAt: new Date().toISOString(),
  };
  fs.writeFileSync(file, `${JSON.stringify(payload, null, 2)}\n`, 'utf8');
  return file;
}

function ensureSideloaded() {
  let manifest;
  try {
    manifest = writeRuntimeManifest(PORT);
  } catch (e) {
    console.log(`[dsh-office-for-mso] ${msg(DEFAULT_LANG, 'logManifestWriteFailed', { message: e.message })}`);
    return;
  }
  if (process.platform !== 'win32') {
    console.log(`[dsh-office-for-mso] ${msg(DEFAULT_LANG, 'logNonWindows', { manifest })}`);
    return;
  }
  const { spawnSync } = require('node:child_process');
  const reg = (args) => spawnSync('reg', args, { stdio: 'ignore', windowsHide: true }).status === 0;
  try {
    const xml = fs.readFileSync(manifest, 'utf8');
    const m = xml.match(/<Id>\s*([0-9a-fA-F-]{36})\s*<\/Id>/);
    if (!m) { console.log(`[dsh-office-for-mso] ${msg(DEFAULT_LANG, 'logManifestNoId')}`); return; }
    const guid = m[1];
    const key = 'HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer';
    // Always refresh the registry value so a port change rewrites the path Office reads.
    if (reg(['add', key, '/v', guid, '/t', 'REG_SZ', '/d', manifest, '/f'])) {
      console.log(`[dsh-office-for-mso] ${msg(DEFAULT_LANG, 'logRegistered', { guid, url: `http://127.0.0.1:${PORT}/taskpane.html` })}`);
      console.log(`[dsh-office-for-mso] ${msg(DEFAULT_LANG, 'logPortChanged')}`);
    } else {
      console.log(`[dsh-office-for-mso] ${msg(DEFAULT_LANG, 'logRegisterFailed', { manifest })}`);
    }
  } catch (e) {
    console.log(`[dsh-office-for-mso] ${msg(DEFAULT_LANG, 'logRegisterSkipped', { message: e.message })}`);
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
function sendBodyError(res, error, lang) {
  if (error && error.code === 'body_too_large') {
    // 请求体没读完就回包，这条连接不能再复用（否则剩余字节会被当成下一个请求）
    sendJson(res, 413, { ok: false, code: 'body_too_large', error: msg(lang, 'bodyTooLarge', { max: MAX_BODY_BYTES }) }, { connection: 'close' });
  } else {
    sendJson(res, 400, { ok: false, code: 'bad_json', error: msg(lang, 'badJson') });
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
  // 本次请求的应答语言（决定所有错误文案与能力表说明用哪种语言）
  const lang = langOf(url, req);

  // ---- 来源校验：任何请求都必须来自本机自身来源（或本机无 Origin 的 agent 调用）----
  if (!isLocalRequest(req)) {
    sendJson(res, 403, { ok: false, code: 'forbidden_origin', error: msg(lang, 'forbiddenOrigin') });
    return;
  }

  // ---- 静态文件（白名单内；API 路径 /office/* 已单独处理）----
  if (req.method === 'GET' && !p.startsWith('/office/')) {
    let file = '';
    try { file = p === '/' ? 'taskpane.html' : decodeURIComponent(p.slice(1)); } catch { /* 非法百分号编码 → 走 404 */ }
    if (!STATIC_FILES.has(file)) { sendJson(res, 404, { error: msg(lang, 'notFound') }); return; }
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
        res.end(msg(lang, 'actionsBundleUnavailable', { message: String(e && e.message || e) }) + '\n');
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
    sendJson(res, 404, { error: msg(lang, 'notFound') });
    return;
  }

  if (req.method === 'OPTIONS') { res.writeHead(204); res.end(); return; }

  // POST /office/command —— DSH 侧发指令，同步等待结果；body.host 可选（Word/Excel/PowerPoint），缺省=任意文档
  if (req.method === 'POST' && p === '/office/command') {
    let body;
    try { body = await readBody(req); } catch (e) { sendBodyError(res, e, lang); return; }
    if (!body.action || typeof body.action !== 'string') { sendJson(res, 400, { ok: false, code: 'bad_args', error: msg(lang, 'actionRequired') }); return; }
    const host = body.host && typeof body.host === 'string' ? body.host : null;
    if (host && !VALID_HOSTS.includes(host)) { sendJson(res, 400, { ok: false, code: 'bad_args', error: msg(lang, 'invalidHost', { host, valid: VALID_HOSTS.join('/') }) }); return; }
    const instance = body.instance && typeof body.instance === 'string' ? body.instance : null; // 必填：指定加载项实例（多文档精确路由，防误执行）
    if (!instance) {
      sendJson(res, 200, { ok: false, code: 'instance_required', error: msg(lang, 'instanceRequired') });
      return;
    }
    if (pending) { sendJson(res, 409, { ok: false, code: 'busy', error: msg(lang, 'busy') }); return; }
    // 快速失败：目标 host/instance 窗格不在线时立即返回明确错误（而不是傻等 90s 超时）
    const now = Date.now();
    pruneInstances(now);
    if (!instances[instance]) {
      sendJson(res, 200, { ok: false, code: 'addin_offline', error: msg(lang, 'addinOfflineInstanceMissing', { instance }) });
      return;
    }
    if (now - instances[instance].lastBeat > OFFLINE_MS) {
      sendJson(res, 200, { ok: false, code: 'addin_offline', error: msg(lang, 'addinOfflineStale') });
      return;
    }
    if (host && (!heartbeats[host] || now - heartbeats[host] > OFFLINE_MS)) {
      sendJson(res, 200, { ok: false, code: 'addin_offline', error: msg(lang, 'addinOfflineHost', { host }) });
      return;
    }
    const commandId = crypto.randomUUID();
    const result = await new Promise((resolve) => {
      const timer = setTimeout(() => resolve({ ok: false, code: 'timeout', error: msg(lang, 'timeout'), timeout: true }), COMMAND_TIMEOUT_MS);
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
    try { body = await readBody(req); } catch (e) { sendBodyError(res, e, lang); return; }
    if (pending && pending.commandId === body.commandId) {
      const r = pending.resolve;
      pending = null;
      // 错误也转发 code/result（如 confirm_required 的 result.preview），否则 AI 拿不到预览
      r(body.ok === false
        ? { ok: false, error: body.error || msg(lang, 'addInError'), code: body.code, result: body.result }
        : { ok: true, result: body.result });
      sendJson(res, 200, { ok: true });
    } else {
      sendJson(res, 404, { ok: false, error: msg(lang, 'unknownCommandId') });
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
      sendJson(res, 400, { ok: false, code: 'bad_args', error: msg(lang, 'confirmModeInvalid') });
    }
    return;
  }

  // GET /office/capabilities —— 能力发现：结构（capabilities.json）+ 该语言的参数说明（i18n）
  // 对外形状与拆分前一致（name/hosts/destructive/requirement/args/description），只是文案按语言取。
  if (req.method === 'GET' && p === '/office/capabilities') {
    const text = I18N[lang].capabilities;
    const actions = CAPABILITIES.map((entry) => {
      const t = text[entry.name];
      return t ? { ...entry, ...t } : { ...entry };
    });
    sendJson(res, 200, { version: '2.0.0', lang, actions });
    return;
  }

  // GET /office/errors —— 错误码发现：返回错误码枚举 + 结果语义约定（"无匹配"≠错误）
  if (req.method === 'GET' && p === '/office/errors') {
    sendJson(res, 200, { version: '1.1.0', lang, codes: errorCatalog(lang), semantics: I18N[lang].server.semantics });
    return;
  }

  // GET /office/i18n —— 窗格 UI 文案。窗格在 Office.onReady 之后按 Office 界面语言取一次；
  // 与 /office/* 其余端点一样受来源校验保护，所以不进静态白名单。
  if (req.method === 'GET' && p === '/office/i18n') {
    sendJson(res, 200, { lang, pane: I18N[lang].pane });
    return;
  }

  sendJson(res, 404, { error: msg(lang, 'unknownEndpoint') });
});

server.listen(PORT, '127.0.0.1', () => {
  const L = DEFAULT_LANG;
  console.log(`[dsh-office-for-mso] ${msg(L, 'logListening', { url: `http://127.0.0.1:${PORT}` })}`);
  console.log(`  ${msg(L, 'logTaskpaneUrl', { url: `http://127.0.0.1:${PORT}/taskpane.html` })}`);
  console.log(`  ${msg(L, 'logCommandEndpoint', { timeout: COMMAND_TIMEOUT_MS })}`);
  console.log(`  ${msg(L, 'logCapabilityCount', { count: CAPABILITIES.length })}`);
  try {
    console.log(`  ${msg(L, 'logBridgeAnnounced', { file: writeBridgeAnnouncement(PORT) })}`);
  } catch (e) {
    console.log(`  ${msg(L, 'logBridgeAnnounceFailed', { message: e.message })}`);
  }
  ensureSideloaded();
});
