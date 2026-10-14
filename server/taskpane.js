// taskpane.js — DSH Office 执行器窗格（actions 热更新架构）
// 职责：窗格 UI（含 i18n）、桥接轮询/心跳、execute 分发、actions.js 热更新加载。
//
// 所有 action 实现都在 /actions.js（服务端把 server/actions/ 下的分片拼成这一个脚本，可热更新）：
// 改任一分片 → 重启桥接服务 → 窗格自动重新加载，无需重开窗格。
//
// 本文件（窗格）应尽量少改：改本文件后，服务端心跳回传的 codeVersion 会与本页注入的
// window.__CODE_VERSION__ 不一致，窗格在下次心跳时自动 reload（见 heartbeat()）。
//
// 文案不写在这里，也不写死在 HTML 里：服务端把 i18n/<lang>.json 的 pane 段经 /office/i18n 发出，
// 本页按 Office 界面语言取一次再填进 DOM。改文案只改 i18n/*.json，不必碰窗格代码。

const BRIDGE =
  typeof location !== "undefined" &&
  location.origin &&
  /^https?:\/\//i.test(location.origin)
    ? location.origin
    : "http://127.0.0.1:3000";
const POLL_MS = 1000;
const HEARTBEAT_MS = 5000;

const statusEl = document.getElementById('status');
const statusTextEl = document.getElementById('ui-status');
const logEl = document.getElementById('log');
const hostValueEl = document.getElementById('ui-host-value');
const docValueEl = document.getElementById('ui-doc-value');
const modeSeg = document.getElementById('mode-seg');

// ================= i18n =================
// 语言判定与项目其余部分一致：en* 用英文，其余一律中文（不做「不认识就回退英文」）。
let LANG = 'zh';
let T = null;

function normalizeLang(value) {
  const tag = String(value || '').trim().toLowerCase();
  if (tag.startsWith('en')) return 'en';
  if (tag.startsWith('zh')) return 'zh';
  return null;
}

function t(key, vars) {
  if (!T) return key;
  let text = T[key];
  if (typeof text !== 'string') return key;
  if (vars) for (const k of Object.keys(vars)) text = text.split('{' + k + '}').join(String(vars[k]));
  return text;
}

// 状态条的文字要等字典到位才能渲染。字典到达前先把「该显示哪个状态」记下来，
// 否则会短暂闪出一个裸的 key（'connected' / 'initializing'）。
let pendingStatus = null;
function setStatus(cls, text) {
  statusEl.className = 'status ' + cls;
  statusTextEl.textContent = text;
}
function setStatusKey(cls, key) {
  pendingStatus = { cls, key };
  if (T) setStatus(cls, t(key));
}
setStatusKey('warn', 'initializing');

/** 把字典填进 DOM。HTML 里刻意不写任何可见文案，避免出现第二份会漂移的副本。 */
function applyI18n() {
  const set = (id, key) => { const el = document.getElementById(id); if (el) el.textContent = t(key); };
  set('ui-title', 'title');
  set('ui-subtitle', 'subtitle');
  set('ui-host-label', 'host');
  set('ui-doc-label', 'document');
  set('ui-mode-label', 'modeLabel');
  set('ui-mode-ask', 'modeAsk');
  set('ui-mode-auto', 'modeAuto');
  set('ui-mode-hint', 'modeHint');
  set('ui-log-label', 'logTitle');
  set('ui-ap-title', 'approvalTitle');
  set('ui-ap-current', 'approvalCurrent');
  set('ui-ap-after', 'approvalAfter');
  set('approval-confirm', 'approvalConfirm');
  set('approval-reject', 'approvalReject');
  document.title = t('title');
  document.documentElement.lang = LANG === 'en' ? 'en' : 'zh-CN';
  logEl.setAttribute('data-empty', t('logEmpty'));
  setStatus(pendingStatus.cls, t(pendingStatus.key));
}

/** 取字典。窗格就是桥接服务发出来的，所以正常情况下这一次请求必定成功；失败则重试。 */
async function loadI18n() {
  LANG = normalizeLang(
    (typeof Office !== 'undefined' && Office.context && Office.context.displayLanguage)
    || (typeof navigator !== 'undefined' && navigator.language)
  ) || 'zh';
  const r = await api('/office/i18n?lang=' + encodeURIComponent(LANG));
  if (!r || !r.pane) throw new Error('i18n payload missing');
  T = r.pane;
  LANG = r.lang || LANG;
  applyI18n();
}

// ================= 公共辅助（暴露给 actions.js 使用） =================

function log(line) {
  const time = new Date().toLocaleTimeString(LANG === 'en' ? 'en-GB' : 'zh-CN', { hour12: false });
  logEl.textContent = `[${time}] ${line}\n` + logEl.textContent;
  if (logEl.textContent.length > 4000) logEl.textContent = logEl.textContent.slice(0, 4000);
}

async function api(path, opts) {
  const r = await fetch(BRIDGE + path, opts);
  if (!r.ok) {
    const error = new Error(`HTTP ${r.status}`);
    error.status = r.status;
    throw error;
  }
  return r.json();
}

function hostName() {
  return String(Office.context.host || '');
}

// 实例标识（多文档精确路由）：每个窗格唯一；docUrl 识别文档（测试文件 vs 正式论文）
// 注意：必须在 Office.onReady 之后才能读到 document.url/title（脚本加载时 context 未就绪、恒为空）
//       → 心跳时实时读取，不用启动时快照（另存为/路径变化也能跟上）
const INSTANCE_ID = 'inst-' + Math.random().toString(36).slice(2, 10) + '-' + Date.now().toString(36);
function docUrlNow() { try { return String(Office.context.document.url || ''); } catch (e) { return ''; } }
function docTitleNow() { try { return String(Office.context.document.title || ''); } catch (e) { return ''; } }

/** 把「当前文档」显示出来：标题优先，退回 URL 的文件名。用户据此确认指令打在哪个文档上。 */
function syncDoc() {
  const url = docUrlNow();
  const title = docTitleNow();
  const fromUrl = url ? (url.split(/[\\/]/).pop() || url) : '';
  docValueEl.textContent = title || fromUrl || t('documentUntitled');
  docValueEl.title = url || title || '';
}

function okResult(result) { return { ok: true, result }; }
function errResult(code, error) { return { ok: false, code, error }; }

function getSelectionTextAsync() {
  return new Promise((resolve) => {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
      (r) => resolve(r.status === Office.AsyncResultStatus.Succeeded
        ? okResult({ text: r.value || '' })
        : errResult('execution', r.error && r.error.message)));
  });
}

function setSelectionTextAsync(text) {
  return new Promise((resolve) => {
    Office.context.document.setSelectedDataAsync(String(text), { coercionType: Office.CoercionType.Text },
      (r) => resolve(r.status === Office.AsyncResultStatus.Succeeded
        ? okResult({ replaced: String(text).length })
        : errResult('execution', r.error && r.error.message)));
  });
}

// Word 内置样式名映射（小写去空格）
function mapBuiltInStyle(name) {
  const n = String(name || '').toLowerCase().replace(/[\s-]/g, '');
  const map = {
    '标题': 'title', title: 'title', '副标题': 'subtitle', subtitle: 'subtitle',
    '标题1': 'heading1', heading1: 'heading1', '标题2': 'heading2', heading2: 'heading2',
    '标题3': 'heading3', heading3: 'heading3', '正文': 'normal', normal: 'normal',
    '引用': 'quote', quote: 'quote', '强调': 'strong', strong: 'strong',
  };
  const heading = /^(?:标题|heading)([1-9])$/.exec(n);
  if (heading) return 'heading' + heading[1];
  if (n === 'caption' || n === '题注') return 'caption';
  return map[n] || null;
}

function mapAlignment(a) {
  const m = { left: Word.Alignment.left, center: Word.Alignment.centered, right: Word.Alignment.right, justify: Word.Alignment.justified };
  return m[String(a || '').toLowerCase()] || null;
}

// Excel：把 "Sheet1!A1:B10" 或 "A1:B10" 地址解析为 Range 对象（Workbook 无 getRange，需经 Worksheet）
function excelRange(ctx, address) {
  const addr = String(address || '');
  const bang = addr.lastIndexOf('!');
  if (bang >= 0) {
    const sheetName = addr.slice(0, bang);
    const rangePart = addr.slice(bang + 1);
    return ctx.workbook.worksheets.getItem(sheetName).getRange(rangePart);
  }
  return ctx.workbook.worksheets.getActiveWorksheet().getRange(addr);
}

// PowerPoint：getFileAsync 全文件拉取（分片 → 回调提取器）
// 注意：getSliceAsync 的第一个参数是「零基分片序号」，不是字节偏移量。
// 早期版本传累计字节数（received）当序号，文件 ≤ 1 片时凑巧能跑通，多片时第 2 片就取不到。
function readPptFile(extractor) {
  return new Promise((resolve) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, (fileResult) => {
      if (fileResult.status !== Office.AsyncResultStatus.Succeeded) {
        return resolve(errResult('execution', fileResult.error && fileResult.error.message));
      }
      const file = fileResult.value;
      const chunks = [];
      let sliceIndex = 0;
      let received = 0;
      const done = (out) => { try { file.closeAsync(() => resolve(out)); } catch (e) { resolve(out); } };
      const onSlice = (sliceResult) => {
        if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
          return done(errResult('execution', sliceResult.error && sliceResult.error.message));
        }
        chunks.push(new Uint8Array(sliceResult.value.data));
        sliceIndex += 1;
        received += sliceResult.value.size;
        // 两个条件都要留：sliceCount 是权威的片数；size 兜底（防止某些宿主 sliceCount 不准）
        if (sliceIndex >= file.sliceCount || received >= file.size) {
          try { done(extractor(chunks)); }
          catch (e) { done(errResult('execution', String(e && e.message || e))); }
        } else {
          file.getSliceAsync(sliceIndex, onSlice);
        }
      };
      file.getSliceAsync(0, onSlice);
    });
  });
}

// 极简 zip 容器解析（local file header 遍历，支持 stored(0) / deflate(8)）
function unzip(bytes) {
  const entries = {};
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const decoder = new TextDecoder('utf-8');
  let i = 0;
  while (i + 30 <= bytes.length) {
    if (view.getUint32(i, true) !== 0x04034b50) break;
    const method = view.getUint16(i + 8, true);
    const compSize = view.getUint32(i + 18, true);
    const nameLen = view.getUint16(i + 26, true);
    const extraLen = view.getUint16(i + 28, true);
    const name = decoder.decode(bytes.subarray(i + 30, i + 30 + nameLen));
    const dataStart = i + 30 + nameLen + extraLen;
    entries[name] = { method, data: bytes.subarray(dataStart, dataStart + compSize) };
    i = dataStart + compSize;
  }
  return entries;
}

function inflateEntry(entry) {
  if (entry.method === 0) return entry.data;      // stored
  if (entry.method === 8) return pako.inflateRaw(entry.data); // deflate
  throw new Error('unsupported zip method: ' + entry.method);
}

window.__UTIL__ = {
  okResult, errResult, hostName,
  excelRange, getSelectionTextAsync, setSelectionTextAsync,
  mapBuiltInStyle, mapAlignment, readPptFile, unzip, inflateEntry,
};

// ================= actions.js 热更新 =================

let loadedActionsVersion = '';

function loadActionsScript() {
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = '/actions.js?ts=' + Date.now();
    s.onload = () => resolve();
    s.onerror = () => reject(new Error('actions.js load failed'));
    document.head.appendChild(s);
  });
}

// 用户定制 action 文件（user-actions.js，gitignore）：每次热更新一并重载，定制优先于内置
function loadUserActionsScript() {
  return new Promise((resolve) => {
    const s = document.createElement('script');
    s.src = '/user-actions.js?ts=' + Date.now();
    s.onload = () => resolve();
    s.onerror = () => resolve(); // 无定制文件时忽略（server 兜底返回空对象则不会 error）
    document.head.appendChild(s);
  });
}

async function ensureActions() {
  try {
    const r = await api('/office/actions-version');
    if (r && r.version) {
      if (r.version === loadedActionsVersion) return;
      await loadActionsScript();
      await loadUserActionsScript(); // 用户定制（可在 actions.js 之后覆盖/新增）
      // 验证新脚本确实执行：版本标记匹配 + __EXECUTE__ 已定义（防"加载成功但 IIFE 抛错/旧值残留"）
      if (window.__ACTIONS_VERSION__ === r.version && typeof window.__EXECUTE__ === 'function') {
        loadedActionsVersion = r.version;
        log(t('logActionsUpdated', { version: r.version.slice(0, 10) }));
      } else {
        log(t('logActionsFailed'));
      }
    }
  } catch (e) {
    log(t('logVersionCheckFailed', { message: (e && e.message) || e }));
  }
}

// ================= 分发器 =================

function execute(action, args) {
  // 转发给 actions.js 的 __EXECUTE__（含机制级 re-read confirm 拦截，可热更新）
  const fn = window.__EXECUTE__;
  if (typeof fn !== 'function') return errResult('unknown_action', 'actions not loaded yet');
  return Promise.resolve(fn(action, args));
}

// ================= 确认模式与窗格审批 =================

// 初值必须是 'ask'，与服务端默认（CONFIRM_MODE 默认 ask）一致。
// 本变量在 /office/config 返回之前就已参与判断：若初值为 'auto'，从窗格加载到配置同步之间存在
// 「服务端要求审批、窗格却按自动模式直行」的时间窗，覆盖类操作会被无审批执行。
let confirmMode = 'ask';
window.__CONFIRM_MODE__ = 'ask';
let approvalActive = false;
let modeUpdatePending = false;
let pendingConfirmMode = null;
let approvalPending = null; // 当前审批对象，独立持有 timer / resolve
let createApprovedExecution = null;
window.__REGISTER_APPROVAL_EXECUTOR__ = (factory) => { createApprovedExecution = factory; };

function actionMeta(action) {
  return (window.__USER_ACTIONS__ || {})[action] || (window.__ACTIONS__ || {})[action] || {};
}

async function startCommand(cmd) {
  const expiry = typeof cmd.expiresAt === 'number' ? cmd.expiresAt : Date.parse(cmd.expiresAt);
  if (!Number.isFinite(expiry) || expiry <= Date.now()) {
    const error = new Error('command expired');
    error.status = 409;
    throw error;
  }
  const started = await api('/office/commands/' + encodeURIComponent(cmd.commandId) + '/start', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ instance: INSTANCE_ID }),
  });
  if (!started || !started.ok || started.commandId !== cmd.commandId || started.status !== 'executing') {
    throw new Error('command cannot start');
  }
}

function postResult(cmd, out) {
  return api('/office/result', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ commandId: cmd.commandId, instance: INSTANCE_ID, ok: out.ok,
      result: out.result, error: out.error, code: out.code }),
  }).catch((e) => {
    if (e.status === 404 || e.status === 409) return { ok: false, code: 'rejected', error: e.message };
    throw e;
  });
}

function syncModeUI() {
  if (!modeSeg) return;
  for (const b of modeSeg.querySelectorAll('button')) {
    b.setAttribute('aria-checked', String(b.dataset.mode === confirmMode));
    b.disabled = busy || approvalActive || modeUpdatePending || reloadingForCodeVersion;
  }
}

function applyConfirmMode(mode) {
  if ((mode !== 'ask' && mode !== 'auto') || modeUpdatePending) return false;
  if (busy || approvalActive) { pendingConfirmMode = mode; return false; }
  pendingConfirmMode = null;
  confirmMode = mode;
  window.__CONFIRM_MODE__ = mode;
  syncModeUI();
  return true;
}

async function setConfirmMode(mode) {
  if ((mode !== 'ask' && mode !== 'auto') || busy || approvalActive || modeUpdatePending || reloadingForCodeVersion) return false;
  pendingConfirmMode = null;
  modeUpdatePending = true;
  syncModeUI();
  try {
    const saved = await api('/office/config', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ confirmMode: mode }),
    });
    if (!saved || saved.ok !== true || (saved.confirmMode !== 'ask' && saved.confirmMode !== 'auto')) throw new Error('Invalid confirmation mode response');
    modeUpdatePending = false;
    if (!applyConfirmMode(saved.confirmMode)) return false;
    log(t('logMode', { mode: saved.confirmMode === 'auto' ? t('modeAuto') : t('modeAsk') }));
    return true;
  } catch (e) {
    // A failed save leaves the current command policy and UI selection unchanged.
    log(t('logFailed', { action: t('modeLabel'), error: String(e && e.message || e) }));
    return false;
  } finally {
    modeUpdatePending = false;
    syncModeUI();
    flushIdleState();
  }
}

// 按 action 提取"修改后内容"（审批面板展示用）由 actions 分片提供，暴露为 window.__AFTER_CONTENT__：
// 它和 action 语义强绑定，放窗格里会导致"新增 action 还得改窗格"（改窗格要重载）。
// 窗格只做兜底：actions 未加载时退化为原样打印 args。

function renderApproval() {
  const box = document.getElementById('approval');
  if (!box) return;
  if (!approvalPending) { box.hidden = true; return; }
  const confirmButton = document.getElementById('approval-confirm');
  if (confirmButton) confirmButton.disabled = !!approvalPending.reading || !approvalPending.previewOk;
  box.hidden = false;
  document.getElementById('ap-action').textContent =
    `${approvalPending.cmd.action} @ ${approvalPending.cmd.host || hostName()}`;
  // 当前状态预览：有 text 直接显示文本；replace_all 的 hits 列表友好显示（将替换 N 处 + 命中内容）
  let previewText;
  if (approvalPending.reading) {
    previewText = t('approvalReading');
  } else {
    previewText = t('approvalUnavailable');
    const p = approvalPending.preview;
    if (p) {
      if (typeof p.text === 'string') previewText = p.text;
      else if (Array.isArray(p.hits) && p.hits.length) previewText = t('approvalWillReplace', { count: p.wouldReplace }) + '\n' + p.hits.map((h, i) => `${i + 1}. ${h}`).join('\n');
      else if (p.wouldReplace !== undefined) previewText = t('approvalWillReplace', { count: p.wouldReplace });
      else if (p.text !== undefined) previewText = String(p.text);
      else previewText = JSON.stringify(p, null, 2);
    }
  }
  document.getElementById('ap-preview').textContent = previewText.slice(0, 1500);
  const after = document.getElementById('ap-after');
  if (after) {
    // "修改后内容"由 actions.js 提供（新增/修改 action 不必改窗格）
    const afterFn = window.__AFTER_CONTENT__;
    after.textContent = typeof afterFn === 'function'
      ? afterFn(approvalPending.cmd.action, approvalPending.cmd.args || {})
      : JSON.stringify(approvalPending.cmd.args || {});
  }
}

// 审批中定位：让用户看到将操作的位置（选中目标，而非移动内容）
// 统一走 actions.js 的 selectTarget 机制（可热更新）：
//   replace_all → 选中第一处命中；write_range/format_range → 选中目标区域；
//   insert_paragraph/append_text → 选中插入点（文末最后一段；afterSelection 保持选区不动）。
// 无 selectTarget 的动作（如 write_selection 针对当前选区）不动光标。
// ask 模式由窗格持有一次性批准执行入口；调用方提供的 confirm 参数不授予权限。
async function locateInsertionPoint(cmd) {
  const meta = actionMeta(cmd.action);
  if (meta.selectTarget) {
    try { await meta.selectTarget(hostName(), cmd.args || {}); } catch (e) { /* 定位失败不影响审批 */ }
  }
}

async function showApproval(cmd) {
  approvalActive = true;
  const meta = actionMeta(cmd.action);
  const pending = { cmd, preview: null, meta, reading: true, previewOk: false, timer: null };
  approvalPending = pending;
  const expiry = typeof cmd.expiresAt === 'number' ? cmd.expiresAt : Date.parse(cmd.expiresAt);
  const remaining = Number.isFinite(expiry) ? Math.max(0, expiry - Date.now()) : 0;
  const decisionP = new Promise((resolve) => {
    let settled = false;
    pending.resolve = (decision) => {
      if (settled || (decision === 'confirm' && !pending.previewOk)) return;
      settled = true;
      clearTimeout(pending.timer);
      resolve(decision);
    };
    pending.timer = setTimeout(() => pending.resolve('timeout'), Math.min(120000, remaining));
  });
  renderApproval();
  let out;
  let reportResult = true;
  try {
    const previewP = Promise.resolve().then(async () => {
      await locateInsertionPoint(cmd);
      if (approvalPending !== pending) return errResult('rejected', 'approval closed');
      if (typeof meta.preview !== 'function') return errResult('execution', 'approval preview unavailable');
      return meta.preview(hostName(), cmd.args || {});
    }).catch((e) => errResult('execution', String(e && e.message || e)));
    const first = await Promise.race([
      previewP.then((preview) => ({ preview })),
      decisionP.then((decision) => ({ decision })),
    ]);
    if (first.decision) {
      out = errResult('rejected', first.decision === 'reject' ? t('rejectMessage') : t('timeoutMessage'));
    } else if (!first.preview || !first.preview.ok) {
      out = errResult(first.preview && first.preview.code || 'execution', first.preview && first.preview.error || 'approval preview failed');
    } else {
      pending.preview = first.preview.result;
      pending.previewOk = true;
      pending.reading = false;
      renderApproval();
      const decision = await decisionP;
      if (decision !== 'confirm') {
        out = errResult('rejected', decision === 'reject' ? t('rejectMessage') : t('timeoutMessage'));
      } else {
        if (!createApprovedExecution) throw new Error('approval executor unavailable');
        const runApproved = createApprovedExecution(cmd);
        try { await startCommand(cmd); }
        catch (e) { reportResult = false; throw e; }
        out = await runApproved();
        if (out.ok && out.result && typeof out.result === 'object') out.result.approvedBy = 'user';
      }
    }
  } catch (e) {
    out = errResult('rejected', String(e && e.message || e));
  } finally {
    clearTimeout(pending.timer);
    if (approvalPending === pending) approvalPending = null;
    approvalActive = false;
    renderApproval();
  }
  return reportResult ? postResult(cmd, out) : out;
}

// ================= 后台循环 =================

// 连接状态恢复：poll/heartbeat 成功时把状态改回"已连接"（断联重连后不再停留"未连接"）
function setConnected() {
  if (statusEl.className !== 'status ok') setStatusKey('ok', 'connected');
}

let busy = false;
async function poll() {
  if (busy || approvalActive || modeUpdatePending || reloadingForCodeVersion) return; // 审批中暂停轮询，避免重复取同一指令
  busy = true;
  syncModeUI();
  try {
    await ensureActions();
    const host = hostName();
    const cmd = await api('/office/poll?host=' + encodeURIComponent(host) + '&instance=' + encodeURIComponent(INSTANCE_ID));
    setConnected(); // 轮询成功 = 桥接已连上，恢复状态
    if (cmd && cmd.commandId) {
      const meta = actionMeta(cmd.action);
      cmd.args = Object.assign({}, cmd.args || {});
      delete cmd.args.confirm;
      if (meta.destructive && confirmMode === 'ask') {
        log(t('logPendingApproval', { action: cmd.action }));
        await showApproval(cmd); // ask 模式：窗格挂起等用户审批
      } else {
        log(t('logExecuting', { action: cmd.action, args: JSON.stringify(cmd.args || {}).slice(0, 80) }));
        await startCommand(cmd);
        const out = await execute(cmd.action, cmd.args || {});
        await postResult(cmd, out);
        log(out.ok ? t('logDone', { action: cmd.action }) : t('logFailed', { action: cmd.action, error: out.error }));
      }
    }
  } catch (e) {
    if (e.status !== 404 && e.status !== 409) setStatusKey('err', 'disconnected');
  } finally {
    busy = false;
    syncModeUI();
    flushIdleState();
  }
}

// 窗格代码版本：服务端把 taskpane.js（+ i18n 字典）的 mtime 作为 __CODE_VERSION__ 注入本页（见 server.js 静态文件分支），
// 并在每次心跳响应里回传当前值。两者不一致 = 磁盘上的窗格已改、本页跑的是旧代码 → 自动 reload。
// 这样改窗格（taskpane.js）或改文案（i18n/*.json）后重启桥接服务即可生效，不必手动关掉再打开窗格。
let reloadingForCodeVersion = false;
let pendingCodeVersion = null;

function flushIdleState() {
  if (busy || approvalActive || modeUpdatePending) return;
  if (pendingConfirmMode !== null) applyConfirmMode(pendingConfirmMode);
  if (!pendingCodeVersion || !window.__CODE_VERSION__ || pendingCodeVersion === window.__CODE_VERSION__ || reloadingForCodeVersion) return;
  reloadingForCodeVersion = true;
  syncModeUI();
  log(t('logReloading'));
  setTimeout(() => {
    if (busy || approvalActive || modeUpdatePending) { reloadingForCodeVersion = false; syncModeUI(); return; }
    try { location.reload(); } catch (e) { reloadingForCodeVersion = false; syncModeUI(); }
  }, 300);
}

function heartbeat() {
  return api('/office/heartbeat', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ host: hostName(), instanceId: INSTANCE_ID, docUrl: docUrlNow(), docTitle: docTitleNow() }),
  }).then((r) => {
    setConnected();
    syncDoc();
    applyConfirmMode(r && r.confirmMode);
    // 重载前必须避开「审批面板正挂着」和「指令正在执行」：否则会丢掉用户的审批选择或正在跑的结果。
    // 忙碌时保留版本与模式，poll finally 在结果回传后处理，避免固定周期同相造成永久饥饿。
    if (r && typeof r.codeVersion === 'string' && r.codeVersion) pendingCodeVersion = r.codeVersion;
    flushIdleState();
  }).catch(() => { /* 忽略，下次再试 */ });
}

Office.onReady((info) => {
  const hostNameStr = Office.HostType[info.host] || String(info.host);
  hostValueEl.textContent = `${hostNameStr} · ${Office.version}`;
  hostValueEl.title = hostNameStr;

  // 先取文案再启动循环：取不到就每秒重试，但轮询照常跑（窗格的功能不该被一次文案请求卡住）。
  // 日志行放在取到字典之后写，否则会先把裸 key 打进日志。
  (async function bootstrapI18n() {
    for (let attempt = 0; attempt < 30; attempt++) {
      try {
        await loadI18n();
        log(t('logReady', { host: hostNameStr }));
        return;
      } catch (e) { /* 稍后重试 */ }
      await new Promise((r) => setTimeout(r, 1000));
    }
  })();

  setStatusKey('ok', 'connected');
  api('/office/hello', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    // 上报本页实际跑的代码版本（= 服务端注入的 taskpane.js mtime）。
    // 每次 reload 后该值会更新，GET /office/status 的 hellos 可用来验证自动重载确实发生。
    body: JSON.stringify({ host: hostName(), version: String(window.__CODE_VERSION__ || 'unknown') }),
  }).catch(() => {});

  if (modeSeg) {
    modeSeg.addEventListener('click', (e) => {
      const btn = e.target.closest('button[data-mode]');
      if (btn) setConfirmMode(btn.dataset.mode);
    });
  }
  api('/office/config').then((c) => {
    applyConfirmMode(c && c.confirmMode);
  }).catch(() => {});
  syncModeUI();

  const btnConfirm = document.getElementById('approval-confirm');
  const btnReject = document.getElementById('approval-reject');
  if (btnConfirm) btnConfirm.addEventListener('click', () => { if (approvalPending && approvalPending.resolve) approvalPending.resolve('confirm'); });
  if (btnReject) btnReject.addEventListener('click', () => { if (approvalPending && approvalPending.resolve) approvalPending.resolve('reject'); });

  syncDoc();
  setInterval(poll, POLL_MS);
  setInterval(heartbeat, HEARTBEAT_MS);
  heartbeat();
});
