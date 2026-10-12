// taskpane.js — DSH Office 执行器窗格（actions 热更新架构）
// 职责：窗格 UI、桥接轮询/心跳、execute 分发、actions.js 热更新加载。
// 所有 action 实现都在 /actions.js（服务端把 server/actions/ 下的分片拼成这一个脚本，可热更新）：
// 改任一分片 → 重启桥接服务 → 窗格自动重新加载，无需重开窗格。
// 本文件（窗格）应尽量少改：改本文件后，服务端心跳回传的 codeVersion 会与本页注入的
// window.__CODE_VERSION__ 不一致，窗格在下次心跳时自动 reload（见 heartbeat()）。

const BRIDGE =
  typeof location !== "undefined" &&
  location.origin &&
  /^https?:\/\//i.test(location.origin)
    ? location.origin
    : "http://127.0.0.1:3000";
const POLL_MS = 1000;
const HEARTBEAT_MS = 5000;

const statusEl = document.getElementById('status');
const logEl = document.getElementById('log');
const hostEl = document.getElementById('host');

function setStatus(cls, text) {
  statusEl.className = 'status ' + cls;
  statusEl.textContent = text;
}
function log(line) {
  const t = new Date().toLocaleTimeString('zh-CN', { hour12: false });
  logEl.textContent = `[${t}] ${line}\n` + logEl.textContent;
  if (logEl.textContent.length > 4000) logEl.textContent = logEl.textContent.slice(0, 4000);
}

async function api(path, opts) {
  const r = await fetch(BRIDGE + path, opts);
  if (!r.ok) throw new Error(`HTTP ${r.status}`);
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

function okResult(result) { return { ok: true, result }; }
function errResult(code, error) { return { ok: false, code, error }; }

// ================= 公共辅助（暴露给 actions.js 使用） =================

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
  return map[n] || null;
}

function mapAlignment(a) {
  const m = { left: Word.Alignment.left, center: Word.Alignment.center, right: Word.Alignment.right, justify: Word.Alignment.justified };
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
        log('actions 已热更新: ' + r.version.slice(0, 10));
      } else {
        log('⚠️ actions 热更新失败（版本不匹配或未执行），将重试');
      }
    }
  } catch (e) {
    log('⚠️ actions 版本检查失败: ' + (e && e.message || e));
  }
}

// ================= 分发器 =================

function execute(action, args) {
  // 转发给 actions.js 的 __EXECUTE__（含机制级 re-read confirm 拦截，可热更新）
  const fn = window.__EXECUTE__;
  if (typeof fn !== 'function') return errResult('unknown_action', 'actions 尚未加载');
  return Promise.resolve(fn(action, args));
}

// ================= 后台循环 =================

// 连接状态恢复：poll/heartbeat 成功时把状态改回"已连接"（断联重连后不再停留"未连接"）
function setConnected() {
  if (statusEl.className !== 'status ok') setStatus('ok', '已连接：等待 DSH 指令');
}

// ---------- 确认模式与窗格审批 ----------
// 初值必须是 'ask'，与服务端默认（CONFIRM_MODE 默认 ask）一致。
// 本变量在 /office/config 返回之前就已参与判断：若初值为 'auto'，从窗格加载到配置同步之间存在
// 「服务端要求审批、窗格却按自动模式直行」的时间窗，覆盖类操作会被无审批执行。
let confirmMode = 'ask';
window.__CONFIRM_MODE__ = 'ask';
let approvalActive = false;
let approvalPending = null; // { cmd, preview, meta, resolve }

function setConfirmMode(mode) {
  confirmMode = mode;
  window.__CONFIRM_MODE__ = mode;
  const sel = document.getElementById('mode-select');
  if (sel) sel.value = mode;
  api('/office/config', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ confirmMode: mode }),
  }).catch(() => { /* 忽略 */ });
  log(`确认模式: ${mode}`);
}

// 按 action 提取"修改后内容"（审批面板展示用）由 actions 分片提供，暴露为 window.__AFTER_CONTENT__：
// 它和 action 语义强绑定，放窗格里会导致"新增 action 还得改窗格"（改窗格要重载）。
// 窗格只做兜底：actions 未加载时退化为原样打印 args。

function renderApproval() {
  const box = document.getElementById('approval');
  if (!box) return;
  if (!approvalPending) { box.style.display = 'none'; return; }
  box.style.display = 'block';
  document.getElementById('approval-action').textContent =
    `${approvalPending.cmd.action} @ ${approvalPending.cmd.host || hostName()}`;
  // 当前状态预览：有 text 直接显示文本；replace_all 的 hits 列表友好显示（将替换 N 处 + 命中内容）
  let previewText = '(预览不可用)';
  const p = approvalPending.preview;
  if (p) {
    if (typeof p.text === 'string') previewText = p.text;
    else if (Array.isArray(p.hits) && p.hits.length) previewText = `将替换 ${p.wouldReplace} 处：\n` + p.hits.map((h, i) => `${i + 1}. ${h}`).join('\n');
    else if (p.wouldReplace !== undefined) previewText = `将替换 ${p.wouldReplace} 处`;
    else if (p.text !== undefined) previewText = String(p.text);
    else previewText = JSON.stringify(p, null, 2);
  }
  document.getElementById('approval-preview').textContent = previewText.slice(0, 1500);
  const after = document.getElementById('approval-after');
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
// ask 模式由窗格保持审批语义：调用方不应预先传 confirm:true；仅用户点击批准后，窗格才为该次执行附加它。
async function locateInsertionPoint(cmd) {
  const meta = (window.__ACTIONS__ || {})[cmd.action] || {};
  if (meta.selectTarget) {
    try { await meta.selectTarget(hostName(), cmd.args || {}); } catch (e) { /* 定位失败不影响审批 */ }
  }
}

async function showApproval(cmd) {
  approvalActive = true;
  const meta = (window.__ACTIONS__ || {})[cmd.action] || {};
  const previewP = (meta.preview
    ? Promise.resolve(meta.preview(hostName(), cmd.args || {}))
    : Promise.resolve(okResult(null))
  ).catch((e) => errResult('execution', String(e && e.message || e)));
  const p = await previewP;
  approvalPending = { cmd, preview: p.ok ? p.result : null, meta };
  renderApproval();
  // 预选闪烁必须放在审批面板弹出之后：闪烁在面板之前发生的话，用户看到面板时闪烁早已结束（"没闪"）
  // 面板已显示 → 再闪烁定位目标（不阻塞审批等待；用户点击确认时闪烁可能未结束，无副作用）
  locateInsertionPoint(cmd).catch(() => {});
  // 等待用户在窗格点击（确认/拒绝）；审批挂起 120s 未操作 → 自动解除（视为拒绝），防 poll 永久卡死
  const decision = await new Promise((resolve) => {
    approvalPending.resolve = resolve;
    setTimeout(() => { if (approvalPending && approvalPending.resolve) approvalPending.resolve('timeout'); }, 120000);
  });
  const { cmd: c } = approvalPending;
  approvalPending = null;
  approvalActive = false;
  renderApproval();
  if (decision === 'reject' || decision === 'timeout') {
    return api('/office/result', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ commandId: c.commandId, ok: false, error: decision === 'reject' ? '用户在窗格拒绝了该操作' : '审批超时未操作，已取消', code: 'rejected' }),
    });
  }
  const out = await execute(c.action, Object.assign({}, c.args || {}, { confirm: true }));
  // 标记：该操作经用户审批后执行（区别于自动模式直行）
  if (out.ok && out.result && typeof out.result === 'object') out.result.approvedBy = 'user';
  return api('/office/result', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ commandId: c.commandId, ok: out.ok, result: out.result, error: out.error, code: out.code }),
  });
}

let busy = false;
async function poll() {
  if (busy || approvalActive) return; // 审批中暂停轮询，避免重复取同一指令
  busy = true;
  try {
    await ensureActions();
    const host = hostName();
    const cmd = await api('/office/poll?host=' + encodeURIComponent(host) + '&instance=' + encodeURIComponent(INSTANCE_ID));
    setConnected(); // 轮询成功 = 桥接已连上，恢复状态
    if (cmd && cmd.commandId) {
      const meta = (window.__ACTIONS__ || {})[cmd.action] || {};
      if (meta.destructive && confirmMode === 'ask' && !(cmd.args || {}).confirm) {
        log(`待审批: ${cmd.action}`);
        await showApproval(cmd); // ask 模式：窗格挂起等用户审批
      } else {
        log(`执行: ${cmd.action} ${JSON.stringify(cmd.args || {}).slice(0, 80)}`);
        const out = await execute(cmd.action, cmd.args || {});
        await api('/office/result', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ commandId: cmd.commandId, ok: out.ok, result: out.result, error: out.error, code: out.code }),
        });
        log(out.ok ? `完成: ${cmd.action}` : `失败: ${cmd.action} -> ${out.error}`);
      }
    }
  } catch (e) {
    setStatus('err', '桥接服务未连接，等待重试…');
  } finally {
    busy = false;
  }
}

// 窗格代码版本：服务端把 taskpane.js 的 mtime 作为 __CODE_VERSION__ 注入本页（见 server.js 静态文件分支），
// 并在每次心跳响应里回传当前值。两者不一致 = 磁盘上的窗格已改、本页跑的是旧代码 → 自动 reload。
// 这样改窗格（taskpane.js）后重启桥接服务即可生效，不必手动关掉再打开窗格。
let reloadingForCodeVersion = false;
function heartbeat() {
  api('/office/heartbeat', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ host: hostName(), instanceId: INSTANCE_ID, docUrl: docUrlNow(), docTitle: docTitleNow() }),
  }).then((r) => {
    setConnected();
    // 重载前必须避开「审批面板正挂着」和「指令正在执行」：否则会丢掉用户的审批选择或正在跑的结果。
    // 错过这一拍不要紧，下一次心跳（5s 后）会再判一次。
    if (!r || !r.codeVersion || reloadingForCodeVersion) return;
    if (!window.__CODE_VERSION__ || r.codeVersion === window.__CODE_VERSION__) return;
    if (approvalActive || busy) return;
    reloadingForCodeVersion = true;
    log('检测到窗格代码已更新，自动重载…');
    setTimeout(() => { try { location.reload(); } catch (e) { reloadingForCodeVersion = false; } }, 300);
  }).catch(() => { /* 忽略，下次再试 */ });
}

Office.onReady((info) => {
  const hostNameStr = Office.HostType[info.host] || String(info.host);
  hostEl.textContent = `host: ${hostNameStr} (${Office.version})`;
  setStatus('ok', '已连接：等待 DSH 指令');
  log(`Office 就绪 (${hostNameStr})`);
  api('/office/hello', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    // 上报本页实际跑的代码版本（= 服务端注入的 taskpane.js mtime）。
    // 每次 reload 后该值会更新，GET /office/status 的 hellos 可用来验证自动重载确实发生。
    body: JSON.stringify({ host: hostName(), version: String(window.__CODE_VERSION__ || 'unknown') }),
  }).catch(() => {});
  // 模式开关 + 审批按钮
  const modeSel = document.getElementById('mode-select');
  if (modeSel) {
    modeSel.addEventListener('change', (e) => setConfirmMode(e.target.value));
    api('/office/config').then((c) => {
      if (c && (c.confirmMode === 'auto' || c.confirmMode === 'ask')) {
        confirmMode = c.confirmMode;
        window.__CONFIRM_MODE__ = c.confirmMode;
        modeSel.value = c.confirmMode;
      }
    }).catch(() => {});
  }
  const btnConfirm = document.getElementById('approval-confirm');
  const btnReject = document.getElementById('approval-reject');
  if (btnConfirm) btnConfirm.addEventListener('click', () => { if (approvalPending && approvalPending.resolve) approvalPending.resolve('confirm'); });
  if (btnReject) btnReject.addEventListener('click', () => { if (approvalPending && approvalPending.resolve) approvalPending.resolve('reject'); });
  setInterval(poll, POLL_MS);
  setInterval(heartbeat, HEARTBEAT_MS);
  heartbeat();
});
