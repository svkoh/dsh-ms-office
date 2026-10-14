// registry.js — 装配分片（必须最后加载）
//
// 职责只有三件：把「按宿主分片」的实现拍平成统一注册表、跟 capabilities.json 的元数据对账、
// 暴露窗格要用的分发器。这里**不放任何宿主专属逻辑** —— 那样又会把三个宿主混回来。
(function () {
  const A = window.__OFFICE_ACTIONS__;
  const U = A.util;
  const errResult = U.errResult;
  const HOSTS = A.hosts;
  const COMMON = HOSTS.Common || { actions: {} };

  // 某宿主上某 action 的实现：先看该宿主自己的分片，再看宿主无关的 common 分片
  function implFor(host, name) {
    const t = HOSTS[host];
    return (t && t.actions && t.actions[name]) || (COMMON.actions && COMMON.actions[name]) || null;
  }
  function previewFor(host, name) {
    const t = HOSTS[host];
    return (t && t.previewReaders && t.previewReaders[name]) || (COMMON.previewReaders && COMMON.previewReaders[name]) || null;
  }
  function selectTargetFor(host, name) {
    const t = HOSTS[host];
    return (t && t.selectTargets && t.selectTargets[name]) || null;
  }

  // ---- 1. 拍平：action 名 → { impl(host, args) } ----
  const ACTIONS_TABLE = {};
  for (const t of Object.values(HOSTS)) {
    for (const name of Object.keys(t.actions || {})) {
      if (ACTIONS_TABLE[name]) continue; // 同名 = 多宿主 action（如 read_selection），已登记
      ACTIONS_TABLE[name] = {
        impl: (host, args) => {
          const fn = implFor(host, name);
          if (!fn) return errResult('unsupported_host', `${name} not supported in ${host}`);
          return fn(args || {});
        },
      };
    }
  }

  // ---- 2. 与 capabilities.json 对账（唯一真源，服务端注入为 window.__ACTIONS_META__）----
  // 三个方向都要查，任何一条不成立都拒绝加载 —— 静默降级在这里是安全相关的失效：
  //   缺 hosts      → 平台校验失效（在不支持的宿主上尝试执行）
  //   缺 destructive → 覆盖类操作会绕过审批直接执行
  //   缺实现        → AI 能看到这个 action，调用却永远失败
  //   缺预览读取器  → destructive 的 confirm 分支永远走不通（等于功能不可用）
  const META = window.__ACTIONS_META__ || {};
  const problems = [];

  for (const [name, entry] of Object.entries(ACTIONS_TABLE)) {
    const m = META[name];
    if (!m || !Array.isArray(m.hosts) || typeof m.destructive !== 'boolean') {
      problems.push(`${name} 有实现但 capabilities.json 里没有它的 hosts/destructive`);
      continue;
    }
    entry.hosts = m.hosts;
    entry.destructive = m.destructive;
  }

  for (const [name, m] of Object.entries(META)) {
    if (!ACTIONS_TABLE[name]) { problems.push(`${name} 在 capabilities.json 里声明了，但没有宿主实现它`); continue; }
    for (const host of m.hosts) {
      if (!implFor(host, name)) problems.push(`${name} 声明支持 ${host}，但 ${host} 分片里没有实现`);
      if (m.destructive && !previewFor(host, name)) problems.push(`${name} 在 ${host} 上是覆盖类操作，但没有预览读取器`);
    }
  }

  if (problems.length) {
    throw new Error('actions 分片与 capabilities.json 不一致，拒绝加载：' + problems.join('；'));
  }

  // ---- 3. 注入 preview / selectTarget（只在真的存在时挂，语义与拆分前一致）----
  for (const [name, entry] of Object.entries(ACTIONS_TABLE)) {
    if (!entry.destructive) continue;
    entry.preview = (host, args) => {
      const fn = previewFor(host, name);
      if (!fn) return errResult('unsupported_host', `${name} has no preview reader for ${host}`);
      return fn(args || {});
    };
  }
  for (const [name, entry] of Object.entries(ACTIONS_TABLE)) {
    if (!Object.keys(HOSTS).some((h) => selectTargetFor(h, name))) continue;
    entry.selectTarget = (host, args) => {
      const fn = selectTargetFor(host, name);
      return fn ? fn(args || {}) : undefined; // 该宿主不做预选
    };
  }

  // ---- 4. 审批面板「修改后内容」：合并各分片提供的文案 ----
  const AFTER_CONTENT = {};
  for (const t of Object.values(HOSTS)) Object.assign(AFTER_CONTENT, t.afterContent || {});
  window.__AFTER_CONTENT__ = function (action, args) {
    const fn = AFTER_CONTENT[action];
    return fn ? fn(args || {}) : JSON.stringify(args || {});
  };

  window.__ACTIONS__ = ACTIONS_TABLE;

  // ---- 5. 分发器（窗格 execute 转发到这里，便于热更新拦截逻辑）----
  // 外部参数不构成批准。批准入口仅交给窗格闭包，授权绑定该次 commandId。
  function executeInternal(action, inputArgs, commandId) {
    const args = Object.assign({}, inputArgs || {});
    delete args.confirm;
    const approved = typeof commandId === 'string' && commandId.length > 0;
    // 用户定制优先（user-actions.js 可同名覆盖内置 action 或新增）。
    // 注意：user-actions.js 里的 action 不经过 capabilities.json，必须自带 hosts/destructive/preview。
    const userActions = window.__USER_ACTIONS__ || {};
    const meta = userActions[action] || ACTIONS_TABLE[action];
    if (!meta) {
      const normalize = name => String(name).replace(/^(word|excel|ppt)_/, '');
      const requested = normalize(action), tokens = new Set(requested.split('_'));
      const candidates = [...new Set([...Object.keys(ACTIONS_TABLE),...Object.keys(userActions)])]
        .map(name => ({name,score:normalize(name) === requested ? 100 : normalize(name).split('_').filter(token => tokens.has(token)).length}))
        .filter(candidate => candidate.score > 0).sort((a,b) => b.score - a.score || a.name.localeCompare(b.name)).slice(0,5).map(candidate => candidate.name);
      return errResult('unknown_action', `unknown action: ${action}; candidates: ${JSON.stringify(candidates)}; use GET /office/capabilities for exact names and arguments`);
    }
    const host = String(Office.context.host || '');
    if (!meta.hosts.includes(host)) return errResult('unsupported_host', `${action} not supported in ${host}`);
    // 只读文档保护（Common API mode 探测）：readOnly/view 模式拒绝覆盖类操作，明确告知原因而非执行失败
    if (meta.destructive) {
      try {
        const m = String(Office.context.document.mode || '');
        if (m === 'readOnly' || m === 'view') {
          return errResult('readonly', `文档处于只读/查看模式（mode=${m}），覆盖类操作被拒绝；请另存为可编辑副本或切换编辑模式后重试`);
        }
      } catch (e) { /* 探测失败不拦截 */ }
    }
    const run = () => Promise.resolve().then(() => meta.impl(host, args || {})).catch((e) => {
      const debug = (e && e.debugInfo) ? (' [debug] ' + JSON.stringify(e.debugInfo)) : '';
      return errResult('execution', String(e && e.message || e) + debug);
    });
    // 写后验证：覆盖操作执行成功后，再读一次被操作区域（afterState）+ 附整个文档全文（避免盲人摸象）
    const withAfterState = async (r) => {
      if (meta.destructive && r.ok && r.result && typeof r.result === 'object') {
        const warnings = [];
        if (meta.preview) {
          try {
            const a = await meta.preview(host, args);
            if (a && a.ok) r.result.afterState = a.result;
            else warnings.push({ stage: 'afterState', error: a && a.error || 'verification failed' });
          } catch (e) { warnings.push({ stage: 'afterState', error: String(e && e.message || e) }); }
        }
        if (host !== 'PowerPoint') {
          const readDoc = implFor(host, 'read_document');
          if (readDoc) {
            try {
              const d = await readDoc({});
              if (d && d.ok) r.result.document = d.result.text;
              else warnings.push({ stage: 'document', error: d && d.error || 'verification failed' });
            } catch (e) { warnings.push({ stage: 'document', error: String(e && e.message || e) }); }
          }
        }
        if (warnings.length) r.result.verificationWarnings = [...(r.result.verificationWarnings || []),...warnings];
      }
      return r;
    };
    // 覆盖类操作：先取确认模式（缺省 ask=审批，与服务端默认一致；拉取失败也兜底 ask，绝不悄悄回退 auto）
    const modeP = window.__CONFIRM_MODE__ !== undefined
      ? Promise.resolve(window.__CONFIRM_MODE__)
      : fetch('/office/config').then((r) => r.json()).then((c) => { window.__CONFIRM_MODE__ = (c && c.confirmMode) || 'ask'; return window.__CONFIRM_MODE__; }).catch(() => { window.__CONFIRM_MODE__ = 'ask'; return 'ask'; });
    if (meta.destructive && !approved) {
      if (!meta.preview) return errResult('confirm_required', `覆盖类操作 ${action} 需窗格预览确认`);
      return modeP.then((mode) => Promise.resolve().then(() => meta.preview(host, args)).then((p) => {
        if (!p.ok) return p;
        if (mode === 'auto') {
          // 自动模式：改前 re-read（previousState）→ 执行 → 改后验证（afterState）；标记 approvedBy=auto
          return run().then((r) => {
            if (r.ok && r.result && typeof r.result === 'object') {
              r.result.previousState = p.result;
              r.result.approvedBy = 'auto';
            }
            return withAfterState(r);
          });
        }
        // ask 模式：必须用户确认；先选中将要修改的目标（审批可视化，失败不阻塞）
        const preSelect = meta.selectTarget ? Promise.resolve().then(() => meta.selectTarget(host, args)).catch(() => {}) : Promise.resolve();
        return preSelect.then(() => ({ ok: false, code: 'confirm_required', error: '覆盖类操作需在 Office 窗格中批准', result: { preview: p.result, action, args, selectedTarget: true } }));
      })).catch((e) => errResult('execution', String(e && e.message || e)));
    }
    return run().then(withAfterState);
  }
  window.__EXECUTE__ = (action, args) => executeInternal(action, args, null);
  if (typeof window.__REGISTER_APPROVAL_EXECUTOR__ === 'function') {
    window.__REGISTER_APPROVAL_EXECUTOR__((cmd) => {
      // 每次注册的入口只签发一次性、绑定命令与参数快照的执行函数。
      const commandId = cmd.commandId;
      const action = cmd.action;
      const args = Object.assign({}, cmd.args || {});
      let used = false;
      return () => {
        if (used || typeof commandId !== 'string' || !commandId) return errResult('rejected', 'invalid or consumed approval');
        used = true;
        return executeInternal(action, args, commandId);
      };
    });
  }
})();
