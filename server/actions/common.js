// common.js — 宿主无关的 action（只用 Office.context.document / requirements，不碰 Word.run / Excel.run / PowerPoint.run）
//
// 这里的 action 对三个宿主都可用（hosts 由 capabilities.json 声明）。
// 判断标准很硬：**实现里出现 `Office.HostType` 之外任何宿主专属对象模型 = 它不属于这里**。
// 某个宿主需要自己的版本时，把实现放进那个宿主的分片（例如 write_selection 的预览读取器
// 就必须各自调本宿主的 read_selection，所以它挂在各宿主的 previewReaders 上）。
(function () {
  const A = window.__OFFICE_ACTIONS__;
  const U = A.util;
  const okResult = U.okResult, errResult = U.errResult, setSelectionTextAsync = U.setSelectionTextAsync;
  const isSetSupported = A.isSetSupported;

  // ================= 选区写入 =================
  // 三宿主都用 Common API（setSelectedDataAsync）：Word 段落标记由 \n 转换，行为一致。
  async function writeSelection(args) {
    if (args.text === undefined || args.text === null) return errResult('bad_args', 'text required');
    // 段落规范化：\r\n / \r → \n（setSelectedDataAsync 对 \n 转 Word 段落标记）
    const text = String(args.text).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    return setSelectionTextAsync(text);
  }

  // ================= 文档信息 / 文档设置（Common API）=================
  // 文档信息：url / title / mode（只读判断）/ settings（文档持久化 key-value）
  async function getDocumentInfo(args) {
    try {
      const d = Office.context.document;
      const mode = String(d.mode || '');
      const out = {
        host: String(Office.context.host || ''),
        url: String(d.url || ''),
        title: String(d.title || ''),
        mode, // 'edit' | 'readOnly' | 'view'
        readOnly: mode === 'readOnly' || mode === 'view',
        settings: {},
      };
      const keys = args && Array.isArray(args.keys) ? args.keys.map(String) : null;
      if (d.settings) {
        if (keys && keys.length) {
          for (const k of keys) {
            try { out.settings[k] = d.settings.get(k); } catch (e) { out.settings[k] = undefined; }
          }
        } else {
          // 无 keys 时尝试枚举宿主暴露的可枚举属性；多数宿主不暴露 → 提示用 keys 精确读取
          try { for (const k in d.settings) { if (typeof d.settings[k] !== 'function') out.settings[k] = d.settings[k]; } } catch (e) { /* 枚举失败忽略 */ }
          out.settingsNote = '未传 args.keys 时仅能枚举宿主暴露的属性；精确读取请传 keys 数组';
        }
      }
      return okResult(out);
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  // 文档设置持久化：settings.set + saveAsync（随文档保存，跨会话可用）
  async function setSetting(args) {
    const key = args && args.key !== undefined ? String(args.key) : null;
    if (!key) return errResult('bad_args', 'key required');
    try {
      const d = Office.context.document;
      if (!d.settings) return errResult('requirement', 'settings API not available');
      const value = args.value; // 任意 JSON 值
      d.settings.set(key, value);
      const saved = await new Promise((resolve) => {
        d.settings.saveAsync((r) => resolve(r.status === Office.AsyncResultStatus.Succeeded));
      });
      if (!saved) return errResult('execution', 'settings.saveAsync failed');
      return okResult({ key, saved: true, value });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  // ================= 环境诊断 =================
  // 基础部分是宿主无关的（Common API + requirement sets）；
  // 宿主专属的探测（例如 Word 对象模型属性是否存在）由该宿主分片的 environmentProbe 追加。
  async function getEnvironment() {
    const base = {
      host: String(Office.context.host || ''),
      platform: String(Office.context.platform || ''),
      officeVersion: String(Office.version || ''),
      // Common API 探测：Office.context.document 可用能力（与 Word.js 对象模型是两套体系，一直可用）
      commonApi: (function () {
        try {
          const d = Office.context.document;
          const proto = Object.getOwnPropertyNames(Object.getPrototypeOf(d) || {}).concat(Object.keys(d));
          return {
            url: String(d.url || ''),
            title: String(d.title || ''),
            mode: String(d.mode || ''),
            settings: !!d.settings,
            customProperties: !!d.customProperties,
            getFileAsync: typeof d.getFileAsync === 'function',
            getSelectedDataAsync: typeof d.getSelectedDataAsync === 'function',
            setSelectedDataAsync: typeof d.setSelectedDataAsync === 'function',
            addHandlerAsync: typeof d.addHandlerAsync === 'function',
            availableKeys: proto.filter((k) => typeof d[k] !== 'function').slice(0, 30),
          };
        } catch (e) { return { error: String(e && e.message || e) }; }
      })(),
      diagnostics: (function () {
        try {
          const d = Office.context.diagnostics || {};
          return {
            officeContext: String(d.OfficeContext || ''),
            hostAppVersion: String(d.HostApplicationVersion || ''),
            hostPlatformVersion: String(d.HostPlatformVersion || ''),
            hostSession: String(d.HostSession || ''),
            hasDiagnostics: !!Office.context.diagnostics,
          };
        } catch (e) { return { error: String(e && e.message || e) }; }
      })(),
      requirementSets: {
        WordApi_1_1: isSetSupported('WordApi', '1.1'),
        WordApi_1_3: isSetSupported('WordApi', '1.3'),
        WordApi_1_4: isSetSupported('WordApi', '1.4'),
        WordApi_1_5: isSetSupported('WordApi', '1.5'),
        WordApi_1_8: isSetSupported('WordApi', '1.8'),
        ExcelApi_1_1: isSetSupported('ExcelApi', '1.1'),
        ExcelApi_1_6: isSetSupported('ExcelApi', '1.6'),
        ExcelApi_1_8: isSetSupported('ExcelApi', '1.8'),
        ExcelApi_1_9: isSetSupported('ExcelApi', '1.9'),
        ExcelApi_1_10: isSetSupported('ExcelApi', '1.10'),
        PowerPointApi_1_1: isSetSupported('PowerPointApi', '1.1'),
        PowerPointApi_1_3: isSetSupported('PowerPointApi', '1.3'),
        PowerPointApi_1_4: isSetSupported('PowerPointApi', '1.4'),
        PowerPointApi_1_5: isSetSupported('PowerPointApi', '1.5'),
        PowerPointApi_1_8: isSetSupported('PowerPointApi', '1.8'),
        PowerPointApi_1_9: isSetSupported('PowerPointApi', '1.9'),
        PowerPointApi_1_10: isSetSupported('PowerPointApi', '1.10'),
        ImageCoercion_1_1: isSetSupported('ImageCoercion', '1.1'),
      },
    };
    const probe = A.hosts[String(Office.context.host || '')];
    if (probe && typeof probe.environmentProbe === 'function') {
      try { return okResult({ ...base, probe: await probe.environmentProbe() }); }
      catch (e) { return errResult('execution', String(e && e.message || e)); }
    }
    return okResult(base);
  }

  A.hosts.Common = {
    actions: {
      write_selection: writeSelection,
      get_document_info: getDocumentInfo,
      set_setting: setSetting,
      get_environment: getEnvironment,
    },
    // 这三个 action 在多个宿主上语义完全一致，所以预览读取器也放这里
    previewReaders: {
      set_setting: (args) => getDocumentInfo({ keys: [args.key] }),
    },
    // 审批面板「修改后内容」的文案（窗格的 renderApproval 通过 window.__AFTER_CONTENT__ 调用）。
    // 跟 impl 放一起，才能做到「改一个 action 只改一处」。
    afterContent: {
      write_selection: (a) => (a.text !== undefined ? a.text : '(未提供)'),
      replace_all: (a) => {
        // "修改后"只显示替换结果：直接写新内容；replace 为空（删除匹配）时写【删去选中】
        const rep = a.replace !== undefined && a.replace !== null ? String(a.replace) : '';
        return rep === '' ? '【删去选中】' : rep;
      },
      format_selection: (a) => JSON.stringify(a),
    },
  };
})();
