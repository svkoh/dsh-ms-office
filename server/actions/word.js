// word.js — Word（docx）宿主的全部 action 实现
//
// 分片契约：本文件是自包含 IIFE，只通过 window.__OFFICE_ACTIONS__ 与其它分片交换东西，
// 与其它分片不共享词法作用域。挂载点：
//   A.hosts.Word = { actions, selectTargets, previewReaders, afterContent, environmentProbe? }
// 这里只放 Word 的代码：分片之后「按宿主分叉」的写法既不需要（文件名就是宿主）也不允许（函数签名里已经没有 host 参数）。
(function () {
  const A = window.__OFFICE_ACTIONS__;
  const U = A.util;
  const okResult = U.okResult, errResult = U.errResult, mapBuiltInStyle = U.mapBuiltInStyle, mapAlignment = U.mapAlignment;
  const isSetSupported = A.isSetSupported;

  // ================= 跨宿主 action 的 Word 实现 =================
  // read_selection / read_document / read_styles / replace_all / format_selection / add_comment /
  // read_comments / read_properties / locate_select 在多个宿主各有一份实现；拆分前它们是同一个
  // 函数里按宿主分叉的那几段，现在每个宿主只保留自己那一份。
  // （write_selection 三宿主实现完全一致，在 common.js。）
  const sleep = A.sleep, notFound = A.notFound;

  async function readSelection(args) {
    const a = args || {};
    if (a.withStyles) {
      try {
        return await Word.run(async (ctx) => {
          const sel = ctx.document.getSelection();
          sel.load('text, font/name, font/size, font/bold, font/italic, font/color, font/underline, font/highlightColor');
          await ctx.sync();
          const f = sel.font;
          return okResult({ host: 'Word', text: sel.text || '', styles: { font: { name: f.name, size: f.size, bold: f.bold, italic: f.italic, color: f.color, underline: f.underline, highlightColor: f.highlightColor } } });
        });
      } catch (e) { return errResult('execution', String(e && e.message || e)); }
    }
    // 有选中返回选中文本；未选中返回光标左右各 5 字（XXXXX | XXXXX，| = 光标位置）
    try {
      return await Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        sel.load('text, start');
        await ctx.sync();
        const text = sel.text || '';
        if (text) return okResult({ text });
        const pos = sel.start;
        if (pos === undefined || pos === null) return okResult({ text: '' });
        // 未选中：取光标所在段落，段内定位取左右各 5 字；定位失败则退回报整段
        const para = sel.paragraphs.getFirst();
        const pr = para.getRange();
        pr.load('start');
        para.load('text');
        await ctx.sync();
        const pt = para.text || '';
        if (!pt) return okResult({ text: '' });
        if (pr.start !== undefined) {
          const rel = Math.max(0, Math.min(pt.length, pos - pr.start));
          const s = Math.max(0, rel - 5);
          const e = Math.min(pt.length, rel + 5);
          return okResult({ text: pt.slice(s, rel) + ' | ' + pt.slice(rel, e) });
        }
        return okResult({ text: pt });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function readDocument() {
    try {
      return await Word.run(async (ctx) => {
        const body = ctx.document.body;
        body.load('text');
        await ctx.sync();
        return okResult({ text: body.text || '', host: 'Word', characterCount: (body.text || '').length });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function readStyles() {
    try {
      return await Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        sel.load('text, font/name, font/size, font/bold, font/italic, font/color, font/underline, font/highlightColor');
        await ctx.sync();
        const f = sel.font;
        return okResult({ host: 'Word', text: sel.text || '', styles: { font: { name: f.name, size: f.size, bold: f.bold, italic: f.italic, color: f.color, underline: f.underline, highlightColor: f.highlightColor } } });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function replaceAll(args) {
    const search = String(args.search ?? '');
    const replace = String(args.replace ?? '');
    if (!search) return errResult('bad_args', 'search required');
    try {
      return await Word.run(async (ctx) => {
        const results = ctx.document.body.search(search, { matchCase: false, matchWholeWord: false, ignoreSpace: true });
        if (args.dryRun) {
          // 预览：返回命中数量 + 命中文本上下文（审批面板可看清将替换什么）
          results.load('items/text');
          await ctx.sync();
          const hits = results.items.map((r) => r.text).slice(0, 20);
          return okResult({ host: 'Word', dryRun: true, wouldReplace: results.items.length, hits });
        }
        results.load('length');
        await ctx.sync();
        let count = 0;
        let lastRange = null;
        for (let i = 0; i < results.items.length; i++) {
          lastRange = results.items[i];
          lastRange.insertText(replace, Word.InsertLocation.replace);
          count++;
        }
        await ctx.sync();
        // 改完自动选中最后一处修改（保持高亮反馈；Office.js 只能单选区，多处无法同时选中）
        if (lastRange) {
          try { lastRange.select(); await ctx.sync(); } catch (e) { /* 选中失败不影响结果 */ }
        }
        return okResult({ replaced: count, host: 'Word', selected: lastRange ? true : false });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function appendText(args) {
    const text = String(args.text ?? '');
    if (!text) return errResult('bad_args', 'text required');
    try {
      return await Word.run(async (ctx) => {
        const para = ctx.document.body.insertParagraph(text, Word.InsertLocation.end);
        await ctx.sync();
        // 改完自动选中追加的段落
        try { para.getRange().select(); await ctx.sync(); } catch (e) { /* 选中失败不影响结果 */ }
        return okResult({ appended: text.length, host: 'Word', selected: true });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function formatSelection(args) {
    try {
      return await Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        const f = sel.font;
        if (args.font) f.name = String(args.font);
        if (args.size) f.size = Number(args.size);
        if (args.bold !== undefined) f.bold = !!args.bold;
        if (args.italic !== undefined) f.italic = !!args.italic;
        if (args.color) f.color = String(args.color);
        if (args.highlight) f.highlightColor = String(args.highlight);
        await ctx.sync();
        return okResult({ host: 'Word', formatted: true });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function addComment(args) {
    const text = String(args.text ?? '');
    if (!text) return errResult('bad_args', 'text required');
    try {
      return await Word.run(async (ctx) => {
        const comments = ctx.document.body.comments;
        if (!comments) return errResult('requirement', 'Word comments API (WordApi 1.4+) not available');
        const sel = ctx.document.getSelection();
        comments.add(sel, text);
        await ctx.sync();
        return okResult({ host: 'Word', added: true });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function readComments() {
    try {
      return await Word.run(async (ctx) => {
        const comments = ctx.document.body.comments;
        if (!comments) return errResult('requirement', 'Word comments API (WordApi 1.4+) not available');
        comments.load('items/content, items/author, items/createdDate, items/resolved');
        await ctx.sync();
        const list = comments.items.map((c) => ({ content: c.content, author: c.author, createdDate: c.createdDate, resolved: c.resolved }));
        return okResult({ host: 'Word', comments: list });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function readProperties() {
    try {
      return await Word.run(async (ctx) => {
        const p = ctx.document.properties;
        p.load('title, author, subject, keywords, comments, category, lastModifiedBy');
        await ctx.sync();
        return okResult({ host: 'Word', properties: { title: p.title, author: p.author, subject: p.subject, keywords: p.keywords, comments: p.comments, category: p.category, lastModifiedBy: p.lastModifiedBy } });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function locateSelect(args) {
    const blinks = Math.max(0, Math.min(5, Number(args.blinks) || 0));
    const interval = Math.max(100, Math.min(800, Number(args.interval) || 300));
    const reps = Math.max(1, blinks);
    const bookmark = (args.bookmark !== undefined ? String(args.bookmark) : args.anchor !== undefined ? String(args.anchor) : null);
    const text = args.text !== undefined ? String(args.text) : null;
    if (!bookmark && !text) return errResult('bad_args', 'locate_select requires one of: bookmark/anchor, text');
    try {
      let anySelected = false;
      for (let i = 0; i < reps; i++) {
        if (bookmark) {
          // 主路径：goToByIdAsync 滚动并选中书签内容（getBookmark 在部分宿主不可用，仅作增强）
          const navigated = await new Promise((resolve) => {
            Office.context.document.goToByIdAsync(bookmark, Office.GoToType.Bookmark, (res) => resolve(res.status === Office.AsyncResultStatus.Succeeded));
          });
          let precise = false;
          try {
            await Word.run(async (ctx) => {
              const bm = ctx.document.getBookmark(bookmark);
              bm.load('name');
              await ctx.sync();
              bm.getRange().select();
              await ctx.sync();
            });
            precise = true;
          } catch (e) { /* getBookmark 不可用，goToByIdAsync 已选中 */ }
          if (!navigated && !precise) throw notFound(`bookmark not found: ${bookmark}`);
          if (navigated || precise) anySelected = true;
        } else {
          await Word.run(async (ctx) => {
            const results = ctx.document.body.search(text, { matchCase: false, ignoreSpace: true });
            results.load('length');
            await ctx.sync();
            if (!results.items || results.items.length === 0) throw notFound(`text not found: ${text}`);
            results.items[0].select();
            await ctx.sync();
          });
          anySelected = true;
        }
        if (blinks > 0) {
          await sleep(interval);
          await Word.run(async (ctx) => {
            // 取消整块高亮：折叠为零宽选区（留在原位）。枚举可能不可用，用字面量 + 降级链。
            const sel = ctx.document.getSelection();
            try { sel.collapse('start'); }
            catch (e) {
              try { sel.getRange('start').select(); }
              catch (e2) { /* 保持选中 */ }
            }
            await ctx.sync();
          });
          await sleep(interval);
        }
      }
      return okResult({ host: 'Word', located: bookmark || text, selected: anySelected, blinks });
    } catch (e) {
      if (e && e.notFound) return errResult('not_found', `locate_select: ${e.message}`);
      return errResult('execution', String(e && e.message || e));
    }
  }

  // ===== 审批预选（selectTarget）：ask 确认模式下，预览时主动选中将要修改的内容 =====
  // 注意：Office.js 只能选中一个连续区域，"一次性多处"无法同时选中 —— 多处替换场景选第一处示意。
  // 不做闪烁：审批时用户焦点在 DSH 窗格（Word 失焦），闪烁根本看不到；只需选中保持，切回 Word 即可见，
  //         执行完成后还有"改完自动选中"兜底。
  async function selectReplaceFirst(args) {
    if (!args.search) return;
    try {
      await Word.run(async (ctx) => {
        const results = ctx.document.body.search(String(args.search), { matchCase: false, ignoreSpace: true });
        results.load('length');
        await ctx.sync();
        if (results.items && results.items.length > 0) { results.items[0].select(); await ctx.sync(); }
      });
    } catch (e) { /* 预选失败不阻塞审批 */ }
  }

  // 仅正文边界插入可移动预选。调用者按动作默认值解析语义；选区/搜索目标保持原选区。
  async function selectWordInsertPoint(location) {
    if (location !== 'start' && location !== 'end') return;
    try {
      await Word.run(async (ctx) => {
        const paras = ctx.document.body.paragraphs;
        paras.load('items');
        await ctx.sync();
        const items = paras.items;
        if (items && items.length) {
          const paragraph = location === 'start' ? items[0] : items[items.length - 1];
          paragraph.getRange(location === 'start' ? 'Start' : 'End').select();
          await ctx.sync();
        }
      });
    } catch (e) { /* 定位失败不阻塞审批 */ }
  }

  // ================= Word 组 =================

  async function readTables() {
    try {
      return await Word.run(async (ctx) => {
        const tables = ctx.document.body.tables;
        tables.load('items/rowCount,items/style,items/styleBuiltIn');
        await ctx.sync();
        const out = [];
        for (let ti = 0; ti < tables.items.length; ti++) {
          const t = tables.items[ti];
          const range = t.getRange();
          range.load('text');
          await ctx.sync();
          const raw = range.text || '';
          const lines = raw.split('\r').filter((l) => l.length > 0);
          const cells = lines.map((l) => l.split('\t').map((v) => v.trim()));
          out.push({ index: ti + 1, rowCount: t.rowCount, columnCount: cells.length ? cells[0].length : 0, style: t.style || null, styleBuiltIn: t.styleBuiltIn || null, cells });
        }
        return okResult({ host: 'Word', tables: out });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function readTableBorders() {
    try {
      return await Word.run(async (ctx) => {
        const tables = ctx.document.body.tables;
        tables.load('items');
        await ctx.sync();
        const locations = ['Top', 'Bottom', 'Left', 'Right', 'InsideHorizontal', 'InsideVertical'];
        const refs = [];
        for (let ti = 0; ti < tables.items.length; ti++) {
          for (const location of locations) {
            const border = tables.items[ti].getBorder(location);
            border.load('color,type,width');
            refs.push({ table: ti + 1, location, border });
          }
        }
        await ctx.sync();
        return okResult({
          host: 'Word',
          tableCount: tables.items.length,
          borders: refs.map((ref) => ({
            table: ref.table,
            location: ref.location,
            color: ref.border.color,
            type: ref.border.type,
            width: ref.border.width,
          })),
        });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function setTableBorders(args) {
    const color = String(args.color || '');
    if (!/^#[0-9a-fA-F]{6}$/.test(color)) return errResult('bad_args', 'color required in #RRGGBB format');
    try {
      return await Word.run(async (ctx) => {
        const tables = ctx.document.body.tables;
        tables.load('items');
        await ctx.sync();
        const allLocations = ['Top', 'Bottom', 'Left', 'Right', 'InsideHorizontal', 'InsideVertical'];
        const locations = Array.isArray(args.locations) ? args.locations.filter((v) => allLocations.includes(v)) : allLocations;
        if (!locations.length) return errResult('bad_args', 'locations contains no valid border location');
        const requested = Array.isArray(args.tableIndexes) ? args.tableIndexes.map(Number) : tables.items.map((_, i) => i + 1);
        const indexes = [...new Set(requested)].filter((n) => Number.isInteger(n) && n >= 1 && n <= tables.items.length);
        if (!indexes.length) return errResult('bad_args', 'tableIndexes contains no valid table index');
        const width = args.width === undefined ? 0.5 : Number(args.width);
        if (!Number.isFinite(width) || width <= 0) return errResult('bad_args', 'width must be a positive number');
        const borders = [];
        for (const index of indexes) {
          const table = tables.items[index - 1];
          if (args.enableGrid === true) table.styleBuiltIn = 'TableGrid';
          for (const location of locations) borders.push(table.getBorder(location));
        }
        if (args.enableGrid === true) await ctx.sync();
        for (const border of borders) border.color = color;
        await ctx.sync();
        return okResult({ host: 'Word', tableIndexes: indexes, borderCount: indexes.length * locations.length, color: color.toUpperCase(), note: '保留原有边框线型和粗细' });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function setFont(args) {
    const font = String(args.font || '');
    if (!font) return errResult('bad_args', 'font required');
    try {
      return await Word.run(async (ctx) => {
        const paras = ctx.document.body.paragraphs;
        paras.load('items/font/name, items/text');
        await ctx.sync();
        let count = 0;
        for (const p of paras.items) {
          p.font.name = font;
          count++;
        }
        await ctx.sync();
        return okResult({ host: 'Word', paragraphs: count, font });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function wordListParagraphs(args) {
    const filter = String(args.filter || 'all');
    if (!['all', 'empty', 'nonEmpty'].includes(filter)) return errResult('bad_args', 'filter must be all|empty|nonEmpty');
    try { return await Word.run(async (ctx) => {
      const paragraphs = ctx.document.body.paragraphs; paragraphs.load('items/text,items/style,items/styleBuiltIn,items/inlinePictures/items'); await ctx.sync();
      const items = paragraphs.items.map((p, index) => ({ index, text: p.text || '', empty: !String(p.text || '').trim(), style: p.style || null, styleBuiltIn: p.styleBuiltIn || null, inlinePictureCount: p.inlinePictures.items.length, final: index === paragraphs.items.length - 1 })).filter((p) => filter === 'all' || (filter === 'empty' ? p.empty : !p.empty));
      return okResult({ host: 'Word', count: items.length, paragraphs: items });
    }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function wordDeleteParagraphs(args) {
    const targets = args.targets;
    if (!Array.isArray(targets) || !targets.length || targets.some((t) => !t || !Number.isInteger(Number(t.index)) || Number(t.index) < 0 || t.expectedText === undefined)) return errResult('bad_args', 'targets must be [{index,expectedText}] with zero-based indexes');
    try { return await Word.run(async (ctx) => {
      const paragraphs = ctx.document.body.paragraphs; paragraphs.load('items/text,items/inlinePictures/items'); await ctx.sync();
      const indexes = targets.map((t) => Number(t.index)); if (new Set(indexes).size !== indexes.length) return errResult('bad_args', 'target indexes must be unique');
      for (const target of targets) { const index = Number(target.index), p = paragraphs.items[index]; if (!p) return errResult('not_found', `paragraph ${index} not found`); if (String(p.text || '') !== String(target.expectedText)) return errResult('not_found', `paragraph ${index} text changed`); if (index === paragraphs.items.length - 1 || p.inlinePictures.items.length) return errResult('bad_args', `paragraph ${index} cannot be safely deleted`); }
      for (const index of indexes.slice().sort((a, b) => b - a)) paragraphs.items[index].delete(); await ctx.sync();
      return okResult({ host: 'Word', deleted: indexes.length, indexes });
    }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function insertParagraph(args) {
    const text = String(args.text ?? '');
    if (!text) return errResult('bad_args', 'text required');
    try {
      return await Word.run(async (ctx) => {
        const loc = args.location === 'start' ? Word.InsertLocation.start : (args.location === 'afterSelection' ? Word.InsertLocation.after : Word.InsertLocation.end);
        let para;
        if (args.location === 'afterSelectionParagraph') {
          const source = ctx.document.getSelection().paragraphs.getFirst();
          para = source.insertParagraph(text, Word.InsertLocation.after);
        } else if (args.location === 'afterSelection') {
          const sel = ctx.document.getSelection();
          para = sel.insertParagraph(text, Word.InsertLocation.after);
        } else {
          para = ctx.document.body.insertParagraph(text, loc);
        }
        const style = mapBuiltInStyle(args.style);
        if (style) { para.styleBuiltIn = Word.BuiltInStyleName[style]; }
        await ctx.sync();
        // 改完自动选中插入的段落
        try { para.getRange().select(); await ctx.sync(); } catch (e) { /* 选中失败不影响结果 */ }
        return okResult({ host: 'Word', inserted: text.length, style: style || null, location: args.location || 'end', selected: true });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  function validateWordTableRows(rows) {
    if (!Array.isArray(rows) || rows.length === 0 || !Array.isArray(rows[0]) || rows[0].length === 0) {
      throw new Error('rows must be a non-empty 2D array');
    }
    const columnCount = rows[0].length;
    if (rows.length > 1000 || columnCount > 63 || rows.length * columnCount > 10000) {
      throw new Error('table exceeds limit (1000 rows, 63 columns, 10000 cells)');
    }
    return rows.map((row, rowIndex) => {
      if (!Array.isArray(row) || row.length !== columnCount) throw new Error(`rows[${rowIndex}] must contain exactly ${columnCount} cells`);
      return row.map((value) => String(value ?? ''));
    });
  }

  async function insertTable(args) {
    if (!isSetSupported('WordApi', '1.3')) return errResult('requirement', 'insert_table requires WordApi 1.3+');
    let stringRows;
    try { stringRows = validateWordTableRows(args && args.rows); }
    catch (e) { return errResult('bad_args', String(e && e.message || e)); }
    const location = String(args.location || 'end');
    if (!['start', 'end', 'beforeSelection', 'afterSelection'].includes(location)) {
      return errResult('bad_args', 'location must be start, end, beforeSelection, or afterSelection');
    }
    if (args.style && args.styleBuiltIn) return errResult('bad_args', 'style and styleBuiltIn are mutually exclusive');
    if (args.headerRow !== undefined && typeof args.headerRow !== 'boolean') return errResult('bad_args', 'headerRow must be boolean');
    try {
      return await Word.run(async (ctx) => {
        const rowCount = stringRows.length;
        const columnCount = stringRows[0].length;
        let table;
        if (location === 'beforeSelection' || location === 'afterSelection') {
          const selection = ctx.document.getSelection();
          const insertLocation = location === 'beforeSelection' ? Word.InsertLocation.before : Word.InsertLocation.after;
          table = selection.insertTable(rowCount, columnCount, insertLocation, stringRows);
        } else {
          const insertLocation = location === 'start' ? Word.InsertLocation.start : Word.InsertLocation.end;
          table = ctx.document.body.insertTable(rowCount, columnCount, insertLocation, stringRows);
        }
        if (args.style) table.style = String(args.style);
        if (args.styleBuiltIn) table.styleBuiltIn = String(args.styleBuiltIn);
        if (args.headerRow !== undefined) table.headerRowCount = args.headerRow ? 1 : 0;
        table.load('rowCount,values,headerRowCount,style,styleBuiltIn');
        await ctx.sync();
        let selected = false;
        try { table.getRange().select(); await ctx.sync(); selected = true; } catch (e) { /* Selection is optional. */ }
        return okResult({ host: 'Word', rows: table.rowCount, cols: table.values && table.values[0] ? table.values[0].length : columnCount, values: table.values, location, style: table.style || null, styleBuiltIn: table.styleBuiltIn || null, headerRowCount: table.headerRowCount, selected, requirementSet: 'WordApi 1.3' });
      });
    } catch (e) {
      const debug = (e && e.debugInfo) ? (' [debug] ' + JSON.stringify(e.debugInfo)) : '';
      return errResult('execution', `${e && e.message || e}${debug}`);
    }
  }

  const WORD_HEADER_FOOTER_TYPES = { primary: 'Primary', firstPage: 'FirstPage', evenPages: 'EvenPages' };
  const WORD_BREAK_TYPES = { page: 'Page', line: 'Line', sectionNext: 'SectionNext', sectionContinuous: 'SectionContinuous', sectionEven: 'SectionEven', sectionOdd: 'SectionOdd' };

  async function wordListSections(args) {
    if (!isSetSupported('WordApi', '1.1')) return errResult('requirement', 'word_list_sections requires WordApi 1.1+');
    try {
      return await Word.run(async (ctx) => {
        const sections = ctx.document.sections;
        sections.load('items');
        await ctx.sync();
        const limit = Math.min(5000, Math.max(50, Number(args.textLimit) || 500));
        const refs = sections.items.map((section) => { const body = section.body; body.load('text'); return body; });
        await ctx.sync();
        return okResult({ host: 'Word', count: refs.length, sections: refs.map((body, index) => ({ index, text: String(body.text || '').slice(0, limit), truncated: String(body.text || '').length > limit })) });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  function validateHeaderFooterArgs(args, requireWrite) {
    const section = args.section === undefined ? 0 : args.section;
    if (section !== 'all' && (!Number.isInteger(Number(section)) || Number(section) < 0)) throw new Error('section must be a zero-based integer or all');
    const types = args.types || [args.type || 'primary'];
    if (!Array.isArray(types) || !types.length || types.some((type) => !WORD_HEADER_FOOTER_TYPES[type])) throw new Error('type(s) must be primary, firstPage, or evenPages');
    if (requireWrite && !['header', 'footer'].includes(args.kind)) throw new Error('kind must be header or footer');
    if (requireWrite && args.text === undefined) throw new Error('text required');
    if (requireWrite && !['replace', 'append', 'prepend'].includes(args.mode || 'replace')) throw new Error('mode must be replace, append, or prepend');
    return { section, types: Array.from(new Set(types)) };
  }

  async function wordReadHeadersFooters(args) {
    if (!isSetSupported('WordApi', '1.1')) return errResult('requirement', 'word_read_headers_footers requires WordApi 1.1+');
    let spec; try { spec = validateHeaderFooterArgs(args || {}, false); } catch (e) { return errResult('bad_args', e.message); }
    try {
      return await Word.run(async (ctx) => {
        const sections = ctx.document.sections; sections.load('items'); await ctx.sync();
        const indexes = spec.section === 'all' ? sections.items.map((_, i) => i) : [Number(spec.section)];
        if (indexes.some((i) => i >= sections.items.length)) return errResult('bad_args', `section index out of range (count ${sections.items.length})`);
        const refs = [];
        for (const index of indexes) for (const type of spec.types) {
          const section = sections.items[index];
          const header = section.getHeader(WORD_HEADER_FOOTER_TYPES[type]);
          const footer = section.getFooter(WORD_HEADER_FOOTER_TYPES[type]);
          header.load('text'); footer.load('text'); refs.push({ index, type, header, footer });
        }
        await ctx.sync();
        return okResult({ host: 'Word', sectionCount: sections.items.length, items: refs.map((r) => ({ section: r.index, type: r.type, header: r.header.text || '', footer: r.footer.text || '' })), warning: 'Office.js cannot inspect or disable Link to Previous; first/even body existence does not prove its page-layout switch is enabled.' });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function wordTransferContent(args) {
    if (!isSetSupported('WordApi', '1.1')) return errResult('requirement', 'word_transfer_content requires WordApi 1.1+');
    const source = args.source || {}, destination = args.destination || {};
    const operation = String(args.operation || 'copy'), mode = String(args.mode || 'append');
    if (!['selection', 'table', 'search'].includes(source.kind)) return errResult('bad_args', 'source.kind must be selection|table|search');
    if (source.kind === 'search' && !String(source.query || '')) return errResult('bad_args', 'search source requires query');
    if (!['selection', 'body', 'header', 'footer'].includes(destination.kind)) return errResult('bad_args', 'destination.kind must be selection|body|header|footer');
    if (!['copy', 'move'].includes(operation) || !['replace', 'append', 'prepend', 'before', 'after'].includes(mode)) return errResult('bad_args', 'invalid operation or mode');
    try {
      return await Word.run(async (ctx) => {
        let sourceRange, targetRange;
        const tables = ctx.document.body.tables, sections = ctx.document.sections;
        if (source.kind === 'table') tables.load('items');
        if (source.kind === 'search') sourceRange = ctx.document.body.search(String(source.query || ''), { matchCase: !!source.matchCase });
        if (sourceRange) sourceRange.load('items');
        if (destination.kind === 'header' || destination.kind === 'footer') sections.load('items');
        await ctx.sync();
        if (source.kind === 'selection') sourceRange = ctx.document.getSelection();
        if (source.kind === 'table') { const index = Number(source.index); if (!Number.isInteger(index) || index < 0 || index >= tables.items.length) return errResult('not_found', 'source table index not found'); sourceRange = tables.items[index].getRange(); }
        if (source.kind === 'search') { const index = Number(source.occurrence === undefined ? 0 : source.occurrence); if (!Number.isInteger(index) || index < 0 || index >= sourceRange.items.length) return errResult('not_found', 'source search occurrence not found'); sourceRange = sourceRange.items[index]; }
        if (destination.kind === 'selection') targetRange = ctx.document.getSelection();
        else if (destination.kind === 'body') targetRange = ctx.document.body;
        else { const index = Number(destination.section === undefined ? 0 : destination.section), type = String(destination.type || 'primary'); if (!Number.isInteger(index) || index < 0 || index >= sections.items.length || !WORD_HEADER_FOOTER_TYPES[type]) return errResult('bad_args', 'invalid destination section or type'); targetRange = destination.kind === 'header' ? sections.items[index].getHeader(WORD_HEADER_FOOTER_TYPES[type]) : sections.items[index].getFooter(WORD_HEADER_FOOTER_TYPES[type]); }
        sourceRange.load('text'); targetRange.load('text'); const ooxml = sourceRange.getOoxml(); await ctx.sync(); const transferredText = sourceRange.text || '';
        if (args.expectedText !== undefined && String(transferredText).trim() !== String(args.expectedText).trim()) return errResult('not_found', 'source text changed since it was read');
        if (mode === 'replace' && String(targetRange.text || '').trim() && args.overwrite !== true) return errResult('bad_args', 'destination is not empty; set overwrite:true');
        const locations = { replace: Word.InsertLocation.replace, append: Word.InsertLocation.end, prepend: Word.InsertLocation.start, before: Word.InsertLocation.before, after: Word.InsertLocation.after };
        targetRange.insertOoxml(ooxml.value, locations[mode]); await ctx.sync();
        if (operation === 'move') { sourceRange.delete(); await ctx.sync(); }
        return okResult({ host: 'Word', operation, source, destination, mode, text: transferredText });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function wordSetHeaderFooter(args) {
    if (!isSetSupported('WordApi', '1.1')) return errResult('requirement', 'word_set_header_footer requires WordApi 1.1+');
    let spec; try { spec = validateHeaderFooterArgs(args || {}, true); } catch (e) { return errResult('bad_args', e.message); }
    try {
      return await Word.run(async (ctx) => {
        const sections = ctx.document.sections; sections.load('items'); await ctx.sync();
        const indexes = spec.section === 'all' ? sections.items.map((_, i) => i) : [Number(spec.section)];
        if (indexes.some((i) => i >= sections.items.length)) return errResult('bad_args', `section index out of range (count ${sections.items.length})`);
        const type = spec.types[0]; const refs = [];
        for (const index of indexes) {
          const body = args.kind === 'header' ? sections.items[index].getHeader(WORD_HEADER_FOOTER_TYPES[type]) : sections.items[index].getFooter(WORD_HEADER_FOOTER_TYPES[type]);
          const location = (args.mode || 'replace') === 'prepend' ? 'Start' : (args.mode || 'replace') === 'append' ? 'End' : 'Replace';
          body.insertText(String(args.text), location); body.load('text'); refs.push({ index, body });
        }
        await ctx.sync();
        return okResult({ host: 'Word', kind: args.kind, type, sections: refs.map((r) => ({ index: r.index, text: r.body.text || '' })), warning: 'Header/footer may be linked to a previous section; the portable Office.js API cannot unlink it.' });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function wordInsertBreak(args) {
    if (!isSetSupported('WordApi', '1.1')) return errResult('requirement', 'word_insert_break requires WordApi 1.1+');
    const type = String(args.type || ''); const location = String(args.location || 'afterSelection');
    if (!WORD_BREAK_TYPES[type]) return errResult('bad_args', `type must be ${Object.keys(WORD_BREAK_TYPES).join('|')}`);
    if (!['start', 'end', 'beforeSelection', 'afterSelection'].includes(location)) return errResult('bad_args', 'invalid location');
    try {
      return await Word.run(async (ctx) => {
        const sections = ctx.document.sections; sections.load('items'); await ctx.sync(); const before = sections.items.length;
        if (location === 'start' || location === 'end') ctx.document.body.insertBreak(WORD_BREAK_TYPES[type], location === 'start' ? 'Start' : 'End');
        else ctx.document.getSelection().insertBreak(WORD_BREAK_TYPES[type], location === 'beforeSelection' ? 'Before' : 'After');
        await ctx.sync(); sections.load('items'); await ctx.sync();
        return okResult({ host: 'Word', type, location, sectionCountBefore: before, sectionCountAfter: sections.items.length });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function wordInsertCaption(args) {
    const dynamic = args.dynamic === true;
    if (!isSetSupported('WordApi', dynamic ? '1.5' : '1.1')) return errResult('requirement', `word_insert_caption ${dynamic ? 'dynamic' : 'static'} mode requires WordApi ${dynamic ? '1.5' : '1.1'}+`);
    if (dynamic && Office.context.platform === Office.PlatformType.OfficeOnline) return errResult('unsupported', 'dynamic SEQ field insertion is not supported in Word on the web; use dynamic:false with an explicit number');
    const label = String(args.label || '').trim(), title = String(args.title || '').trim(), sequenceId = String(args.sequenceId || 'Figure');
    const separator = args.separator === undefined ? ': ' : String(args.separator); const position = String(args.position || 'afterSelection');
    if (!label || !title || title.length > 1000) return errResult('bad_args', 'non-empty label and title (max 1000 chars) required');
    if (dynamic && !/^[A-Za-z][A-Za-z0-9_]{0,39}$/.test(sequenceId)) return errResult('bad_args', 'sequenceId must match [A-Za-z][A-Za-z0-9_]{0,39}');
    const staticNumber = Number(args.number === undefined ? 1 : args.number);
    if (!dynamic && (!Number.isInteger(staticNumber) || staticNumber < 0)) return errResult('bad_args', 'number must be a non-negative integer in static mode');
    if (/\r|\n/.test(label + sequenceId + separator) || !['beforeSelection', 'afterSelection'].includes(position)) return errResult('bad_args', 'invalid label, separator, or position');
    try {
      return await Word.run(async (ctx) => {
        const selection = ctx.document.getSelection();
        const paragraph = selection.insertParagraph((dynamic ? '' : String(staticNumber)) + separator + title, position === 'beforeSelection' ? 'Before' : 'After');
        paragraph.styleBuiltIn = 'Caption'; await ctx.sync();
        if (dynamic) {
          const start = paragraph.getRange('Start'); start.insertField('After', 'Seq', sequenceId, true); await ctx.sync();
        }
        paragraph.getRange('Start').insertText(label + ' ', 'Start'); await ctx.sync();
        const range = paragraph.getRange(); range.load('text'); await ctx.sync(); range.select(); await ctx.sync();
        return okResult({ host: 'Word', text: range.text || '', label, sequenceId: dynamic ? sequenceId : null, fieldType: dynamic ? 'Seq' : null, dynamic, position, selected: true });
      });
    } catch (e) { return errResult('execution', `${dynamic ? 'dynamic SEQ field' : 'static caption'}: ${String(e && e.message || e)}`); }
  }

  const WORD_MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
  const WORD_MAIN_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

  function mathEscapeXml(value) {
    return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
  }

  function mathNodeToOmml(node, state, depth) {
    if (depth > 20 || ++state.count > 500) throw new Error('equation is too complex (max depth 20, max nodes 500)');
    if (typeof node === 'string' || typeof node === 'number') node = { type: 'text', value: String(node) };
    if (!node || typeof node !== 'object') throw new Error('equation node must be an object, string, or number');
    const type = String(node.type || 'text');
    if (type === 'text') {
      const value = String(node.value === undefined ? '' : node.value);
      if (value.length > 2000) throw new Error('equation text node too long (max 2000 characters)');
      const normal = node.normal === true ? '<m:nor/>' : '';
      return `<m:r><m:rPr>${normal}</m:rPr><m:t xml:space="preserve">${mathEscapeXml(value)}</m:t></m:r>`;
    }
    if (type === 'group') {
      if (!Array.isArray(node.children) || !node.children.length) throw new Error('group.children must be a non-empty array');
      return node.children.map((child) => mathNodeToOmml(child, state, depth + 1)).join('');
    }
    if (type === 'fraction') {
      if (node.numerator === undefined || node.denominator === undefined) throw new Error('fraction requires numerator and denominator');
      return `<m:f><m:fPr><m:type m:val="bar"/></m:fPr><m:num>${mathNodeToOmml(node.numerator, state, depth + 1)}</m:num><m:den>${mathNodeToOmml(node.denominator, state, depth + 1)}</m:den></m:f>`;
    }
    if (type === 'subscript' || type === 'superscript') {
      if (node.base === undefined || node.script === undefined) throw new Error(`${type} requires base and script`);
      const tag = type === 'subscript' ? 'sSub' : 'sSup';
      const scriptTag = type === 'subscript' ? 'sub' : 'sup';
      return `<m:${tag}><m:e>${mathNodeToOmml(node.base, state, depth + 1)}</m:e><m:${scriptTag}>${mathNodeToOmml(node.script, state, depth + 1)}</m:${scriptTag}></m:${tag}>`;
    }
    throw new Error(`unsupported equation node type: ${type}`);
  }

  function buildEquationOoxml(expression, display) {
    const math = mathNodeToOmml(expression, { count: 0 }, 0);
    const mathBlock = display === false ? `<m:oMath>${math}</m:oMath>` : `<m:oMathPara><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr><m:oMath>${math}</m:oMath></m:oMathPara>`;
    const paragraph = `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${mathBlock}</w:p>`;
    const relationships = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
    return `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>${relationships}</pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="${WORD_MAIN_NS}" xmlns:m="${WORD_MATH_NS}"><w:body>${paragraph}<w:sectPr/></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
  }

  async function wordInsertEquation(args) {
    if (!isSetSupported('WordApi', '1.1')) return errResult('requirement', 'word_insert_equation requires WordApi 1.1+ Range.insertOoxml');
    const target = args.target || { kind: (args.location === 'start' || args.location === 'end') ? 'body' : 'selection' };
    const kind = String(target.kind || 'selection');
    const mode = String(args.mode || ({ beforeSelection: 'before', afterSelection: 'after', replaceSelection: 'replace', start: 'prepend', end: 'append' }[args.location] || 'after'));
    if (!['selection', 'body', 'search'].includes(kind) || !['replace', 'before', 'after', 'prepend', 'append'].includes(mode)) return errResult('bad_args', 'invalid target.kind or mode');
    const query = String(target.query || ''), occurrence = Number(target.occurrence === undefined ? 0 : target.occurrence);
    if (kind === 'search' && (!query || !Number.isInteger(occurrence) || occurrence < 0)) return errResult('bad_args', 'search target requires query and zero-based occurrence');
    let ooxml; try { ooxml = buildEquationOoxml(args.expression, args.display !== false); } catch (e) { return errResult('bad_args', String(e && e.message || e)); }
    try { return await Word.run(async (ctx) => {
      let range;
      if (kind === 'body') range = ctx.document.body;
      else if (kind === 'selection') range = ctx.document.getSelection();
      else { const hits = ctx.document.body.search(query, { matchCase: !!target.matchCase, matchWholeWord: !!target.matchWholeWord }); hits.load('items/text'); await ctx.sync(); if (occurrence >= hits.items.length) return errResult('not_found', 'search occurrence not found'); range = hits.items[occurrence]; }
      range.load('text'); await ctx.sync(); const previousText = range.text || '';
      const locations = { replace: Word.InsertLocation.replace, before: Word.InsertLocation.before, after: Word.InsertLocation.after, prepend: Word.InsertLocation.start, append: Word.InsertLocation.end };
      range.insertOoxml(ooxml, locations[mode]); await ctx.sync();
      return okResult({ host: 'Word', inserted: true, nativeEquation: true, target: { kind, ...(kind === 'search' ? { query, occurrence } : {}) }, mode, previousText, display: args.display !== false });
    }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function wordReadSelectionParagraph() {
    try {
      return await Word.run(async (ctx) => {
        const paragraph = ctx.document.getSelection().paragraphs.getFirst();
        const range = paragraph.getRange();
        range.load('text');
        const ooxml = range.getOoxml();
        await ctx.sync();
        const xml = String(ooxml.value || '');
        const tokens = [];
        const re = /<m:t[^>]*>([\s\S]*?)<\/m:t>/g;
        let match;
        while ((match = re.exec(xml))) tokens.push(match[1]);
        return okResult({ host: 'Word', text: range.text || '', mathText: tokens.join(''), hasEquation: tokens.length > 0, ooxml: xml });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function wordReadEquations() {
    if (!isSetSupported('WordApi', '1.1')) return errResult('requirement', 'word_read_equations requires WordApi 1.1+ Range.getOoxml');
    try {
      return await Word.run(async (ctx) => {
        const result = ctx.document.body.getOoxml();
        await ctx.sync();
        const xml = String(result.value || '');
        const doc = new DOMParser().parseFromString(xml, 'application/xml');
        const parseErrors = doc.getElementsByTagName('parsererror');
        if (parseErrors.length) return errResult('execution', 'Word returned malformed OOXML while reading equations');
        const nodes = Array.from(doc.getElementsByTagNameNS(WORD_MATH_NS, 'oMath'));
        const equations = nodes.map((node, index) => ({ index: index + 1, linearText: node.textContent || '', omml: new XMLSerializer().serializeToString(node) }));
        return okResult({ host: 'Word', count: equations.length, equations });
      });
    } catch (e) {
      const debug = (e && e.debugInfo) ? (' [debug] ' + JSON.stringify(e.debugInfo)) : '';
      return errResult('execution', String(e && e.message || e) + debug);
    }
  }

  async function wordInsertImage(args) {
    const b64 = String(args.base64 || '').replace(/^data:image\/\w+;base64,/, '');
    const mode = String(args.mode || 'after');
    if (!b64 || !['before', 'after', 'replace'].includes(mode)) return errResult('bad_args', 'base64 and mode before|after|replace required');
    const width = args.width === undefined ? null : Number(args.width), height = args.height === undefined ? null : Number(args.height);
    if ((width !== null && (!Number.isFinite(width) || width <= 0)) || (height !== null && (!Number.isFinite(height) || height <= 0))) return errResult('bad_args', 'width and height must be positive numbers');
    try { return await Word.run(async (ctx) => {
      const range = ctx.document.getSelection(); range.load('text'); await ctx.sync(); const previousText = range.text || '';
      const pic = range.insertInlinePictureFromBase64(b64, Word.InsertLocation[mode]); if (width !== null) pic.width = width; if (height !== null) pic.height = height; await ctx.sync();
      return okResult({ host: 'Word', inserted: true, mode, replacedText: mode === 'replace' ? previousText : '' });
    }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function wordInsertHtml(args) {
    if (!isSetSupported('WordApi', '1.1')) return errResult('requirement', 'word_insert_html requires WordApi 1.1+');
    const html = String(args && args.html || ''), target = String(args.target || 'body'), mode = String(args.mode || 'append');
    if (!html || html.length > 8 * 1024 * 1024 || !['body', 'selection'].includes(target) || !['replace', 'append', 'prepend', 'before', 'after'].includes(mode)) return errResult('bad_args', 'html, target body|selection, and valid mode required');
    try { return await Word.run(async (ctx) => {
      const range = target === 'body' ? ctx.document.body : ctx.document.getSelection(); range.load('text'); await ctx.sync(); const previousText = range.text || '';
      const locations = { replace: Word.InsertLocation.replace, append: Word.InsertLocation.end, prepend: Word.InsertLocation.start, before: Word.InsertLocation.before, after: Word.InsertLocation.after };
      range.insertHtml(html, locations[mode]); await ctx.sync();
      return okResult({ host: 'Word', target, mode, previousText, htmlCharacters: html.length });
    }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function wordSetPageSetup(args) {
    const fields = ['topMargin', 'bottomMargin', 'leftMargin', 'rightMargin', 'headerDistance', 'footerDistance'];
    const values = {}; try { for (const key of fields) if (args[key] !== undefined) { const value = Number(args[key]); if (!Number.isFinite(value) || value < 0) throw new Error(`${key} must be a non-negative point value`); values[key] = value; } } catch (e) { return errResult('bad_args', e.message); }
    if (!Object.keys(values).length) return errResult('bad_args', `at least one of ${fields.join(', ')} is required`);
    const requested = args.sections === undefined || args.sections === 'all' ? 'all' : args.sections;
    if (requested !== 'all' && (!Array.isArray(requested) || requested.some((i) => !Number.isInteger(Number(i)) || Number(i) < 0))) return errResult('bad_args', 'sections must be all or an array of zero-based indexes');
    try { return await Word.run(async (ctx) => {
      const sections = ctx.document.sections; sections.load('items'); await ctx.sync(); const indexes = requested === 'all' ? sections.items.map((_, i) => i) : Array.from(new Set(requested.map(Number)));
      if (indexes.some((i) => i >= sections.items.length)) return errResult('not_found', 'section index not found');
      for (const index of indexes) Object.assign(sections.items[index].pageSetup, values); await ctx.sync();
      return okResult({ host: 'Word', sections: indexes, values });
    }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  // getSliceAsync 的第一个参数是「零基分片序号」，不是字节偏移量（早期版本传累计字节数，
  // 文件 ≤ 1 片时凑巧能跑通，多片时第 2 片就取不到 → 导出结果被静默截断）。
  function wordExportFile(args) {
    const type = String(args && args.type || 'docx').toLowerCase();
    if (!['docx', 'pdf'].includes(type)) return Promise.resolve(errResult('bad_args', 'type must be docx or pdf'));
    const fileType = type === 'pdf' ? Office.FileType.Pdf : Office.FileType.Compressed;
    return new Promise((resolve) => {
      Office.context.document.getFileAsync(fileType, { sliceSize: 4194304 }, (fileResult) => {
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
          if (sliceIndex >= file.sliceCount || received >= file.size) {
            const mime = type === 'pdf' ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
            const size = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
            const reader = new FileReader();
            reader.onerror = () => done(errResult('execution', String(reader.error || 'FileReader failed')));
            reader.onload = () => done(okResult({ host: 'Word', fileType: type, size, base64: String(reader.result || '').replace(/^data:.*?;base64,/, '') }));
            reader.readAsDataURL(new Blob(chunks, { type: mime }));
          } else file.getSliceAsync(sliceIndex, onSlice);
        };
        file.getSliceAsync(0, onSlice);
      });
    });
  }

  async function applyStyle(args) {
    const style = mapBuiltInStyle(args.style);
    if (!style) return errResult('bad_args', `unknown style: ${args.style}`);
    try {
      return await Word.run(async (ctx) => {
        const bs = Word.BuiltInStyleName[style];
        let count = 0;
        if (args.scope === 'all') {
          const paras = ctx.document.body.paragraphs;
          paras.load('items');
          await ctx.sync();
          for (const p of paras.items) { p.styleBuiltIn = bs; count++; }
        } else {
          const sel = ctx.document.getSelection();
          sel.styleBuiltIn = bs;
          count = 1;
        }
        await ctx.sync();
        return okResult({ host: 'Word', applied: count, style, scope: args.scope || 'selection' });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function setParagraphFormat(args) {
    try {
      return await Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        const pf = sel.paragraphFormat;
        if (!pf) return errResult('requirement', 'Word paragraphFormat API not available in this environment');
        if (args.alignment) {
          const al = mapAlignment(args.alignment);
          if (al) pf.alignment = al;
        }
        if (args.indent !== undefined) pf.leftIndent = Number(args.indent) * 28.35;
        if (args.lineSpacing !== undefined) pf.lineSpacing = Number(args.lineSpacing);
        if (args.listType && args.listType !== 'none') {
          const paras = sel.paragraphs;
          paras.load('items');
          await ctx.sync();
          ctx.document.body.lists.create(paras.items, args.listType === 'numbered' ? Word.ListType.numbered : Word.ListType.bulleted);
        }
        await ctx.sync();
        return okResult({ host: 'Word', formatted: true, listType: args.listType || null });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function search(args) {
    const query = String(args.query ?? '');
    if (!query) return errResult('bad_args', 'query required');
    try {
      return await Word.run(async (ctx) => {
        const results = ctx.document.body.search(query, { matchCase: !!args.matchCase, matchWildcards: !!args.wildcard });
        results.load('items/text');
        await ctx.sync();
        const hits = results.items.map((r, i) => ({ index: i + 1, text: r.text }));
        return okResult({ host: 'Word', hits, count: hits.length });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  A.hosts.Word = {
    actions: {
      // 跨宿主
      read_selection: readSelection,
      read_document: readDocument,
      read_styles: readStyles,
      replace_all: replaceAll,
      format_selection: formatSelection,
      add_comment: addComment,
      read_comments: readComments,
      read_properties: readProperties,
      locate_select: locateSelect,
      // Word 专有
      append_text: appendText,
      read_tables: readTables,
      read_table_borders: readTableBorders,
      set_table_borders: setTableBorders,
      set_font: setFont,
      word_list_paragraphs: wordListParagraphs,
      word_delete_paragraphs: wordDeleteParagraphs,
      insert_paragraph: insertParagraph,
      insert_table: insertTable,
      word_insert_html: wordInsertHtml,
      word_set_page_setup: wordSetPageSetup,
      word_insert_equation: wordInsertEquation,
      word_read_equations: wordReadEquations,
      word_read_selection_paragraph: wordReadSelectionParagraph,
      word_export_file: wordExportFile,
      word_insert_image: wordInsertImage,
      apply_style: applyStyle,
      set_paragraph_format: setParagraphFormat,
      search,
      word_list_sections: wordListSections,
      word_read_headers_footers: wordReadHeadersFooters,
      word_set_header_footer: wordSetHeaderFooter,
      word_transfer_content: wordTransferContent,
      word_insert_break: wordInsertBreak,
      word_insert_caption: wordInsertCaption,
    },
    selectTargets: {
      replace_all: selectReplaceFirst,
      insert_paragraph: (a) => selectWordInsertPoint(a.location || 'end'),
      insert_table: (a) => selectWordInsertPoint(a.location || 'end'),
      word_insert_equation: (a) => {
        const kind = a.target ? (a.target.kind || 'selection') : (['start', 'end'].includes(a.location) ? 'body' : 'selection');
        if (kind !== 'body') return;
        const mode = a.mode || ({ start: 'prepend', end: 'append' }[a.location] || 'after');
        return selectWordInsertPoint(mode === 'prepend' ? 'start' : mode === 'append' ? 'end' : null);
      },
      word_insert_break: (a) => selectWordInsertPoint(a.location || 'afterSelection'),
      word_insert_caption: () => undefined, // position 的两个合法值均相对当前选区
    },
    // 机制级 re-read：覆盖类操作先预览，ask 模式经窗格内部授权入口执行。
    // 每个覆盖类 action 映射到「读当前状态」的函数，预览结果随 confirm_required 返回给 AI。
    previewReaders: {
      write_selection: () => readSelection({}),
      replace_all: (a) => replaceAll({ ...a, dryRun: true }),
      append_text: () => readSelection({}),
      set_table_borders: (a) => readTableBorders(a),
      set_font: (a) => okResult({ host: 'Word', wouldSetFont: a.font, note: '影响全文所有段落（含表格内文字），属覆盖类操作' }),
      word_delete_paragraphs: () => wordListParagraphs({ filter: 'all' }),
      insert_paragraph: () => readSelection({}),
      insert_table: () => readSelection({}),
      word_insert_html: (a) => (a.target === 'selection' ? readSelection({}) : readDocument({})),
      word_set_page_setup: () => wordListSections({}),
      word_insert_equation: (a) => (a.target && a.target.kind === 'search' ? search({ query: a.target.query, matchCase: a.target.matchCase }) : readSelection({ withStyles: true })),
      word_insert_image: () => readSelection({ withStyles: true }),
      apply_style: () => readSelection({}),
      format_selection: () => readSelection({ withStyles: true }),
      set_paragraph_format: () => readSelection({ withStyles: true }),
      add_comment: () => readSelection({}),
      word_set_header_footer: (a) => wordReadHeadersFooters({ section: a.section, types: [a.type || 'primary'] }),
      word_insert_break: () => wordListSections({}),
      word_insert_caption: () => readSelection({}),
      word_transfer_content: (a) => Promise.all([
        a.source && a.source.kind === 'table' ? readTables({}) : readSelection({}),
        a.destination && ['header', 'footer'].includes(a.destination.kind) ? wordReadHeadersFooters({ section: a.destination.section === undefined ? 0 : a.destination.section, types: [a.destination.type || 'primary'] }) : readDocument({}),
      ]).then(([source, destination]) => source.ok && destination.ok ? okResult({ source: source.result, destination: destination.result }) : (source.ok ? destination : source)),
    },
    afterContent: {
      insert_paragraph: (a) => (a.text !== undefined ? a.text : '(未提供)'),
      append_text: (a) => (a.text !== undefined ? a.text : '(未提供)'),
      insert_table: (a) => JSON.stringify(a.rows || []),
      word_delete_paragraphs: (a) => `删除段落（零基索引）：${(a.targets || []).map((t) => t.index).join(', ')}`,
      word_insert_html: (a) => `插入 HTML（${a.target || 'body'} / ${a.mode || 'append'}，${String(a.html || '').length} 字符）`,
      word_set_page_setup: (a) => `设置页面（节：${a.sections === undefined ? 'all' : JSON.stringify(a.sections)}）`,
      word_insert_image: (a) => `插入图片（${a.mode || 'after'}）`,
      word_transfer_content: (a) => `${a.operation || 'copy'}：${JSON.stringify(a.source || {})} → ${JSON.stringify(a.destination || {})}`,
      set_font: (a) => `全文设为字体：${a.font}`,
      apply_style: (a) => `应用样式：${a.style}（${a.scope || 'selection'}）`,
    },
    // get_environment 的 Word 专属探测：对象模型属性是否真实存在（区分「API 不存在」与「存在但调用失败」）
    environmentProbe: async () => {
      return await Word.run(async (ctx) => {
        const body = ctx.document.body;
        const sel = ctx.document.getSelection();
        return {
          bodyHasComments: !!body.comments,
          selHasParagraphFormat: !!sel.paragraphFormat,
          bodyInsertTableType: typeof body.insertTable,
          bodyInsertOoxmlType: typeof body.insertOoxml,
          selHasFont: !!sel.font,
        };
      });
    },
  };
})();
