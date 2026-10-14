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
        tables.load('items/rowCount,items/values,items/style,items/styleBuiltIn');
        await ctx.sync();
        const out = tables.items.map((t, ti) => {
          const cells = t.values.map(row => row.slice());
          return { index: ti + 1, rowCount: t.rowCount, columnCount: cells.length ? cells[0].length : 0, style: t.style || null, styleBuiltIn: t.styleBuiltIn || null, cells };
        });
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
    const color = args && args.color;
    if (typeof color !== 'string' || !/^#[0-9a-fA-F]{6}$/.test(color)) return errResult('bad_args', 'color required in #RRGGBB format');
    const allLocations = ['Top', 'Bottom', 'Left', 'Right', 'InsideHorizontal', 'InsideVertical'];
    const locations = args.locations === undefined ? allLocations : args.locations;
    if (!Array.isArray(locations) || !locations.length || locations.some((value) => !allLocations.includes(value)) || new Set(locations).size !== locations.length) return errResult('bad_args', 'locations must contain unique valid border locations');
    if (args.tableIndexes !== undefined && (!Array.isArray(args.tableIndexes) || !args.tableIndexes.length || args.tableIndexes.some((value) => !Number.isInteger(value) || value < 1) || new Set(args.tableIndexes).size !== args.tableIndexes.length)) return errResult('bad_args', 'tableIndexes must contain unique positive one-based integers');
    if (args.width !== undefined && (typeof args.width !== 'number' || !Number.isFinite(args.width) || args.width <= 0)) return errResult('bad_args', 'width must be a positive finite number in points');
    if (args.enableGrid !== undefined && typeof args.enableGrid !== 'boolean') return errResult('bad_args', 'enableGrid must be boolean');
    if (!isSetSupported('WordApi', '1.3')) return errResult('requirement', 'set_table_borders requires WordApi 1.3+');
    try {
      return await Word.run(async (ctx) => {
        const tables = ctx.document.body.tables;
        tables.load('items'); await ctx.sync();
        const indexes = args.tableIndexes === undefined ? tables.items.map((_, index) => index + 1) : args.tableIndexes;
        if (!indexes.length) return errResult('not_found', 'document contains no tables');
        if (indexes.some((index) => index > tables.items.length)) return errResult('not_found', 'table index not found');
        const refs = [];
        for (const index of indexes) for (const location of locations) {
          const border = tables.items[index - 1].getBorder(location);
          border.load('color,type,width'); refs.push({ table: index, location, border });
        }
        // Resolve every border before applying a table style or border setter.
        await ctx.sync();
        if (args.enableGrid === true) {
          for (const index of indexes) tables.items[index - 1].styleBuiltIn = 'TableGrid';
          await ctx.sync();
        }
        for (const { border } of refs) {
          border.color = color;
          if (args.width !== undefined) border.width = args.width;
        }
        await ctx.sync();
        for (const { border } of refs) border.load('color,type,width');
        await ctx.sync();
        return okResult({ host: 'Word', tableIndexes: indexes, borderCount: refs.length, color: color.toUpperCase(), width: args.width === undefined ? null : args.width, borders: refs.map(({ table, location, border }) => ({ table, location, color: border.color, type: border.type, width: border.width })), note: 'width is the requested value; borders contains the actual readback. Line type is preserved unless enableGrid applies TableGrid.', ...(args.width === undefined ? {} : { warning: 'Office.js does not apply width to border types with fixed widths; inspect the actual borders readback.' }) });
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
    if (args.paragraphs !== undefined) return errResult('bad_args', 'Use targets, not paragraphs: {targets:[{index:0,expectedText:"exact text"}]}');
    const targets = args.targets;
    if (!Array.isArray(targets) || !targets.length || targets.some((t) => !t || !Number.isInteger(t.index) || t.index < 0 || typeof t.expectedText !== 'string')) return errResult('bad_args', 'targets must be [{index,expectedText}] with zero-based integer indexes and exact string text');
    try { return await Word.run(async (ctx) => {
      const paragraphs = ctx.document.body.paragraphs; paragraphs.load('items/text,items/inlinePictures/items'); await ctx.sync();
      const indexes = targets.map((t) => Number(t.index)); if (new Set(indexes).size !== indexes.length) return errResult('bad_args', 'target indexes must be unique');
      for (const target of targets) {
        const index = target.index, p = paragraphs.items[index];
        if (!p) return errResult('not_found', `paragraph ${index} not found`);
        if (p.text !== target.expectedText) return errResult('not_found', `paragraph ${index} text changed; reread word_list_paragraphs`);
        if (index === paragraphs.items.length - 1) return errResult('bad_args', `paragraph ${index} cannot be safely deleted: last_paragraph`);
        if (p.inlinePictures.items.length) return errResult('bad_args', `paragraph ${index} cannot be safely deleted: contains_inline_picture`);
      }
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
    if (requireWrite && args.text === undefined && args.field === undefined) throw new Error('text or field required');
    if (requireWrite && args.field !== undefined) {
      if (args.text !== undefined) throw new Error('text and field are mutually exclusive');
      normalizeWordField(args.field);
    }
    if (requireWrite && !['replace', 'append', 'prepend'].includes(args.mode || 'replace')) throw new Error('mode must be replace, append, or prepend');
    return { section, types: Array.from(new Set(types)) };
  }

  async function wordReadHeadersFooters(args) {
    if (!isSetSupported('WordApi', '1.1')) return errResult('requirement', 'word_read_headers_footers requires WordApi 1.1+');
    let spec; try { spec = validateHeaderFooterArgs(args || {}, false); } catch (e) { return errResult('bad_args', e.message); }
    const withFields = args.includeFields === true;
    if (withFields && !isSetSupported('WordApi','1.5')) return errResult('requirement','Field readback requires WordApi 1.5');
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
          header.load('text'); footer.load('text');
          const headerFields = withFields ? header.getRange().fields : null, footerFields = withFields ? footer.getRange().fields : null;
          if (withFields && (!headerFields || !footerFields)) return errResult('requirement','Native header/footer fields unavailable');
          if (withFields) { headerFields.load('items'); footerFields.load('items'); }
          refs.push({ index, type, header, footer, headerFields, footerFields });
        }
        await ctx.sync();
        if (withFields) { for (const r of refs) for (const f of [...r.headerFields.items,...r.footerFields.items]) loadNativeField(f); await ctx.sync(); }
        return okResult({ host: 'Word', sectionCount: sections.items.length, items: refs.map((r) => ({ section: r.index, type: r.type, header: r.header.text || '', footer: r.footer.text || '', ...(withFields ? {headerFields:r.headerFields.items.map(nativeFieldSnapshot),footerFields:r.footerFields.items.map(nativeFieldSnapshot)} : {}) })), warning: 'Office.js cannot inspect or disable Link to Previous; first/even body existence does not prove its page-layout switch is enabled.' });
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

  // Create only a blank native document; no binary document upload or arbitrary code.
  async function wordCreateDocument(args) {
    if (Object.keys(args || {}).length) return errResult('bad_args','word_create_document accepts no arguments');
    if (!isSetSupported('WordApi','1.3')) return errResult('requirement','Blank document creation requires WordApi 1.3');
    try { return await Word.run(async ctx => {
      if (!ctx.application || typeof ctx.application.createDocument !== 'function') return errResult('requirement','Application.createDocument unavailable');
      const document = ctx.application.createDocument();
      document.open(); await ctx.sync();
      return okResult({host:'Word',created:true,opened:true,blank:true,requiresNewInstance:true});
    }); } catch (e) { return errResult('execution',String(e.message || e)+(e.debugInfo ? ' [debug] '+JSON.stringify(e.debugInfo) : '')); }
  }

  function normalizeWordField(raw) {
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) throw new Error("field must be {type,switches?}");
    const type = {toc:"TOC",page:"Page",numpages:"NumPages"}[String(raw.type || "").toLowerCase()];
    if (!type) throw new Error("field.type must be TOC|Page|NumPages");
    const switches = raw.switches === undefined ? (type === "TOC" ? String.fromCharCode(92)+"o \"1-3\" "+String.fromCharCode(92)+"h" : "") : raw.switches;
    if (typeof switches !== "string" || switches.length > 1000 || (/[\r\n{}]/.test(switches) || switches.includes(String.fromCharCode(0)))) throw new Error("invalid field switches");
    return {type,switches};
  }
  function nativeFieldRequirement() {
    if (!isSetSupported("WordApi","1.5")) return errResult("requirement","Native fields require WordApi 1.5");
    if (typeof Office !== "undefined" && Office.context.platform === "OfficeOnline") return errResult("unsupported","Native fields require Word desktop");
    return null;
  }
  function loadNativeField(f) { f.load("type,code,result/text"); }
  function nativeFieldSnapshot(f) { return {type:f.type,code:f.code,text:f.result.text || ""}; }
  async function wordReadFields(args) {
    if (!isSetSupported("WordApi","1.5")) return errResult("requirement","Fields require WordApi 1.5");
    const target = args.target || "body";
    if (!["body","selection"].includes(target)) return errResult("bad_args","invalid target");
    try { return await Word.run(async ctx => {
      const range = target === "body" ? ctx.document.body.getRange() : ctx.document.getSelection();
      if (!range.fields) return errResult("requirement","Range.fields unavailable");
      range.load("text"); range.fields.load("items"); await ctx.sync();
      for (const f of range.fields.items) loadNativeField(f); await ctx.sync();
      return okResult({host:"Word",target,text:range.text || "",fields:range.fields.items.map(nativeFieldSnapshot)});
    }); } catch (e) { return errResult("execution",String(e.message || e)); }
  }
  async function wordInsertField(args,previewOnly = false) {
    let spec; try { spec = normalizeWordField(args); } catch(e) { return errResult("bad_args",e.message); }
    const target = args.target || "selection", mode = args.mode || "after";
    if (!["body","selection"].includes(target) || !["replace","prepend","append","before","after"].includes(mode) || (target === "body" && ["before","after"].includes(mode))) return errResult("bad_args","invalid target or mode");
    const anchor=args.paragraph;
    if (anchor !== undefined && (!anchor || !Number.isInteger(anchor.index) || anchor.index < 0 || typeof anchor.expectedText !== "string" || target !== "selection")) return errResult("bad_args","paragraph requires exact zero-based index and expectedText on selection target");
    const unavailable = nativeFieldRequirement(); if (unavailable) return unavailable;
    if (previewOnly) return wordReadFields({target});
    try { return await Word.run(async ctx => {
      let range, paragraph;
      if (anchor) {const paragraphs=ctx.document.body.paragraphs;paragraphs.load("items/text");await ctx.sync();paragraph=paragraphs.items[anchor.index];if(!paragraph || paragraph.text!==anchor.expectedText)return errResult("bad_args","Field paragraph anchor changed");if(typeof paragraph.getRange!=="function")return errResult("requirement","Paragraph.getRange unavailable");range=paragraph.getRange("Whole");}
      else range = target === "body" ? ctx.document.body.getRange() : ctx.document.getSelection();
      if (typeof range.insertField !== "function") return errResult("requirement","Range.insertField unavailable");
      range.load("text"); await ctx.sync();
      if (spec.type === "TOC" && anchor && isSetSupported("WordApiDesktop","1.4") && typeof ctx.document.tablesOfContents?.add === "function") {
        if (args.switches !== undefined) return errResult("bad_args","Paragraph-anchored native TOC uses built-in heading levels 1-3; omit switches");
        if (mode === "replace" && anchor.expectedText) {const content=paragraph.getRange("Content");if(typeof content.insertText !== "function")return errResult("requirement","Paragraph content replacement unavailable");content.insertText("","Replace");await ctx.sync();}
        ctx.document.tablesOfContents.add(paragraph.getRange(["append","after"].includes(mode)?"End":"Start"),{upperHeadingLevel:1,lowerHeadingLevel:3,useBuiltInHeadingStyles:true,includePageNumbers:true,rightAlignPageNumbers:true}); await ctx.sync();
        const fields=ctx.document.body.getRange().fields;fields.load("items");await ctx.sync();for(const field of fields.items)loadNativeField(field);await ctx.sync();
        const field=fields.items.find(f=>/^TOC\b/i.test(f.code.trim()));if(!field)return errResult("execution","Native tablesOfContents.add did not expose a TOC field");
        return okResult({host:"Word",nativeField:true,nativeTocCollection:true,target,mode,field:nativeFieldSnapshot(field)});
      }
      const location={replace:"Replace",prepend:"Start",append:"End",before:"Before",after:"After"}[mode];
      const f=spec.type === "TOC" && anchor ? range.insertField(location,spec.type) : range.insertField(location,spec.type,spec.switches,true);await ctx.sync();
      if (typeof f.updateResult === "function") { f.updateResult(); await ctx.sync(); }
      loadNativeField(f); await ctx.sync(); return okResult({host:"Word",nativeField:true,target,mode,field:nativeFieldSnapshot(f)});
    }); } catch(e) { return errResult("execution",String(e.message || e)+(e.debugInfo ? "; "+JSON.stringify(e.debugInfo) : "")); }
  }

  // Global definitions; use exact localized names from word_read_styles.
  function scriptFontFields(font) { return isSetSupported('WordApiDesktop','1.3') ? ['nameAscii','nameFarEast','nameOther','nameBidirectional'].filter(key => key in font) : []; }
  function scriptFontSnapshot(font) { return Object.fromEntries(scriptFontFields(font).map(key => [key,font[key]])); }
  function normalizedAlignment(value) {
    const key = String(value ?? '').toLowerCase();
    return {left:'left',center:'center',centered:'center',right:'right',justify:'justify',justified:'justify'}[key] || value;
  }
  // Exact document names win. Alias matching uses the same built-in map as paragraph formatting.
  async function resolveParagraphStyles(ctx, names) {
    if (typeof ctx.document.getStyles !== 'function') throw Object.assign(new Error('Document.getStyles unavailable'), {code:'requirement'});
    const styles = ctx.document.getStyles();
    if (typeof styles.getByNameOrNullObject !== 'function') throw Object.assign(new Error('StyleCollection.getByNameOrNullObject unavailable'), {code:'requirement'});
    const refs = names.map(name => styles.getByNameOrNullObject(name));
    for (const ref of refs) ref.load('isNullObject,type,nameLocal');
    await ctx.sync();
    if (refs.some(ref => ref.isNullObject)) {
      styles.load('items/nameLocal,items/type,items/builtIn'); await ctx.sync();
      for (let i = 0; i < refs.length; i++) {
        if (!refs[i].isNullObject) continue;
        const alias = mapBuiltInStyle?.(names[i]);
        const matches = alias ? styles.items.filter(style => style.builtIn && mapBuiltInStyle?.(style.nameLocal) === alias) : [];
        if (matches.length === 1) refs[i] = matches[0];
        else throw Object.assign(new Error('style not found: ' + names[i] + '; available paragraph styles: ' + JSON.stringify(styles.items.filter(s => s.type === 'Paragraph').map(s => s.nameLocal))), {code:'not_found'});
      }
    }
    if (refs.some(ref => ref.type !== 'Paragraph')) throw Object.assign(new Error('only paragraph styles supported'), {code:'bad_args'});
    return refs;
  }
  function stylePatchVerification(target, snapshot) {
    const fields = [];
    const compare = (field, requested, actual) => {
      const normalize = value => field.endsWith('.color') && typeof value === 'string' ? value.toUpperCase() : field === 'paragraphFormat.alignment' ? normalizedAlignment(value) : value;
      const status = actual === undefined || actual === null ? 'unverified' : normalize(requested) === normalize(actual) ? 'applied' : 'mismatch';
      fields.push({field,requested,actual,status});
    };
    for (const [key,value] of Object.entries(target.patch.font || {})) compare('font.' + key, value, snapshot.font[key]);
    if (target.patch.font?.name !== undefined) for (const key of Object.keys(snapshot.font).filter(k => k.startsWith('name') && k !== 'name')) compare('font.' + key, target.patch.font.name, snapshot.font[key]);
    for (const key of paragraphFields) if (target.patch[key] !== undefined) compare('paragraphFormat.' + key, target.patch[key], snapshot.paragraphFormat[key]);
    return {requestedName:target.name,name:snapshot.name,fields};
  }
  function styleSnapshot(style) {
    return {name:style.nameLocal,type:style.type,builtIn:style.builtIn,font:{name:style.font.name,size:style.font.size,color:style.font.color,bold:style.font.bold,italic:style.font.italic,...scriptFontSnapshot(style.font)},paragraphFormat:Object.fromEntries(paragraphFields.map(k => [k,k === 'alignment' ? normalizedAlignment(style.paragraphFormat[k]) : style.paragraphFormat[k]]))};
  }
  function loadStyleSnapshot(style) { style.load("nameLocal,type,builtIn"); style.font.load(['name','size','color','bold','italic',...scriptFontFields(style.font)].join(',')); style.paragraphFormat.load(paragraphFields.join(",")); }
  async function wordReadStyles(args) {
    if (!isSetSupported("WordApi","1.5")) return errResult("requirement","Styles require WordApi 1.5");
    const names = args.names;
    if (names !== undefined && (!Array.isArray(names) || !names.length || names.some(n => typeof n !== "string" || !n.trim()) || new Set(names).size !== names.length)) return errResult("bad_args","names must be unique nonempty names");
    try { return await Word.run(async ctx => {
      if (typeof ctx.document.getStyles !== "function") return errResult("requirement","Document.getStyles unavailable");
      const styles = ctx.document.getStyles();
      if (typeof styles.getByNameOrNullObject !== "function") return errResult("requirement","StyleCollection.getByNameOrNullObject unavailable");
      let refs;
      if (names) refs = await resolveParagraphStyles(ctx, names);
      else { styles.load("items/nameLocal,items/type,items/builtIn"); await ctx.sync(); refs = styles.items.filter(s => s.type === "Paragraph"); }
      if (refs.some(r => !r.font || !r.paragraphFormat)) return errResult("requirement","Style font/paragraphFormat unavailable");
      for (const ref of refs) loadStyleSnapshot(ref); await ctx.sync();
      return okResult({host:"Word",styles:refs.map(styleSnapshot)});
    }); } catch(e) { return errResult(['not_found','bad_args','requirement'].includes(e.code) ? e.code : 'execution',String(e.message || e)); }
  }
  async function wordCreateBodyStyle(previewOnly) {
    if (!isSetSupported("WordApi","1.5")) return errResult("requirement","Global body preset requires WordApi 1.5");
    try {
      const before=await Word.run(async ctx=>{
        if (typeof ctx.document.getStyles !== "function" || typeof ctx.document.body.insertOoxml !== "function" || typeof ctx.document.getSelection().delete !== "function") return {error:errResult("requirement","Native style/insertOoxml/delete unavailable")};
        const collection=ctx.document.getStyles(); if (typeof collection.getByNameOrNullObject !== "function") return {error:errResult("requirement","Native style lookup unavailable")};
        const ref=collection.getByNameOrNullObject("范本正文"); ref.load("isNullObject"); await ctx.sync();
        if (!ref.isNullObject) {loadStyleSnapshot(ref); await ctx.sync(); return {existing:styleSnapshot(ref)};}
        const paragraphs=ctx.document.body.paragraphs; paragraphs.load("items/text,items/styleBuiltIn"); await ctx.sync();
        const paragraph=paragraphs.items.find(p=>p.styleBuiltIn === "Normal" && p.text.trim());
        return {target:paragraph?.text};
      });
      if (before.error) return before.error;
      if (before.existing) return okResult({host:"Word",global:true,preset:"body-1.5",created:false,styles:[before.existing],requiresNativeExportVerification:true});
      if (previewOnly) return okResult({host:"Word",global:true,preset:"body-1.5",name:"范本正文",exists:false,bodyAnchor:before.target});
      if (!before.target) return errResult("bad_args","Add a body paragraph before creating the reusable body preset");
      const created=await Word.run(async ctx=>{
        const ooxml="<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\"><pkg:part pkg:name=\"/_rels/.rels\" pkg:contentType=\"application/vnd.openxmlformats-package.relationships+xml\"><pkg:xmlData><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name=\"/word/document.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"><pkg:xmlData><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p/><w:p><w:pPr><w:pStyle w:val=\"DSHAtlasBody\"/></w:pPr><w:r><w:t>DSH_NATIVE_BODY_STYLE_ANCHOR</w:t></w:r></w:p><w:p/></w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name=\"/word/_rels/document.xml.rels\" pkg:contentType=\"application/vnd.openxmlformats-package.relationships+xml\"><pkg:xmlData><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rIdStyles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name=\"/word/styles.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"><pkg:xmlData><w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:style w:type=\"paragraph\" w:customStyle=\"1\" w:styleId=\"DSHAtlasBody\"><w:name w:val=\"范本正文\"/><w:next w:val=\"DSHAtlasBody\"/><w:qFormat/><w:pPr><w:spacing w:before=\"0\" w:after=\"0\" w:line=\"360\" w:lineRule=\"auto\"/><w:ind w:firstLine=\"480\"/><w:jc w:val=\"both\"/></w:pPr><w:rPr><w:rFonts w:ascii=\"宋体\" w:hAnsi=\"宋体\" w:eastAsia=\"宋体\" w:cs=\"宋体\"/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/><w:color w:val=\"000000\"/></w:rPr></w:style></w:styles></pkg:xmlData></pkg:part></pkg:package>";
        const inserted=ctx.document.body.insertOoxml(ooxml,"End"); await ctx.sync();
        const ref=ctx.document.getStyles().getByNameOrNullObject("范本正文"); ref.load("isNullObject"); await ctx.sync();
        if (ref.isNullObject) { inserted.delete(); await ctx.sync(); return errResult("execution","Native style import did not register the style before anchor cleanup"); }
        const paragraphs=ctx.document.body.paragraphs; paragraphs.load("items/text"); await ctx.sync();
        const target=paragraphs.items.find(p=>p.text===before.target);
        if (!target) {inserted.delete(); await ctx.sync(); return errResult("execution","Body anchor changed during style import");}
        target.style="范本正文"; await ctx.sync();
        inserted.delete(); await ctx.sync();
        target.styleBuiltIn="Normal"; await ctx.sync();
        loadStyleSnapshot(ref); await ctx.sync();
        return okResult({host:"Word",global:true,preset:"body-1.5",created:true,styles:[styleSnapshot(ref)],temporaryBodyAnchorRestored:true,requiresNativeExportVerification:true});
      });
      if (!created.ok) return created;
      const actual=await wordReadStyles({names:["范本正文"]});
      return actual.ok ? okResult({...created.result,styles:actual.result.styles}) : actual;
    } catch(e) { return errResult("execution",String(e.message || e)); }
  }
  async function wordRenderPages(args) {
    if (!Number.isInteger(args.expectedPageCount) || args.expectedPageCount < 1 || args.expectedPageCount > 10) return errResult("bad_args","expectedPageCount must be an integer 1..10");
    if (!isSetSupported("WordApi","1.1")) return errResult("requirement","Word body text requires WordApi 1.1");
    try {
      const bodyText = await Word.run(async ctx => { const body=ctx.document.body; body.load("text"); await ctx.sync(); return body.text; });
      const response = await fetch("/office/native-word-pages",{method:"POST",headers:{"Content-Type":"application/json; charset=utf-8"},body:JSON.stringify({bodyText,expectedPageCount:args.expectedPageCount})});
      const out = await response.json();
      if (!response.ok || out.ok === false) return errResult(out.code || "execution",out.error || "Native Word page rendering failed");
      return out.ok === true ? out : okResult(out);
    } catch(e) { return errResult("execution",String(e.message || e)); }
  }
  async function wordSetStyles(args,previewOnly = false) {
    if (Array.isArray(args.styles) && args.styles.some(raw => raw && raw.preset !== undefined)) {
      if (args.styles.length !== 1 || args.styles[0].name !== "范本正文" || args.styles[0].preset !== "body-1.5" || Object.keys(args.styles[0]).some(key => !["name","preset"].includes(key))) return errResult("bad_args","Only the bounded 范本正文/body-1.5 preset is supported");
      return wordCreateBodyStyle(previewOnly);
    }
    let targets;
    try {
      if (!Array.isArray(args.styles) || !args.styles.length || args.styles.length > 100) throw new Error("styles must contain 1..100 patches");
      const seen = new Set();
      targets = args.styles.map(raw => {
        if (!raw || typeof raw.name !== 'string' || !raw.name.trim() || seen.has(raw.name)) throw new Error('each style requires a unique nonempty name and font/paragraphFormat patch');
        const unknown = Object.keys(raw).find(k => !['name','font','paragraphFormat'].includes(k));
        if (unknown) throw new Error('unsupported style field: ' + unknown + '; allowed: name, font, paragraphFormat; type is read-only');
        seen.add(raw.name);
        if (raw.paragraphFormat !== undefined && (!raw.paragraphFormat || typeof raw.paragraphFormat !== "object" || Array.isArray(raw.paragraphFormat) || Object.keys(raw.paragraphFormat).some(k => !paragraphFields.includes(k)))) throw new Error("unsupported paragraphFormat field; lineSpacing is points");
        return {name:raw.name,patch:normalizeParagraphPatch({...raw.paragraphFormat,...(raw.font === undefined ? {} : {font:raw.font})})};
      });
    } catch(e) { return errResult("bad_args",e.message); }
    if (!isSetSupported("WordApi","1.5")) return errResult("requirement","Styles require WordApi 1.5");
    try { return await Word.run(async ctx => {
      const refs = await resolveParagraphStyles(ctx, targets.map(t => t.name));
      if (new Set(refs.map(ref => ref.nameLocal)).size !== refs.length) return errResult('bad_args','style names resolve to the same definition; use unique styles');
      for (const ref of refs) {
        if (!ref.font || !ref.paragraphFormat) return errResult("requirement","Style font/paragraphFormat unavailable");
        if (targets[refs.indexOf(ref)].patch.font?.name !== undefined && isSetSupported('WordApiDesktop','1.3') && !('nameFarEast' in ref.font)) return errResult('requirement','Native Far East style font API unavailable');
        loadStyleSnapshot(ref);
      }
      await ctx.sync();
      if (previewOnly) return okResult({host:"Word",styles:refs.map(styleSnapshot)});
      refs.forEach((ref,i) => {
        for (const [key,value] of Object.entries(targets[i].patch.font || {})) ref.font[key] = value;
        if (targets[i].patch.font?.name !== undefined) for (const key of scriptFontFields(ref.font)) ref.font[key] = targets[i].patch.font.name;
        for (const key of paragraphFields) if (targets[i].patch[key] !== undefined) ref.paragraphFormat[key] = targets[i].patch[key];
      });
      await ctx.sync(); for (const ref of refs) loadStyleSnapshot(ref); await ctx.sync();
      const snapshots = refs.map(styleSnapshot);
      const verification = targets.map((target, i) => stylePatchVerification(target, snapshots[i]));
      const verificationWarnings = verification.flatMap(style => style.fields.filter(field => field.status !== 'applied').map(field => ({name:style.name,...field})));
      return okResult({host:'Word',global:true,updated:refs.length,styles:snapshots,verification,verified:verificationWarnings.length === 0,...(verificationWarnings.length ? {verificationWarnings} : {})});
    }); } catch(e) { return errResult(e.code === 'not_found' || e.code === 'bad_args' || e.code === 'requirement' ? e.code : 'execution',String(e.message || e)); }
  }

  async function wordSetHeaderFooter(args) {
    if (!isSetSupported('WordApi', '1.1')) return errResult('requirement', 'word_set_header_footer requires WordApi 1.1+');
    let spec; try { spec = validateHeaderFooterArgs(args || {}, true); } catch (e) { return errResult('bad_args', e.message); }
    const fieldSpec = args.field === undefined ? null : normalizeWordField(args.field);
    if (fieldSpec) { const unavailable = nativeFieldRequirement(); if (unavailable) return unavailable; }
    try {
      return await Word.run(async (ctx) => {
        const sections = ctx.document.sections; sections.load('items'); await ctx.sync();
        const indexes = spec.section === 'all' ? sections.items.map((_, i) => i) : [Number(spec.section)];
        if (indexes.some((i) => i >= sections.items.length)) return errResult('bad_args', `section index out of range (count ${sections.items.length})`);
        const type = spec.types[0]; const refs = [];
        for (const index of indexes) {
          const body = args.kind === 'header' ? sections.items[index].getHeader(WORD_HEADER_FOOTER_TYPES[type]) : sections.items[index].getFooter(WORD_HEADER_FOOTER_TYPES[type]);
          const location = (args.mode || 'replace') === 'prepend' ? 'Start' : (args.mode || 'replace') === 'append' ? 'End' : 'Replace';
          const range = fieldSpec ? body.getRange() : null;
          if (fieldSpec && typeof range.insertField !== 'function') return errResult('requirement','Range.insertField unavailable');
          if (fieldSpec && (!range.fields || typeof range.fields.load !== 'function' || typeof body.insertText !== 'function')) return errResult('requirement','Native footer field/text APIs unavailable');
          refs.push({ index, body, range, location });
        }
        // Read every existing field before deciding which native targets need replacement.
        if (fieldSpec) {
          for (const ref of refs) {ref.body.load("text");ref.range.fields.load("items");} await ctx.sync();
          for (const ref of refs) for (const field of ref.range.fields.items) loadNativeField(field); await ctx.sync();
          const desired=(fieldSpec.type.toUpperCase()+" "+fieldSpec.switches).replace(/\s+/g," ").trim();
          for (const ref of refs) {
            const current=ref.range.fields.items;
            if (ref.location === "Replace" && current.length === 1 && current[0].type === fieldSpec.type && current[0].code.replace(/\s+/g," ").trim() === desired && ref.body.text.trim() === current[0].result.text.trim()) ref.field=current[0];
          }
        }
        for (const ref of refs) {
          if (fieldSpec && !ref.field) {
            if (ref.location === "Replace") {ref.body.insertText("","Replace");await ctx.sync();}
            ref.field=ref.body.getRange().insertField(ref.location === "Replace" ? "Start" : ref.location,fieldSpec.type,fieldSpec.switches,true);await ctx.sync();
            if (!["Page","NumPages"].includes(fieldSpec.type) && typeof ref.field.updateResult === "function") {ref.field.updateResult();await ctx.sync();}
          } else if (!fieldSpec) ref.body.insertText(String(args.text),ref.location);
        }
        await ctx.sync();
        for (const ref of refs) {ref.body.load("text");if(fieldSpec){ref.actualFields=ref.body.getRange().fields;ref.actualFields.load("items");}} await ctx.sync();
        for (const ref of refs) if (fieldSpec) {ref.field=ref.actualFields.items[0];if(ref.field)loadNativeField(ref.field);} await ctx.sync();

        return okResult({ host: 'Word', kind: args.kind, type, sections: refs.map((r) => ({ index: r.index, text: r.body.text || '', ...(r.field ? {field:nativeFieldSnapshot(r.field)} : {}) })), warning: 'Header/footer may be linked to a previous section; the portable Office.js API cannot unlink it.' });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)+(e.debugInfo ? '; '+JSON.stringify(e.debugInfo) : '')); }
  }

  async function wordInsertBreak(args) {
    if (!isSetSupported('WordApi', '1.1')) return errResult('requirement', 'word_insert_break requires WordApi 1.1+');
    const type = String(args.type || ''); const location = String(args.location || 'afterSelection');
    if (!WORD_BREAK_TYPES[type]) return errResult('bad_args', `type must be ${Object.keys(WORD_BREAK_TYPES).join('|')}`);
    if (!['start', 'end', 'beforeSelection', 'afterSelection'].includes(location)) return errResult('bad_args', 'invalid location');
    const anchor = args.paragraph;
    if (anchor !== undefined && (!anchor || !Number.isInteger(anchor.index) || anchor.index < 0 || typeof anchor.expectedText !== 'string' || !['beforeSelection','afterSelection'].includes(location))) return errResult('bad_args','paragraph anchor requires nonnegative index and exact expectedText with beforeSelection|afterSelection');
    if (anchor && !isSetSupported('WordApi','1.3')) return errResult('requirement','Exact paragraph break anchor requires WordApi 1.3');
    try {
      return await Word.run(async (ctx) => {
        const sections = ctx.document.sections; sections.load('items'); await ctx.sync(); const before = sections.items.length;
        let destination;
        if (anchor) {
          const paragraphs = ctx.document.body.paragraphs; paragraphs.load('items/text'); await ctx.sync();
          const paragraph = paragraphs.items[anchor.index];
          if (!paragraph || paragraph.text !== anchor.expectedText) return errResult('bad_args','Paragraph break anchor changed; reread paragraphs');
          if (typeof paragraph.getRange !== 'function') return errResult('requirement','Paragraph.getRange unavailable');
          destination = paragraph.getRange();
          if (typeof destination.insertBreak !== 'function') return errResult('requirement','Range.insertBreak unavailable');
        }
        if (location === 'start' || location === 'end') ctx.document.body.insertBreak(WORD_BREAK_TYPES[type], location === 'start' ? 'Start' : 'End');
        else (destination || ctx.document.getSelection()).insertBreak(WORD_BREAK_TYPES[type], location === 'beforeSelection' ? 'Before' : 'After');
        await ctx.sync(); sections.load('items'); await ctx.sync();
        return okResult({ host: 'Word', type, location, ...(anchor ? {paragraph:anchor} : {}), sectionCountBefore: before, sectionCountAfter: sections.items.length });
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
    if (!args || typeof args.html !== 'string') return errResult('bad_args', 'html must be a string; example: {html:"<p>Text</p>",target:"body",mode:"append"}');
    if (Object.keys(args).some(key => !['html','target','mode','confirm'].includes(key))) return errResult('bad_args', 'word_insert_html accepts only html, target, mode');
    if (!isSetSupported('WordApi', '1.1')) return errResult('requirement', 'word_insert_html requires WordApi 1.1+');
    const html = args.html, target = args.target || 'body', mode = args.mode || 'append';
    if (!html || html.length > 8 * 1024 * 1024 || !['body', 'selection'].includes(target) || !['replace', 'append', 'prepend', 'before', 'after'].includes(mode)) return errResult('bad_args', 'html, target body|selection, and valid mode required');
    try { return await Word.run(async (ctx) => {
      const range = target === 'body' ? ctx.document.body : ctx.document.getSelection(); range.load('text'); await ctx.sync(); const previousText = range.text || '';
      const locations = { replace: Word.InsertLocation.replace, append: Word.InsertLocation.end, prepend: Word.InsertLocation.start, before: Word.InsertLocation.before, after: Word.InsertLocation.after };
      if (typeof range.insertHtml !== 'function' || (target === 'body' && mode === 'replace' && typeof range.clear !== 'function')) return errResult('requirement','Native HTML insertion/clear API unavailable');
      // Word's documented setup clears complex content before inserting a fragment.
      // Body.insertHtml(Replace) can throw GeneralException on existing TOC/table bodies.
      if (target === 'body' && mode === 'replace') { range.clear(); range.insertHtml(html, Word.InsertLocation.end); }
      else range.insertHtml(html, locations[mode]);
      await ctx.sync();
      return okResult({ host: 'Word', target, mode, previousText, htmlCharacters: html.length });
    }); } catch (e) { return errResult('execution', String(e && e.message || e) + (e && e.code ? ' [code] '+e.code : '') + (e && e.debugInfo ? ' [debug] '+JSON.stringify(e.debugInfo) : '')); }
  }

  async function wordSetPageSetup(args, previewOnly = false) {
    if (!isSetSupported('WordApiDesktop','1.3')) return errResult('requirement','Word page setup requires WordApiDesktop 1.3');
    const pointFields = ['topMargin','bottomMargin','leftMargin','rightMargin','headerDistance','footerDistance','pageWidth','pageHeight'];
    const fields = [...pointFields,'differentFirstPageHeaderFooter'];
    const values = {};
    try {
      for (const key of pointFields) if (args[key] !== undefined) {
        const value = args[key];
        if (typeof value !== 'number' || !Number.isFinite(value) || value < 0 || value > 1584 || ((key === 'pageWidth' || key === 'pageHeight') && value === 0)) throw new Error(key + ' must be a finite nonnegative point value <=1584');
        values[key] = value;
      }
      if (args.differentFirstPageHeaderFooter !== undefined) {
        if (typeof args.differentFirstPageHeaderFooter !== 'boolean') throw new Error('differentFirstPageHeaderFooter must be boolean');
        values.differentFirstPageHeaderFooter = args.differentFirstPageHeaderFooter;
      }
    } catch (e) { return errResult('bad_args',e.message); }
    if (!Object.keys(values).length) return errResult('bad_args','at least one page setup field is required');
    const requested = args.sections === undefined || args.sections === 'all' ? 'all' : args.sections;
    if (requested !== 'all' && (!Array.isArray(requested) || !requested.length || requested.some(i => !Number.isInteger(i) || i < 0))) return errResult('bad_args','sections must be all or a nonempty array of zero-based indexes');
    try { return await Word.run(async ctx => {
      const sections = ctx.document.sections; sections.load('items'); await ctx.sync();
      const indexes = requested === 'all' ? sections.items.map((_,i) => i) : [...new Set(requested)];
      if (indexes.some(i => i >= sections.items.length)) return errResult('not_found','section index not found');
      if (indexes.some(i => !sections.items[i].pageSetup)) return errResult('requirement','Section.pageSetup unavailable');
      for (const index of indexes) sections.items[index].pageSetup.load(fields.join(','));
      await ctx.sync();
      const snapshot = () => indexes.map(index => ({index,values:Object.fromEntries(fields.map(k => [k,sections.items[index].pageSetup[k]]))}));
      if (previewOnly) return okResult({host:'Word',sections:snapshot()});
      for (const index of indexes) Object.assign(sections.items[index].pageSetup,values);
      await ctx.sync();
      for (const index of indexes) sections.items[index].pageSetup.load(fields.join(','));
      await ctx.sync();
      return okResult({host:'Word',sections:indexes,values,pageSetup:snapshot()});
    }); } catch (e) { return errResult('execution',String(e && e.message || e)); }
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

  const paragraphFields = ['alignment', 'spaceBefore', 'spaceAfter', 'lineSpacing', 'leftIndent', 'rightIndent', 'firstLineIndent'];
  const paragraphLoad = 'text,style,styleBuiltIn,alignment,spaceBefore,spaceAfter,lineSpacing,leftIndent,rightIndent,firstLineIndent';

  function normalizeParagraphPatch(raw) {
    const patch = {};
    if (raw.style !== undefined) {
      if (typeof raw.style !== 'string' || !raw.style.trim()) throw new Error('style must be a nonempty document style name or built-in alias');
      const style = mapBuiltInStyle(raw.style);
      if (style && Word.BuiltInStyleName?.[style]) patch.styleBuiltIn = Word.BuiltInStyleName[style];
      patch.styleName = raw.style;
    }
    if (raw.alignment !== undefined) {
      const align = mapAlignment(normalizedAlignment(raw.alignment));
      if (!align) throw new Error('alignment must be left|center|right|justify');
      patch.alignment = align;
    }
    for (const key of paragraphFields.slice(1)) {
      if (raw[key] === undefined) continue;
      const value = raw[key];
      const minimum = ['leftIndent', 'rightIndent', 'firstLineIndent'].includes(key) ? -1584 : 0;
      if (typeof value !== 'number' || !Number.isFinite(value) || value < minimum || value > 1584 || (key === 'lineSpacing' && value === 0)) throw new Error(key + ' must be a finite point value in range');
      patch[key] = value;
    }
    if (raw.font !== undefined) {
      if (!raw.font || typeof raw.font !== 'object' || Array.isArray(raw.font)) throw new Error('font must be an object');
      patch.font = {};
      for (const [key,value] of Object.entries(raw.font)) {
        if (!['name','size','color','bold','italic'].includes(key)) throw new Error('unsupported font field: ' + key);
        if (key === 'name' && (typeof value !== 'string' || !value.trim() || value.length > 128)) throw new Error('font.name must be a nonempty name');
        if (key === 'size' && (typeof value !== 'number' || !Number.isFinite(value) || value < 1 || value > 1638)) throw new Error('font.size must be 1..1638 points');
        if (key === 'color' && (typeof value !== 'string' || !/^#[0-9a-f]{6}$/i.test(value))) throw new Error('font.color must be #RRGGBB');
        if (['bold','italic'].includes(key) && typeof value !== 'boolean') throw new Error('font.' + key + ' must be boolean');
        patch.font[key] = value;
      }
      if (!Object.keys(patch.font).length) throw new Error('font cannot be empty');
    }
    if (!Object.keys(patch).length) throw new Error('at least one paragraph format field is required');
    return patch;
  }

  function paragraphSnapshot(p, index, withStyle = true) {
    return {index,text:p.text,...(p.style === undefined ? {} : {styleName:p.style}),...(withStyle ? {styleBuiltIn:p.styleBuiltIn} : {styleRead:false}), ...Object.fromEntries(paragraphFields.map(key => [key,key === 'alignment' ? normalizedAlignment(p[key]) : p[key]])),
      font:{name:p.font.name,size:p.font.size,color:p.font.color,bold:p.font.bold,italic:p.font.italic}};
  }

  function loadParagraphSnapshot(p, withStyle = true) {
    p.load(withStyle ? paragraphLoad : paragraphLoad.replace('styleBuiltIn,',''));
    p.font.load('name,size,color,bold,italic');
  }

  function writeParagraphPatch(p, patch) {
    // Apply semantic style first, then explicit overrides; absent fields stay untouched.
    if (patch.styleName !== undefined) p.style = patch.styleName;
    if (patch.styleBuiltIn !== undefined) p.styleBuiltIn = patch.styleBuiltIn;
    for (const key of paragraphFields) if (patch[key] !== undefined) p[key] = patch[key];
    for (const [key,value] of Object.entries(patch.font || {})) p.font[key] = value;
  }

  async function validateNamedParagraphStyles(ctx, patches) {
    const names = [...new Set(patches.map(p => p.styleName).filter(Boolean))];
    if (!names.length) return null;
    if (!isSetSupported('WordApi','1.5') || typeof ctx.document.getStyles !== 'function') {
      if (patches.some(patch => patch.styleName && !patch.styleBuiltIn)) return errResult('requirement','Document style names require WordApi 1.5');
      for (const patch of patches) delete patch.styleName;
      return null;
    }
    try {
      const refs = await resolveParagraphStyles(ctx, names);
      for (const patch of patches) if (patch.styleName) {
        patch.styleName = refs[names.indexOf(patch.styleName)].nameLocal;
        delete patch.styleBuiltIn;
      }
      return null;
    } catch (e) { return errResult(e.code || 'bad_args', e.message); }
  }
  async function wordFormatParagraphs(args, previewOnly = false) {
    let targets;
    try {
      if (args.style !== undefined) throw new Error('style must be inside each target: {targets:[{index:0,expectedText:"exact text",style:"Normal"}]}');
      const unknown = Object.keys(args).find(key => !['targets','confirm'].includes(key));
      if (unknown) throw new Error('unsupported word_format_paragraphs argument: ' + unknown + '; format fields belong inside targets[]');
      if (!Array.isArray(args.targets) || !args.targets.length || args.targets.length > 500) throw new Error('targets must contain 1..500 explicit paragraphs');
      const seen = new Set();
      targets = args.targets.map(t => {
        if (!t || typeof t !== 'object' || !Number.isInteger(t.index) || t.index < 0 || typeof t.expectedText !== 'string' || seen.has(t.index)) throw new Error('each target needs a unique zero-based index and exact expectedText');
        if (Object.keys(t).some(k => !['index','expectedText','style','font',...paragraphFields].includes(k))) throw new Error('unsupported paragraph target field');
        seen.add(t.index);
        return {index:t.index,expectedText:t.expectedText,patch:normalizeParagraphPatch(t)};
      });
    } catch (e) { return errResult('bad_args', e.message); }
    if (!isSetSupported('WordApi','1.3')) return errResult('requirement','word_format_paragraphs requires WordApi 1.3');
    try {
      return await Word.run(async ctx => {
        const unavailable = await validateNamedParagraphStyles(ctx, targets.map(t => t.patch)); if (unavailable) return unavailable;
        const paragraphs = ctx.document.body.paragraphs;
        paragraphs.load('items/text');
        await ctx.sync();
        // Validate the whole target set before queuing any writes.
        for (const t of targets) {
          const p = paragraphs.items[t.index];
          if (!p || p.text !== t.expectedText) return errResult('bad_args', 'paragraph changed or missing at index ' + t.index + '; reread paragraphs before editing');
          loadParagraphSnapshot(p);
        }
        await ctx.sync();
        const previous = targets.map(t => paragraphSnapshot(paragraphs.items[t.index],t.index));
        if (previewOnly) return okResult({host:'Word',paragraphs:previous});
        for (const t of targets) writeParagraphPatch(paragraphs.items[t.index], t.patch);
        await ctx.sync();
        for (const t of targets) loadParagraphSnapshot(paragraphs.items[t.index]);
        await ctx.sync();
        return okResult({host:'Word',formatted:targets.length,paragraphs:targets.map(t => paragraphSnapshot(paragraphs.items[t.index],t.index))});
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function setParagraphFormat(args) {
    let patch;
    const listType = args.listType || null;
    try {
      if (listType && !['none','numbered','bulleted'].includes(listType)) throw new Error('listType must be none|numbered|bulleted');
      const raw = {...args};
      if (args.indent !== undefined) {
        if (typeof args.indent !== 'number' || !Number.isFinite(args.indent)) throw new Error('indent must be a finite number in cm');
        raw.leftIndent = args.indent * 72 / 2.54;
      }
      patch = Object.keys(raw).some(k => ['style','font',...paragraphFields].includes(k)) ? normalizeParagraphPatch(raw) : {};
      if (!Object.keys(patch).length && !listType) throw new Error('at least one paragraph setting is required');
    } catch (e) { return errResult('bad_args', e.message); }
    if (patch.styleBuiltIn !== undefined && !isSetSupported('WordApi','1.3')) return errResult('requirement','WordApi 1.3 is required for built-in paragraph styles');
    if (listType && !isSetSupported('WordApi','1.3')) return errResult('requirement','WordApi 1.3 is required for list formatting');
    try {
      return await Word.run(async ctx => {
        const unavailable = await validateNamedParagraphStyles(ctx, [patch]); if (unavailable) return unavailable;
        const paras = ctx.document.getSelection().paragraphs;
        paras.load(listType ? 'items/isListItem' : 'items');
        await ctx.sync();
        if (listType && paras.items.some(p => typeof p.detachFromList !== 'function' || typeof p.startNewList !== 'function' || typeof p.attachToList !== 'function')) return errResult('requirement','Word paragraph list API unavailable');
        if (listType && listType !== 'none' && paras.items.some(p => p.isListItem)) return errResult('bad_args','Selection already contains list items; explicitly remove existing list formatting first');
        for (const p of paras.items) writeParagraphPatch(p,patch);
        if (listType === 'none') {
          for (const p of paras.items) if (p.isListItem) p.detachFromList();
        } else if (listType && paras.items.length) {
          const list = paras.items[0].startNewList();
          if (listType === 'numbered') list.setLevelNumbering(0,Word.ListNumbering.arabic);
          else list.setLevelBullet(0,Word.ListBullet.solid);
          list.load('id');
          await ctx.sync();
          for (const p of paras.items.slice(1)) p.attachToList(list.id,0);
        }
        await ctx.sync();
        const withStyle = isSetSupported('WordApi','1.3');
        for (const p of paras.items) loadParagraphSnapshot(p,withStyle);
        await ctx.sync();
        return okResult({host:'Word',formatted:paras.items.length,listType,paragraphs:paras.items.map((p,i) => paragraphSnapshot(p,i,withStyle))});
      });
    } catch (e) { return errResult('execution',String(e && e.message || e)); }
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
      word_format_paragraphs: wordFormatParagraphs,
      word_render_pages: wordRenderPages,
      word_create_document: wordCreateDocument,
      word_set_styles: wordSetStyles,
      word_read_styles: wordReadStyles,
      word_insert_field: wordInsertField,
      word_read_fields: wordReadFields,
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
      word_insert_field: (a) => a.target === 'body' ? selectWordInsertPoint(a.mode === 'prepend' ? 'start' : a.mode === 'append' ? 'end' : null) : undefined,
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
      word_format_paragraphs: (a) => wordFormatParagraphs(a,true),
      word_create_document: () => readDocument({}),
      word_set_styles: (a) => wordSetStyles(a,true),
      word_insert_field: (a) => wordInsertField(a,true),
      insert_paragraph: () => readSelection({}),
      insert_table: () => readSelection({}),
      word_insert_html: (a) => (a.target === 'selection' ? readSelection({}) : readDocument({})),
      word_set_page_setup: (a) => wordSetPageSetup(a,true),
      word_insert_equation: (a) => (a.target && a.target.kind === 'search' ? search({ query: a.target.query, matchCase: a.target.matchCase }) : readSelection({ withStyles: true })),
      word_insert_image: () => readSelection({ withStyles: true }),
      apply_style: () => readSelection({}),
      format_selection: () => readSelection({ withStyles: true }),
      set_paragraph_format: () => readSelection({ withStyles: true }),
      add_comment: () => readSelection({}),
      word_set_header_footer: (a) => wordReadHeadersFooters({ section: a.section, types: [a.type || 'primary'], includeFields: a.field !== undefined }),
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
      word_format_paragraphs: (a) => JSON.stringify(a.targets || []),
      word_create_document: () => 'Create a new blank Word document and open it',
      word_set_styles: (a) => JSON.stringify(a.styles || []),
      word_insert_field: (a) => `Native ${a.type} field (${a.target || 'selection'} / ${a.mode || 'after'})`,
      word_set_header_footer: (a) => a.field ? `Native ${a.field.type} field (${a.kind})` : String(a.text || ''),
    },
    // get_environment 的 Word 专属探测：对象模型属性是否真实存在（区分「API 不存在」与「存在但调用失败」）
    environmentProbe: async () => {
      return await Word.run(async (ctx) => {
        const body = ctx.document.body;
        const sel = ctx.document.getSelection();
        const diagnostics = {};
        if (isSetSupported('WordApiDesktop','1.4') && 'isReadOnly' in ctx.document) {
          try { ctx.document.load('isReadOnly,isFinal,protectionType,isWriteReserved'); await ctx.sync(); diagnostics.document = {isReadOnly:ctx.document.isReadOnly,isFinal:ctx.document.isFinal,protectionType:ctx.document.protectionType,isWriteReserved:ctx.document.isWriteReserved}; }
          catch (e) { diagnostics.documentError = String(e.message || e); }
        }
        if (body.contentControls) {
          try { body.contentControls.load('items/id,items/title,items/tag,items/cannotDelete,items/cannotEdit'); await ctx.sync(); diagnostics.contentControls = body.contentControls.items.map(c => ({id:c.id,title:c.title,tag:c.tag,cannotDelete:c.cannotDelete,cannotEdit:c.cannotEdit})); }
          catch (e) { diagnostics.contentControlsError = String(e.message || e); }
        }
        if (isSetSupported('WordApiDesktop','1.2') && typeof body.getRange === 'function') {
          try { const pages = body.getRange().pages; if (pages) { pages.load('items/index,items/width,items/height'); await ctx.sync(); diagnostics.pagination = {count:pages.items.length,pages:pages.items.map(p => ({index:p.index,width:p.width,height:p.height}))}; } }
          catch (e) { diagnostics.paginationError = String(e.message || e); }
        }
        return {
          nativeDiagnostics: diagnostics,
          bodyHasComments: !!body.comments,
          selHasParagraphFormat: !!sel.paragraphFormat,
          bodyInsertTableType: typeof body.insertTable,
          atlasBodyPresetSupported: true,
          atlasBodyPresetVersion: 3,
          footerFieldReuseSupported: true,
          nativeFieldParagraphAnchorSupported: true,
          nativeTocCollectionAddType: isSetSupported("WordApiDesktop","1.4") ? typeof ctx.document.tablesOfContents?.add : "undefined",
          bodyInsertOoxmlType: typeof body.insertOoxml,
          selHasFont: !!sel.font,
          documentGetStylesType: typeof ctx.document.getStyles,
          selectionInsertFieldType: typeof sel.insertField,
          selectionHasFields: !!sel.fields,
          supportedDesktopSets: ['1.1','1.2','1.3','1.4'].filter(version => isSetSupported('WordApiDesktop',version)),
        };
      });
    },
  };
})();
