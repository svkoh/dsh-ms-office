// excel.js — Excel（xlsx）宿主的全部 action 实现
//
// 分片契约：本文件是自包含 IIFE，只通过 window.__OFFICE_ACTIONS__ 与其它分片交换东西，
// 与其它分片不共享词法作用域。挂载点：
//   A.hosts.Excel = { actions, selectTargets, previewReaders, afterContent, environmentProbe? }
// 这里只放 Excel 的代码：分片之后「按宿主分叉」的写法既不需要（文件名就是宿主）也不允许（函数签名里已经没有 host 参数）。
(function () {
  const A = window.__OFFICE_ACTIONS__;
  const U = A.util;
  const okResult = U.okResult, errResult = U.errResult, excelRange = U.excelRange, getSelectionTextAsync = U.getSelectionTextAsync;
  const isSetSupported = A.isSetSupported, validateMatrix = A.validateMatrix;

  // ================= 跨宿主 action 的 Excel 实现 =================
  // read_selection / read_document / read_styles / replace_all / format_selection / add_comment /
  // read_comments / read_properties / locate_select 在多个宿主各有一份实现；拆分前它们是同一个
  // 函数里按宿主分叉的那几段，现在每个宿主只保留自己那一份。
  // （write_selection 三宿主实现完全一致，在 common.js。）
  const sleep = A.sleep, notFound = A.notFound;

  async function readSelection(args) {
    const a = args || {};
    if (a.withStyles) {
      try {
        return await Excel.run(async (ctx) => {
          const range = ctx.workbook.getSelectedRange();
          range.load('rowCount, columnCount');
          await ctx.sync();
          const rows = Math.min(range.rowCount, 10);
          const cols = Math.min(range.columnCount, 10);
          const cellRefs = [];
          for (let r = 0; r < rows; r++) {
            for (let c = 0; c < cols; c++) {
              const cell = range.getCell(r, c);
              cell.load('text, format/fill/color, format/font/name, format/font/size, format/font/bold, format/font/italic, format/font/color, numberFormat');
              cellRefs.push(cell);
            }
          }
          await ctx.sync();
          const cells = [];
          for (let r = 0; r < rows; r++) {
            for (let c = 0; c < cols; c++) {
              const cell = cellRefs[r * cols + c];
              const f = cell.format.font;
              cells.push({ row: r + 1, col: c + 1, value: cell.text, font: { name: f.name, size: f.size, bold: f.bold, italic: f.italic, color: f.color }, fill: cell.format.fill.color, numberFormat: cell.numberFormat });
            }
          }
          return okResult({ host: 'Excel', rows, cols, total: range.rowCount * range.columnCount, cells });
        });
      } catch (e) { return errResult('execution', String(e && e.message || e)); }
    }
    // 无 withStyles：Excel 选区没有更多可读的样式信息，走 Common API 取选区文本
    return getSelectionTextAsync();
  }

  async function readDocument(args) {
    try {
      return await Excel.run(async (ctx) => {
        const sheetName = args.sheet ? String(args.sheet) : null;
        const ws = sheetName ? ctx.workbook.worksheets.getItem(sheetName) : ctx.workbook.worksheets.getActiveWorksheet();
        ws.load('name');
        const range = ws.getUsedRange(true);
        range.load('rowCount, columnCount');
        await ctx.sync();
        const MAX_CELLS = 5000;
        const rowCount = range.rowCount, columnCount = range.columnCount;
        const bounded = boundedReadRange(range, MAX_CELLS);
        bounded.load('text');
        await ctx.sync();
        const rows = bounded.text || [];
        return okResult({ text: rows.map((row) => row.join('\t')).join('\n'), host: 'Excel', sheet: ws.name, rows: rowCount, cols: columnCount, truncated: rowCount * columnCount > MAX_CELLS });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function readStyles() {
    try {
      return await Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.load('rowCount, columnCount');
        await ctx.sync();
        const rows = Math.min(range.rowCount, 10);
        const cols = Math.min(range.columnCount, 10);
        const cellRefs = [];
        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < cols; c++) {
            const cell = range.getCell(r, c);
            cell.load('text, format/fill/color, format/font/name, format/font/size, format/font/bold, format/font/italic, format/font/color, numberFormat');
            cellRefs.push(cell);
          }
        }
        await ctx.sync();
        const cells = [];
        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < cols; c++) {
            const cell = cellRefs[r * cols + c];
            const f = cell.format.font;
            cells.push({ row: r + 1, col: c + 1, value: cell.text, font: { name: f.name, size: f.size, bold: f.bold, italic: f.italic, color: f.color }, fill: cell.format.fill.color, numberFormat: cell.numberFormat });
          }
        }
        return okResult({ host: 'Excel', rows, cols, total: range.rowCount * range.columnCount, cells });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function replaceAll(args) {
    const search = String(args.search ?? '');
    const replace = String(args.replace ?? '');
    if (!search) return errResult('bad_args', 'search required');
    // Excel 的 range.replaceAll 没有 dryRun 语义、也给不出命中数 → 明确拒绝，不假装预览
    if (args.dryRun) return errResult('unsupported', 'dryRun not supported in Excel replace_all');
    try {
      return await Excel.run(async (ctx) => {
        const range = ctx.workbook.worksheets.getActiveWorksheet().getUsedRange(true);
        range.replaceAll(search, replace, { completeMatch: false, matchCase: false });
        await ctx.sync();
        return okResult({ replaced: -1, host: 'Excel' });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function formatSelection(args) {
    try {
      return await Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        const fmt = range.format;
        if (args.font) fmt.font.name = String(args.font);
        if (args.size) fmt.font.size = Number(args.size);
        if (args.bold !== undefined) fmt.font.bold = !!args.bold;
        if (args.italic !== undefined) fmt.font.italic = !!args.italic;
        if (args.color) fmt.font.color = String(args.color);
        await ctx.sync();
        return okResult({ host: 'Excel', formatted: true });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function addComment(args) {
    const text = String(args.text ?? '');
    if (!text) return errResult('bad_args', 'text required');
    if (!args.cell) return errResult('bad_args', 'cell (e.g. A1 or Sheet1!A1) required for Excel');
    try {
      return await Excel.run(async (ctx) => {
        // cell 可带表名（Sheet1!A1）→ 加到目标表；否则加到激活表
        let cellAddr = String(args.cell);
        let ws;
        const bang = cellAddr.lastIndexOf('!');
        if (bang >= 0) {
          ws = ctx.workbook.worksheets.getItem(cellAddr.slice(0, bang));
          cellAddr = cellAddr.slice(bang + 1);
        } else {
          ws = ctx.workbook.worksheets.getActiveWorksheet();
        }
        ws.comments.add(cellAddr, text);
        await ctx.sync();
        return okResult({ host: 'Excel', added: true, cell: cellAddr, sheet: ws.name });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function readComments() {
    try {
      return await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getActiveWorksheet();
        const comments = ws.comments;
        // 注：此环境 Comment.getRange() 不存在、cellAddress 可能为空 → 尽力返回
        comments.load('items/author, items/content, items/resolved, items/cellAddress');
        await ctx.sync();
        const list = comments.items.map((c) => ({ cell: c.cellAddress || '', content: c.content, author: c.author, resolved: c.resolved }));
        return okResult({ host: 'Excel', comments: list });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function readProperties() {
    try {
      return await Excel.run(async (ctx) => {
        const p = ctx.workbook.properties;
        p.load('title, author, subject, keywords, comments, category');
        await ctx.sync();
        return okResult({ host: 'Excel', properties: { title: p.title, author: p.author, subject: p.subject, keywords: p.keywords, comments: p.comments, category: p.category } });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function locateSelect(args) {
    const blinks = Math.max(0, Math.min(5, Number(args.blinks) || 0));
    const interval = Math.max(100, Math.min(800, Number(args.interval) || 300));
    const reps = Math.max(1, blinks);
    const bookmark = (args.bookmark !== undefined ? String(args.bookmark) : args.anchor !== undefined ? String(args.anchor) : null);
    const text = args.text !== undefined ? String(args.text) : null;
    const address = (args.range !== undefined ? String(args.range) : args.address !== undefined ? String(args.address) : null);
    if (!bookmark && !text && !address) return errResult('bad_args', 'locate_select requires one of: bookmark/anchor, text, range/address');
    try {
      const sheetName = args.sheet !== undefined ? String(args.sheet) : null;
      // 先解析目标地址（一次），避免每轮重复扫描
      let targetAddress = null;
      if (bookmark) {
        await Excel.run(async (ctx) => {
          const nm = ctx.workbook.names.getItem(bookmark);
          const r = nm.getRange();
          r.load('address');
          await ctx.sync();
          targetAddress = r.address;
        });
      } else if (address) {
        targetAddress = address;
      } else {
        await Excel.run(async (ctx) => {
          const ws = sheetName ? ctx.workbook.worksheets.getItem(sheetName) : ctx.workbook.worksheets.getActiveWorksheet();
          const used = ws.getUsedRange(true);
          used.load('text, rowCount, columnCount');
          await ctx.sync();
          let fr = -1, fc = -1;
          for (let r = 0; r < used.rowCount && fr < 0; r++) {
            for (let c = 0; c < used.columnCount && fr < 0; c++) {
              if (String((used.text[r] && used.text[r][c]) || '').indexOf(text) >= 0) { fr = r; fc = c; }
            }
          }
          if (fr < 0) throw notFound(`text not found: ${text}`);
          const cell = used.getCell(fr, fc);
          cell.load('address');
          await ctx.sync();
          targetAddress = cell.address;
        });
      }
      if (!targetAddress) throw notFound('could not resolve target');
      const [sheetPart, addrPart] = targetAddress.indexOf('!') >= 0 ? targetAddress.split('!') : [sheetName, targetAddress];
      for (let i = 0; i < reps; i++) {
        await Excel.run(async (ctx) => {
          const ws = sheetPart ? ctx.workbook.worksheets.getItem(sheetPart) : ctx.workbook.worksheets.getActiveWorksheet();
          ws.getRange(addrPart).select(); // 整块选中（绿色边框）
          await ctx.sync();
        });
        if (blinks > 0) {
          await sleep(interval);
          await Excel.run(async (ctx) => {
            const ws = sheetPart ? ctx.workbook.worksheets.getItem(sheetPart) : ctx.workbook.worksheets.getActiveWorksheet();
            ws.getRange(addrPart.split(':')[0]).select(); // 选回首格：取消整块高亮、留在原位
            await ctx.sync();
          });
          await sleep(interval);
        }
      }
      return okResult({ host: 'Excel', located: bookmark || text || targetAddress, selected: true, blinks });
    } catch (e) {
      if (e && e.notFound) return errResult('not_found', `locate_select: ${e.message}`);
      return errResult('execution', String(e && e.message || e));
    }
  }

  // ===== 审批预选（selectTarget）：ask 确认模式下把将要修改的区域选中给用户看 =====
  async function selectExcelTargetRange(args) {
    if (!args.address) return;
    try {
      await Excel.run(async (ctx) => { excelRange(ctx, String(args.address)).select(); await ctx.sync(); });
    } catch (e) { /* 预选失败不阻塞审批 */ }
  }

  // ================= Excel 组 =================

  async function listSheets() {
    try {
      return await Excel.run(async (ctx) => {
        const sheets = ctx.workbook.worksheets;
        sheets.load('items/name, items/position, items/visibility');
        await ctx.sync();
        const list = sheets.items.map((s) => ({ name: s.name, position: s.position, visibility: s.visibility }));
        return okResult({ host: 'Excel', sheets: list });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  // A wide range may exceed the budget in its first row; cap columns as well.
  function boundedReadRange(range, limit) {
    if (range.rowCount * range.columnCount <= limit) return range;
    const cols = Math.min(range.columnCount, limit);
    const rows = Math.min(range.rowCount, Math.max(1, Math.floor(limit / cols)));
    return range.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
  }

  async function readRange(args) {
    const address = String(args.address || '');
    if (!address) return errResult('bad_args', 'address (e.g. Sheet1!A1:B10) required');
    try {
      return await Excel.run(async (ctx) => {
        const range = excelRange(ctx, address);
        range.load('address, rowCount, columnCount');
        await ctx.sync();
        const requestedLimit = args.limit === undefined ? 5000 : Number(args.limit);
        if (!Number.isSafeInteger(requestedLimit) || requestedLimit < 1) return errResult('bad_args', 'limit must be a positive integer');
        const limit = Math.min(requestedLimit, 5000);
        const total = (range.rowCount || 0) * (range.columnCount || 0);
        const bounded = boundedReadRange(range, limit);
        bounded.load('values, formulas, numberFormat');
        const dimensions = isSetSupported('ExcelApi', '1.2');
        if (args.withLayout) range.format.load(`wrapText,horizontalAlignment,verticalAlignment${dimensions ? ',columnWidth,rowHeight' : ''}`);
        await ctx.sync();
        return okResult({ host: 'Excel', address: range.address || address, rows: range.rowCount, cols: range.columnCount, truncated: total > limit, values: bounded.values || [], formulas: bounded.formulas || [], numberFormat: bounded.numberFormat || [], ...(args.withLayout ? { layout: excelLayoutSnapshot(range.format, dimensions) } : {}) });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function excelRenderRange(args) {
    if (!args || typeof args.address !== 'string' || !args.address.trim()) return errResult('bad_args', 'address required');
    const maxCells = args.maxCells === undefined ? 5000 : args.maxCells;
    if (!Number.isInteger(maxCells) || maxCells < 1 || maxCells > 10000) return errResult('bad_args', 'maxCells must be an integer from 1 to 10000');
    if (!isSetSupported('ExcelApi', '1.7')) return errResult('requirement', 'excel_render_range requires ExcelApi 1.7+');
    try { return await Excel.run(async (ctx) => {
      const range = excelRange(ctx, args.address);
      range.load('rowCount,columnCount,address'); await ctx.sync();
      const rows = range.rowCount, cols = range.columnCount;
      if (!Number.isInteger(rows) || !Number.isInteger(cols) || rows < 1 || cols < 1) return errResult('execution', 'Range dimensions are unavailable or invalid');
      if (rows * cols > maxCells) return errResult('bad_args', `Range contains ${rows * cols} cells, exceeding maxCells ${maxCells}; request a smaller range (no image was rendered)`);
      const image = range.getImage(); await ctx.sync();
      return okResult({ host: 'Excel', address: range.address, rows, cols, mimeType: 'image/png', base64: image.value });
    }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function excelRenderChart(args) {
    if (!args || typeof args.sheet !== 'string' || !args.sheet.trim() || typeof args.name !== 'string' || !args.name.trim()) return errResult('bad_args', 'sheet and name required');
    if (!isSetSupported('ExcelApi', '1.9')) return errResult('requirement', 'excel_render_chart requires ExcelApi 1.9+');
    try { return await Excel.run(async (ctx) => {
      const chart = ctx.workbook.worksheets.getItem(args.sheet).charts.getItem(args.name);
      chart.load('name,width,height'); const image = chart.getImage(); await ctx.sync();
      return okResult({ host: 'Excel', sheet: args.sheet, name: chart.name, width: chart.width, height: chart.height, mimeType: 'image/png', base64: image.value });
    }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function writeRange(args) {
    const address = String(args.address || '');
    if (!address) return errResult('bad_args', 'address required');
    if ((args.values !== undefined) === (args.formulas !== undefined)) return errResult('bad_args', 'provide exactly one of values or formulas');
    let data; try { data = validateMatrix(args.values !== undefined ? args.values : args.formulas, args.values !== undefined ? 'values' : 'formulas', 50000); } catch (e) { return errResult('bad_args', e.message); }
    try {
      return await Excel.run(async (ctx) => {
        const targetRange = excelRange(ctx, address);
        targetRange.load('rowCount,columnCount,address');
        await ctx.sync();
        if (targetRange.rowCount !== data.length || targetRange.columnCount !== data[0].length) return errResult('bad_args', `matrix is ${data.length}x${data[0].length}, target range is ${targetRange.rowCount}x${targetRange.columnCount}`);
        const totalCells = data.length * data[0].length;
        const BATCH = 5000;
        if (totalCells <= BATCH) {
          const range = targetRange;
          if (args.values) range.values = args.values;
          if (args.formulas) range.formulas = args.formulas;
          await ctx.sync();
          range.load('values,formulas,numberFormat'); await ctx.sync();
          const actual = args.values ? range.values : range.formulas;
          const mismatches = [];
          for (let r = 0; r < data.length && mismatches.length < 20; r++) for (let c = 0; c < data[0].length && mismatches.length < 20; c++) if (String(actual[r][c] ?? '') !== String(data[r][c] ?? '')) mismatches.push({ row: r, col: c, expected: data[r][c], actual: actual[r][c] });
          try { range.select(); await ctx.sync(); } catch (e) { /* 选中失败不影响结果 */ }
          return okResult({ host: 'Excel', address: range.address || address, rows: data.length, cols: data[0].length, batched: 1, selected: true, verified: mismatches.length === 0, mismatches, values: range.values, formulas: range.formulas, numberFormat: range.numberFormat });
        }
        const perChunk = Math.max(1, Math.floor(BATCH / data[0].length));
        let chunks = 0;
        for (let i = 0; i < data.length; i += perChunk) {
          const chunk = data.slice(i, i + perChunk);
          const base = excelRange(ctx, address).getCell(i, 0);
          const target = base.getResizedRange(chunk.length - 1, data[0].length - 1);
          if (args.values) target.values = chunk;
          if (args.formulas) target.formulas = chunk;
          await ctx.sync();
          chunks++;
        }
        targetRange.load('values,formulas,numberFormat'); await ctx.sync();
        const actual = args.values ? targetRange.values : targetRange.formulas; const mismatches = [];
        for (let r = 0; r < data.length && mismatches.length < 20; r++) for (let c = 0; c < data[0].length && mismatches.length < 20; c++) if (String(actual[r][c] ?? '') !== String(data[r][c] ?? '')) mismatches.push({ row: r, col: c, expected: data[r][c], actual: actual[r][c] });
        try { targetRange.select(); await ctx.sync(); } catch (e) { /* 选中失败不影响结果 */ }
        return okResult({ host: 'Excel', address: targetRange.address || address, rows: data.length, cols: data[0].length, batched: chunks, selected: true, verified: mismatches.length === 0, mismatches, values: targetRange.values, formulas: targetRange.formulas, numberFormat: targetRange.numberFormat });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  const EXCEL_HORIZONTAL_ALIGNMENTS = ['General', 'Left', 'Center', 'Right', 'Fill', 'Justify', 'CenterAcrossSelection', 'Distributed'];
  const EXCEL_VERTICAL_ALIGNMENTS = ['Top', 'Center', 'Bottom', 'Justify', 'Distributed'];

  function excelLayoutSnapshot(fmt, dimensions) {
    return { wrapText: fmt.wrapText, horizontalAlignment: fmt.horizontalAlignment, verticalAlignment: fmt.verticalAlignment, ...(dimensions ? { columnWidth: fmt.columnWidth, rowHeight: fmt.rowHeight } : {}) };
  }

  function validateRangeLayout(args) {
    const layout = {};
    for (const key of ['columnWidth', 'rowHeight']) if (args[key] !== undefined) {
      if (typeof args[key] !== 'number' || !Number.isFinite(args[key]) || args[key] <= 0) throw new Error(`${key} must be a positive finite number in points`);
      layout[key] = args[key];
    }
    for (const key of ['autoFit', 'autoFitColumns', 'autoFitRows', 'wrapText', 'bold']) if (args[key] !== undefined && typeof args[key] !== 'boolean') throw new Error(`${key} must be boolean`);
    if (args.wrapText !== undefined) layout.wrapText = args.wrapText;
    for (const [key, allowed] of [['horizontalAlignment', EXCEL_HORIZONTAL_ALIGNMENTS], ['verticalAlignment', EXCEL_VERTICAL_ALIGNMENTS]]) if (args[key] !== undefined) {
      const value = typeof args[key] === 'string' && allowed.find((item) => item.toLowerCase() === args[key].toLowerCase());
      if (!value) throw new Error(`${key} must be ${allowed.join('|')}`);
      layout[key] = value;
    }
    if ((args.autoFit === true || args.autoFitColumns === true) && args.columnWidth !== undefined) throw new Error('autoFit/autoFitColumns and columnWidth are mutually exclusive');
    if (args.autoFitRows === true && args.rowHeight !== undefined) throw new Error('autoFitRows and rowHeight are mutually exclusive');
    if (args.size !== undefined && (typeof args.size !== 'number' || !Number.isFinite(args.size) || args.size <= 0)) throw new Error('size must be a positive finite number');
    if (args.tableStyle && !args.tableName) throw new Error('tableName required when tableStyle is provided');
    return layout;
  }

  async function formatRange(args, previewOnly = false) {
    const address = String(args.address || '');
    if (!address) return errResult('bad_args', 'address required');
    let layout; try { layout = validateRangeLayout(args); } catch (e) { return errResult('bad_args', e.message); }
    const dimensions = isSetSupported('ExcelApi', '1.2');
    if (!dimensions && (args.columnWidth !== undefined || args.rowHeight !== undefined || (args.autoFit === true || args.autoFitColumns === true) || args.autoFitRows === true)) return errResult('requirement', 'range dimensions and autofit require ExcelApi 1.2+');
    if (previewOnly) {
      const before = await readRange({ address, withLayout: true });
      return before.ok ? okResult({ ...before.result, previewOnly: true, requested: { ...args } }) : before;
    }
    try {
      return await Excel.run(async (ctx) => {
        const range = excelRange(ctx, address);
        const fmt = range.format;
        // Resolve both targets before queuing any writes (including a table style).
        range.load('address,rowCount,columnCount');
        const table = args.tableStyle ? ctx.workbook.tables.getItem(String(args.tableName)) : null;
        if (table) table.load('name');
        await ctx.sync();
        if (args.font) fmt.font.name = String(args.font);
        if (args.size !== undefined) fmt.font.size = args.size;
        if (args.bold !== undefined) fmt.font.bold = args.bold;
        if (args.fill) fmt.fill.color = String(args.fill);
        if (args.color) fmt.font.color = String(args.color);
        if (args.numberFormat) range.numberFormat = [[String(args.numberFormat)]];
        Object.assign(fmt, layout);
        if ((args.autoFit === true || args.autoFitColumns === true)) fmt.autofitColumns();
        if (args.autoFitRows === true) fmt.autofitRows();
        if (table) table.style = String(args.tableStyle);
        await ctx.sync();
        const sample = boundedReadRange(range, 5000);
        sample.load('numberFormat');
        range.load('address,format/fill/color,format/font/name,format/font/size,format/font/bold,format/font/color');
        fmt.load(`wrapText,horizontalAlignment,verticalAlignment${dimensions ? ',columnWidth,rowHeight' : ''}`);
        await ctx.sync();
        return okResult({ host: 'Excel', formatted: true, address: range.address || address, verification: { numberFormat: sample.numberFormat, numberFormatTruncated: range.rowCount * range.columnCount > 5000, fill: fmt.fill.color, font: { name: fmt.font.name, size: fmt.font.size, bold: fmt.font.bold, color: fmt.font.color }, layout: excelLayoutSnapshot(fmt, dimensions), tableName: args.tableName || null, tableStyle: args.tableStyle || null } });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  function validateFreezePanes(args) {
    if (!args || typeof args.sheet !== 'string' || !args.sheet.trim()) throw new Error('sheet required');
    if (!['rows', 'columns', 'range', 'none'].includes(args.mode)) throw new Error('mode must be rows|columns|range|none');
    if (args.mode === 'rows' || args.mode === 'columns') {
      const max = args.mode === 'rows' ? 1048576 : 16384;
      if (!Number.isInteger(args.count) || args.count < 1 || args.count > max) throw new Error(`count must be a positive integer <= ${max}`);
      if (args.range !== undefined) throw new Error('range is only valid in range mode');
    } else if (args.count !== undefined) throw new Error('count is only valid in rows/columns mode');
    if (args.mode === 'range') {
      if (typeof args.range !== 'string' || !/^\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6})?$/.test(args.range)) throw new Error('range must be a sheet-local A1 cell or rectangular cell range');
      for (const cell of args.range.split(':')) {
        const parts = cell.replace(/\$/g, '').match(/^([A-Za-z]+)(\d+)$/);
        const col = parts[1].toUpperCase().split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
        if (col > 16384 || Number(parts[2]) > 1048576) throw new Error('range exceeds worksheet bounds');
      }
    } else if (args.range !== undefined) throw new Error('range is only valid in range mode');
    return args;
  }

  async function excelFreezePanes(args, preview) {
    let spec; try { spec = validateFreezePanes(args); } catch (e) { return errResult('bad_args', e.message); }
    if (!isSetSupported('ExcelApi', '1.7')) return errResult('requirement', 'excel_set_freeze_panes requires ExcelApi 1.7+');
    try { return await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(spec.sheet);
      sheet.load('name');
      const range = spec.mode === 'range' ? sheet.getRange(spec.range) : null;
      if (range) range.load('address');
      await ctx.sync();
      const panes = sheet.freezePanes;
      if (!preview) {
        if (spec.mode === 'rows') panes.freezeRows(spec.count);
        else if (spec.mode === 'columns') panes.freezeColumns(spec.count);
        else if (spec.mode === 'range') panes.freezeAt(range);
        else panes.unfreeze();
        await ctx.sync();
      }
      const frozen = panes.getLocationOrNullObject(); frozen.load('isNullObject,address'); await ctx.sync();
      return okResult({ host: 'Excel', sheet: sheet.name, ...(preview ? { previewOnly: true } : { mode: spec.mode }), frozenRange: frozen.isNullObject ? null : frozen.address, requirementSet: 'ExcelApi 1.7' });
    }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function excelPrintLayout(args, preview = false) {
    if (!args || typeof args.sheet !== 'string' || !args.sheet.trim() || typeof args.address !== 'string') return errResult('bad_args','sheet and sheet-local address required');
    try { validateFreezePanes({sheet:args.sheet,mode:'range',range:args.address}); } catch (e) { return errResult('bad_args',e.message); }
    const spec = {address:args.address,orientation:args.orientation || 'Landscape',paperSize:args.paperSize || 'A4',fitToPagesWide:args.fitToPagesWide === undefined ? 1 : args.fitToPagesWide,fitToPagesTall:args.fitToPagesTall === undefined ? 1 : args.fitToPagesTall};
    if (!['Landscape','Portrait'].includes(spec.orientation) || !['A4','Letter','A3'].includes(spec.paperSize)) return errResult('bad_args','orientation must be Landscape|Portrait; paperSize must be A4|Letter|A3');
    for (const key of ['fitToPagesWide','fitToPagesTall']) if (!Number.isInteger(spec[key]) || spec[key] < 1 || spec[key] > 100) return errResult('bad_args',key+' must be integer 1-100');
    if (!isSetSupported('ExcelApi','1.9')) return errResult('requirement','excel_set_print_layout requires ExcelApi 1.9+');
    try { return await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(args.sheet), range = sheet.getRange(spec.address), layout = sheet.pageLayout;
      sheet.load('name'); range.load('address,rowCount,columnCount');
      const fields = 'orientation,paperSize,zoom,printGridlines,printHeadings,centerHorizontally,draftMode,blackAndWhite';
      layout.load(fields); const beforeArea = layout.getPrintAreaOrNullObject(); beforeArea.load('isNullObject,address'); await ctx.sync();
      if (range.rowCount * range.columnCount > 10000) return errResult('bad_args','Print area must contain at most 10000 cells');
      const snapshot = area => ({printArea:area.isNullObject ? null : area.address,orientation:layout.orientation,paperSize:layout.paperSize,zoom:layout.zoom,printGridlines:layout.printGridlines,printHeadings:layout.printHeadings,centerHorizontally:layout.centerHorizontally,draftMode:layout.draftMode,blackAndWhite:layout.blackAndWhite});
      const previousState = snapshot(beforeArea);
      if (preview) return okResult({host:'Excel',sheet:sheet.name,previewOnly:true,requested:spec,previousState});
      layout.setPrintArea(range); layout.orientation = spec.orientation; layout.paperSize = spec.paperSize; layout.zoom = {horizontalFitToPages:spec.fitToPagesWide,verticalFitToPages:spec.fitToPagesTall};
      layout.printGridlines = false; layout.printHeadings = false; layout.centerHorizontally = true; layout.draftMode = false; layout.blackAndWhite = false;
      await ctx.sync(); layout.load(fields); const afterArea = layout.getPrintAreaOrNullObject(); afterArea.load('isNullObject,address'); await ctx.sync();
      return okResult({host:'Excel',sheet:sheet.name,requested:spec,previousState,afterState:snapshot(afterArea)});
    }); } catch (e) { return errResult('execution',String(e && e.message || e)); }
  }

  async function listTables(args) {
    if (!isSetSupported('ExcelApi', '1.1')) return errResult('requirement', 'list_tables requires ExcelApi 1.1+');
    try { return await Excel.run(async (ctx) => {
      const collection = args.sheet ? ctx.workbook.worksheets.getItem(String(args.sheet)).tables : ctx.workbook.tables;
      collection.load('items/id,items/name,items/style,items/showHeaders,items/showTotals'); await ctx.sync();
      const refs = collection.items.map((table) => { const range = table.getRange(); range.load('address,rowCount,columnCount'); return { table, range }; }); await ctx.sync();
      const limit = Math.min(5000, Math.max(1, Number(args.limit) || 500));
      const withValues = refs.filter(({ range }) => args.includeValues && range.rowCount * range.columnCount <= limit);
      if (withValues.length) { for (const { range } of withValues) range.load('values'); await ctx.sync(); }
      return okResult({ host: 'Excel', tables: refs.map(({ table, range }) => ({ id: table.id, name: table.name, style: table.style, showHeaders: table.showHeaders, showTotals: table.showTotals, address: range.address, rows: range.rowCount, cols: range.columnCount, values: args.includeValues && range.rowCount * range.columnCount <= limit ? range.values : undefined })) });
    }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function createTable(args) {
    if (!isSetSupported('ExcelApi', '1.1')) return errResult('requirement', 'create_table requires ExcelApi 1.1+');
    const address = String(args.address || ''); if (!address || typeof args.hasHeaders !== 'boolean') return errResult('bad_args', 'address and boolean hasHeaders required');
    try { return await Excel.run(async (ctx) => {
      const table = ctx.workbook.tables.add(address, args.hasHeaders);
      if (args.name) table.name = String(args.name); if (args.style) table.style = String(args.style);
      for (const key of ['showTotals','showBandedRows','showBandedColumns','highlightFirstColumn','highlightLastColumn','showFilterButton']) if (args[key] !== undefined) { if (typeof args[key] !== 'boolean') return errResult('bad_args', `${key} must be boolean`); table[key] = args[key]; }
      table.load('id,name,style,showHeaders,showTotals'); const range = table.getRange(); range.load('address,rowCount,columnCount,values'); await ctx.sync(); range.select(); await ctx.sync();
      return okResult({ host: 'Excel', table: { id: table.id, name: table.name, style: table.style, showHeaders: table.showHeaders, showTotals: table.showTotals, address: range.address, rows: range.rowCount, cols: range.columnCount, values: range.values }, selected: true });
    }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function changeTableLifecycle(args, convert) {
    const action = convert ? 'convert_table_to_range' : 'delete_table';
    const requirement = convert ? '1.2' : '1.1'; if (!isSetSupported('ExcelApi', requirement)) return errResult('requirement', `${action} requires ExcelApi ${requirement}+`);
    const name = String(args.name || ''); if (!name) return errResult('bad_args', 'name required');
    try { return await Excel.run(async (ctx) => { const table = ctx.workbook.tables.getItem(name); const range = table.getRange(); range.load('address,values'); await ctx.sync(); const snapshot = { name, address: range.address, values: range.values }; if (convert) table.convertToRange(); else table.delete(); await ctx.sync(); return okResult({ host: 'Excel', action, previous: snapshot, tableRemoved: true, dataPreserved: convert }); }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function readDataValidation(args) {
    if (!isSetSupported('ExcelApi', '1.8')) return errResult('requirement', 'read_data_validation requires ExcelApi 1.8+');
    const address = String(args.address || ''); if (!address) return errResult('bad_args', 'address required');
    try { return await Excel.run(async (ctx) => { const range = excelRange(ctx, address); const dv = range.dataValidation; range.load('address'); dv.load('type,rule,errorAlert,ignoreBlanks,prompt,valid'); await ctx.sync(); return okResult({ host: 'Excel', address: range.address, dataValidation: { type: dv.type, rule: dv.rule, errorAlert: dv.errorAlert, ignoreBlanks: dv.ignoreBlanks, prompt: dv.prompt, valid: dv.valid } }); }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function setDataValidation(args) {
    if (!isSetSupported('ExcelApi', '1.8')) return errResult('requirement', 'set_data_validation requires ExcelApi 1.8+');
    const address = String(args.address || ''), rule = args.rule; const keys = rule && Object.keys(rule).filter((k) => ['wholeNumber','decimal','textLength','date','time','list','custom'].includes(k));
    if (!address || !rule || keys.length !== 1) return errResult('bad_args', 'address and rule with exactly one validation type required');
    try { return await Excel.run(async (ctx) => { const range = excelRange(ctx, address); const dv = range.dataValidation; dv.rule = rule; if (args.ignoreBlanks !== undefined) dv.ignoreBlanks = !!args.ignoreBlanks; if (args.errorAlert) dv.errorAlert = args.errorAlert; if (args.prompt) dv.prompt = args.prompt; dv.load('type,rule,errorAlert,ignoreBlanks,prompt,valid'); range.load('address'); await ctx.sync(); range.select(); await ctx.sync(); return okResult({ host: 'Excel', address: range.address, dataValidation: { type: dv.type, rule: dv.rule, errorAlert: dv.errorAlert, ignoreBlanks: dv.ignoreBlanks, prompt: dv.prompt, valid: dv.valid }, selected: true }); }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function clearDataValidation(args) {
    if (!isSetSupported('ExcelApi', '1.8')) return errResult('requirement', 'clear_data_validation requires ExcelApi 1.8+');
    const address = String(args.address || ''); if (!address) return errResult('bad_args', 'address required');
    try { return await Excel.run(async (ctx) => { const range = excelRange(ctx, address); range.dataValidation.clear(); await ctx.sync(); range.dataValidation.load('type,valid'); await ctx.sync(); return okResult({ host: 'Excel', address, cleared: true, type: range.dataValidation.type, valid: range.dataValidation.valid }); }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function readConditionalFormats(args) {
    if (!isSetSupported('ExcelApi', '1.6')) return errResult('requirement', 'read_conditional_formats requires ExcelApi 1.6+');
    const address = String(args.address || ''); if (!address) return errResult('bad_args', 'address required');
    try { return await Excel.run(async (ctx) => { const range = excelRange(ctx, address); const formats = range.conditionalFormats; formats.load('items/id,items/type,items/priority,items/stopIfTrue'); range.load('address'); await ctx.sync(); return okResult({ host: 'Excel', address: range.address, intersectionSemantics: true, formats: formats.items.map((cf) => ({ id: cf.id, type: cf.type, priority: cf.priority, stopIfTrue: cf.stopIfTrue })) }); }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function addConditionalFormat(args) {
    if (!isSetSupported('ExcelApi', '1.6')) return errResult('requirement', 'add_conditional_format requires ExcelApi 1.6+');
    const address = String(args.address || ''), type = String(args.type || ''); if (!address || !['Custom','CellValue','ContainsText'].includes(type) || !args.rule) return errResult('bad_args', 'address, supported type, and rule required');
    try { return await Excel.run(async (ctx) => { const range = excelRange(ctx, address); const cf = range.conditionalFormats.add(type); const branch = type === 'Custom' ? cf.custom : type === 'CellValue' ? cf.cellValue : cf.textComparison; branch.rule = args.rule; if (args.format) { if (args.format.fillColor) branch.format.fill.color = args.format.fillColor; if (args.format.fontColor) branch.format.font.color = args.format.fontColor; if (args.format.bold !== undefined) branch.format.font.bold = !!args.format.bold; }
      if (args.stopIfTrue !== undefined) cf.stopIfTrue = !!args.stopIfTrue; if (args.priority !== undefined) cf.priority = Number(args.priority); cf.load('id,type,priority,stopIfTrue'); await ctx.sync(); range.select(); await ctx.sync(); return okResult({ host: 'Excel', address, conditionalFormat: { id: cf.id, type: cf.type, priority: cf.priority, stopIfTrue: cf.stopIfTrue }, intersectionSemantics: true, selected: true }); }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function validateData(args) {
    if (!isSetSupported('ExcelApi', '1.8')) return errResult('requirement', 'validate_data requires ExcelApi 1.8+');
    const address = String(args.address || ''); if (!address) return errResult('bad_args', 'address required');
    try { return await Excel.run(async (ctx) => { const range = excelRange(ctx, address); const dv = range.dataValidation; dv.load('valid,type'); await ctx.sync(); const result = { host: 'Excel', address, valid: dv.valid, type: dv.type }; if (args.details) { if (!isSetSupported('ExcelApi','1.9')) return errResult('requirement','validate_data details requires ExcelApi 1.9+'); const invalid = dv.getInvalidCellsOrNullObject(); invalid.load('isNullObject,address,areaCount,cellCount'); await ctx.sync(); result.invalid = invalid.isNullObject ? null : { address: invalid.address, areas: invalid.areaCount, cells: invalid.cellCount }; } return okResult(result); }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function deleteConditionalFormat(args) {
    if (!isSetSupported('ExcelApi', '1.6')) return errResult('requirement', 'delete_conditional_format requires ExcelApi 1.6+');
    const address = String(args.address || ''), id = String(args.id || ''); if (!address || !id) return errResult('bad_args', 'address and id required');
    try { return await Excel.run(async (ctx) => { const range = excelRange(ctx, address); range.conditionalFormats.getItem(id).delete(); await ctx.sync(); range.conditionalFormats.load('items/id'); await ctx.sync(); return okResult({ host: 'Excel', address, deletedId: id, remainingIds: range.conditionalFormats.items.map((cf)=>cf.id), intersectionSemantics: true }); }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function clearConditionalFormats(args) {
    if (!isSetSupported('ExcelApi', '1.6')) return errResult('requirement', 'clear_conditional_formats requires ExcelApi 1.6+');
    const address = String(args.address || ''); if (!address) return errResult('bad_args', 'address required');
    try { return await Excel.run(async (ctx) => { const range = excelRange(ctx, address); range.conditionalFormats.clearAll(); await ctx.sync(); range.conditionalFormats.load('items'); await ctx.sync(); return okResult({ host: 'Excel', address, remaining: range.conditionalFormats.items.length, intersectionSemantics: true }); }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function excelExportFile(args) {
    const type = String(args && args.type || 'xlsx').toLowerCase();
    if (type === 'pdf') {
      if (!isSetSupported('ExcelApi','1.4')) return errResult('requirement','Native PDF nonblank preflight requires ExcelApi 1.4+');
      try {
        const printable = await Excel.run(async (ctx) => {
          const sheets = ctx.workbook.worksheets; sheets.load('items/name'); await ctx.sync();
          const refs = sheets.items.map(sheet => { const range = sheet.getUsedRangeOrNullObject(true); range.load('isNullObject'); sheet.charts.load('items/name'); return {range,sheet}; });
          await ctx.sync();
          return refs.some(({range,sheet}) => !range.isNullObject || sheet.charts.items.length > 0);
        });
        if (!printable) return errResult('bad_args','Native PDF export requires nonblank cells or charts; blank workbooks trigger the Office empty-print dialog');
      } catch (e) { return errResult('execution','Native PDF preflight failed: '+String(e && e.message || e)); }
    }
    if (!['xlsx', 'pdf'].includes(type)) return Promise.resolve(errResult('bad_args', 'type must be xlsx or pdf'));
    if (typeof Office.context.document.getFileAsync !== 'function') return Promise.resolve(errResult('requirement','Native file export is unavailable in this Excel runtime'));
    const fileType = type === 'pdf' ? Office.FileType.Pdf : Office.FileType.Compressed;
    return new Promise((resolve) => {
      let settled = false;
      const finish = (out) => { if (settled) return; settled = true; clearTimeout(timer); resolve(out); };
      const timer = setTimeout(() => finish(errResult('execution', `Office getFileAsync(${type}) did not call back within 30 seconds; native export support is unverified`)), 30000);
      Office.context.document.getFileAsync(fileType, { sliceSize: 4194304 }, (fileResult) => {
        if (settled) { if (fileResult.value) try { fileResult.value.closeAsync(() => {}); } catch (e) {} return; }
        if (fileResult.status !== Office.AsyncResultStatus.Succeeded) {
          return finish(errResult('execution', JSON.stringify(fileResult.error || { message: 'Office getFileAsync failed' })));
        }
        const file = fileResult.value;
        const chunks = [];
        let sliceIndex = 0;
        let received = 0;
        const done = (out) => { try { file.closeAsync(() => {}); } catch (e) {} finish(out); };
        const onSlice = (sliceResult) => {
          if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
            return done(errResult('execution', sliceResult.error && sliceResult.error.message));
          }
          chunks.push(new Uint8Array(sliceResult.value.data));
          sliceIndex += 1;
          received += sliceResult.value.size;
          if (sliceIndex >= file.sliceCount || received >= file.size) {
            const mime = type === 'pdf' ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
            const size = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
            const reader = new FileReader();
            reader.onerror = () => done(errResult('execution', String(reader.error || 'FileReader failed')));
            reader.onload = () => done(okResult({ host: 'Excel', fileType: type, size, base64: String(reader.result || '').replace(/^data:.*?;base64,/, '') }));
            reader.readAsDataURL(new Blob(chunks, { type: mime }));
          } else file.getSliceAsync(sliceIndex, onSlice);
        };
        file.getSliceAsync(0, onSlice);
      });
    });
  }

  const CHART_TYPES = {
    line: 'Line', column: 'ColumnClustered', bar: 'BarClustered', pie: 'Pie',
    area: 'Area', scatter: 'XYScatter', doughnut: 'Doughnut', radar: 'Radar',
  };

  async function insertChart(args, previewOnly = false) {
    const dataRange = args.dataRange;
    if (typeof dataRange !== 'string' || !dataRange.trim()) return errResult('bad_args','dataRange (e.g. Sheet1!A1:B10) required');
    const typeKey = args.type === undefined ? 'column' : typeof args.type === 'string' ? args.type.toLowerCase() : '';
    const typeName = CHART_TYPES[typeKey];
    if (!typeName) return errResult('bad_args','type must be one of: '+Object.keys(CHART_TYPES).join(', '));
    if (args.title !== undefined && (typeof args.title !== 'string' || args.title.length > 500)) return errResult('bad_args','title must be a string up to 500 characters');
    let position;
    if (args.position !== undefined) {
      const p = args.position;
      if (!p || typeof p !== 'object' || Array.isArray(p) || Object.keys(p).some(k => !['startCell','endCell'].includes(k))) return errResult('bad_args','position must be {startCell,endCell}');
      const cell = value => {
        if (typeof value !== 'string') return null;
        const m = /^\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})$/.exec(value);
        if (!m) return null;
        const column = [...m[1].toUpperCase()].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0), row=Number(m[2]);
        return column<=16384 && row<=1048576 ? {column,row,address:m[1].toUpperCase()+row} : null;
      };
      const first=cell(p.startCell),last=cell(p.endCell);
      if (!first || !last || last.column < first.column || last.row < first.row) return errResult('bad_args','position needs local start/end A1 cells with positive width/height within worksheet bounds');
      position={startCell:first.address,endCell:last.address};
    }
    if (args.font !== undefined && (typeof args.font !== 'string' || !args.font.trim() || args.font.length>200)) return errResult('bad_args','font must be a font family name');
    if (args.seriesColors !== undefined && (!Array.isArray(args.seriesColors) || args.seriesColors.length < 1 || args.seriesColors.length > 20 || args.seriesColors.some(c => typeof c !== 'string' || !/^#[0-9a-f]{6}$/i.test(c)))) return errResult('bad_args','seriesColors must contain 1-20 #RRGGBB colors');
    for (const key of ['fill','color']) if (args[key] !== undefined && (typeof args[key] !== 'string' || !/^#[0-9a-f]{6}$/i.test(args[key]))) return errResult('bad_args',key+' must be #RRGGBB');
    if (!isSetSupported('ExcelApi','1.1')) return errResult('requirement','insert_chart requires ExcelApi 1.1');
    try {
      const sourceRead=previewOnly ? await readRange({address:dataRange,limit:5000}) : null;
      if (sourceRead && !sourceRead.ok) return sourceRead;
      return await Excel.run(async ctx => {
        const source=excelRange(ctx,dataRange),ws=source.worksheet;
        source.load('address,rowCount,columnCount');ws.load('name');
        if (previewOnly) ws.charts.load('items/name,items/left,items/top,items/width,items/height');
        await ctx.sync();
        if (previewOnly) return okResult({host:'Excel',sheet:ws.name,dataRange:source.address,source:sourceRead.result,existingCharts:ws.charts.items.map(c=>({name:c.name,left:c.left,top:c.top,width:c.width,height:c.height})),requested:{type:typeName,title:args.title||null,position:position||null}});
        const chart=ws.charts.add(typeName,source,'Auto');
        if (args.title !== undefined) {chart.title.text=args.title;chart.title.visible=!!args.title;}
        if (position) chart.setPosition(position.startCell,position.endCell);
        if (args.font) chart.format.font.name=args.font;
        if (args.color) chart.format.font.color=args.color;
        if (args.fill) chart.format.fill.setSolidColor(args.fill);
        if (args.seriesColors) {
          chart.series.load('items/name'); await ctx.sync();
          chart.series.items.forEach((series,i) => series.format.fill.setSolidColor(args.seriesColors[i % args.seriesColors.length]));
        }
        if (args.font) chart.format.font.load('name');
        chart.load('name,left,top,width,height');
        chart.title.load('text,visible');
        await ctx.sync();
        return okResult({host:'Excel',chart:{name:chart.name,font:args.font?chart.format.font.name:null,type:typeName,sheet:ws.name,dataRange:source.address,title:chart.title.text,titleVisible:chart.title.visible,requestedPosition:position||null,geometry:{left:chart.left,top:chart.top,width:chart.width,height:chart.height}}});
      });
    } catch(e) {return errResult('execution',String(e&&e.message||e));}
  }

  async function addSheet(args) {
    const name = String(args.name || '');
    if (!name) return errResult('bad_args', 'name required');
    try {
      return await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.add(name);
        ws.load('name');
        await ctx.sync();
        return okResult({ host: 'Excel', added: ws.name });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function renameSheet(args) {
    const oldName = String(args.oldName || '');
    const newName = String(args.newName || '');
    if (!oldName || !newName) return errResult('bad_args', 'oldName and newName required');
    try {
      return await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(oldName);
        ws.name = newName;
        await ctx.sync();
        return okResult({ host: 'Excel', renamed: { from: oldName, to: newName } });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function deleteSheet(args) {
    const name = String(args.name || '');
    if (!name) return errResult('bad_args', 'name required');
    try {
      return await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(name);
        ws.load('name, position, visibility');
        await ctx.sync();
        if (args.dryRun) {
          return okResult({ host: 'Excel', dryRun: true, wouldDelete: { name: ws.name, position: ws.position, visibility: ws.visibility } });
        }
        ws.delete();
        await ctx.sync();
        return okResult({ host: 'Excel', deleted: name });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function applySort(args) {
    const address = String(args.address || '');
    const fields = args.fields;
    if (!address) return errResult('bad_args', 'address required');
    if (!Array.isArray(fields) || fields.length === 0) return errResult('bad_args', 'fields [{column,ascending}] required');
    try {
      return await Excel.run(async (ctx) => {
        const range = excelRange(ctx, address);
        range.sort.apply(fields.map((f) => ({ key: Number(f.column) - 1, ascending: f.ascending !== false })), false);
        await ctx.sync();
        return okResult({ host: 'Excel', sorted: fields.length });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function applyFilter(args) {
    const address = String(args.address || '');
    if (!address) return errResult('bad_args', 'address required');
    try {
      return await Excel.run(async (ctx) => {
        // autoFilter 属于 worksheet，且 range 必须在同一工作表上 → 从 address 解析目标表
        const range = excelRange(ctx, address);
        const bang = address.lastIndexOf('!');
        const ws = bang >= 0 ? ctx.workbook.worksheets.getItem(address.slice(0, bang)) : ctx.workbook.worksheets.getActiveWorksheet();
        if (!ws.autoFilter) return errResult('requirement', 'Excel autoFilter API (ExcelApi 1.9+) not available');
        if (args.columns !== undefined && !Array.isArray(args.columns)) return errResult('bad_args', 'columns must be an array');
        const filters = [];
        const filterTypes = ['Values', 'Custom', 'Dynamic', 'TopItems', 'TopPercent', 'BottomItems', 'BottomPercent', 'CellColor', 'FontColor', 'Icon'];
        for (const c of args.columns || []) {
          if (!c || !Number.isInteger(c.column) || c.column < 1) return errResult('bad_args', 'column must be a positive 1-based integer');
          const filterType = c.filterType || 'Values';
          if (!filterTypes.includes(filterType)) return errResult('bad_args', 'invalid filterType');
          let criteria = c.criteria;
          if (filterType === 'Values' && Array.isArray(criteria)) criteria = { values: criteria };
          if (!criteria || typeof criteria !== 'object' || Array.isArray(criteria)) return errResult('bad_args', 'criteria must be a FilterCriteria object (or an array for Values)');
          if (criteria.filterOn !== undefined && criteria.filterOn !== filterType) return errResult('bad_args', 'criteria.filterOn must match filterType');
          if (filterType === 'Values' && (!Array.isArray(criteria.values) || criteria.values.length === 0 || criteria.values.some((v) => typeof v !== 'string'))) return errResult('bad_args', 'Values criteria.values must be a non-empty string array');
          if (filterType === 'Custom' && typeof criteria.criterion1 !== 'string') return errResult('bad_args', 'Custom criteria.criterion1 must be a string');
          filters.push({ columnIndex: c.column - 1, criteria: { ...criteria, filterOn: filterType } });
        }
        range.load('columnCount'); await ctx.sync();
        if (filters.some((f) => f.columnIndex >= range.columnCount)) return errResult('bad_args', 'column exceeds range columnCount');
        if (filters.length) { for (const f of filters) ws.autoFilter.apply(range, f.columnIndex, f.criteria); }
        else ws.autoFilter.apply(range);
        await ctx.sync();
        return okResult({ host: 'Excel', filtered: true, columns: filters.length });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  // 单参数统计函数白名单：SUM/AVERAGE/COUNT/MAX/MIN/PRODUCT（用 workbook.functions 类型化求值，不写单元格）
  const EVAL_FUNCTIONS = { SUM: 'sum', AVERAGE: 'average', COUNT: 'count', MAX: 'max', MIN: 'min', PRODUCT: 'product' };

  async function evaluateFormula(args) {
    const formula = String(args.formula || '').trim();
    if (!formula) return errResult('bad_args', 'formula required (e.g. SUM(A1:A10))');
    const m = formula.match(/^([A-Z_]+)\((.*)\)$/i);
    const fnUpper = m ? m[1].toUpperCase() : '';
    const fnName = EVAL_FUNCTIONS[fnUpper];
    if (!fnName) return errResult('bad_args', `unsupported formula: ${formula}；支持 ${Object.keys(EVAL_FUNCTIONS).join('/')}(区域)`);
    try {
      return await Excel.run(async (ctx) => {
        const argStr = m[2].trim();
        if (!argStr) return errResult('bad_args', '函数缺少区域参数');
        // 逗号分隔的多区域是合法用法（如 SUM(A1:A5, C1:C5)），必须全部参与求值。
        // Excel.Functions 这批方法是变参签名（sum(...values)），所以用 ...rangeObjs 展开。
        // 早期版本只取 ranges[0]，多区域调用会静默返回「只算了第一段」的不完整结果。
        const ranges = argStr.split(',').map((s) => s.trim()).filter((s) => s.length > 0);
        if (!ranges.length) return errResult('bad_args', '函数缺少区域参数');
        const funcs = ctx.workbook.functions;
        if (!funcs) return errResult('requirement', 'Excel Functions API not available');
        const rangeObjs = ranges.map((s) => excelRange(ctx, s));
        const result = funcs[fnName](...rangeObjs);
        result.load('value');
        await ctx.sync();
        const value = result.value;
        return okResult({ host: 'Excel', formula, value, function: fnUpper, ranges });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }


  A.hosts.Excel = {
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
      // Excel 专有
      list_sheets: listSheets,
      excel_export_file: excelExportFile,
      excel_set_print_layout: (args) => excelPrintLayout(args, false),
      excel_render_range: excelRenderRange,
      excel_render_chart: excelRenderChart,
      read_range: readRange,
      write_range: writeRange,
      excel_set_freeze_panes: (args) => excelFreezePanes(args, false),
      format_range: formatRange,
      list_tables: listTables,
      create_table: createTable,
      delete_table: (args) => changeTableLifecycle(args, false),
      convert_table_to_range: (args) => changeTableLifecycle(args, true),
      read_data_validation: readDataValidation,
      set_data_validation: setDataValidation,
      clear_data_validation: clearDataValidation,
      read_conditional_formats: readConditionalFormats,
      add_conditional_format: addConditionalFormat,
      delete_conditional_format: deleteConditionalFormat,
      clear_conditional_formats: clearConditionalFormats,
      insert_chart: insertChart,
      add_sheet: addSheet,
      rename_sheet: renameSheet,
      delete_sheet: deleteSheet,
      apply_sort: applySort,
      apply_filter: applyFilter,
      evaluate_formula: evaluateFormula,
      validate_data: validateData,
    },
    selectTargets: {
      write_range: selectExcelTargetRange,
      format_range: selectExcelTargetRange,
      create_table: selectExcelTargetRange,
      add_conditional_format: selectExcelTargetRange,
      delete_conditional_format: selectExcelTargetRange,
      clear_conditional_formats: selectExcelTargetRange,
      set_data_validation: selectExcelTargetRange,
      clear_data_validation: selectExcelTargetRange,
    },
    // 机制级 re-read：覆盖类操作必须先预览当前状态，确认后带 confirm:true 才执行。
    previewReaders: {
      write_selection: () => readSelection({}),
      replace_all: async (a) => {
        const preview = await readDocument({});
        if (!preview.ok) return preview;
        const note = preview.result.truncated ? '[仅预览当前状态：已截断至最多 5000 个单元格；非完整匹配结果]' : '[仅预览当前状态；非匹配结果]';
        return okResult({ ...preview.result, text: note + '\n' + preview.result.text, search: String(a.search ?? ''), replace: String(a.replace ?? ''), previewOnly: true, matching: 'Office replaceAll matching is evaluated only during execution; text is a bounded current-state preview.' });
      },
      format_selection: () => readSelection({ withStyles: true }),
      add_comment: () => readSelection({}),
      write_range: (a) => readRange({ address: a.address }),
      excel_set_freeze_panes: (a) => excelFreezePanes(a, true),
      excel_set_print_layout: (a) => excelPrintLayout(a, true),
      format_range: (a) => formatRange(a, true),
      insert_chart: (a) => insertChart(a,true),
      add_sheet: () => listSheets({}),
      rename_sheet: () => listSheets({}),
      delete_sheet: (a) => deleteSheet({ ...a, dryRun: true }),
      apply_sort: (a) => readRange({ address: a.address }),
      apply_filter: (a) => readRange({ address: a.address }),
      create_table: (a) => readRange({ address: a.address }),
      delete_table: (a) => listTables(a),
      convert_table_to_range: (a) => listTables(a),
      add_conditional_format: (a) => readConditionalFormats(a),
      delete_conditional_format: (a) => readConditionalFormats(a),
      clear_conditional_formats: (a) => readConditionalFormats(a),
      set_data_validation: (a) => readDataValidation(a),
      clear_data_validation: (a) => readDataValidation(a),
    },
    afterContent: {
      write_range: (a) => JSON.stringify(a.values || a.formulas || []),
      excel_set_freeze_panes: (a) => JSON.stringify({ sheet: a.sheet, mode: a.mode, count: a.count, range: a.range }),
      format_range: (a) => JSON.stringify(a),
      excel_set_print_layout: (a) => JSON.stringify(a),
      rename_sheet: (a) => `${a.oldName} → ${a.newName}`,
      delete_sheet: (a) => `删除工作表：${a.name}`,
      apply_sort: (a) => JSON.stringify(a),
      apply_filter: (a) => JSON.stringify(a),
    },
  };
})();
