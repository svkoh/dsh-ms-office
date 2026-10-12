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
              cell.load('text, format/fill/color, format/font/name, format/font/size, format/font/bold, format/font/italic, format/font/color, format/numberFormat');
              cellRefs.push(cell);
            }
          }
          await ctx.sync();
          const cells = [];
          for (let r = 0; r < rows; r++) {
            for (let c = 0; c < cols; c++) {
              const cell = cellRefs[r * cols + c];
              const f = cell.format.font;
              cells.push({ row: r + 1, col: c + 1, value: cell.text, font: { name: f.name, size: f.size, bold: f.bold, italic: f.italic, color: f.color }, fill: cell.format.fill.color, numberFormat: cell.format.numberFormat });
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
        range.load('text, rowCount, columnCount');
        await ctx.sync();
        const rows = range.text || [];
        const MAX_CELLS = 5000;
        const out = [];
        let cellCount = 0;
        let truncated = false;
        for (let r = 0; r < rows.length; r++) {
          cellCount += rows[r].length;
          if (cellCount > MAX_CELLS) { truncated = true; break; }
          out.push(rows[r].join('\t'));
        }
        return okResult({ text: out.join('\n'), host: 'Excel', sheet: ws.name, rows: rows.length, cols: range.columnCount, truncated });
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
            cell.load('text, format/fill/color, format/font/name, format/font/size, format/font/bold, format/font/italic, format/font/color, format/numberFormat');
            cellRefs.push(cell);
          }
        }
        await ctx.sync();
        const cells = [];
        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < cols; c++) {
            const cell = cellRefs[r * cols + c];
            const f = cell.format.font;
            cells.push({ row: r + 1, col: c + 1, value: cell.text, font: { name: f.name, size: f.size, bold: f.bold, italic: f.italic, color: f.color }, fill: cell.format.fill.color, numberFormat: cell.format.numberFormat });
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

  async function readRange(args) {
    const address = String(args.address || '');
    if (!address) return errResult('bad_args', 'address (e.g. Sheet1!A1:B10) required');
    try {
      return await Excel.run(async (ctx) => {
        const range = excelRange(ctx, address);
        range.load('address, values, formulas, numberFormat, rowCount, columnCount');
        await ctx.sync();
        const limit = args.limit ? Number(args.limit) : 5000;
        const total = (range.rowCount || 0) * (range.columnCount || 0);
        const truncated = total > limit;
        const sliceRows = truncated ? Math.max(1, Math.floor(limit / (range.columnCount || 1))) : range.rowCount;
        const values = (range.values || []).slice(0, sliceRows);
        return okResult({ host: 'Excel', address: range.address || address, rows: range.rowCount, cols: range.columnCount, truncated, values, formulas: (range.formulas || []).slice(0, sliceRows), numberFormat: (range.numberFormat || []).slice(0, sliceRows) });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
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

  async function formatRange(args) {
    const address = String(args.address || '');
    if (!address) return errResult('bad_args', 'address required');
    try {
      return await Excel.run(async (ctx) => {
        const range = excelRange(ctx, address);
        const fmt = range.format;
        if (args.font) fmt.font.name = String(args.font);
        if (args.size) fmt.font.size = Number(args.size);
        if (args.bold !== undefined) fmt.font.bold = !!args.bold;
        if (args.fill) fmt.fill.color = String(args.fill);
        if (args.numberFormat) fmt.numberFormat = String(args.numberFormat);
        if (args.autoFit) fmt.autoFitColumns();
        if (args.tableStyle) {
          if (!args.tableName) return errResult('bad_args', 'tableName required when tableStyle is provided');
          ctx.workbook.tables.getItem(String(args.tableName)).style = String(args.tableStyle);
        }
        await ctx.sync();
        range.load('address,numberFormat,format/fill/color,format/font/name,format/font/size,format/font/bold'); await ctx.sync();
        return okResult({ host: 'Excel', formatted: true, address: range.address || address, verification: { numberFormat: range.numberFormat, fill: fmt.fill.color, font: { name: fmt.font.name, size: fmt.font.size, bold: fmt.font.bold }, tableName: args.tableName || null, tableStyle: args.tableStyle || null } });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function listTables(args) {
    if (!isSetSupported('ExcelApi', '1.1')) return errResult('requirement', 'list_tables requires ExcelApi 1.1+');
    try { return await Excel.run(async (ctx) => {
      const collection = args.sheet ? ctx.workbook.worksheets.getItem(String(args.sheet)).tables : ctx.workbook.tables;
      collection.load('items/id,items/name,items/style,items/showHeaders,items/showTotals'); await ctx.sync();
      const refs = collection.items.map((table) => { const range = table.getRange(); range.load('address,rowCount,columnCount,values'); return { table, range }; }); await ctx.sync();
      const limit = Math.min(5000, Math.max(1, Number(args.limit) || 500));
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

  const CHART_TYPES = {
    line: 'Line', column: 'ColumnClustered', bar: 'BarClustered', pie: 'Pie',
    area: 'Area', scatter: 'XYScatter', doughnut: 'Doughnut', radar: 'Radar',
  };

  async function insertChart(args) {
    const dataRange = String(args.dataRange || '');
    if (!dataRange) return errResult('bad_args', 'dataRange (e.g. Sheet1!A1:B10) required');
    try {
      return await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getActiveWorksheet();
        const typeName = CHART_TYPES[String(args.type || '').toLowerCase()] || 'ColumnClustered';
        // sourceData 传 Range 对象（字符串跨表引用在此环境报 sourceData 错误）
        const source = excelRange(ctx, dataRange);
        const chart = ws.charts.add(typeName, source, 'auto');
        if (args.title) {
          chart.title.text = args.title;
        }
        await ctx.sync();
        return okResult({ host: 'Excel', chart: { type: typeName, dataRange, title: args.title || null } });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
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
        const criteria = Array.isArray(args.columns)
          ? args.columns.map((c) => ({ column: Number(c.column), filterType: c.filterType || 'Values', criteria1: c.criteria1, criteria2: c.criteria2 }))
          : undefined;
        ws.autoFilter.apply(range, criteria);
        await ctx.sync();
        return okResult({ host: 'Excel', filtered: true, columns: criteria ? criteria.length : 0 });
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
      read_range: readRange,
      write_range: writeRange,
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
      replace_all: (a) => replaceAll({ ...a, dryRun: true }),
      format_selection: () => readSelection({ withStyles: true }),
      add_comment: () => readSelection({}),
      write_range: (a) => readRange({ address: a.address }),
      format_range: (a) => readRange({ address: a.address }),
      insert_chart: () => listSheets({}),
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
      format_range: (a) => JSON.stringify(a),
      rename_sheet: (a) => `${a.oldName} → ${a.newName}`,
      delete_sheet: (a) => `删除工作表：${a.name}`,
      apply_sort: (a) => JSON.stringify(a),
      apply_filter: (a) => JSON.stringify(a),
    },
  };
})();
