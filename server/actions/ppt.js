// ppt.js — PowerPoint（pptx）宿主的全部 action 实现
//
// 分片契约：本文件是自包含 IIFE，只通过 window.__OFFICE_ACTIONS__ 与其它分片交换东西，
// 与其它分片不共享词法作用域。挂载点：
//   A.hosts.PowerPoint = { actions, selectTargets, previewReaders, afterContent, environmentProbe? }
// 这里只放 PowerPoint 的代码：分片之后「按宿主分叉」的写法既不需要（文件名就是宿主）也不允许（函数签名里已经没有 host 参数）。
(function () {
  const A = window.__OFFICE_ACTIONS__;
  const U = A.util;
  const okResult = U.okResult, errResult = U.errResult, readPptFile = U.readPptFile, unzip = U.unzip, inflateEntry = U.inflateEntry, getSelectionTextAsync = U.getSelectionTextAsync;
  const isSetSupported = A.isSetSupported, validateMatrix = A.validateMatrix;

  // ================= 跨宿主 action 的 PowerPoint 实现 =================
  // PowerPoint 没有 Word/Excel 那样的对象模型选区，两个 action 都退到 Common API。
  // （write_selection 三宿主实现完全一致，在 common.js。）

  async function readSelection(args) {
    const a = args || {};
    if (a.withStyles) return errResult('unsupported_host', 'read_selection withStyles not supported in PowerPoint');
    return getSelectionTextAsync();
  }

  // PPT 没有"全文"概念：read_document = 逐页抽文字（见 readPptDocument）
  async function readDocument() {
    return readPptDocument();
  }

  // ================= PPT 组 =================

  const PPT_DESIGN_WIDTH = 960;
  const PPT_DESIGN_HEIGHT = 540;
  const PPT_MAX_SLIDES = 50;
  const PPT_MAX_ELEMENTS = 500;
  const PPT_COLOR_RE = /^#[0-9A-F]{6}$/i;
  const PPT_ELEMENT_TYPES = ['text', 'shape', 'line'];
  const PPT_ALIGNMENTS = ['Left', 'Center', 'Right', 'Justify', 'JustifyLow', 'Distributed', 'ThaiDistributed'];
  const PPT_VERTICAL_ALIGNMENTS = ['Top', 'Middle', 'Bottom', 'TopCentered', 'MiddleCentered', 'BottomCentered'];
  const PPT_CONNECTORS = ['Straight', 'Elbow', 'Curve'];
  const PPT_DASH_STYLES = ['Dash', 'DashDot', 'DashDotDot', 'LongDash', 'LongDashDot', 'RoundDot', 'Solid', 'SquareDot', 'LongDashDotDot', 'SystemDash', 'SystemDot', 'SystemDashDot'];
  const PPT_TABLE_STYLES = new Set(['NoStyleNoGrid','NoStyleTableGrid','LightStyle1','LightStyle2','LightStyle3','MediumStyle1','MediumStyle2','MediumStyle3','MediumStyle4','DarkStyle1','DarkStyle2']);
  for (const base of ['ThemedStyle1','ThemedStyle2','LightStyle1','LightStyle2','LightStyle3','MediumStyle1','MediumStyle2','MediumStyle3','MediumStyle4']) for (let accent = 1; accent <= 6; accent++) PPT_TABLE_STYLES.add(`${base}Accent${accent}`);
  for (const base of ['DarkStyle1','DarkStyle2']) for (let accent = 1; accent <= 3; accent++) PPT_TABLE_STYLES.add(`${base}Accent${accent}`);

  function pptFiniteNumber(value, path, min, max) {
    const n = Number(value);
    if (!Number.isFinite(n) || n < min || n > max) throw new Error(`${path} must be a number from ${min} to ${max}`);
    return n;
  }

  function pptOptionalColor(value, path) {
    if (value === undefined || value === null || value === '') return null;
    const color = String(value);
    if (!PPT_COLOR_RE.test(color)) throw new Error(`${path} must be #RRGGBB`);
    return color.toUpperCase();
  }

  function validatePptBuildArgs(args) {
    if (!args || !Array.isArray(args.slides) || args.slides.length === 0) throw new Error('slides must be a non-empty array');
    if (args.slides.length > PPT_MAX_SLIDES) throw new Error(`slides exceeds limit ${PPT_MAX_SLIDES}`);
    const designWidth = pptFiniteNumber(args.designWidth === undefined ? PPT_DESIGN_WIDTH : args.designWidth, 'designWidth', 100, 10000);
    const designHeight = pptFiniteNumber(args.designHeight === undefined ? PPT_DESIGN_HEIGHT : args.designHeight, 'designHeight', 100, 10000);
    let totalElements = 0;
    const slides = args.slides.map((slide, slideIndex) => {
      if (!slide || !Array.isArray(slide.elements)) throw new Error(`slides[${slideIndex}].elements must be an array`);
      totalElements += slide.elements.length;
      if (totalElements > PPT_MAX_ELEMENTS) throw new Error(`total elements exceeds limit ${PPT_MAX_ELEMENTS}`);
      const elements = slide.elements.map((element, elementIndex) => {
        const path = `slides[${slideIndex}].elements[${elementIndex}]`;
        if (!element || !PPT_ELEMENT_TYPES.includes(element.type)) throw new Error(`${path}.type must be ${PPT_ELEMENT_TYPES.join('|')}`);
        const out = {
          type: element.type,
          x: pptFiniteNumber(element.x, `${path}.x`, 0, designWidth),
          y: pptFiniteNumber(element.y, `${path}.y`, 0, designHeight),
          w: pptFiniteNumber(element.w, `${path}.w`, 0.1, designWidth),
          h: pptFiniteNumber(element.h, `${path}.h`, 0.1, designHeight),
          rotation: element.rotation === undefined ? null : pptFiniteNumber(element.rotation, `${path}.rotation`, -360, 360),
        };
        if (out.x + out.w > designWidth + 0.01 || out.y + out.h > designHeight + 0.01) throw new Error(`${path} exceeds design bounds ${designWidth}x${designHeight}`);
        if (element.type === 'text' || element.type === 'shape') {
          out.text = element.text === undefined || element.text === null ? '' : String(element.text);
          out.font = element.font === undefined || element.font === null ? null : String(element.font);
          out.size = pptFiniteNumber(element.size === undefined ? 18 : element.size, `${path}.size`, 1, 400);
          out.bold = element.bold === undefined ? false : !!element.bold;
          out.italic = element.italic === undefined ? false : !!element.italic;
          out.color = pptOptionalColor(element.color === undefined ? '#000000' : element.color, `${path}.color`);
          out.fill = pptOptionalColor(element.fill, `${path}.fill`);
          out.fillTransparency = pptFiniteNumber(element.fillTransparency === undefined ? 0 : element.fillTransparency, `${path}.fillTransparency`, 0, 1);
          out.lineColor = pptOptionalColor(element.lineColor, `${path}.lineColor`);
          out.lineWidth = pptFiniteNumber(element.lineWidth === undefined ? 1 : element.lineWidth, `${path}.lineWidth`, 0, 50);
          out.align = element.align === undefined ? 'Left' : String(element.align);
          out.valign = element.valign === undefined ? 'Top' : String(element.valign);
          if (!PPT_ALIGNMENTS.includes(out.align)) throw new Error(`${path}.align is invalid`);
          if (!PPT_VERTICAL_ALIGNMENTS.includes(out.valign)) throw new Error(`${path}.valign is invalid`);
          const margins = element.margins || {};
          out.margins = {
            left: pptFiniteNumber(margins.left === undefined ? 0 : margins.left, `${path}.margins.left`, 0, 100),
            right: pptFiniteNumber(margins.right === undefined ? 0 : margins.right, `${path}.margins.right`, 0, 100),
            top: pptFiniteNumber(margins.top === undefined ? 0 : margins.top, `${path}.margins.top`, 0, 100),
            bottom: pptFiniteNumber(margins.bottom === undefined ? 0 : margins.bottom, `${path}.margins.bottom`, 0, 100),
          };
          out.autoFit = element.autoFit === undefined ? false : !!element.autoFit;
          out.runs = Array.isArray(element.runs) ? element.runs.map((run, runIndex) => {
            const runPath = `${path}.runs[${runIndex}]`;
            if (!run || typeof run !== 'object') throw new Error(`${runPath} must be an object`);
            const start = pptFiniteNumber(run.start, `${runPath}.start`, 0, out.text.length);
            const length = pptFiniteNumber(run.length, `${runPath}.length`, 1, Math.max(1, out.text.length));
            if (!Number.isInteger(start) || !Number.isInteger(length) || start + length > out.text.length) throw new Error(`${runPath} exceeds text bounds`);
            return {
              start, length,
              color: pptOptionalColor(run.color, `${runPath}.color`),
              bold: run.bold === undefined ? null : !!run.bold,
              italic: run.italic === undefined ? null : !!run.italic,
              size: run.size === undefined ? null : pptFiniteNumber(run.size, `${runPath}.size`, 1, 400),
              font: run.font === undefined || run.font === null ? null : String(run.font),
            };
          }) : [];
          if (element.type === 'shape') {
            out.shapeType = element.shapeType === undefined ? 'Rectangle' : String(element.shapeType);
            if (!/^[A-Za-z][A-Za-z0-9]*$/.test(out.shapeType)) throw new Error(`${path}.shapeType is invalid`);
          }
        } else {
          out.connectorType = element.connectorType === undefined ? 'Straight' : String(element.connectorType);
          if (!PPT_CONNECTORS.includes(out.connectorType)) throw new Error(`${path}.connectorType is invalid`);
          out.color = pptOptionalColor(element.color === undefined ? '#808080' : element.color, `${path}.color`);
          out.lineWidth = pptFiniteNumber(element.lineWidth === undefined ? 1 : element.lineWidth, `${path}.lineWidth`, 0.1, 50);
          out.dashStyle = element.dashStyle === undefined ? 'Solid' : String(element.dashStyle);
          if (!PPT_DASH_STYLES.includes(out.dashStyle)) throw new Error(`${path}.dashStyle is invalid`);
          out.transparency = pptFiniteNumber(element.transparency === undefined ? 0 : element.transparency, `${path}.transparency`, 0, 1);
        }
        return out;
      });
      return { elements };
    });
    return { replaceAll: !!args.replaceAll, designWidth, designHeight, totalElements, slides };
  }

  function pptShapeOptions(element, sx, sy) {
    return { left: element.x * sx, top: element.y * sy, width: element.w * sx, height: element.h * sy };
  }

  function applyPptTextFormat(shape, element) {
    const frame = shape.textFrame;
    frame.leftMargin = element.margins.left;
    frame.rightMargin = element.margins.right;
    frame.topMargin = element.margins.top;
    frame.bottomMargin = element.margins.bottom;
    frame.wordWrap = true;
    frame.verticalAlignment = element.valign;
    frame.autoSizeSetting = element.autoFit ? 'AutoSizeTextToFitShape' : 'AutoSizeNone';
    const range = frame.textRange;
    if (element.font) range.font.name = element.font;
    range.font.size = element.size;
    range.font.bold = element.bold;
    range.font.italic = element.italic;
    range.font.color = element.color;
    range.paragraphFormat.horizontalAlignment = element.align;
    for (const run of element.runs || []) {
      const runRange = range.getSubstring(run.start, run.length);
      if (run.font) runRange.font.name = run.font;
      if (run.size !== null) runRange.font.size = run.size;
      if (run.bold !== null) runRange.font.bold = run.bold;
      if (run.italic !== null) runRange.font.italic = run.italic;
      if (run.color) runRange.font.color = run.color;
    }
  }

  function applyPptShapeFormat(shape, element) {
    if (element.fill) {
      shape.fill.setSolidColor(element.fill);
      shape.fill.transparency = element.fillTransparency;
    } else {
      shape.fill.clear();
    }
    shape.lineFormat.visible = !!element.lineColor && element.lineWidth > 0;
    if (shape.lineFormat.visible) {
      shape.lineFormat.color = element.lineColor;
      shape.lineFormat.weight = element.lineWidth;
    }
  }

  async function pptAddSlides(args) {
    if (!isSetSupported('PowerPointApi', '1.4')) return errResult('requirement', 'ppt_add_slides requires PowerPointApi 1.4+');
    const count = Number(args.count === undefined ? 1 : args.count), clearPlaceholders = args.clearPlaceholders === true;
    if (!Number.isInteger(count) || count < 1 || count > 50) return errResult('bad_args', 'count must be an integer from 1 to 50');
    try { return await PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides; slides.load('items'); await ctx.sync(); const start = slides.items.length;
      for (let i = 0; i < count; i++) slides.add(); await ctx.sync(); slides.load('items'); await ctx.sync();
      const created = slides.items.slice(start); for (const slide of created) { slide.load('id'); if (clearPlaceholders) slide.shapes.load('items'); } await ctx.sync();
      if (clearPlaceholders) { for (const slide of created) for (const shape of slide.shapes.items) shape.delete(); await ctx.sync(); }
      return okResult({ host: 'PowerPoint', added: created.map((slide) => ({ id: slide.id })), slideCount: slides.items.length, clearPlaceholders });
    }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function pptAddElements(args) {
    if (!isSetSupported('PowerPointApi', '1.4')) return errResult('requirement', 'ppt_add_elements requires PowerPointApi 1.4+');
    const coordinateSpace = args.coordinateSpace || {}, targetSize = args.targetSize || {}; let spec, targetWidth, targetHeight;
    try { spec = validatePptBuildArgs({ designWidth: coordinateSpace.width, designHeight: coordinateSpace.height, slides: [{ elements: args.elements }] }); targetWidth = pptFiniteNumber(targetSize.width, 'targetSize.width', 100, 10000); targetHeight = pptFiniteNumber(targetSize.height, 'targetSize.height', 100, 10000); } catch (e) { return errResult('bad_args', String(e && e.message || e)); }
    const slideId = String(args.slideId || ''); if (!slideId) return errResult('bad_args', 'slideId required');
    try { return await PowerPoint.run(async (ctx) => {
      const slide = ctx.presentation.slides.getItem(slideId), sx = targetWidth / spec.designWidth, sy = targetHeight / spec.designHeight, counts = { text: 0, shape: 0, line: 0 }, created = [];
      for (let index = 0; index < spec.slides[0].elements.length; index++) { const element = spec.slides[0].elements[index]; let shape;
        try { if (element.type === 'text') { shape = slide.shapes.addTextBox(element.text, pptShapeOptions(element, sx, sy)); applyPptTextFormat(shape, element); } else if (element.type === 'shape') { shape = slide.shapes.addGeometricShape(element.shapeType, pptShapeOptions(element, sx, sy)); applyPptShapeFormat(shape, element); shape.textFrame.textRange.text = element.text; applyPptTextFormat(shape, element); } else { shape = slide.shapes.addLine(element.connectorType, pptShapeOptions(element, sx, sy)); shape.lineFormat.visible = true; shape.lineFormat.color = element.color; shape.lineFormat.weight = element.lineWidth; shape.lineFormat.dashStyle = element.dashStyle; shape.lineFormat.transparency = element.transparency; } if (element.rotation !== null) shape.rotation = element.rotation; shape.load('id'); created.push(shape); counts[element.type]++; } catch (e) { throw new Error(`elements[${index}]: ${e && e.message || e}`); }
      }
      await ctx.sync(); return okResult({ host: 'PowerPoint', slideId, elements: created.map((shape) => ({ id: shape.id })), counts, coordinateSpace: { width: spec.designWidth, height: spec.designHeight }, targetSize: { width: targetWidth, height: targetHeight } });
    }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function pptDeleteShapes(args) {
    if (!isSetSupported('PowerPointApi', '1.4')) return errResult('requirement', 'ppt_delete_shapes requires PowerPointApi 1.4+');
    const slideId = String(args && args.slideId || '');
    const shapeIds = Array.isArray(args && args.shapeIds) ? Array.from(new Set(args.shapeIds.map(String).filter(Boolean))) : [];
    if (!slideId || !shapeIds.length) return errResult('bad_args', 'slideId and non-empty shapeIds are required');
    try { return await PowerPoint.run(async (ctx) => {
      const slide = ctx.presentation.slides.getItem(slideId); const shapes = slide.shapes; shapes.load('items/id'); await ctx.sync();
      const byId = new Map(shapes.items.map((shape) => [String(shape.id), shape])); const missing = shapeIds.filter((id) => !byId.has(id));
      if (missing.length) return errResult('not_found', `shape IDs not found on slide ${slideId}: ${missing.join(', ')}`);
      for (const id of shapeIds) byId.get(id).delete();
      await ctx.sync(); shapes.load('items/id'); await ctx.sync();
      return okResult({ host: 'PowerPoint', slideId, deletedShapeIds: shapeIds, remainingShapeIds: shapes.items.map((shape) => String(shape.id)), remainingCount: shapes.items.length });
    }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function pptDeleteShapesPreview(args) {
    if (!isSetSupported('PowerPointApi', '1.4')) return errResult('requirement', 'ppt_delete_shapes requires PowerPointApi 1.4+');
    const slideId = String(args && args.slideId || ''); const shapeIds = Array.isArray(args && args.shapeIds) ? Array.from(new Set(args.shapeIds.map(String).filter(Boolean))) : [];
    if (!slideId || !shapeIds.length) return errResult('bad_args', 'slideId and non-empty shapeIds are required');
    try { return await PowerPoint.run(async (ctx) => { const shapes = ctx.presentation.slides.getItem(slideId).shapes; shapes.load('items/id,name,type,left,top,width,height'); await ctx.sync(); const byId = new Map(shapes.items.map((shape) => [String(shape.id), shape])); const matched = shapeIds.filter((id) => byId.has(id)).map((id) => { const shape = byId.get(id); return { id, name: shape.name, type: shape.type, left: shape.left, top: shape.top, width: shape.width, height: shape.height }; }); const missing = shapeIds.filter((id) => !byId.has(id)); return okResult({ host: 'PowerPoint', slideId, currentCount: shapes.items.length, matchedShapes: matched, missingShapeIds: missing, remainingCount: shapes.items.length - matched.length }); }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function pptDeleteSlides(args) {
    if (!isSetSupported('PowerPointApi', '1.2')) return errResult('requirement', 'ppt_delete_slides requires PowerPointApi 1.2+');
    const slideIds = Array.isArray(args && args.slideIds) ? Array.from(new Set(args.slideIds.map(String).filter(Boolean))) : [];
    if (!slideIds.length) return errResult('bad_args', 'slideIds must be a non-empty array');
    try { return await PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides; slides.load('items/id'); await ctx.sync();
      const byId = new Map(slides.items.map((slide) => [String(slide.id), slide]));
      const missing = slideIds.filter((id) => !byId.has(id));
      if (missing.length) return errResult('not_found', `slide IDs not found: ${missing.join(', ')}`);
      if (slideIds.length >= slides.items.length) return errResult('bad_args', 'cannot delete every slide in the presentation');
      for (const id of slideIds) byId.get(id).delete();
      await ctx.sync(); slides.load('items/id'); await ctx.sync();
      return okResult({ host: 'PowerPoint', deletedSlideIds: slideIds, remainingSlideIds: slides.items.map((slide) => String(slide.id)), remainingCount: slides.items.length });
    }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function pptDeleteSlidesPreview(args) {
    if (!isSetSupported('PowerPointApi', '1.2')) return errResult('requirement', 'ppt_delete_slides requires PowerPointApi 1.2+');
    const slideIds = Array.isArray(args && args.slideIds) ? Array.from(new Set(args.slideIds.map(String).filter(Boolean))) : [];
    if (!slideIds.length) return errResult('bad_args', 'slideIds must be a non-empty array');
    try { return await PowerPoint.run(async (ctx) => { const slides = ctx.presentation.slides; slides.load('items/id'); await ctx.sync(); const existing = new Set(slides.items.map((slide) => String(slide.id))); const matched = slideIds.filter((id) => existing.has(id)); const missing = slideIds.filter((id) => !existing.has(id)); return okResult({ host: 'PowerPoint', currentCount: slides.items.length, matchedSlideIds: matched, missingSlideIds: missing, remainingCount: slides.items.length - matched.length }); }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function pptInspectSlides(args) {
    if (!isSetSupported('PowerPointApi', '1.4')) return errResult('requirement', 'ppt_inspect_slides requires PowerPointApi 1.4+');
    try { return await PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides; slides.load('items/id'); await ctx.sync();
      const wanted = Array.isArray(args.slideIds) && args.slideIds.length ? new Set(args.slideIds.map(String)) : null;
      const refs = slides.items.filter((s) => !wanted || wanted.has(String(s.id)));
      const hasZOrder = isSetSupported('PowerPointApi', '1.8');
      for (const slide of refs) slide.shapes.load(hasZOrder ? 'items/id,items/name,items/type,items/left,items/top,items/width,items/height,items/zOrderPosition' : 'items/id,items/name,items/type,items/left,items/top,items/width,items/height');
      await ctx.sync();
      return okResult({ host: 'PowerPoint', slides: refs.map((slide) => ({ id: slide.id, shapes: slide.shapes.items.map((shape) => ({ id: shape.id, name: shape.name, type: shape.type, left: shape.left, top: shape.top, width: shape.width, height: shape.height, ...(hasZOrder ? { zOrderPosition: shape.zOrderPosition } : {}) })) })) });
    }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  function pptInsertImage(args) {
    if (!isSetSupported('ImageCoercion', '1.1')) return Promise.resolve(errResult('requirement', 'ppt_insert_image requires ImageCoercion 1.1+'));
    const base64 = String(args.base64 || '').replace(/^data:image\/\w+;base64,/, ''); if (!base64) return Promise.resolve(errResult('bad_args', 'base64 required'));
    const mode = String(args.mode || 'positioned'); if (!['positioned','selectionOrPlaceholder'].includes(mode)) return Promise.resolve(errResult('bad_args', 'mode must be positioned or selectionOrPlaceholder'));
    const options = { coercionType: Office.CoercionType.Image };
    try {
      if (mode === 'positioned') {
        if (args.left === undefined || args.top === undefined) throw new Error('positioned mode requires left and top');
        options.imageLeft = pptFiniteNumber(args.left, 'left', -10000, 10000); options.imageTop = pptFiniteNumber(args.top, 'top', -10000, 10000);
      }
      if (args.width !== undefined) options.imageWidth = pptFiniteNumber(args.width, 'width', 1, 10000);
      if (args.height !== undefined) options.imageHeight = pptFiniteNumber(args.height, 'height', 1, 10000);
    } catch (e) { return Promise.resolve(errResult('bad_args', e.message)); }
    return new Promise((resolve) => Office.context.document.setSelectedDataAsync(base64, options, (r) => resolve(r.status === Office.AsyncResultStatus.Succeeded ? okResult({ host: 'PowerPoint', inserted: true, mode, geometry: { left: options.imageLeft, top: options.imageTop, width: options.imageWidth, height: options.imageHeight }, cropSupported: false }) : errResult('execution', r.error && r.error.message))));
  }

  async function pptInsertTable(args) {
    if (!isSetSupported('PowerPointApi', '1.8')) return errResult('requirement', 'ppt_insert_table requires PowerPointApi 1.8+');
    let values; try { values = validateMatrix(args.values, 'values', 10000); } catch (e) { return errResult('bad_args', e.message); }
    if (args.style !== undefined && (typeof args.style !== 'string' || !PPT_TABLE_STYLES.has(args.style))) return errResult('bad_args', 'style must be a valid PowerPoint.TableStyle enum string');
    if (args.style && !isSetSupported('PowerPointApi', '1.9')) return errResult('requirement', 'PowerPoint table style requires PowerPointApi 1.9+');
    try { return await PowerPoint.run(async (ctx) => {
      const slide = ctx.presentation.slides.getItem(String(args.slideId));
      const normalizedValues = values.map((row) => row.map((value) => String(value)));
      const rowCount = normalizedValues.length, columnCount = normalizedValues[0].length;
      if (args.columns !== undefined && (!Array.isArray(args.columns) || args.columns.length !== columnCount)) return errResult('bad_args', `columns must contain ${columnCount} items`);
      if (args.rows !== undefined && (!Array.isArray(args.rows) || args.rows.length !== rowCount)) return errResult('bad_args', `rows must contain ${rowCount} items`);
      if (args.specificCellProperties !== undefined) { try { validateMatrix(args.specificCellProperties, 'specificCellProperties', 10000); } catch (e) { return errResult('bad_args', e.message); } if (args.specificCellProperties.length !== rowCount || args.specificCellProperties[0].length !== columnCount) return errResult('bad_args', `specificCellProperties must be ${rowCount}x${columnCount}`); }
      const options = { specificCellProperties: normalizedValues.map((row, r) => row.map((text, c) => ({ ...((args.specificCellProperties && args.specificCellProperties[r][c]) || {}), text }))) };
      for (const key of ['left','top','width','height']) if (args[key] !== undefined) options[key] = pptFiniteNumber(args[key], key, 0, 10000);
      for (const key of ['columns','rows','mergedAreas','uniformCellProperties','style']) if (args[key] !== undefined) options[key] = args[key];
      const shape = slide.shapes.addTable(rowCount, columnCount, options); shape.load('id,name,type,left,top,width,height'); const table = shape.getTable(); table.load('rowCount,columnCount,values'); await ctx.sync();
      return okResult({ host: 'PowerPoint', slideId: args.slideId, shape: { id: shape.id, name: shape.name, type: shape.type, left: shape.left, top: shape.top, width: shape.width, height: shape.height }, table: { rows: table.rowCount, cols: table.columnCount, values: table.values }, verified: JSON.stringify(table.values) === JSON.stringify(normalizedValues) });
    }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function pptListPlaceholders(args) {
    if (!isSetSupported('PowerPointApi', '1.8')) return errResult('requirement', 'ppt_list_placeholders requires PowerPointApi 1.8+');
    try { return await PowerPoint.run(async (ctx) => { const slides=ctx.presentation.slides; slides.load('items/id'); await ctx.sync(); const refs=args.slideId?[slides.getItem(String(args.slideId))]:slides.items; for(const slide of refs)slide.shapes.load('items/id,items/name,items/type,items/left,items/top,items/width,items/height,items/placeholderFormat/type,items/placeholderFormat/containedType'); await ctx.sync(); const out=[]; for(const slide of refs)for(const shape of slide.shapes.items){try{const p=shape.placeholderFormat;if(p&&p.type)out.push({slideId:slide.id,id:shape.id,name:shape.name,left:shape.left,top:shape.top,width:shape.width,height:shape.height,type:p.type,containedType:p.containedType});}catch(e){}} return okResult({host:'PowerPoint',placeholders:out,warning:'Office.js cannot guarantee insertion into an arbitrary placeholder ID; selectionOrPlaceholder uses current PowerPoint selection.'}); }); } catch(e){return errResult('execution',String(e&&e.message||e));}
  }

  async function pptGroupShapes(args) {
    if (!isSetSupported('PowerPointApi', '1.8')) return errResult('requirement', 'ppt_group_shapes requires PowerPointApi 1.8+');
    if (!Array.isArray(args.shapeIds) || args.shapeIds.length < 2) return errResult('bad_args', 'shapeIds requires at least two IDs');
    try { return await PowerPoint.run(async (ctx) => { const slide = ctx.presentation.slides.getItem(String(args.slideId)); const group = slide.shapes.addGroup(args.shapeIds.map(String)); if (args.name) group.name = String(args.name); group.load('id,name,type,left,top,width,height,zOrderPosition'); await ctx.sync(); return okResult({ host: 'PowerPoint', slideId: args.slideId, group: { id: group.id, name: group.name, type: group.type, left: group.left, top: group.top, width: group.width, height: group.height, zOrderPosition: group.zOrderPosition }, members: args.shapeIds }); }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function pptUngroupShape(args) {
    if (!isSetSupported('PowerPointApi','1.8')) return errResult('requirement','ppt_ungroup_shape requires PowerPointApi 1.8+');
    if (!args.slideId || !args.groupShapeId) return errResult('bad_args','slideId and groupShapeId required');
    try{return await PowerPoint.run(async(ctx)=>{const slide=ctx.presentation.slides.getItem(String(args.slideId));const shape=slide.shapes.getItem(String(args.groupShapeId));shape.group.ungroup();await ctx.sync();slide.shapes.load('items/id,items/name,items/type,items/left,items/top,items/width,items/height');await ctx.sync();return okResult({host:'PowerPoint',slideId:args.slideId,ungrouped:args.groupShapeId,shapes:slide.shapes.items.map((s)=>({id:s.id,name:s.name,type:s.type,left:s.left,top:s.top,width:s.width,height:s.height}))});});}catch(e){return errResult('execution',String(e&&e.message||e));}
  }

  async function pptSetZOrder(args) {
    if (!isSetSupported('PowerPointApi', '1.8')) return errResult('requirement', 'ppt_set_z_order requires PowerPointApi 1.8+');
    const operation = String(args.operation || ''); if (!['BringForward','BringToFront','SendBackward','SendToBack'].includes(operation) || !Array.isArray(args.shapeIds) || !args.shapeIds.length) return errResult('bad_args', 'valid operation and shapeIds required');
    try { return await PowerPoint.run(async (ctx) => { const slide = ctx.presentation.slides.getItem(String(args.slideId)); const refs = args.shapeIds.map((id) => slide.shapes.getItem(String(id))); refs.forEach((shape) => { shape.setZOrder(operation); shape.load('id,zOrderPosition'); }); await ctx.sync(); return okResult({ host: 'PowerPoint', slideId: args.slideId, operation, shapes: refs.map((s) => ({ id: s.id, zOrderPosition: s.zOrderPosition })) }); }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function pptAlignDistribute(args, distribute) {
    const action = distribute ? 'ppt_distribute_shapes' : 'ppt_align_shapes';
    if (!isSetSupported('PowerPointApi', '1.4')) return errResult('requirement', `${action} requires PowerPointApi 1.4+`);
    if (!Array.isArray(args.shapeIds) || args.shapeIds.length < 2) return errResult('bad_args', 'shapeIds requires at least two IDs');
    const mode = String(args.mode || '');
    if ((!distribute && !['left','center','right','top','middle','bottom'].includes(mode)) || (distribute && !['horizontal','vertical'].includes(mode))) return errResult('bad_args', 'invalid mode');
    if (distribute && args.shapeIds.length < 3) return errResult('bad_args', 'distribution requires at least three shapes');
    try { return await PowerPoint.run(async (ctx) => { const slide = ctx.presentation.slides.getItem(String(args.slideId)); const refs = args.shapeIds.map((id) => slide.shapes.getItem(String(id))); refs.forEach((s) => s.load('id,left,top,width,height')); await ctx.sync();
      if (distribute) { const horizontal = mode === 'horizontal'; const sorted = refs.slice().sort((a,b) => (horizontal ? a.left-b.left : a.top-b.top)); const start = horizontal ? sorted[0].left : sorted[0].top; const end = horizontal ? sorted[sorted.length-1].left+sorted[sorted.length-1].width : sorted[sorted.length-1].top+sorted[sorted.length-1].height; const total = sorted.reduce((sum,s) => sum+(horizontal?s.width:s.height),0); const gap=(end-start-total)/(sorted.length-1); let cursor=start; for(const s of sorted){ if(horizontal)s.left=cursor; else s.top=cursor; cursor+=(horizontal?s.width:s.height)+gap; } }
      else { const first=refs[0]; for(const s of refs.slice(1)){ if(mode==='left')s.left=first.left; if(mode==='center')s.left=first.left+(first.width-s.width)/2; if(mode==='right')s.left=first.left+first.width-s.width; if(mode==='top')s.top=first.top; if(mode==='middle')s.top=first.top+(first.height-s.height)/2; if(mode==='bottom')s.top=first.top+first.height-s.height; } }
      await ctx.sync(); refs.forEach((s)=>s.load('id,left,top,width,height')); await ctx.sync(); return okResult({host:'PowerPoint',slideId:args.slideId,mode,shapes:refs.map((s)=>({id:s.id,left:s.left,top:s.top,width:s.width,height:s.height}))}); }); } catch(e){ return errResult('execution',String(e&&e.message||e)); }
  }

  function readSlides() {
    return new Promise((resolve) => {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
        (r) => resolve(r.status === Office.AsyncResultStatus.Succeeded
          ? okResult({ host: 'PowerPoint', slides: r.value && r.value.slides ? r.value.slides.map((s) => ({ id: s.id, title: s.title || '' })) : [] })
          : errResult('execution', r.error && r.error.message)));
    });
  }

  function readPptDocument() {
    return readPptFile(extractPptText);
  }

  function readPptNotes() {
    return readPptFile(extractPptNotes);
  }

  async function pptRenderSlides(args) {
    if (!isSetSupported('PowerPointApi', '1.8')) return errResult('requirement', 'ppt_render_slides requires PowerPointApi 1.8+');
    let width, height;
    try {
      width = pptFiniteNumber(args && args.width === undefined ? 960 : args.width, 'width', 64, 4096);
      height = pptFiniteNumber(args && args.height === undefined ? 540 : args.height, 'height', 64, 4096);
    } catch (e) { return errResult('bad_args', String(e && e.message || e)); }
    try {
      return await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load('items/id');
        await ctx.sync();
        const rendered = slides.items.map((slide) => ({ id: slide.id, image: slide.getImageAsBase64({ width, height }) }));
        await ctx.sync();
        return okResult({ host: 'PowerPoint', width, height, slides: rendered.map((item, index) => ({ index: index + 1, id: item.id, base64: item.image.value })) });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  function chunksToBase64(chunks) {
    return new Promise((resolve, reject) => {
      try {
        const reader = new FileReader();
        reader.onerror = () => reject(reader.error || new Error('FileReader failed'));
        reader.onload = () => resolve(String(reader.result || '').replace(/^data:.*?;base64,/, ''));
        reader.readAsDataURL(new Blob(chunks, { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }));
      } catch (e) { reject(e); }
    });
  }

  function pptExportFile() {
    return readPptFile(async (chunks) => {
      const size = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
      const base64 = await chunksToBase64(chunks);
      return okResult({ host: 'PowerPoint', fileType: 'pptx', size, base64 });
    });
  }

  function extractPptText(chunks) {
    const entries = concatChunksToZip(chunks);
    const slideNames = Object.keys(entries)
      .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
      .sort((a, b) => parseInt(a.match(/slide(\d+)/)[1], 10) - parseInt(b.match(/slide(\d+)/)[1], 10));
    const decoder = new TextDecoder('utf-8');
    const parts = [];
    for (const name of slideNames) {
      const xml = decoder.decode(inflateEntry(entries[name]));
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      const aT = doc.getElementsByTagName('a:t');
      let slideText = '';
      for (const t of aT) slideText += t.textContent;
      const num = name.match(/slide(\d+)/)[1];
      parts.push(`--- 幻灯片 ${num} ---\n${slideText.trim()}`);
    }
    return okResult({ host: 'PowerPoint', slides: slideNames.length, text: parts.join('\n\n') });
  }

  function extractPptNotes(chunks) {
    const entries = concatChunksToZip(chunks);
    const decoder = new TextDecoder('utf-8');
    const slideNums = Object.keys(entries)
      .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
      .map((n) => parseInt(n.match(/slide(\d+)/)[1], 10));
    const notesMap = {};
    for (const sn of slideNums) {
      const relsName = `ppt/slides/_rels/slide${sn}.xml.rels`;
      const relsEntry = entries[relsName];
      if (!relsEntry) continue;
      const relsDoc = new DOMParser().parseFromString(decoder.decode(inflateEntry(relsEntry)), 'application/xml');
      const rels = relsDoc.getElementsByTagName('Relationship');
      let notesTarget = null;
      for (const rel of rels) {
        if (rel.getAttribute('Type') && rel.getAttribute('Type').indexOf('notesSlide') >= 0) {
          notesTarget = rel.getAttribute('Target');
          break;
        }
      }
      if (!notesTarget) continue;
      const notesName = 'ppt/notesSlides/' + notesTarget.split('/').pop();
      const notesEntry = entries[notesName];
      if (!notesEntry) continue;
      const notesDoc = new DOMParser().parseFromString(decoder.decode(inflateEntry(notesEntry)), 'application/xml');
      const aT = notesDoc.getElementsByTagName('a:t');
      let text = '';
      for (const t of aT) text += t.textContent;
      notesMap[sn] = text.trim();
    }
    const parts = Object.keys(notesMap).sort((a, b) => Number(a) - Number(b)).map((sn) => `--- 幻灯片 ${sn} 备注 ---\n${notesMap[sn] || '(无备注文本)'}`);
    return okResult({ host: 'PowerPoint', notes: parts.join('\n\n'), slidesWithNotes: parts.length });
  }

  function concatChunksToZip(chunks) {
    const total = chunks.reduce((a, c) => a + c.length, 0);
    const buf = new Uint8Array(total);
    let off = 0;
    for (const c of chunks) { buf.set(c, off); off += c.length; }
    return unzip(buf);
  }


  A.hosts.PowerPoint = {
    actions: {
      // 跨宿主
      read_selection: readSelection,
      read_document: readDocument,
      // PowerPoint 专有
      read_slides: readSlides,
      ppt_read_notes: readPptNotes,
      ppt_add_slides: pptAddSlides,
      ppt_add_elements: pptAddElements,
      ppt_delete_slides: pptDeleteSlides,
      ppt_delete_shapes: pptDeleteShapes,
      ppt_render_slides: pptRenderSlides,
      ppt_export_file: pptExportFile,
      ppt_inspect_slides: pptInspectSlides,
      ppt_list_placeholders: pptListPlaceholders,
      ppt_insert_image: pptInsertImage,
      ppt_insert_table: pptInsertTable,
      ppt_group_shapes: pptGroupShapes,
      ppt_ungroup_shape: pptUngroupShape,
      ppt_set_z_order: pptSetZOrder,
      ppt_align_shapes: (args) => pptAlignDistribute(args, false),
      ppt_distribute_shapes: (args) => pptAlignDistribute(args, true),
    },
    // 机制级 re-read：覆盖类操作必须先预览当前状态，确认后带 confirm:true 才执行。
    previewReaders: {
      write_selection: () => readSelection({}),
      ppt_add_slides: () => pptInspectSlides({}),
      ppt_add_elements: (a) => pptInspectSlides({ slideIds: [a.slideId] }),
      ppt_delete_slides: (a) => pptDeleteSlidesPreview(a),
      ppt_delete_shapes: (a) => pptDeleteShapesPreview(a),
      ppt_insert_image: () => readSlides({}),
      ppt_insert_table: (a) => pptInspectSlides({ slideIds: [a.slideId] }),
      ppt_group_shapes: (a) => pptInspectSlides({ slideIds: [a.slideId] }),
      ppt_ungroup_shape: (a) => pptInspectSlides({ slideIds: [a.slideId] }),
      ppt_set_z_order: (a) => pptInspectSlides({ slideIds: [a.slideId] }),
      ppt_align_shapes: (a) => pptInspectSlides({ slideIds: [a.slideId] }),
      ppt_distribute_shapes: (a) => pptInspectSlides({ slideIds: [a.slideId] }),
    },
    afterContent: {
      ppt_add_slides: (a) => `新增 ${a.count || 1} 张幻灯片${a.clearPlaceholders ? '（清除占位符）' : ''}`,
      ppt_add_elements: (a) => `向 slideId=${a.slideId || '(未提供)'} 添加 ${(a.elements || []).length} 个元素；坐标 ${JSON.stringify(a.coordinateSpace || {})} → 目标 ${JSON.stringify(a.targetSize || {})} pt`,
    },
  };
})();
