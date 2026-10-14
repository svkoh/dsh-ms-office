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
  const PPT_ALIGN_MODES = ['left', 'center', 'right', 'top', 'middle', 'bottom'];
  const PPT_DISTRIBUTE_MODES = ['horizontal', 'vertical'];
  const PPT_ZORDER_OPS = ['BringForward', 'BringToFront', 'SendBackward', 'SendToBack'];
  // shapeType 的合法值由 Office 决定，插件不能凭空枚举。下面这组是在加载项里逐个试过的，
  // 其余 PascalCase 名字仍然放行（不回归），但执行报错时会把这份清单作为提示附上。
  const PPT_SHAPE_TYPES = ['Rectangle', 'Ellipse', 'Triangle', 'RightTriangle', 'RightArrow', 'Pentagon', 'Parallelogram', 'Diamond', 'Hexagon', 'Chevron', 'Heart', 'Plaque', 'Trapezoid', 'RoundRectangle', 'SnipRoundRectangle'];
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

  // 枚举按大小写不敏感接受，但返回注册表里的规范写法：调用方写 "left" 也能通过。
  // 报错时必须列出合法值——只说 "is invalid" 的话调用方只能靠猜，这是之前踩过的坑。
  function pptEnum(value, path, allowed, fallback) {
    if (value === undefined || value === null || value === '') return fallback === undefined ? null : fallback;
    const raw = String(value);
    const match = allowed.find((item) => item.toLowerCase() === raw.toLowerCase());
    if (!match) throw new Error(`${path} is invalid: "${raw}" (valid: ${allowed.join(', ')})`);
    return match;
  }

  // shapeType 是 Office 的几何类型名（PascalCase）。只放宽首字母，其余字符仍要求纯字母。
  function pptShapeTypeName(value, path) {
    const raw = String(value);
    const canonical = raw.charAt(0).toUpperCase() + raw.slice(1);
    if (!/^[A-Za-z][A-Za-z0-9]*$/.test(canonical)) throw new Error(`${path} is invalid: "${raw}" (must be an Office shape name, e.g. ${PPT_SHAPE_TYPES.slice(0, 4).join(', ')})`);
    return canonical;
  }

  // 执行层的 InvalidArgument 常常只剩一个光秃秃的错误名，调用方看不出是哪项参数错了。
  // shapeType 合不合法只有 Office 知道，插件不能凭空枚举，所以不拦未验证的名字，
  // 只在执行报错时把「已验证可用」的清单补进消息里，让调用方知道下一步该试什么。
  function shapeTypeHint(spec) {
    const unverified = (spec.slides || []).flatMap((slide) => slide.elements)
      .filter((element) => element.type === 'shape' && !PPT_SHAPE_TYPES.includes(element.shapeType));
    if (unverified.length === 0) return '';
    const names = [...new Set(unverified.map((element) => element.shapeType))].join(', ');
    return ` (shapeType "${names}" is accepted but unverified; verified values: ${PPT_SHAPE_TYPES.join(', ')})`;
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
        if (!element) throw new Error(`${path} must be an object`);
        // 别名兼容：kind→type、width→w、height→h、fontSize→size。这几个是调用方最常写错的。
        const alias = (primary, secondary) => (element[primary] !== undefined ? element[primary] : element[secondary]);
        const type = pptEnum(alias('type', 'kind'), `${path}.type`, PPT_ELEMENT_TYPES);
        const out = {
          type,
          x: pptFiniteNumber(element.x, `${path}.x`, 0, designWidth),
          y: pptFiniteNumber(element.y, `${path}.y`, 0, designHeight),
          w: pptFiniteNumber(alias('w', 'width'), `${path}.w`, type === 'line' ? 0 : 0.1, designWidth),
          h: pptFiniteNumber(alias('h', 'height'), `${path}.h`, type === 'line' ? 0 : 0.1, designHeight),
          rotation: element.rotation === undefined ? null : pptFiniteNumber(element.rotation, `${path}.rotation`, -360, 360),
        };
        if (out.x + out.w > designWidth + 0.01 || out.y + out.h > designHeight + 0.01) throw new Error(`${path} exceeds design bounds ${designWidth}x${designHeight}`);
        if (out.type === 'text' || out.type === 'shape') {
          out.text = element.text === undefined || element.text === null ? '' : String(element.text);
          out.font = element.font === undefined || element.font === null ? null : String(element.font);
          const sizeValue = alias('size', 'fontSize');
          out.size = pptFiniteNumber(sizeValue === undefined ? 18 : sizeValue, `${path}.size`, 1, 400);
          out.bold = element.bold === undefined ? false : !!element.bold;
          out.italic = element.italic === undefined ? false : !!element.italic;
          out.color = pptOptionalColor(element.color === undefined ? '#000000' : element.color, `${path}.color`);
          out.fill = pptOptionalColor(element.fill, `${path}.fill`);
          out.fillTransparencyProvided = element.fillTransparency !== undefined;
          out.fillTransparency = pptFiniteNumber(element.fillTransparency === undefined ? 0 : element.fillTransparency, `${path}.fillTransparency`, 0, 1);
          out.lineColor = pptOptionalColor(element.lineColor, `${path}.lineColor`);
          out.lineWidthProvided = element.lineWidth !== undefined;
          out.lineWidth = pptFiniteNumber(element.lineWidth === undefined ? 1 : element.lineWidth, `${path}.lineWidth`, 0, 50);
          out.align = pptEnum(element.align, `${path}.align`, PPT_ALIGNMENTS, 'Left');
          out.valign = pptEnum(element.valign, `${path}.valign`, PPT_VERTICAL_ALIGNMENTS, 'Top');
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
          if (out.type === 'shape') {
            out.shapeType = element.shapeType === undefined ? 'Rectangle' : pptShapeTypeName(element.shapeType, `${path}.shapeType`);
          }
        } else {
          out.connectorType = pptEnum(element.connectorType, `${path}.connectorType`, PPT_CONNECTORS, 'Straight');
          out.color = pptOptionalColor(element.color === undefined ? '#808080' : element.color, `${path}.color`);
          out.lineWidth = pptFiniteNumber(element.lineWidth === undefined ? 1 : element.lineWidth, `${path}.lineWidth`, 0.1, 50);
          out.dashStyle = pptEnum(element.dashStyle, `${path}.dashStyle`, PPT_DASH_STYLES, 'Solid');
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

  // ShapeFont.name is script-sensitive. Native export verifies the mapping.
  function applyPptScriptFont(range, text, name) {
    range.font.name = name;
    for (const match of String(text || '').matchAll(/[\u2E80-\u9FFF\uF900-\uFAFF\u3000-\u303F\uFF00-\uFFEF]+|[^\u2E80-\u9FFF\uF900-\uFAFF\u3000-\u303F\uFF00-\uFFEF]+/g)) {
      range.getSubstring(match.index, match[0].length).font.name = name;
    }
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
    if (element.font) applyPptScriptFont(range, element.text, element.font);
    range.font.size = element.size;
    range.font.bold = element.bold;
    range.font.italic = element.italic;
    range.font.color = element.color;
    range.paragraphFormat.horizontalAlignment = element.align;
    for (const run of element.runs || []) {
      const runRange = range.getSubstring(run.start, run.length);
      if (run.font) applyPptScriptFont(runRange, element.text.slice(run.start, run.start + run.length), run.font);
      if (run.size !== null) runRange.font.size = run.size;
      if (run.bold !== null) runRange.font.bold = run.bold;
      if (run.italic !== null) runRange.font.italic = run.italic;
      if (run.color) runRange.font.color = run.color;
    }
  }

  function applyPptTextBoxShapeFormat(shape, element) {
    if (element.fill) {
      shape.fill.setSolidColor(element.fill);
      shape.fill.transparency = element.fillTransparency;
    } else if (element.fillTransparencyProvided) {
      shape.fill.transparency = element.fillTransparency;
    }
    if (element.lineColor) {
      shape.lineFormat.visible = element.lineWidth > 0;
      shape.lineFormat.color = element.lineColor;
      if (element.lineWidth > 0) shape.lineFormat.weight = element.lineWidth;
    } else if (element.lineWidthProvided) {
      if (element.lineWidth === 0) shape.lineFormat.visible = false;
      else shape.lineFormat.weight = element.lineWidth;
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
        try { if (element.type === 'text') { shape = slide.shapes.addTextBox(element.text, pptShapeOptions(element, sx, sy)); applyPptTextBoxShapeFormat(shape, element); applyPptTextFormat(shape, element); } else if (element.type === 'shape') { shape = slide.shapes.addGeometricShape(element.shapeType, pptShapeOptions(element, sx, sy)); applyPptShapeFormat(shape, element); shape.textFrame.textRange.text = element.text; applyPptTextFormat(shape, element); } else { shape = slide.shapes.addLine(element.connectorType, pptShapeOptions(element, sx, sy)); shape.width = element.w * sx; shape.height = element.h * sy; shape.lineFormat.visible = true; shape.lineFormat.color = element.color; shape.lineFormat.weight = element.lineWidth; shape.lineFormat.dashStyle = element.dashStyle; shape.lineFormat.transparency = element.transparency; } if (element.rotation !== null) shape.rotation = element.rotation; shape.load('id'); created.push(shape); counts[element.type]++; } catch (e) { throw new Error(`elements[${index}]: ${e && e.message || e}`); }
      }
      await ctx.sync(); return okResult({ host: 'PowerPoint', slideId, elements: created.map((shape) => ({ id: shape.id })), counts, coordinateSpace: { width: spec.designWidth, height: spec.designHeight }, targetSize: { width: targetWidth, height: targetHeight } });
    }); } catch (e) { return errResult('execution', String(e && e.message || e) + shapeTypeHint(spec)); }
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

  const PPT_EDITABLE_STYLE_TYPES = ['TextBox', 'GeometricShape'];
  const PPT_GEOMETRY_FIELDS = ['left', 'top', 'width', 'height'];

  function pptShapeGeometry(shape, hasZOrder) {
    return { id: shape.id, name: shape.name, type: shape.type, left: shape.left, top: shape.top, width: shape.width, height: shape.height, ...(hasZOrder ? { zOrderPosition: shape.zOrderPosition } : {}) };
  }

  async function pptReadShapeStyles(ctx, shapes) {
    // Unsupported shape types must remain readable geometrically. Do not ask
    // their textFrame to load: Office throws InvalidArgument for those types.
    const readable = shapes.filter((shape) => PPT_EDITABLE_STYLE_TYPES.includes(shape.type));
    for (const shape of readable) {
      shape.textFrame.textRange.load('text,font/name,font/size,font/color,font/bold,font/italic');
      shape.fill.load('type,transparency');
    }
    if (readable.length) await ctx.sync();
    for (const shape of readable) if (shape.fill.type === 'Solid') shape.fill.load('foregroundColor');
    if (readable.some((shape) => shape.fill.type === 'Solid')) await ctx.sync();
    return new Map(shapes.map((shape) => {
      if (!PPT_EDITABLE_STYLE_TYPES.includes(shape.type)) return [shape, { stylesRead: false, stylesReadReason: `Style reading is not supported for shape type ${shape.type}` }];
      const range = shape.textFrame.textRange, font = range.font, fill = shape.fill;
      return [shape, { stylesRead: true, text: range.text, font: { name: font.name, size: font.size, color: font.color, bold: font.bold, italic: font.italic }, fill: { type: fill.type, transparency: fill.transparency, ...(fill.type === 'Solid' ? { color: fill.foregroundColor } : {}) } }];
    }));
  }

  async function pptInspectSlides(args) {
    if (!isSetSupported('PowerPointApi', '1.4')) return errResult('requirement', 'ppt_inspect_slides requires PowerPointApi 1.4+');
    if (args.withStyles !== undefined && typeof args.withStyles !== 'boolean') return errResult('bad_args', 'withStyles must be boolean');
    try { return await PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides; slides.load('items/id'); await ctx.sync();
      const wanted = Array.isArray(args.slideIds) && args.slideIds.length ? new Set(args.slideIds.map(String)) : null;
      const refs = slides.items.filter((s) => !wanted || wanted.has(String(s.id)));
      const hasZOrder = isSetSupported('PowerPointApi', '1.8');
      for (const slide of refs) slide.shapes.load(hasZOrder ? 'items/id,items/name,items/type,items/left,items/top,items/width,items/height,items/zOrderPosition' : 'items/id,items/name,items/type,items/left,items/top,items/width,items/height');
      await ctx.sync();
      const styles = args.withStyles ? await pptReadShapeStyles(ctx, refs.flatMap((slide) => slide.shapes.items)) : null;
      return okResult({ host: 'PowerPoint', slides: refs.map((slide) => ({ id: slide.id, shapes: slide.shapes.items.map((shape) => ({ ...pptShapeGeometry(shape, hasZOrder), ...(styles ? styles.get(shape) : {}) })) })) });
    }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  function validatePptUpdates(args) {
    if (!args || typeof args.slideId !== 'string' || !args.slideId.trim()) throw new Error('slideId required');
    if (!Array.isArray(args.updates) || !args.updates.length || args.updates.length > PPT_MAX_ELEMENTS) throw new Error(`updates must contain 1-${PPT_MAX_ELEMENTS} items`);
    const seen = new Set();
    const finite = (value, key, min, max) => {
      if (typeof value !== 'number' || !Number.isFinite(value) || value < min || value > max) throw new Error(`${key} must be a finite number from ${min} to ${max}`);
      return value;
    };
    const object = (value) => value && typeof value === 'object' && !Array.isArray(value);
    return args.updates.map((update, index) => {
      const path = `updates[${index}]`;
      if (!object(update) || typeof update.shapeId !== 'string' || !update.shapeId.trim()) throw new Error(`${path}.shapeId required`);
      if (seen.has(update.shapeId)) throw new Error('shapeId values must be unique');
      seen.add(update.shapeId);
      const allowed = ['shapeId', 'text', 'font', 'fill', ...PPT_GEOMETRY_FIELDS];
      if (Object.keys(update).some((key) => !allowed.includes(key))) throw new Error(`${path} contains an unsupported property; runs patching is not supported`);
      const out = { shapeId: update.shapeId };
      if (update.text !== undefined) {
        if (typeof update.text !== 'string' || update.text.length > 100000) throw new Error(`${path}.text must be a string of at most 100000 characters`);
        out.text = update.text;
      }
      for (const key of PPT_GEOMETRY_FIELDS) if (update[key] !== undefined) out[key] = finite(update[key], `${path}.${key}`, key === 'width' || key === 'height' ? 0 : -10000, 10000);
      if (update.font !== undefined) {
        if (!object(update.font) || !Object.keys(update.font).length || Object.keys(update.font).some((key) => !['name', 'size', 'color', 'bold', 'italic'].includes(key))) throw new Error(`${path}.font must contain supported font properties`);
        out.font = {};
        if (update.font.name !== undefined) { if (typeof update.font.name !== 'string' || !update.font.name.trim()) throw new Error(`${path}.font.name must be a non-empty string`); out.font.name = update.font.name; }
        if (update.font.size !== undefined) out.font.size = finite(update.font.size, `${path}.font.size`, 1, 400);
        if (update.font.color !== undefined) { if (typeof update.font.color !== 'string' || !PPT_COLOR_RE.test(update.font.color)) throw new Error(`${path}.font.color must be #RRGGBB`); out.font.color = update.font.color.toUpperCase(); }
        for (const key of ['bold', 'italic']) if (update.font[key] !== undefined) { if (typeof update.font[key] !== 'boolean') throw new Error(`${path}.font.${key} must be boolean`); out.font[key] = update.font[key]; }
        if (!Object.keys(out.font).length) throw new Error(`${path}.font must contain a defined property`);
      }
      if (update.fill !== undefined) {
        if (!object(update.fill) || Object.keys(update.fill).some((key) => !['color', 'transparency'].includes(key)) || typeof update.fill.color !== 'string' || !PPT_COLOR_RE.test(update.fill.color)) throw new Error(`${path}.fill requires color in #RRGGBB format`);
        out.fill = { color: update.fill.color.toUpperCase() };
        if (update.fill.transparency !== undefined) out.fill.transparency = finite(update.fill.transparency, `${path}.fill.transparency`, 0, 1);
      }
      if (Object.keys(out).length === 1) throw new Error(`${path} must contain a property to update`);
      return out;
    });
  }

  async function pptUpdateShapes(args, preview) {
    let updates; try { updates = validatePptUpdates(args); } catch (e) { return errResult('bad_args', e.message); }
    if (!isSetSupported('PowerPointApi', '1.4')) return errResult('requirement', 'ppt_update_shapes requires PowerPointApi 1.4+');
    try { return await PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides; slides.load('items/id'); await ctx.sync();
      const slide = slides.items.find((item) => String(item.id) === args.slideId);
      if (!slide) return errResult('not_found', `slide ID not found: ${args.slideId}`);
      slide.shapes.load('items/id,items/name,items/type,items/left,items/top,items/width,items/height'); await ctx.sync();
      const byId = new Map(slide.shapes.items.map((shape) => [String(shape.id), shape]));
      const missing = updates.filter((update) => !byId.has(update.shapeId)).map((update) => update.shapeId);
      if (missing.length) return errResult('not_found', `shape IDs not found: ${missing.join(', ')}`);
      for (const update of updates) if ((update.text !== undefined || update.font || update.fill) && !PPT_EDITABLE_STYLE_TYPES.includes(byId.get(update.shapeId).type)) return errResult('unsupported', `text/font/fill editing is not supported for shape type ${byId.get(update.shapeId).type} (${update.shapeId})`);
      for (const update of updates) if (byId.get(update.shapeId).type !== 'Line' && (update.width === 0 || update.height === 0)) return errResult('bad_args', 'Only native Line shapes may have zero width or height');
      const selected = updates.map((update) => byId.get(update.shapeId));
      // Resolve all requested objects/properties before queuing any setter.
      const before = await pptReadShapeStyles(ctx, selected);
      if (preview) return okResult({ host: 'PowerPoint', slideId: args.slideId, previewOnly: true, shapes: selected.map((shape) => ({ ...pptShapeGeometry(shape, false), ...before.get(shape) })) });
      for (const update of updates) {
        const shape = byId.get(update.shapeId);
        for (const key of PPT_GEOMETRY_FIELDS) if (update[key] !== undefined) shape[key] = update[key];
        if (update.text !== undefined) shape.textFrame.textRange.text = update.text;
        if (update.font) for (const [key, value] of Object.entries(update.font)) shape.textFrame.textRange.font[key] = value;
        if (update.fill) {
          shape.fill.setSolidColor(update.fill.color);
          const transparency = update.fill.transparency === undefined ? before.get(shape).fill.transparency : update.fill.transparency;
          if (typeof transparency === 'number') shape.fill.transparency = transparency;
        }
      }
      await ctx.sync();
      for (const shape of selected) shape.load('id,name,type,left,top,width,height');
      await ctx.sync();
      const after = await pptReadShapeStyles(ctx, selected);
      return okResult({ host: 'PowerPoint', slideId: args.slideId, updated: updates.length, shapes: selected.map((shape) => ({ ...pptShapeGeometry(shape, false), ...after.get(shape) })), requirementSet: 'PowerPointApi 1.4' });
    }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function pptInsertImage(args) {
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
    if (args.slideId !== undefined) {
      if (typeof args.slideId !== 'string' || !args.slideId.trim()) return errResult('bad_args', 'slideId must be a non-empty string');
      if (!isSetSupported('PowerPointApi', '1.2')) return errResult('requirement', 'targeted ppt_insert_image requires PowerPointApi 1.2+');
      try { await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides; slides.load('items/id'); await ctx.sync();
        if (!slides.items.some((slide) => String(slide.id) === args.slideId)) throw new Error('Target slide not found: ' + args.slideId);
        ctx.presentation.setSelectedSlides([args.slideId]); await ctx.sync();
      }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
    }
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
    let operation;
    try { operation = pptEnum(args.operation, 'operation', PPT_ZORDER_OPS); } catch (e) { return errResult('bad_args', e.message); }
    if (!Array.isArray(args.shapeIds) || !args.shapeIds.length) return errResult('bad_args', 'shapeIds must be a non-empty array');
    try { return await PowerPoint.run(async (ctx) => { const slide = ctx.presentation.slides.getItem(String(args.slideId)); const refs = args.shapeIds.map((id) => slide.shapes.getItem(String(id))); refs.forEach((shape) => { shape.setZOrder(operation); shape.load('id,zOrderPosition'); }); await ctx.sync(); return okResult({ host: 'PowerPoint', slideId: args.slideId, operation, shapes: refs.map((s) => ({ id: s.id, zOrderPosition: s.zOrderPosition })) }); }); } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function pptAlignDistribute(args, distribute) {
    const action = distribute ? 'ppt_distribute_shapes' : 'ppt_align_shapes';
    if (!isSetSupported('PowerPointApi', '1.4')) return errResult('requirement', `${action} requires PowerPointApi 1.4+`);
    if (!Array.isArray(args.shapeIds) || args.shapeIds.length < 2) return errResult('bad_args', 'shapeIds requires at least two IDs');
    let mode;
    try { mode = pptEnum(args.mode, 'mode', distribute ? PPT_DISTRIBUTE_MODES : PPT_ALIGN_MODES); } catch (e) { return errResult('bad_args', e.message); }
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
    args = args || {};
    if (!isSetSupported('PowerPointApi', '1.8')) return errResult('requirement', 'ppt_render_slides requires PowerPointApi 1.8+');
    let width, height, wanted;
    try {
      width = pptFiniteNumber(args.width === undefined ? 960 : args.width, 'width', 64, 4096);
      height = pptFiniteNumber(args.height === undefined ? 540 : args.height, 'height', 64, 4096);
      if (args.slideIds !== undefined) {
        if (!Array.isArray(args.slideIds) || !args.slideIds.length || args.slideIds.some((id) => typeof id !== 'string' || !id.trim())) throw new Error('slideIds must be a non-empty array of slide ID strings');
        wanted = new Set(args.slideIds);
        if (wanted.size !== args.slideIds.length) throw new Error('slideIds must be unique');
      }
    } catch (e) { return errResult('bad_args', String(e && e.message || e)); }
    try {
      return await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load('items/id');
        await ctx.sync();
        if (wanted) {
          const existing = new Set(slides.items.map((slide) => String(slide.id)));
          const missing = [...wanted].filter((id) => !existing.has(id));
          if (missing.length) return errResult('not_found', `slide IDs not found: ${missing.join(', ')}`);
        }
        const rendered = slides.items.map((slide, index) => ({ slide, index: index + 1 })).filter(({ slide }) => !wanted || wanted.has(String(slide.id))).map(({ slide, index }) => ({ index, id: slide.id, image: slide.getImageAsBase64({ width, height }) }));
        await ctx.sync();
        return okResult({ host: 'PowerPoint', width, height, slides: rendered.map((item) => ({ index: item.index, id: item.id, base64: item.image.value })) });
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

  function pptExportFile(args) {
    const type = String(args && args.type || 'pptx').toLowerCase();
    if (!['pptx', 'pdf'].includes(type)) return Promise.resolve(errResult('bad_args', 'type must be pptx or pdf'));
    if (type === 'pptx') return readPptFile(async (chunks) => {
      const size = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
      return okResult({ host: 'PowerPoint', fileType: type, size, base64: await chunksToBase64(chunks) });
    });
    // PDF bytes come directly from the open Office document.
    return new Promise((resolve) => {
      try { Office.context.document.getFileAsync(Office.FileType.Pdf, { sliceSize: 4194304 }, (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) return resolve(errResult('execution', result.error && result.error.message));
        const file = result.value, chunks = []; let index = 0, closed = false;
        const done = (out) => { if (closed) return; closed = true; try { file.closeAsync(() => resolve(out)); } catch (e) { resolve(out); } };
        const onSlice = (slice) => {
          try {
            if (slice.status !== Office.AsyncResultStatus.Succeeded) return done(errResult('execution', slice.error && slice.error.message));
            chunks.push(new Uint8Array(slice.value.data)); index++;
            if (index < file.sliceCount) { file.getSliceAsync(index, onSlice); return; }
            const reader = new FileReader();
            reader.onerror = () => done(errResult('execution', String(reader.error || 'FileReader failed')));
            reader.onload = () => done(okResult({ host: 'PowerPoint', fileType: 'pdf', size: chunks.reduce((n, c) => n + c.length, 0), base64: String(reader.result || '').replace(/^data:.*?;base64,/, '') }));
            reader.readAsDataURL(new Blob(chunks, { type: 'application/pdf' }));
          } catch (e) { done(errResult('execution', String(e && e.message || e))); }
        };
        try { if (!file.sliceCount) return done(errResult('execution', 'Native PDF contains no slices')); file.getSliceAsync(0, onSlice); } catch (e) { done(errResult('execution', String(e && e.message || e))); }
      }); } catch (e) { resolve(errResult('execution', String(e && e.message || e))); }
    });
  }

  const PPT_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const PPT_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

  function pptXml(entries, name) {
    if (!entries[name]) throw new Error(`Missing PPT part: ${name}`);
    const doc = new DOMParser().parseFromString(new TextDecoder('utf-8').decode(inflateEntry(entries[name])), 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error(`Invalid PPT XML: ${name}`);
    return doc;
  }

  // Resolve OPC part names without fetching external relationships or leaving the ZIP.
  function pptPartTarget(source, target) {
    if (!target || /[\\?#]/.test(target) || /^[a-z][a-z0-9+.-]*:/i.test(target)) throw new Error('Invalid PPT relationship target');
    const parts = target.startsWith('/') ? [] : source.split('/').slice(0, -1);
    for (const part of target.split('/')) {
      if (!part || part === '.') continue;
      if (part === '..') { if (!parts.length) throw new Error('PPT relationship escapes package'); parts.pop(); }
      else parts.push(part);
    }
    return parts.join('/');
  }

  function pptRelationships(entries, source) {
    const parts = source.split('/'), file = parts.pop();
    const name = [...parts, '_rels', file + '.rels'].join('/');
    if (!entries[name]) return [];
    return Array.from(pptXml(entries, name).getElementsByTagNameNS('*', 'Relationship'));
  }

  function orderedPptSlides(entries) {
    const source = 'ppt/presentation.xml';
    const doc = pptXml(entries, source);
    const rels = new Map(pptRelationships(entries, source).map((rel) => [rel.getAttribute('Id'), rel]));
    const list = doc.getElementsByTagNameNS(PPT_NS, 'sldIdLst')[0];
    if (!list) return [];
    return Array.from(list.getElementsByTagNameNS(PPT_NS, 'sldId')).map((slide, index) => {
      const rel = rels.get(slide.getAttributeNS(PPT_REL_NS, 'id'));
      if (!rel || rel.getAttribute('TargetMode') === 'External' || !/\/slide$/.test(rel.getAttribute('Type') || '')) throw new Error('Missing or invalid PPT slide relationship');
      const name = pptPartTarget(source, rel.getAttribute('Target'));
      if (!entries[name]) throw new Error(`Missing PPT part: ${name}`);
      return { name, index: index + 1, id: slide.getAttribute('id') };
    });
  }

  function pptPartText(entries, name) {
    const text = pptXml(entries, name).getElementsByTagNameNS(DRAWING_NS, 't');
    return Array.from(text, (t) => t.textContent).join('').trim();
  }

  function pptReadPresentationInfo() {
    return readPptFile((chunks) => {
      const entries = concatChunksToZip(chunks);
      const doc = pptXml(entries, 'ppt/presentation.xml');
      const size = doc.getElementsByTagNameNS(PPT_NS, 'sldSz')[0];
      const cx = size && Number(size.getAttribute('cx')), cy = size && Number(size.getAttribute('cy'));
      if (!Number.isSafeInteger(cx) || !Number.isSafeInteger(cy) || cx <= 0 || cy <= 0) return errResult('execution', 'Missing or invalid presentation slide size');
      return okResult({ host: 'PowerPoint', slideSize: { width: cx / 12700, height: cy / 12700, unit: 'pt' }, aspectRatio: cx / cy, slideIds: orderedPptSlides(entries).map((slide) => slide.id) });
    });
  }

  function extractPptText(chunks) {
    const entries = concatChunksToZip(chunks);
    const slides = orderedPptSlides(entries);
    const parts = slides.map((slide) => `--- 幻灯片 ${slide.index} ---\n${pptPartText(entries, slide.name)}`);
    return okResult({ host: 'PowerPoint', slides: slides.length, text: parts.join('\n\n'), slideIds: slides.map((slide) => slide.id) });
  }

  function extractPptNotes(chunks) {
    const entries = concatChunksToZip(chunks);
    const slides = orderedPptSlides(entries);
    const parts = [], slideIds = [];
    for (const slide of slides) {
      const rel = pptRelationships(entries, slide.name).find((item) => /\/notesSlide$/.test(item.getAttribute('Type') || ''));
      if (!rel || rel.getAttribute('TargetMode') === 'External') continue;
      const name = pptPartTarget(slide.name, rel.getAttribute('Target'));
      const text = pptPartText(entries, name);
      parts.push(`--- 幻灯片 ${slide.index} 备注 ---\n${text || '(无备注文本)'}`);
      slideIds.push(slide.id);
    }
    return okResult({ host: 'PowerPoint', notes: parts.join('\n\n'), slidesWithNotes: parts.length, slideIds });
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
      ppt_read_presentation_info: pptReadPresentationInfo,
      ppt_update_shapes: (a) => pptUpdateShapes(a, false),
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
      ppt_update_shapes: (a) => pptUpdateShapes(a, true),
      ppt_add_elements: (a) => pptInspectSlides({ slideIds: [a.slideId] }),
      ppt_delete_slides: (a) => pptDeleteSlidesPreview(a),
      ppt_delete_shapes: (a) => pptDeleteShapesPreview(a),
      ppt_insert_image: (a) => a.slideId ? pptInspectSlides({ slideIds: [a.slideId] }) : readSlides({}),
      ppt_insert_table: (a) => pptInspectSlides({ slideIds: [a.slideId] }),
      ppt_group_shapes: (a) => pptInspectSlides({ slideIds: [a.slideId] }),
      ppt_ungroup_shape: (a) => pptInspectSlides({ slideIds: [a.slideId] }),
      ppt_set_z_order: (a) => pptInspectSlides({ slideIds: [a.slideId] }),
      ppt_align_shapes: (a) => pptInspectSlides({ slideIds: [a.slideId] }),
      ppt_distribute_shapes: (a) => pptInspectSlides({ slideIds: [a.slideId] }),
    },
    afterContent: {
      ppt_update_shapes: (a) => JSON.stringify({ slideId: a.slideId, updates: a.updates }),
      ppt_add_slides: (a) => `新增 ${a.count || 1} 张幻灯片${a.clearPlaceholders ? '（清除占位符）' : ''}`,
      ppt_add_elements: (a) => `向 slideId=${a.slideId || '(未提供)'} 添加 ${(a.elements || []).length} 个元素；坐标 ${JSON.stringify(a.coordinateSpace || {})} → 目标 ${JSON.stringify(a.targetSize || {})} pt`,
    },
  };
})();
