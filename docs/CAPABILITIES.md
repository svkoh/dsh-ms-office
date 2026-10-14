# 完整操作清单

> [English](CAPABILITIES.en.md) · 中文

本插件当前支持的全部 action。**机器可读的那份以运行时为准**，随时可以取：

```bash
curl "$(jq -r .url "$DSH_HOME/dsh-ms-office/bridge.json")/office/capabilities?lang=zh"
```

（`bridge.json` 的说明见[集成方文档](AGENT.md#先拿到桥接地址)。）

**读法**

- **D** = 覆盖类操作（destructive）：执行前会先读一遍当前状态，默认还要你在窗格审批。
- **平台**：W = Word，E = Excel，P = PowerPoint。
- **枚举大小写不敏感**（`align:"left"` 和 `align:"Center"` 都接受），但非法值被拒时会把合法值清单写进报错——照清单改，不要猜。
- **字段名以 `args` 说明为准**。常见误写已做别名兼容：`kind`→`type`、`width`/`height`→`w`/`h`、`fontSize`→`size`；颜色只接受 `#RRGGBB`。

> `server/capabilities.json` 是**结构**的唯一真源（名称 / 平台 / 是否覆盖类 / requirement），
> 人类可读的 `args` 和 `returns` 说明来自 `i18n/<lang>.json`，接口把两者合并后返回。
> 以下表格是给人看的概览；参数表描述请求输入，返回值表描述响应路径。

## 返回结果约定

`POST /office/command` 只返回一层 `{ok,commandId,host,instance,result}` 信封；操作数据在 `response.result`，不是 `response.result.result`。查询命令状态时，先取其中的 `response` 再按此结构读取。能力元数据中的 `returns` 标明信封内的返回路径，例如 `result.tables` 和 `result.styles`，不属于要发送的参数。

写入样式时，`ok:true` 代表执行完成；宣称格式已生效前，还应检查 `result.verified`、逐字段 `applied` / `mismatch` / `unverified` 状态及 `verificationWarnings`。Word 原生渲染只接受唯一精确匹配的文档，不做模糊匹配或选取最接近候选；未找到匹配文档时附匹配诊断，其他失败报告具体原因。宿主支持时，Word 分页诊断在 `result.probe.nativeDiagnostics.pagination`，不是 `result.pagination`。

完整请求只序列化一次，按 UTF-8 字节发送；参见 [Office 技能中的 PowerShell 示例](../skills/ms-office/SKILL.md#4-发指令)。

| action | 返回值（响应路径，不是参数） |
| --- | --- |
| `read_tables` | `result.tables`: [{index,rowCount,columnCount,style,styleBuiltIn,cells:string[][]}]；表索引从1开始；使用原生values保留空单元格、换行和空白，无额外分隔行 |
| `read_table_borders` | `result`: {host,tableCount,borders:[{table,location,color,type,width}]}；table从1开始，width单位pt |
| `word_format_paragraphs` | `result`: {host,formatted,paragraphs:[{index,text,styleName,styleBuiltIn,alignment,font,...}]} |
| `word_read_styles` | `result.styles`: [{name,type,builtIn,font,paragraphFormat}]；alignment统一left\|center\|right\|justify，其他宿主值原样保留 |
| `word_set_styles` | `result`: {host,global,updated,styles,verified,verification:[{requestedName,name,fields:[{field,requested,actual,status:applied\|mismatch\|unverified}]}],verificationWarnings?}；ok=true代表写入执行，verified与逐字段状态确认生效；预设返回created/styles等独立字段 |
| `get_environment` | `result`: 环境诊断对象；Word分页在 `result.probe.nativeDiagnostics.pagination`，section信息在 `result.probe.nativeDiagnostics` 内，字段是否存在取决于宿主支持 |
| `word_render_pages` | `result`: {pages:[{index,width,height,base64,...}],pageCount,renderer,documentFingerprint,...}；只接受唯一精确匹配文档，不使用最接近候选；未匹配时error包含匹配诊断，其他错误给出具体原因 |

## 通用

| action | 平台 | D | 说明 | 条件 | 参数 |
| --- | --- | --- | --- | --- | --- |
| `read_selection` | W/E/P |  |  |  | `withStyles`: boolean 可选：同时返回样式 |
| `write_selection` | W/E/P | D |  |  | `text`: string 必填：替换选区的文本 |
| `read_document` | W/E/P |  |  |  | `sheet`: string 可选(Excel)：指定工作表名 |
| `read_styles` | W/E |  |  |  |  |
| `replace_all` | W/E | D |  |  | `search`: string 必填<br>`replace`: string 必填<br>`dryRun`: boolean 可选：仅返回命中数不执行 |
| `format_selection` | W/E | D |  |  | `font`: string 可选<br>`size`: number 可选<br>`bold`: boolean 可选<br>`italic`: boolean 可选<br>`color`: string 可选(#RRGGBB)<br>`highlight`: string 可选(Word) |
| `add_comment` | W/E | D |  |  | `text`: string 必填<br>`selectionText`: string 可选(Word)：给含该文本的选区加批注<br>`cell`: string 可选(Excel)：A1 单元格 |
| `read_comments` | W/E |  |  |  |  |
| `read_properties` | W/E |  |  |  |  |
| `get_environment` | W/E/P |  | 读取Office环境、原生API能力及确认模式 |  | `syncConfig`: 可选boolean；将本窗格同步到服务端用户已选择的确认模式，不接受模式值或审批参数<br>`reloadPane`: 可选boolean；仅在公开命令状态明确为空后重新载入本窗格，不改变文档 |
| `get_document_info` | W/E/P |  |  |  | `keys`: array 可选：要读取的 settings 键列表（如 ["lang","theme"]） |
| `set_setting` | W/E/P | D |  |  | `key`: string 必填：设置键<br>`value`: any 必填：任意 JSON 值（随文档持久化，跨会话可读） |
| `locate_select` | W/E |  |  |  | `bookmark`: string 可选：Word 书签 / Excel 命名区域名<br>`anchor`: string 可选：同 bookmark<br>`text`: string 可选：定位首个匹配文本<br>`range`: string 可选(Excel)：A1:B5 或 Sheet!A1:B5<br>`address`: string 可选(Excel)：同 range<br>`sheet`: string 可选(Excel)：指定工作表<br>`blinks`: number 可选：闪烁次数(默认0=只选中保持,1-5才闪)<br>`interval`: number 可选：间隔ms(默认300,100-800) |

## Word

| action | 平台 | D | 说明 | 条件 | 参数 |
| --- | --- | --- | --- | --- | --- |
| `append_text` | W | D |  |  | `text`: string 必填：文末追加段落 |
| `read_tables` | W |  |  |  |  |
| `read_table_borders` | W |  |  | WordApi 1.3+ |  |
| `set_table_borders` | W | D |  | WordApi 1.3+ | `color`: string 必填：#RRGGBB<br>`tableIndexes`: number[] 可选：从1开始，默认全部<br>`locations`: Top\|Bottom\|Left\|Right\|InsideHorizontal\|InsideVertical 数组，可选<br>`enableGrid`: boolean 可选：先应用 TableGrid 以启用完整网格<br>`width`: 可选正数边框宽度(pt)；不传保留现有粗细 |
| `set_font` | W | D |  |  | `font`: string 必填：字体名 |
| `word_list_paragraphs` | W |  |  |  | `filter`: all\|empty\|nonEmpty 可选，默认 all |
| `word_delete_paragraphs` | W | D |  |  | `targets`: array 必填：[{index,expectedText}]，index为从0开始的整数，expectedText为完整字符串；参数名不是paragraphs；最后一段及含内嵌图片的段落不可删除，第0段非末段时可删除 |
| `insert_paragraph` | W | D |  |  | `text`: string 必填<br>`style`: string 可选：内置样式如 Heading1<br>`location`: string 可选：end\|start\|afterSelection\|afterSelectionParagraph |
| `insert_table` | W | D |  | WordApi 1.3+ | `rows`: array 必填：矩形二维数组，最多1000行/63列/10000单元格<br>`location`: string 可选：start\|end\|beforeSelection\|afterSelection<br>`style`: string 可选：仅本地化/自定义表格样式名<br>`styleBuiltIn`: string 可选：跨语言内置样式枚举，如 GridTable4_Accent1；不可与 style 同用<br>`headerRow`: boolean 可选：首行为标题行 |
| `word_insert_html` | W | D |  | WordApi 1.1+ | `html`: string 必填：最多8*1024*1024个UTF-16代码单元；对象/数组/数字拒绝，禁止重复JSON序列化；示例`{html:"<p>正文</p>",target:"body",mode:"append"}`<br>`target`: body\|selection 可选，默认 body<br>`mode`: replace\|append\|prepend\|before\|after 可选，默认 append |
| `word_set_page_setup` | W | D |  | WordApiDesktop 1.3+ | `sections`: all\|number[] 可选：从0开始，默认 all<br>`topMargin`: number 可选(points)<br>`bottomMargin`: number 可选(points)<br>`leftMargin`: number 可选(points)<br>`rightMargin`: number 可选(points)<br>`headerDistance`: number 可选(points)<br>`footerDistance`: number 可选(points)<br>`pageWidth`: optional positive pt<br>`pageHeight`: optional positive pt<br>`differentFirstPageHeaderFooter`: optional boolean |
| `word_insert_equation` | W | D |  | WordApi 1.1+ Range.insertOoxml | `expression`: 结构化公式节点：text\|group\|fraction\|subscript\|superscript<br>`target`: object 可选：{kind:selection\|body\|search,query/occurrence/matchCase/matchWholeWord}，occurrence 从0开始<br>`mode`: replace\|before\|after\|prepend\|append 可选，默认 after<br>`display`: boolean 可选：默认 true |
| `word_read_equations` | W |  |  | WordApi 1.1+ Range.getOoxml |  |
| `word_read_selection_paragraph` | W |  |  | WordApi 1.1+ Range.getOoxml |  |
| `word_export_file` | W |  |  |  | `type`: docx\|pdf，默认 docx |
| `word_insert_image` | W | D |  |  | `base64`: string 必填：PNG/JPEG base64<br>`mode`: before\|after\|replace 可选，默认 after<br>`width`: number 可选(px)<br>`height`: number 可选(px) |
| `word_transfer_content` | W | D |  | WordApi 1.1+ | `source`: object 必填：{kind:selection\|table\|search,index/query/occurrence/matchCase}<br>`destination`: object 必填：{kind:selection\|body\|header\|footer,section/type}<br>`operation`: copy\|move 可选，默认 copy<br>`mode`: replace\|append\|prepend\|before\|after 可选，默认 append<br>`expectedText`: string 可选：写前核对源文本<br>`overwrite`: boolean 可选：替换非空目标时须 true |
| `apply_style` | W | D |  |  | `style`: string 必填：Heading1/标题 1/正文等<br>`scope`: string 可选：selection\|all |
| `set_paragraph_format` | W | D |  |  | `alignment`: string 可选：left\|center\|right\|justify<br>`indent`: number 可选(cm)<br>`lineSpacing`: 可选正数，单位pt；18pt不能据此宣称原生1.5倍行距，倍数行距需原生预设及导出核验<br>`listType`: none\|numbered\|bulleted；WordApi1.3；已有列表须先显式移除后再转换 |
| `search` | W |  |  |  | `query`: string 必填<br>`matchCase`: boolean 可选<br>`wildcard`: boolean 可选：通配符 |
| `word_list_sections` | W |  |  | WordApi 1.1+ | `textLimit`: number 可选 |
| `word_read_headers_footers` | W |  |  | WordApi 1.1+ | `section`: number\|all<br>`types`: primary\|firstPage\|evenPages 数组<br>`includeFields`: optional boolean; WordApi1.5; read actual native field type/code/text |
| `word_set_header_footer` | W | D |  | WordApi 1.1+ | `kind`: header\|footer<br>`section`: number\|all<br>`type`: primary\|firstPage\|evenPages<br>`text`: optional string; required unless field is provided<br>`mode`: replace\|append\|prepend<br>`field`: optional {type:TOC\|Page\|NumPages,switches?}; mutually exclusive with text; requires WordApi1.5 |
| `word_insert_break` | W | D |  | WordApi 1.1+ | `type`: page\|line\|sectionNext\|sectionContinuous\|sectionEven\|sectionOdd<br>`location`: start\|end\|beforeSelection\|afterSelection<br>`paragraph`: 可选{index:非负整数,expectedText:字符串}；按原生段落索引及完整文本精确定位 |
| `word_insert_caption` | W | D | 静态编号需 WordApi 1.1+；动态 SEQ 需 WordApi 1.5+（本机不可用） | WordApi 1.1+ | `label`: string<br>`title`: string<br>`sequenceId`: ASCII identifier<br>`dynamic`: boolean 可选，默认 false；true 尝试动态 SEQ<br>`number`: number 可选，静态编号默认 1<br>`position`: beforeSelection\|afterSelection<br>`separator`: string 可选 |
| `word_format_paragraphs` | W | D | 按精确索引及完整文本格式化段落，支持内置样式和文档已存在的段落样式 | WordApi 1.3+ | `targets`: 必填1..500个精确目标[{index,expectedText,style?,font?:{name,size,color,bold,italic},alignment?,spaceBefore?,spaceAfter?,lineSpacing?,leftIndent?,rightIndent?,firstLineIndent?}]；0基索引不重复，先验证全部完整文本；style必须在每个target内，支持仅style修改；支持内置别名及文档已存在段落样式名（文档名需要WordApi1.5）；alignment接受left\|center\|right\|justify和Left\|Centered\|Right\|Justified，读回统一小写；尺寸pt，lineSpacing=18不等于原生1.5倍 |
| `word_read_styles` | W |  | 读取真实全局段落样式及中西文字体属性 | WordApi 1.5+ | `names`: string[] 可选：文档实际段落样式名或受支持的中英内置别名；省略列出全部，推荐使用读回的实际name |
| `word_set_styles` | W | D | 设置真实全局段落样式并逐字段回读核对；可能影响所有使用该样式的段落 | WordApi 1.5+ | `styles`: 必填1..100个{name,font?,paragraphFormat?}；name为已存在文档样式名或受支持中英内置别名（Normal/Title/Heading1等），不存在不自动创建；type是只读字段不可传；font仅name/size/color/bold/italic；paragraphFormat为段落格式字段，alignment接受规范值或Office枚举值；或单独{name:"范本正文",preset:"body-1.5"}导入原生预设；数值lineSpacing为pt，不能代替倍数 |
| `word_read_fields` | W |  | 读取真实Word字段代码及缓存结果 | WordApi 1.5+ | `target`: body\|selection; default body |
| `word_insert_field` | W | D | 插入真实TOC或PAGE字段并读回 | WordApi 1.5+ | `type`: TOC\|Page\|NumPages<br>`switches`: optional single-line field switches<br>`target`: body\|selection; default selection<br>`mode`: replace\|prepend\|append\|before\|after; body accepts only replace/prepend/append<br>`paragraph`: TOC可选{index,expectedText}；正文段落0基索引，完整文本精确匹配；使用可用的Desktop1.4原生目录集合；采用默认1–3级目录，不接受自定义switches；插入后读回真实缓存 |
| `word_create_document` | W | D | 原生新建并打开空白Word文档；不改变原文档，无文件/Base64输入 | WordApi 1.3+ |  |
| `word_render_pages` | W |  | 只读导出Word原生整页PNG，包含页眉页脚；要求Windows桌面Word打印布局 | WordApi 1.1+; Windows desktop Word native PrintLayout | `expectedPageCount`: 必填整数1..10；预期原生页数 |

## Excel

| action | 平台 | D | 说明 | 条件 | 参数 |
| --- | --- | --- | --- | --- | --- |
| `list_sheets` | E |  | 读取所有工作表名称、顺序和可见性 |  |  |
| `read_range` | E |  |  |  | `address`: string 必填：Sheet!A1:B10<br>`limit`: number 可选：正整数，最大读取格数；默认且上限为 5000<br>`withLayout`: 可选boolean：读回实际换行/对齐及受支持的pt尺寸；混合值保留null |
| `write_range` | E | D |  |  | `address`: string 必填<br>`values`: array 可选：二维数组<br>`formulas`: array 可选：二维数组(A1 公式) |
| `format_range` | E | D |  | ExcelApi 1.1+; column/row dimensions and autofit require ExcelApi 1.2+ | `address`: string 必填<br>`font`: string 可选<br>`size`: number 可选<br>`bold`: boolean 可选<br>`fill`: string 可选(#RRGGBB)<br>`numberFormat`: string 可选<br>`autoFit`: 可选boolean：autoFitColumns的兼容别名；ExcelApi 1.2+<br>`tableStyle`: string 可选(如 TableStyleMedium2)<br>`tableName`: 设置 tableStyle 时必填<br>`columnWidth`: 可选正数列宽(pt)，不可与autoFitColumns:true或autoFit:true同用<br>`rowHeight`: 可选正数行高(pt)，不可与autoFitRows:true同用<br>`wrapText`: 可选boolean<br>`horizontalAlignment`: General\|Left\|Center\|Right\|Fill\|Justify\|CenterAcrossSelection\|Distributed<br>`verticalAlignment`: Top\|Center\|Bottom\|Justify\|Distributed<br>`autoFitRows`: 可选boolean：自动行高；ExcelApi1.2<br>`autoFitColumns`: 可选boolean：自动列宽；ExcelApi 1.2+；true不可与columnWidth同用，autoFit为兼容别名。<br>`color`: 可选文字颜色 #RRGGBB |
| `insert_chart` | E | D |  |  | `type`: string 必填：Line\|Column\|Bar\|Pie\|Area\|Scatter 等<br>`dataRange`: string 必填：Sheet!A1:B10<br>`title`: string 可选<br>`position`: 可选{startCell,endCell}：图表所在工作表的A1单元格；start左上、end右下；默认在dataRange的工作表创建，不依赖活动工作表<br>`font`: 可选字体名称；中文默认宋体<br>`seriesColors`: optional #RRGGBB[]<br>`fill`: optional chart background #RRGGBB<br>`color`: optional chart text #RRGGBB |
| `add_sheet` | E | D |  |  | `name`: string 必填 |
| `rename_sheet` | E | D |  |  | `oldName`: string 必填<br>`newName`: string 必填 |
| `delete_sheet` | E | D |  |  | `name`: string 必填<br>`dryRun`: boolean 可选 |
| `apply_sort` | E | D |  |  | `address`: string 必填<br>`fields`: array 必填：[{column:1,ascending:true}] |
| `apply_filter` | E | D |  |  | `address`: string 必填<br>`columns`: array 可选：[{column:1,filterType:'Values',criteria:{values:['a']}}]；column 从 1 开始，Values 可简写 criteria:['a']；Custom 使用 criteria:{criterion1:'>2'} |
| `evaluate_formula` | E |  |  |  | `formula`: string 必填：如 SUM(A1:A10) |
| `list_tables` | E |  |  | ExcelApi 1.1+ | `sheet`: string 可选<br>`includeValues`: boolean 可选<br>`limit`: number 可选 |
| `create_table` | E | D |  | ExcelApi 1.1+ | `address`: string<br>`hasHeaders`: boolean<br>`name`: string 可选<br>`style`: string 可选 |
| `delete_table` | E | D |  | ExcelApi 1.1+ | `name`: string |
| `convert_table_to_range` | E | D |  | ExcelApi 1.2+ | `name`: string |
| `read_conditional_formats` | E |  |  | ExcelApi 1.6+ | `address`: string |
| `add_conditional_format` | E | D |  | ExcelApi 1.6+ | `address`: string<br>`type`: Custom\|CellValue\|ContainsText<br>`rule`: object<br>`format`: object 可选 |
| `delete_conditional_format` | E | D |  | ExcelApi 1.6+ | `address`: string<br>`id`: string |
| `clear_conditional_formats` | E | D |  | ExcelApi 1.6+ | `address`: string |
| `read_data_validation` | E |  |  | ExcelApi 1.8+ | `address`: string |
| `set_data_validation` | E | D |  | ExcelApi 1.8+ | `address`: string<br>`rule`: object 单一验证类型<br>`ignoreBlanks`: boolean 可选<br>`errorAlert`: object 可选<br>`prompt`: object 可选 |
| `clear_data_validation` | E | D |  | ExcelApi 1.8+ | `address`: string |
| `validate_data` | E |  | details 需 ExcelApi 1.9+ | ExcelApi 1.8+ | `address`: string<br>`details`: boolean 可选 |
| `excel_set_freeze_panes` | E | D |  | ExcelApi 1.7+ | `sheet`: 必填工作表名<br>`mode`: 必填rows\|columns\|range\|none<br>`count`: 仅rows/columns：正整数<br>`range`: 仅range：要被冻结的表内A1范围，如A1:B2；不是光标锚点 |
| `excel_render_range` | E |  |  | ExcelApi 1.7+ | `address`: 必填范围地址；将完整区域渲染为PNG用于视觉复核<br>`maxCells`: 可选整数1..10000，默认5000；超限拒绝，不截断 |
| `excel_export_file` | E |  | 通过Office原生文件API导出；支持情况以宿主实际返回为准 | Common File API; actual host support is checked | `type`: xlsx\|pdf; default xlsx |
| `excel_render_chart` | E |  | Office原生Chart.getImage图表截图 | ExcelApi 1.9+ | `sheet`: string required<br>`name`: string required |
| `excel_set_print_layout` | E | D | 设置指定工作表有界打印区域和原生页面布局，并实际读回 | ExcelApi 1.9+ | `sheet`: string required<br>`address`: sheet-local A1 rectangle; at most 10000 cells<br>`orientation`: Landscape\|Portrait; default Landscape<br>`paperSize`: A4\|Letter\|A3; default A4<br>`fitToPagesWide`: integer1..100; default1<br>`fitToPagesTall`: integer1..100; default1 |

## PowerPoint

| action | 平台 | D | 说明 | 条件 | 参数 |
| --- | --- | --- | --- | --- | --- |
| `read_slides` | P |  | 读取当前选中的幻灯片范围（不是全文件页数）；整份演示文稿用 read_document |  |  |
| `ppt_read_notes` | P |  |  |  |  |
| `ppt_add_slides` | P | D |  | PowerPointApi 1.4+ | `count`: number 可选：1-50，默认1<br>`clearPlaceholders`: boolean 可选：默认 false |
| `ppt_add_elements` | P | D |  | PowerPointApi 1.4+ | `slideId`: string 必填：目标幻灯片 ID<br>`coordinateSpace`: object 必填：元素设计坐标 {width,height}<br>`targetSize`: object 必填：实际目标尺寸 points {width,height}<br>`elements`: array 必填：text\|shape\|line 元素；通用 x/y/w/h，文本支持 font/size/color/bold/italic/align/valign/margins/autoFit/runs，shape 支持 shapeType/fill/lineColor，line 支持 connectorType/color/lineWidth/dashStyle |
| `ppt_delete_slides` | P | D |  | PowerPointApi 1.2+ | `slideIds`: string[] 必填：精确幻灯片 ID；拒绝删除全部页面 |
| `ppt_delete_shapes` | P | D |  | PowerPointApi 1.4+ | `slideId`: string 必填<br>`shapeIds`: string[] 必填：精确 shape ID；删除后返回剩余 IDs |
| `ppt_render_slides` | P |  |  | PowerPointApi 1.8+ | `width`: number 可选：PNG宽度，默认960<br>`height`: number 可选：PNG高度，默认540<br>`slideIds`: 可选非空且不重复的原生页ID数组；不传渲染全稿；输出保留文稿实际页序和原1基页码，先验证全部ID再渲染 |
| `ppt_export_file` | P |  |  |  | `type`: pptx\|pdf; default pptx; native Office export |
| `ppt_inspect_slides` | P |  |  | PowerPointApi 1.4+ | `slideIds`: string[] 可选<br>`withStyles`: 可选boolean：支持时读取text/font/fill，不支持类型明确stylesRead:false与原因 |
| `ppt_list_placeholders` | P |  |  | PowerPointApi 1.8+ | `slideId`: string 可选 |
| `ppt_insert_image` | P | D |  | ImageCoercion 1.1+ | `base64`: string<br>`mode`: positioned\|selectionOrPlaceholder<br>`left`: number<br>`top`: number<br>`width`: number 可选<br>`height`: number 可选<br>`slideId`: optional string: select exact native slide before inserting image |
| `ppt_insert_table` | P | D | style 需 PowerPointApi 1.9+ | PowerPointApi 1.8+ | `slideId`: string<br>`values`: 矩形二维数组<br>`left`: number 可选<br>`top`: number 可选<br>`width`: number 可选<br>`height`: number 可选<br>`style`: PowerPoint.TableStyle 枚举字符串可选，如 MediumStyle2Accent1 |
| `ppt_group_shapes` | P | D |  | PowerPointApi 1.8+ | `slideId`: string<br>`shapeIds`: string[]<br>`name`: string 可选 |
| `ppt_ungroup_shape` | P | D |  | PowerPointApi 1.8+ | `slideId`: string<br>`groupShapeId`: string |
| `ppt_set_z_order` | P | D |  | PowerPointApi 1.8+ | `slideId`: string<br>`shapeIds`: string[]<br>`operation`: BringForward\|BringToFront\|SendBackward\|SendToBack |
| `ppt_align_shapes` | P | D |  | PowerPointApi 1.4+ | `slideId`: string<br>`shapeIds`: string[]<br>`mode`: left\|center\|right\|top\|middle\|bottom |
| `ppt_distribute_shapes` | P | D |  | PowerPointApi 1.4+ | `slideId`: string<br>`shapeIds`: string[]<br>`mode`: horizontal\|vertical |
| `ppt_update_shapes` | P | D |  | PowerPointApi 1.4+ | `slideId`: 必填原生幻灯片ID<br>`updates`: 必填1..500个patch：{shapeId,text?,font?:{name,size,color,bold,italic},fill?:{color,transparency},left?,top?,width?,height?}；几何与字号单位pt；仅修改显式字段；text替换整段文本，不支持runs局部编辑；文本/字体/填充限TextBox/GeometricShape。 |
| `ppt_read_presentation_info` | P |  |  | Common API getFileAsync (Compressed) |  |

## 能力边界

有些事 Office.js 就是不提供，插件会如实返回错误码而不是假装成功：

- **平台级没有的 API**：PPT 原生图表创建、图片裁剪、任意 placeholder ID 定向插图；Word 的跨平台 Link to Previous / 剪贴板移动
- **本机运行时缺的**：Word 的 `body.comments` 在部分版本上不存在；段落排版使用真实的 Paragraph 属性，内置样式需1.3、页面设置需Desktop1.3
- **已知的 API 坑**：Excel 没有 `workbook.getRange` / `calculate`（已改用官方替代 API）；`Range.autoFilter` 必须走 worksheet 级
- **shapeType 的合法值由 Office 决定**：插件不猜枚举，任何 PascalCase 名字都放行；执行失败时把「已验证可用」的清单补进报错，不让调用方只看到一个光秃秃的 `InvalidArgument`

## 加自己的 action

内置清单之外，你可以在 `server/user-actions.js` 里加自定义 action 或覆盖内置行为——见[开发文档](DEVELOPMENT.md#用户定制user-actionsjs)。
