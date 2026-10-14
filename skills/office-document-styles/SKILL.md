---
name: office-document-styles
description: 用每种格式一份原生范本展示和复用Word、Excel、PowerPoint的基本元素、样式和版式；按元素选择并核验真实输出。
whenToUse: 用户要求新建、排版、美化或规范 Word、Excel、PowerPoint 文档，或者需要模板、报告样式、预算表风格、演示文稿版式与质量检查时。
---

# Office 基础样式

目标是让信息清楚、结构可编辑、颜色与排版一致。先理解内容，再选配方。未指定品牌时直接采用该文档类型的默认方案；已有品牌规范、公司模板和用户明确偏好优先。

这份技能提供原创样式参数和命令配方；互联网来源用于参考，出处和授权证据见 [sources.md](references/sources.md)。它不携带第三方模板、字体、插画或母版二进制，也不把配色方案称为原生 Office 主题。

## 资源与入口
每种格式只有一份基础样式范本，按元素和版式复用。不同页面用于展示不同做法，不按近似业务场景增加模板数量。
- [style-atlas.json](resources/style-atlas.json)：三份范本的元素覆盖、字号、布局、动作映射和能力边界。
- [Word 范本](resources/templates/word-base-styles.docx)：封面、真实目录、章节/段落/列表、表格、对齐、颜色、图片与图注、页眉页脚/PAGE。
- [Excel 范本](resources/templates/excel-base-styles.xlsx)：基础样式、数据与图表两个 Sheet。
- [PPT 范本](resources/templates/ppt-base-styles.pptx)：图形、原生表格、数据图、图文、流程、时间线、网格与配色。
- [sources.json](resources/sources.json)、[sources.md](references/sources.md)：互联网参考与授权边界；原生文件及插图均为原创，不分发第三方模板。
- [quality-guide.md](references/quality-guide.md)：元素覆盖、实际读回与原生视觉验收。
PDF与原生PNG预览以核验记录列出的实际导出为准；未完成的导出明确标注，不以外部脚本生成文件代替插件验收。先另存工作副本；按需要取用元素，不强制把整本范本写进用户文档。历史业务配方仅保留兼容，不是默认模板目录。
以本技能目录为基准，用 Node 18+ 运行 scripts/style-atlas.mjs list、show word|excel|ppt 或 check。check只检查定义与文件完整性，不证明原生视觉或桥接执行成功。
用户数据不得被演示数字替换；不为凑页数、字数或记录数编造事实。示例持续标注DEMO，用户资料按用户希望的sourceLabel显示，不打印内部USER标签。
## 执行流程

1. 加载 Office 桥接技能 `ms-office`，读其公告和 `/office/status`，定位唯一目标实例；读取 `/office/capabilities` 和 `get_environment`。动作以当前运行时为准，源码中存在不等于已安装运行时支持。未上线或不满足 requirement 时停止执行，说明缺什么，不伪造成功。
2. 读取文档基线与用户需要。Word 读 `word_list_paragraphs`、`read_document`；Excel 读 `list_sheets`；PPT 读 `ppt_read_presentation_info`、`ppt_inspect_slides({withStyles:true})`、`read_document`。PPT 的真实 `slideSize` 为 pt，不能拿导出 PNG 的像素宽高代替。
3. 按需要选择范本中的元素或版式，读取当前内容与样式后限定修改范围。新建空白文档必须实际读回确认；已有正文、表格和用户幻灯片不得因套样式被清空。Excel新表名须不冲突。批量命令如使用计划，仍须检查preconditions与checkPlan。
4. 向用户简述方案与影响范围，按已有授权执行。保持运行时审批模式，不能通过 `confirm:true`、内部批准入口或更改审批设置来绕过用户审批。计划自身不会获得额外批准权。
5. **依次执行**计划步骤。每条请求使用顶层 `instance` 和该步稳定的 `idempotencyKey`；不同实例不可共用绑定结果。带 `bindings` 的步骤先用 `resolveStep(step,results)` 绑定先前成功结果，再发送；`results[from]` 是该动作返回的 `result` 数据。未绑定的 `null`、占位符或绑定对象不能直接发给 Office。不要用 eval 或猜测幻灯片/形状 ID。
6. 任一步拒绝或失败即停止依赖它的后续写入。已成功的步骤记录回执；不宣称整批原子回滚。`timeout`/`execution_unknown` 先查询命令状态并核对文档，保持原幂等 key，不盲重发。
7. 写后读回内容和样式，再做视觉检查。PPT 用 `ppt_render_slides({slideIds:[仅本轮页ID]})`；Excel 若运行时支持 `excel_render_range` 则渲染目标范围；Word 在Windows打印布局中用 `word_render_pages({expectedPageCount:实际页数})` 检查原生整页（含页眉页脚）；可用时另导出原生PDF。PNG/PDF base64 需要解码到工作区后查看，不能凭“导出成功”声称无溢出。
8. 报告实际完成、发现的问题、修正结果，以及仍需人工检查的分页、打印、字体替换、读屏顺序等事项。样式检查和内容正确性分别核对。

## 三类文档的默认做法
**Word**：同一文件覆盖常用文档元素。正文宋体12pt、首缩24pt、真实1.5倍行距；黑体14pt一级标题，文内标题18pt，封面可28pt。范本实际定义“范本正文”/Normal/Title/Heading1/2/3/Caption全局样式；正文使用“范本正文”，由word_set_styles({styles:[{name:"范本正文",preset:"body-1.5"}]})在插件内原生导入；全局Normal的点数行距不用于宣称1.5倍，含真实TOC和PAGE字段、原生表格和嵌入图片。直接段落格式不等于重定义全局样式；运行时支持WordApi1.5时，用word_read_styles/word_set_styles核验并设置全局样式，用word_insert_field和页眉页脚field参数插入真实TOC/PAGE；lineSpacing参数是pt，真实1.5倍行距仍须通过原生导出核对360/auto。正文通常两端对齐，标签左对齐，独立标题/图注居中，签署信息可右对齐；默认黑色为主，强调色少量使用。
**Excel**：中文字体统一宋体，正文11pt；恰好两个Sheet。基础样式展示标题加粗、表头/边框、对齐、数值类型、真实自动列宽、换行/自动行高、Table/筛选/冻结/条件格式；数据与图表展示源数据、SUM/差额/IFERROR公式与关联原生Chart。format_range可用autoFitColumns:true（旧autoFit兼容）、wrapText:true、autoFitRows:true；尺寸/自动布局需ExcelApi1.2。先格式化数据，只对数据范围自动列宽，再限制长备注列宽、换行并自动行高。不同步骤分别执行，不能同时指定互斥固定尺寸与自动布局。打印用excel_set_print_layout配置；原生PDF的页数、A4横向和实际缩放仍须核验。
**PowerPoint**：中文标题黑体、正文及表格/图表标签楷体；一份多页样式范本，按元素和空间布局组织。真实960×540pt；标题30、正文18、标签16、表格/图表14、来源12pt，禁止自动缩小文字。白底/深蓝/青绿统一配色，深色标题和数字，强调色用于数据与方向。包含真正可编辑原生Table、嵌入图片及可编辑形状图；形状数据图不冒称原生Chart。图文双栏、三列、流程、时间线和网格各有实际示例。只追加或修改授权页；真实页面尺寸与运行时能力优先，PPT表格动作仍检查PowerPointApi1.8。
## 质量底线

参数层必须过关：颜色对比、字号、几何边界、内容预算、动作与宿主匹配、表名与绑定合法。内容层必须核对：结论是否有依据、数字单位是否一致、公式是否读回、样例是否明确标记。实际效果必须检查：分页/截断、表头宽度、文字溢出、字体替换、图像清晰度、阅读与演讲顺序。无法观察的项目明确列为待检查，不能把“没有检测到”说成“全部合格”。


桥接地址从 `$DSH_HOME/dsh-ms-office/bridge.json` 的 `url` 读取；默认 `http://127.0.0.1:3000` 仅供识别，实际连接必须遵循桥接技能解析公告。
