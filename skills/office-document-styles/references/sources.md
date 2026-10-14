# 样式来源与使用边界 / Design sources

获取日期 / Retrieved: **2026-10-13**（Asia/Shanghai）。保留六个官方结构参考，并补充 Vertex42 预算差额及微软财务展示参考；第三方模板、字体、图片或母版均不打包。仓库中的 Office/PDF/PNG 是本项目原创原生范本及其导出。

默认交付为Word、Excel、PPT各一份元素型样式范本，见[目录](../resources/style-atlas.json)。历史业务配方保留兼容。以下来源用于链接参考和总结布局原则，**不是复制模板**。本技能自行编写的设计 token、recipe 和实现按 **MIT** 发布；MIT 不改变第三方模板、网页文字、图片、字体、商标的权利。Microsoft 页面上的“免费”“可定制”不能当作插件再分发许可。无明示授权的资产不得打包。

证据范围：本次读取的是页面文本及目录标签，没有下载或检查 DOCX、XLSX、PPTX、OTP 内部，也没有进行像素级视觉审阅。以下具体色号、字号、边距与业务页型均为本项目原创选择，不是来源模板属性；原生范本的实际读回与导出验收另按质量指南执行。

## Word：封面、目录与章节

- ID：`word-business-report`
- 来源：[Microsoft Word report templates — Business report (graphic design)](https://word.cloud.microsoft/create/en/papers-and-reports/)
- 获取：2026-10-13；状态：`reference-only`。
- 原页证据：说明使用一致的 headings、appendix sections 和易扫描、易导航的布局；报告结构包括封面、摘要、目录、正文、图表、结论与建议。Business report (graphic design) 的目录资源标签包含 blue / modern / simple。页面说明可调整 fonts、colors、headings，没有证实可把模板或资产随插件再分发。
- 采用原则：分级标题和证据结构一致，摘要、问题、措施、责任和建议清晰。只引用链接并原创实现，不复制母版、图标、配图或模板文件。
- 原创实际方案：正式工作报告用宋体正文12pt、两端对齐、首行缩进24pt、段前后0、1.5倍行距；黑体18pt居中标题/14pt一级标题，A4/25mm人工核验。原生范本已人工设置宋体10.5pt动态PAGE页码，插件字段API仍P1；不宣称国标公文或自动字段支持。

## Word：语义结构与可访问性

- ID：`word-accessible-document`
- 来源：[GOV.UK — Publishing accessible documents](https://www.gov.uk/guidance/publishing-accessible-documents)
- 获取：2026-10-13；状态：`ogl-guidance-reference`。
- 许可证据：页脚写明 “All content is available under the Open Government Licence v3.0, except where otherwise stated”。[OGL v3.0 原文](https://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/)允许复制、分发、改编和商业使用，但需署名并遵守例外；logo、商标及未获授权的第三方权利等不在许可范围内，不得暗示官方背书。
- 原页原则：无衬线字体至少 12pt、左对齐、单一连续列；用真正的 Heading 样式，不跳标题级别；表格仅用于数据，避免合并/拆分；不能仅靠颜色或形状表达含义。
- 原创实际方案：业务请示沿用普通中文正式办公字体与段落角色，说明事项、依据、成本、风险与审批请求。GOV.UK的无衬线/左对齐建议如上保留为来源事实，但本项目仅参考其语义标题、逻辑结构和非颜色单一编码；宋体两端对齐是独立选择，不能称为GOV推荐。

## Excel：月度预算执行

- ID：`excel-budget`
- 来源：[Microsoft Excel budget templates — Basic personal budget](https://excel.cloud.microsoft/create/en/budget-templates/)
- 获取：2026-10-13；状态：`reference-only`。
- 原页证据：用于跟踪开支、规划月度支出和财务目标，可替换单元格值/名称及自定义配色。Basic personal budget / Basic monthly budget 的目录标签为 green / modern / simple。免费和可定制不构成已证实的插件再分发许可；未检查原模板公式。
- 采用原则：收入、支出、余额层级；明细分组；统一金额格式；输入区与计算区用文字标签加浅底共同区分。只引用链接，不复制工作簿、公式或资产。
- 原创实际方案：3个工作表、16条预算业务记录；预算/实际/差额、分组分析和问题/假设完整。KPI由SUM/SUMIF/COUNTIF/IFERROR等原创公式生成，原生Chart明确source-worksheet；概览740pt、图表D7:H19。A4横向90%及重复行1:5/ID列A:B为人工打印QA，不是插件自动打印能力。

## Excel：项目执行的表格结构参考

- ID：`excel-accessible-reference-table`
- 来源：[Government Analysis Function — Releasing statistics in spreadsheets](https://analysisfunction.civilservice.gov.uk/policy-store/releasing-statistics-in-spreadsheets/)
- 获取：2026-10-13；状态：`ogl-guidance-reference`。
- 许可证据：[Terms of use](https://analysisfunction.civilservice.gov.uk/terms-of-use/)第 6 节说明大部分内容为 Crown copyright、适用 OGL，复制时需注明 AF 来源、保留标题/署名/URL 并遵守单项告知。同一条款明确限制将网站图片保存或用于其预期用途之外。**文字规则与网站图片许可不能混为一谈**；示例 ODS 内部未审计，不打包示例或截图。
- 原页原则：真正命名/标记 Table，单一表头，避免合并/拆分/嵌套单元格；表内不设空行空列；文字左对齐、数字右对齐；列内精度一致、千分位与前导零；换行并调行高列宽；合计放在表边缘。
- 原创实际方案：3个工作表、16条项目记录，明细、加权进度、固定日期状态和问题责任互相勾稽；真实Table/Chart、明确引用范围，金额/人数/比例/日期区别格式化。统计发布语境的冻结/筛选建议不一律套用到项目执行。

## PowerPoint：经营分析复盘的布局参考

- ID：`ppt-scientific-discovery`
- 来源：[Microsoft PowerPoint presentation templates — Scientific discovery](https://powerpoint.cloud.microsoft/create/en/presentation-templates/)
- 获取：2026-10-13；状态：`reference-only`。
- 原页证据：页面说明工作/学校场景可高度自定义布局、颜色、logo 和字体；Scientific discovery 的目录标签包含 blue / modern / simple。未证实幻灯片文件、母版或图片可随插件再分发，字体与具体几何结构未检查。
- 采用原则：统一强调色与页型、清晰安全边距、每页一个主张；仅链接参考，不复刻源插画或母版。
- 原创实际方案：经营分析复盘8页、960×540pt起，暖白#FAF9F6、炭黑#24282D、哑铜#836344；标题30pt、正文18pt、表格/轴14pt、重点标签16pt、来源12pt。含6期趋势、分组贡献、毛利桥接、责任行动、里程碑、风险与审批，数据图/分格表格为可编辑形状和文本，不是原生Chart/Table API。

## PowerPoint：立项提案的几何结构参考

- ID：`ppt-libreoffice-modern-cc0`
- 来源：[LibreOffice Extensions — Modern template](https://extensions.libreoffice.org/en/extensions/show/5139)
- 获取：2026-10-13；状态：`cc0-candidate`。
- 许可证据：官方扩展页称模板有多种配色，Green、Orange、Blue 三个 release 的 License 列分别明确标为 CC0。[CC0 释义页](https://creativecommons.org/publicdomain/zero/1.0/deed.en)说明可以复制、修改、分发和商业使用而无须申请许可；但不涵盖商标、专利与第三方权利，不得暗示作者背书。网站页脚默认的 CC BY-SA 3.0（网站文字/图片）与 release 的 CC0 必须区分，官网截图不能自动视为 CC0。
- 使用边界：本技能仍只引用链接和自行设计，不下载资产。该模板可作为未来实际再分发的候选；若另行决定打包，需先审计 OTP 内部第三方素材、字体与许可，并验证 PPTX 兼容性。本次没有完成这些检查。
- 采用原则：统一信息层级和几何结构；本技能的流程、风险、决策版式与主题独立设计，不以替换蓝/绿/橙配色充当完整模板。
- 原创实际方案：立项提案8页、960×540pt起，冷白#F8FAFC、墨黑#1D2733、雾蓝#44657D；字号角色同经营复盘。涵盖痛点证据、目标、方案比较、流程、投入产出假设、路线、资源风险和审批；拒绝720×405，不缩字号，透明文本框、数字保护、中文避头尾与实际时间线轴均由原创配方实现。

## Excel补充：预算实际差额结构

- ID：`excel-business-budget-vertex42`；[Vertex42 Business Budget](https://www.vertex42.com/ExcelTemplates/business-budget.html) 与 [Private Use许可](https://www.vertex42.com/licensing/EULA_privateuse.html) 均在2026-10-13读取HTTP 200。
- 原页说明 Difference = Actual − Budget，并提醒变更分类后复核SUM合计。只参考预算/实际/差额结构与统计对账；不复制模板公式、工作簿或图像，本站原创公式独立编写。
- Private Use允许个人/组织内部使用和有限私人分享，禁止公开分发/出售模板及其衍生文件；有限截图例外不是插件资产再分发许可。仅浏览公开页和示例，不打包第三方文件、截图、公式或衍生模板。

## PowerPoint补充：微软财务展示结构

- ID：`ppt-microsoft-financial-results`；[Microsoft FY25 Q4 earnings release](https://www.microsoft.com/en-us/Investor/earnings/FY-2025-Q4/press-release-webcast)，2026-10-13读取HTTP 200，状态reference-only。
- 参考结论先行、可比指标、分部贡献、口径解释、风险和假设披露；两份PPT的业务数据均为原创DEMO，不使用微软财务数值、标志、图像或幻灯片资产，不推定其有开放再分发许可。

## 中文字体与许可处理

字体仅引用用户已安装名称：Word正文宋体/SimSun、标题黑体/SimHei；Excel和PPT使用Microsoft YaHei，PPT英文回退Arial。运行时检测字体存在性，不下载、不嵌入、不再分发字体文件。不同 Office/操作系统不保证这些字体都存在，需重新检查换行与版面。

若直接复用 OGL 文字，应保留来源并附上许可链接；默认署名声明为：

> Contains public sector information licensed under the Open Government Licence v3.0.

AF 文字另保留 AF 来源、标题、署名和 URL。本文件的来源记录不表示得到任何机构的推荐或背书。[Microsoft Services Agreement](https://www.microsoft.com/en/servicesagreement)还指出 Microsoft 365 应用受补充使用条款约束；该协议不是模板开源许可。本次结论是“未证实再分发权”，不把其他服务条款泛化为所有 Office 模板的绝对禁止规则。

## English summary

Six original references plus Vertex42 budgeting and Microsoft financial-reporting pages are recorded as retrieved on 2026-10-13. They inform independently authored layout principles and MIT-licensed recipes; no template, screenshot, font or other asset was downloaded or copied. Microsoft references remain reference-only because plugin redistribution rights were not established. GOV.UK/Analysis Function guidance text has OGL conditions and exceptions; Analysis Function site images are expressly restricted. LibreOffice Modern template release rows declare CC0, making the template a candidate for a future asset audit, not a bundled asset in this skill. Vertex42 Private Use forbids public redistribution of templates and derived files; no source workbook, image or formula is bundled. Chinese SimSun justified text is independently chosen, not a GOV.UK recommendation. The native Office examples and their PDF/PNG exports are original; tokens and business content are not extracted template properties.
