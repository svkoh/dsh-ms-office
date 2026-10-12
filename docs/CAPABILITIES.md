# 完整操作清单

> [English](CAPABILITIES.en.md) · 中文

本插件当前支持的全部 action。**机器可读的那份以运行时为准**，随时可以取：

```bash
curl "$(jq -r .url "$DSH_HOME/dsh-office-for-mso/bridge.json")/office/capabilities?lang=zh"
```

（`bridge.json` 的说明见[集成方文档](AGENT.md#先拿到桥接地址)。）

**读法**

- **D** = 覆盖类操作（destructive）：执行前会先读一遍当前状态，默认还要你在窗格审批。
- **平台**：W = Word，E = Excel，P = PowerPoint。
- **枚举大小写不敏感**（`align:"left"` 和 `align:"Center"` 都接受），但非法值被拒时会把合法值清单写进报错——照清单改，不要猜。
- **字段名以 `args` 说明为准**。常见误写已做别名兼容：`kind`→`type`、`width`/`height`→`w`/`h`、`fontSize`→`size`；颜色只接受 `#RRGGBB`。

> `server/capabilities.json` 是**结构**的唯一真源（名称 / 平台 / 是否覆盖类 / requirement），
> 人类可读的参数说明来自 `i18n/<lang>.json`，接口把两者合并后返回。
> 下面这张表只是给人看的概览。

## 通用

| action              | 平台    | D | 说明                                          |
| ------------------- | ----- | - | ------------------------------------------- |
| `read_selection`    | W/E/P |   | 读当前选区；`withStyles` 附带样式                     |
| `write_selection`   | W/E/P | D | 用 `text` 替换当前选区                             |
| `read_document`     | W/E/P |   | Word 全文 / Excel 已用区域（`sheet` 指定表）/ PPT 逐页文本 |
| `get_environment`   | W/E/P |   | 宿主版本、平台、requirementSets 支持情况                |
| `get_document_info` | W/E/P |   | 文档 url / title / mode（只读判断）/ settings 键值    |
| `set_setting`       | W/E/P | D | 往文档里写持久化键值（随文档保存，跨会话可读）                     |

## Word

| action                                                 | 平台  | D | 说明                                                          |
| ------------------------------------------------------ | --- | - | ----------------------------------------------------------- |
| `read_styles`                                          | W/E |   | 选区样式（字体 / 字号 / 加粗 / 颜色 / 高亮等）                               |
| `replace_all`                                          | W/E | D | 全文查找替换 `search` / `replace`；`dryRun` 只返回命中数                 |
| `append_text`                                          | W   | D | 文末追加段落                                                      |
| `read_tables`                                          | W   |   | 结构化读取全部表格及表格样式                                              |
| `read_table_borders` / `set_table_borders`             | W   | D | 读取 / 设置六类表格边框（WordApi 1.3+）                                 |
| `set_font`                                             | W   | D | 全文（含表格内文字）设字体                                               |
| `word_list_paragraphs` / `word_delete_paragraphs`      | W   | D | 先列出带零基 `index` 的段落，再用 `targets:[{index,expectedText}]` 显式删除 |
| `insert_paragraph`                                     | W   | D | 插入段落，可指定内置样式与插入位置                                           |
| `insert_table`                                         | W   | D | WordApi 1.3 原生可编辑表格（`styleBuiltIn` 跨语言）                     |
| `word_insert_html`                                     | W   | D | 往正文或选区插入 HTML                                               |
| `word_set_page_setup`                                  | W   | D | 按节设置页边距、页眉页脚距离（point）                                       |
| `word_insert_equation`                                 | W   | D | 插入原生公式（结构化节点，走 OOXML）                                       |
| `word_read_equations`                                  | W   |   | 读取全文公式 OOXML                                                |
| `word_read_selection_paragraph`                        | W   |   | 读选区所在段落及其中的公式 OOXML                                         |
| `word_export_file`                                     | W   |   | 导出 docx / pdf（base64）                                       |
| `word_insert_image`                                    | W   | D | 选区前 / 后 / 替换插图                                              |
| `word_transfer_content`                                | W   | D | 在选区、正文、表格、搜索命中、页眉页脚之间搬运 OOXML                               |
| `apply_style`                                          | W   | D | 应用内置样式（selection / all）                                     |
| `format_selection`                                     | W/E | D | 选区格式化（字体 / 字号 / 加粗 / 斜体 / 颜色 / 高亮）                          |
| `set_paragraph_format`                                 | W   | D | 段落格式（对齐 / 缩进 / 行距 / 列表）                                     |
| `search`                                               | W   |   | 查找文本，返回命中列表                                                 |
| `add_comment` / `read_comments`                        | W/E | D | 加 / 读批注                                                     |
| `read_properties`                                      | W/E |   | 文档属性（标题 / 作者 / 字数等）                                         |
| `word_list_sections`                                   | W   |   | 列出节                                                         |
| `word_read_headers_footers` / `word_set_header_footer` | W   | D | 读 / 写页眉页脚                                                   |
| `word_insert_break`                                    | W   | D | 插入分页 / 分节符                                                  |
| `word_insert_caption`                                  | W   | D | 插入题注（静态编号已实测；动态 `SEQ` 需显式开启）                                |
| `locate_select`                                        | W/E |   | 定位并选中（书签 / 命名区域 / 文本 / Excel 区域），零副作用                       |

## Excel

| action                                                                                                            | 平台 | D | 说明                                                                            |
| ----------------------------------------------------------------------------------------------------------------- | -- | - | ----------------------------------------------------------------------------- |
| `list_sheets`                                                                                                     | E  |   | 列出工作表（名称 / 位置 / 可见性）                                                          |
| `read_range`                                                                                                      | E  |   | 读区域（值 / 公式 / 数字格式）                                                            |
| `write_range`                                                                                                     | E  | D | 批量写 `values` 或 `formulas`（>5000 格自动分块）                                        |
| `format_range`                                                                                                    | E  | D | 区域格式化并回读实际格式                                                                  |
| `insert_chart`                                                                                                    | E  | D | 数据 → 图表（Column / Line / Bar / Pie / Area / Scatter…）                          |
| `add_sheet` / `rename_sheet` / `delete_sheet`                                                                     | E  | D | 工作表管理（`delete_sheet` 支持 `dryRun`）                                             |
| `apply_sort`                                                                                                      | E  | D | 多字段排序                                                                         |
| `apply_filter`                                                                                                    | E  | D | AutoFilter                                                                    |
| `evaluate_formula`                                                                                                | E  |   | 求值 `SUM(A1:A5, C1:C5)` 这类表达式（白名单 SUM / AVERAGE / COUNT / MAX / MIN / PRODUCT） |
| `list_tables` / `create_table`                                                                                    | E  | D | 枚举 / 创建原生表对象                                                                  |
| `delete_table` / `convert_table_to_range`                                                                         | E  | D | 删表及数据 / 只移除表对象保留数据                                                            |
| `read_conditional_formats` / `add_conditional_format` / `delete_conditional_format` / `clear_conditional_formats` | E  | D | 条件格式生命周期（ExcelApi 1.6+）                                                       |
| `read_data_validation` / `set_data_validation` / `clear_data_validation` / `validate_data`                        | E  | D | 数据验证生命周期（1.8+；`validate_data` 的 details 需 1.9+）                               |

## PowerPoint

| action                                                       | 平台 | D | 说明                                     |
| ------------------------------------------------------------ | -- | - | -------------------------------------- |
| `read_slides`                                                | P  |   | 读当前选中的幻灯片范围（整份用 `read_document`）       |
| `ppt_read_notes`                                             | P  |   | 全文件备注（解析 OOXML 的 notesSlides）          |
| `ppt_inspect_slides`                                         | P  |   | 结构快照：shape ID / 类型 / 名称 / 坐标 / 层序 / 文本 |
| `ppt_list_placeholders`                                      | P  |   | placeholder 类型与边界                      |
| `ppt_add_slides` / `ppt_add_elements`                        | P  | D | 批量建页，再按返回的 `slideId` 添加可编辑元素           |
| `ppt_delete_slides` / `ppt_delete_shapes`                    | P  | D | 按明确 ID 删除（拒绝清空演示文稿）                    |
| `ppt_insert_image`                                           | P  | D | 插图（定位 或 当前选区 / placeholder）            |
| `ppt_insert_table`                                           | P  | D | PowerPointApi 1.8+ 原生可编辑表格             |
| `ppt_group_shapes` / `ppt_ungroup_shape` / `ppt_set_z_order` | P  | D | 分组 / 解组 / 层序                           |
| `ppt_align_shapes` / `ppt_distribute_shapes`                 | P  | D | 对齐 / 分布（按 shape 边界做确定性几何计算）            |
| `ppt_render_slides`                                          | P  |   | 逐页渲染 PNG                               |
| `ppt_export_file`                                            | P  |   | Office 原生导出                            |

## 能力边界

有些事 Office.js 就是不提供，插件会如实返回错误码而不是假装成功：

- **平台级没有的 API**：PPT 原生图表创建、图片裁剪、任意 placeholder ID 定向插图；Word 的跨平台 Link to Previous / 剪贴板移动
- **本机运行时缺的**：Word 的 `body.comments` 与 `paragraphFormat` 在部分版本上不存在（插件返回 `requirement` 错误，不假装成功）
- **已知的 API 坑**：Excel 没有 `workbook.getRange` / `calculate`（已改用官方替代 API）；`Range.autoFilter` 必须走 worksheet 级
- **shapeType 的合法值由 Office 决定**：插件不猜枚举，任何 PascalCase 名字都放行；执行失败时把「已验证可用」的清单补进报错，不让调用方只看到一个光秃秃的 `InvalidArgument`

## 加自己的 action

内置清单之外，你可以在 `server/user-actions.js` 里加自定义 action 或覆盖内置行为——见[开发文档](DEVELOPMENT.md#用户定制user-actionsjs)。
