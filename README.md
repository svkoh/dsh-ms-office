# dsh-office-for-mso

> [English](README.en.md) · 中文

在 [DeepSeek Harness（DSH）](https://github.com/deepseek-ai/deepseek-harness) 会话里，直接操控你**已经打开的** Microsoft Office 文档——Word / Excel / PowerPoint 的读取、写入、格式化、结构操作、图表 / 公式 / 批注。

这是一个**独立的 dsh 插件**：装上它，dsh 就多了一个能读写 Office 文档的 agent 能力。除了官方 dsh，它不依赖任何其他项目。

> **「Office」指 Microsoft Office 桌面版**，不含 WPS 等兼容产品。

```
你 ──▶ dsh 会话(AI) ──POST /office/command──▶ 桥接服务（本机 HTTP，127.0.0.1）
                                                  │ 指令入队（按 host / instance 路由）
                           Office 加载项窗格（文档内轮询，1s）
                                                  │ Office.js 执行（78 个 action）
                                                  ▼
                        结果回传 ──▶ AI 取回 ──▶ 汇报给你
```

**它和「让 AI 生成一个 .docx 文件」不是一回事**：加载项跑在 Office 进程里，操作的是**内存中的活文档**——你正在编辑的那一份。没有文件锁，没有「最后保存者赢」，你的编辑和 AI 的编辑由 Office 统一串行处理。

## 特性

- **操作活文档**：不落盘、不加锁，AI 改的就是你眼前那份
- **78 个 action**：从读选区到插入原生公式、从条件格式到 PPT 分组对齐
- **多文档精确路由**：Word / Excel / PowerPoint 各跑一个窗格实例，指令带 `instance` 指定目标文档，不会打到错误的文档上
- **覆盖类操作有机制级护栏**：写之前先读一遍当前状态，默认还要你在窗格里点「更改」才执行（详见[安全模型](#安全模型)）
- **action 可热更新**：改 `server/actions/*` → 重启桥接服务 → 窗格自动重新加载，不用重开窗格

---

## 兼容性

| | 状态 |
|---|---|
| **Windows** | ✅ 完整支持（加载项自动注册走 WEF 注册表） |
| **macOS** | ✅ 核心可用（Node + 加载项）；自动注册不适用，需在 Office 里手动加载清单 |
| **Linux** | ❌ 没有 Office 桌面版加载项宿主 |

- **Node.js ≥ 18**（桥接服务只用标准库）
- **Microsoft Office 桌面版**（Word / Excel / PowerPoint）
- **dsh**：见 `package.json` 的 `peerDependencies`（对 `@deepseek-ai/dsh-client-ui-slots` 声明了范围，dsh 安装时会校验）

## 安装

### 作为 dsh 插件安装（推荐）

```bash
# 装进某个 profile（web / 你的自定义 profile 名）
dsh plugin --profile web add dsh-office-for-mso

# 或直接从仓库装
dsh plugin --profile web add github:svkoh/dsh-office-for-mso

# 本地开发时指向 checkout
dsh plugin --profile web add /path/to/dsh-office-for-mso
```

> `dsh plugin --profile <name> <args>` 会把参数原样转发给 profile 目录下的 pnpm。

装完后**确认 profile 清单里列了这个包**——dsh 加载的是 `dsh.profile.bundles` 这个有序清单，只装依赖不会被加载：

```jsonc
// $DSH_HOME/profiles/web/package.json
{
  "dsh": {
    "profile": {
      "bundles": ["@deepseek-ai/dsh-base", "@deepseek-ai/dsh-web-app", "dsh-office-for-mso"]
    }
  }
}
```

用 dsh 的插件管理面板安装会自动维护这个清单；用 `dsh plugin ... add` 手装的话需要自己确认一下。

### 让 AI 真的会用这个技能

插件启动时会把自带的技能发布到 `~/.agents/skills/office-for-mso`，dsh 会话据此知道「有这么个能力、什么时候该用」。如果你禁用了技能的自动发布，也可以手动复制：

```powershell
mkdir "$HOME\.agents\skills\office-for-mso" -Force | Out-Null
copy "skills\office-for-mso\SKILL.md" "$HOME\.agents\skills\office-for-mso\"
```

## 首次使用

1. 打开一个 Word / Excel / PowerPoint 文档
2. 打开加载项窗格：
   - **Windows**：桥接服务启动时会自动把加载项注册进 `HKCU\...\WEF\Developer`。**关闭并重新打开** Office 文档即可看到窗格。
   - **macOS**：文件 → 选项 → 自定义功能区 → 勾选「开发人员」→ 打开文档 → 开发人员 → 加载项 → 管理「开发人员加载项」→ 添加 `manifest.xml`（Windows 上手动加载同理）
3. 窗格显示 **已连接：等待 DSH 指令** 即成功
4. 在 dsh 会话里说「读一下当前 Word 的全文」之类即可

窗格可以缩小、拖到角落，但**用的时候必须开着**——Office.js 的执行发生在窗格里，关掉就会得到 `addin_offline`。

手动注册 / 卸载加载项：

```powershell
powershell -ExecutionPolicy Bypass -File server\sideload.ps1            # 注册
powershell -ExecutionPolicy Bypass -File server\sideload.ps1 -Remove    # 卸载
```

## 架构

| 文件 | 职责 |
|---|---|
| `lib/index.js` | 插件宿主侧入口：拉起桥接服务、挂设置面板路由 |
| `lib/sidecar-host.mjs` | 桥接服务的生命周期：spawn、健康检查、端口选择、崩溃重启、退出清理 |
| `lib/client.js` | 插件设置面板（状态灯 / 端口 / 自动启动开关） |
| `cordis.patch.yml` | 挂载本插件的 loader 行 |
| `server/server.js` | 桥接服务：指令队列、路由、心跳、能力发现、静态文件 |
| `server/taskpane.js` | 加载项窗格：轮询 / 心跳 / 分发 / actions 热更新 |
| `server/actions/` | **按宿主分片**的 action 实现：`word.js` / `excel.js` / `ppt.js` / `common.js`（宿主无关）/ `runtime.js`（共享运行时）/ `registry.js`（装配与对账） |
| `server/actions/parts.json` | 分片拼接清单（顺序 = 依赖顺序）；服务端与测试读同一份 |
| `server/capabilities.json` | **能力表的唯一真源**（名称 / 平台 / 是否覆盖类 / 参数说明） |
| `server/taskpane.html` | 窗格 UI（含审批面板） |
| `server/manifest.xml` | Office 加载项清单模板（端口 3000；运行期按实际端口改写后写到 `$DSH_HOME/dsh-office-for-mso/manifest.xml`） |
| `server/pako.min.js` | 解压 OOXML 用的 zlib 实现（第三方，见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)） |
| `skills/office-for-mso/SKILL.md` | 给 AI 看的技能说明（查状态 → 发指令 → 错误处理 → 安全约定） |
| `test/` | 105 项测试（`npm test`，零测试框架依赖，用 `node --test`） |

**多文档模型**：每个窗格有自己的 `instanceId`。`GET /office/status` 返回在线文档列表，含 `docUrl`（文档路径，用来识别是哪一份）和 `docTitle`。发指令必须带 `instance`，否则拒绝——这是防止「多文档同时打开时指令被错误文档执行」的机制。

**热更新**：窗格每次轮询前比对 `/actions.js` 的版本（= 全部分片 mtime 的最大值），变了就动态重载；心跳里还会比对 `server/taskpane.js` 的 mtime（`codeVersion`），变了就自动 reload 整个窗格。所以改完代码**重启桥接服务**即可生效。

**端口**：默认 `127.0.0.1:3000`。可以在插件设置面板里改；端口被占用时插件会另选一个空闲端口，并同步改写加载项清单里的地址。改端口后需要关闭并重新打开 Office 文档。

## 能力（78 个 action）

> `server/capabilities.json` 是唯一真源，`GET /office/capabilities` 返回同一份数据。下面这张表只是给人看的概览。
> **D** = 覆盖类操作（destructive）：写之前会先读当前状态，默认需要你在窗格审批。
> 平台：W = Word，E = Excel，P = PowerPoint。

### 通用

| action | 平台 | D | 说明 |
|---|---|---|---|
| `read_selection` | W/E/P | | 读当前选区；`withStyles` 附带样式 |
| `write_selection` | W/E/P | D | 用 `text` 替换当前选区 |
| `read_document` | W/E/P | | Word 全文 / Excel 已用区域（`sheet` 指定表）/ PPT 逐页文本 |
| `get_environment` | W/E/P | | 宿主版本、平台、requirementSets 支持情况 |
| `get_document_info` | W/E/P | | 文档 url / title / mode（只读判断）/ settings 键值 |
| `set_setting` | W/E/P | D | 往文档里写持久化键值（随文档保存，跨会话可读） |

### Word

| action | 平台 | D | 说明 |
|---|---|---|---|
| `read_styles` | W/E | | 选区样式（字体 / 字号 / 加粗 / 颜色 / 高亮等） |
| `replace_all` | W/E | D | 全文查找替换 `search` / `replace`；`dryRun` 只返回命中数 |
| `append_text` | W | D | 文末追加段落 |
| `read_tables` | W | | 结构化读取全部表格及表格样式 |
| `read_table_borders` / `set_table_borders` | W | D | 读取 / 设置六类表格边框（WordApi 1.3+） |
| `set_font` | W | D | 全文（含表格内文字）设字体 |
| `word_list_paragraphs` / `word_delete_paragraphs` | W | D | 先列出带零基 `index` 的段落，再用 `targets:[{index,expectedText}]` 显式删除 |
| `insert_paragraph` | W | D | 插入段落，可指定内置样式与插入位置 |
| `insert_table` | W | D | WordApi 1.3 原生可编辑表格（`styleBuiltIn` 跨语言） |
| `word_insert_html` | W | D | 往正文或选区插入 HTML |
| `word_set_page_setup` | W | D | 按节设置页边距、页眉页脚距离（point） |
| `word_insert_equation` | W | D | 插入原生公式（结构化节点，走 OOXML） |
| `word_read_equations` | W | | 读取全文公式 OOXML |
| `word_read_selection_paragraph` | W | | 读选区所在段落及其中的公式 OOXML |
| `word_export_file` | W | | 导出 docx / pdf（base64） |
| `word_insert_image` | W | D | 选区前 / 后 / 替换插图 |
| `word_transfer_content` | W | D | 在选区、正文、表格、搜索命中、页眉页脚之间搬运 OOXML |
| `apply_style` | W | D | 应用内置样式（selection / all） |
| `format_selection` | W/E | D | 选区格式化（字体 / 字号 / 加粗 / 斜体 / 颜色 / 高亮） |
| `set_paragraph_format` | W | D | 段落格式（对齐 / 缩进 / 行距 / 列表） |
| `search` | W | | 查找文本，返回命中列表 |
| `add_comment` / `read_comments` | W/E | D | 加 / 读批注 |
| `read_properties` | W/E | | 文档属性（标题 / 作者 / 字数等） |
| `word_list_sections` | W | | 列出节 |
| `word_read_headers_footers` / `word_set_header_footer` | W | D | 读 / 写页眉页脚 |
| `word_insert_break` | W | D | 插入分页 / 分节符 |
| `word_insert_caption` | W | D | 插入题注（静态编号已实测；动态 `SEQ` 需显式开启） |
| `locate_select` | W/E | | 定位并选中（书签 / 命名区域 / 文本 / Excel 区域），零副作用 |

### Excel

| action | 平台 | D | 说明 |
|---|---|---|---|
| `list_sheets` | E | | 列出工作表（名称 / 位置 / 可见性） |
| `read_range` | E | | 读区域（值 / 公式 / 数字格式） |
| `write_range` | E | D | 批量写 `values` 或 `formulas`（>5000 格自动分块） |
| `format_range` | E | D | 区域格式化并回读实际格式 |
| `insert_chart` | E | D | 数据 → 图表（Column / Line / Bar / Pie / Area / Scatter…） |
| `add_sheet` / `rename_sheet` / `delete_sheet` | E | D | 工作表管理（`delete_sheet` 支持 `dryRun`） |
| `apply_sort` | E | D | 多字段排序 |
| `apply_filter` | E | D | AutoFilter |
| `evaluate_formula` | E | | 求值 `SUM(A1:A5, C1:C5)` 这类表达式（白名单 SUM / AVERAGE / COUNT / MAX / MIN / PRODUCT） |
| `list_tables` / `create_table` | E | D | 枚举 / 创建原生表对象 |
| `delete_table` / `convert_table_to_range` | E | D | 删表及数据 / 只移除表对象保留数据 |
| `read_conditional_formats` / `add_conditional_format` / `delete_conditional_format` / `clear_conditional_formats` | E | D | 条件格式生命周期（ExcelApi 1.6+） |
| `read_data_validation` / `set_data_validation` / `clear_data_validation` / `validate_data` | E | D | 数据验证生命周期（1.8+；`validate_data` 的 details 需 1.9+） |

### PowerPoint

| action | 平台 | D | 说明 |
|---|---|---|---|
| `read_slides` | P | | 读当前选中的幻灯片范围（整份用 `read_document`） |
| `ppt_read_notes` | P | | 全文件备注（解析 OOXML 的 notesSlides） |
| `ppt_inspect_slides` | P | | 结构快照：shape ID / 类型 / 名称 / 坐标 / 层序 / 文本 |
| `ppt_list_placeholders` | P | | placeholder 类型与边界 |
| `ppt_add_slides` / `ppt_add_elements` | P | D | 批量建页，再按返回的 `slideId` 添加可编辑元素 |
| `ppt_delete_slides` / `ppt_delete_shapes` | P | D | 按明确 ID 删除（拒绝清空演示文稿） |
| `ppt_insert_image` | P | D | 插图（定位 或 当前选区 / placeholder） |
| `ppt_insert_table` | P | D | PowerPointApi 1.8+ 原生可编辑表格 |
| `ppt_group_shapes` / `ppt_ungroup_shape` / `ppt_set_z_order` | P | D | 分组 / 解组 / 层序 |
| `ppt_align_shapes` / `ppt_distribute_shapes` | P | D | 对齐 / 分布（按 shape 边界做确定性几何计算） |
| `ppt_render_slides` | P | | 逐页渲染 PNG |
| `ppt_export_file` | P | | Office 原生导出 |

### 能力边界

有些事 Office.js 就是不提供，插件会如实返回错误码而不是假装成功：

- **平台级没有的 API**：PPT 原生图表创建、图片裁剪、任意 placeholder ID 定向插图；Word 的跨平台 Link to Previous / 剪贴板移动
- **本机运行时缺的**：Word 的 `body.comments` 与 `paragraphFormat` 在部分版本上不存在（插件返回 `requirement` 错误，不假装成功）
- **已知的 API 坑**：Excel 没有 `workbook.getRange` / `calculate`（已改用官方替代 API）；`Range.autoFilter` 必须走 worksheet 级

## 安全模型

这一层值得单独说，因为它决定了「一个网页能不能借你的 Office 文档做坏事」。

### 桥接服务只认自己人

桥接服务监听 `127.0.0.1`，但**监听回环地址不等于安全**——你浏览器里的任何网页都能向 `127.0.0.1` 发请求。所以：

- **不返回任何 `Access-Control-Allow-*` 头**（没有 CORS 全开）
- **带 `Origin` 的请求，其 `Origin` 必须等于服务自身来源**，否则 `403 forbidden_origin`
- **校验 `Host` 头**，挡 DNS 重绑定（攻击者域名解析到 127.0.0.1 时 `Host` 会是攻击者域名）
- **静态文件走白名单**：只服务窗格真正需要的 5 个文件，`manifest.xml`、`*.ps1`、`capabilities.json`、`package.json` 都不在 HTTP 面上

没有 Origin 的调用（本机 curl / agent 经 pwsh 发的请求）照常放行——这是设计内的使用者。

### 覆盖类操作有护栏

- **写前 re-read（机制级）**：所有覆盖类操作执行前先读一遍当前状态，不凭记忆覆盖
- **默认审批模式**：`confirmMode=ask`（默认）时，窗格弹出审批面板显示「当前状态」和「修改后」，你点「更改」才执行。AI **不能**预先发 `confirm:true` 绕过——`confirm:true` 只由窗格在你批准后附加。审批 120s 无操作视为拒绝
- **只读文档保护**：检测到文档处于 `readOnly` / `view` 模式时，覆盖类操作直接返回 `readonly` 错误
- **显式删除目标**：删除段落要先 `word_list_paragraphs` 拿到零基 `index`，再把 `{index, expectedText}` 传回去；内容已变、含图片、文档结尾段都会被拒绝
- **`dryRun` 预览**：`replace_all` / `delete_sheet` 支持只返回影响预览

想改成自动模式，在窗格顶部的下拉里切到「自动模式」（或 `POST /office/config {"confirmMode":"auto"}`）——自动模式下仍会写前 re-read，结果里带 `previousState` / `afterState` 供核验。

## 用户定制（user-actions.js）

想加自己的 action、或覆盖内置行为，在 `server/` 下建 `user-actions.js`（已在 `.gitignore`，升级不会碰它）：

```js
window.__USER_ACTIONS__ = {
  my_action: {
    hosts: ['Word'],
    destructive: false,
    impl: (host, args) => ({ ok: true, result: { text: String(args.text || '') } }),
  },
  // 同名即可覆盖内置 action
};
```

重启桥接服务后生效（热更新，不用重开窗格）。

> `user-actions.js` 里的 action 不经过 `capabilities.json`，所以要**自己带上 `hosts` / `destructive`**；想要审批 + 写后验证，再补一个 `preview`。

## 排查

| 现象 | 处理 |
|---|---|
| 窗格显示「桥接服务未连接」 | 确认插件在设置面板里是「运行中」；或直接 `curl http://127.0.0.1:3000/office/status` |
| 指令报 `timeout` | 加载项没在线：确认文档打开、窗格显示已连接 |
| 指令立即返回 `addin_offline` | 窗格没开：打开文档 → 开始 / 开发人员 → 加载项 → 开发人员加载项 → 打开「DSH Office 执行器」 |
| 窗格没出现 | Windows：确认服务起过（自动注册需要它跑一次），然后**关闭并重开**文档；确认 Office 不是以管理员身份运行（回环地址例外要求普通权限） |
| 改了 `server/actions/*` 没生效 | 重启桥接服务（窗格会自动热更新） |
| 改了 `taskpane.js` 没生效 | 重启桥接服务，然后等下一次心跳（≤5s）自动 reload；仍不行就手动重开窗格 |
| 改了端口后窗格连不上 | 关闭并重新打开 Office 文档（加载项地址写在清单里，需要重新加载） |
| 修改没生效 / 改错了位置 | 确认执行位置（光标 / 选区）正确；看窗格底部的执行日志 |

## 给 AI 调用方的约定

完整错误码表随时可用 `GET /office/errors` 获取（含每个码的处理建议与是否可重试）。

- 发指令前先 `GET /office/status`，确认目标文档的实例在线且心跳新鲜
- **`instance` 必填**——多文档时省略会导致指令被错误的文档执行
- 收到 `addin_offline` **不要重试**，提醒用户打开窗格
- 收到 `confirm_required` 时，把 `result.preview` 展示给用户，等窗格审批；**不得预先发送 `confirm:true`**
- 收到 `rejected` 就停下，尊重用户决定，不要重发
- **「没找到」不是错误**：`search` / `replace_all(dryRun)` 无命中返回 `ok:true + count:0`，如实告知即可

| code | 含义 | 处理 |
|---|---|---|
| `instance_required` | 缺 `instance` | 查 status 拿 instanceId 重发 |
| `addin_offline` | 窗格未开启 / 离线 | 提醒用户开窗格，**不要重试** |
| `busy` | 上一条指令还在跑 | 稍等重发 |
| `timeout` | 90s 无结果 | 查 status；在线可重试一次 |
| `bad_json` / `body_too_large` / `bad_args` | 请求体非法 / 太大 / 参数非法 | 修正后重发 |
| `unknown_action` | action 名不存在 | 查 capabilities 用正确名称 |
| `unsupported_host` | 该 action 不支持当前应用 | 换 action 或换应用 |
| `confirm_required` | 需要用户审批 | 交给窗格审批，不得预先发 `confirm:true` |
| `rejected` | 用户拒绝 / 审批超时 | 停止，**不要重发** |
| `not_found` | 定位 / 查找目标不存在 | 如实告知，**不要重试同一查找** |
| `readonly` | 文档只读 / 查看模式 | 告知用户另存为可编辑副本 |
| `requirement` / `unsupported` | 当前环境缺这个 API | 如实告知，绝不假装成功 |
| `execution` | 执行期异常 | 把 error 原样报告给用户 |

## 开发

```bash
npm install     # 只装 lint 用的 devDependencies；插件本身运行时零依赖
npm test        # 105 项测试
npm run lint
npm run check   # lint + test

# 单独跑桥接服务（调试用）
PORT=3000 node server/server.js
```

**改完代码后重启桥接服务即可**（窗格会自己热更新）：

- 改了 `server/actions/*` → 重启服务
- 改了 `taskpane.js` / `taskpane.html` → 重启服务，等一次心跳自动 reload
- 改了 `server.js` / `lib/*` → 重启 dsh

### 分片契约（`server/actions/`）

服务端按 `parts.json` 的顺序把分片**拼成一个脚本**响应 `/actions.js`：浏览器侧仍然是「一个 `<script>`」，加载与热更新机制完全不变；磁盘上则按宿主分开，Word / Excel / PowerPoint 的代码不再互相穿插。

每个分片都是**自包含 IIFE**，彼此不共享词法作用域，只通过 `window.__OFFICE_ACTIONS__` 交换：

| 分片 | 挂载点 |
|---|---|
| `runtime.js` | `A.isSetSupported` / `A.validateMatrix` / `A.sleep` / `A.notFound` |
| `word.js` | `A.hosts.Word = { actions, selectTargets, previewReaders, afterContent, environmentProbe }` |
| `excel.js` | `A.hosts.Excel = { … }` |
| `ppt.js` | `A.hosts.PowerPoint = { … }` |
| `common.js` | `A.hosts.Common = { actions, previewReaders, afterContent }`（宿主无关，只用 Common API） |
| `registry.js` | 装配：拍平成注册表、与 `capabilities.json` 对账、暴露 `__ACTIONS__` / `__EXECUTE__` / `__AFTER_CONTENT__` |

三条边界由测试强制（`test/actions-boundary.test.mjs`），它们就是「拆开」这件事本身的意义：

1. 宿主分片里**不能**出现按宿主分叉的判断——函数签名里已经没有 `host` 参数
2. `word.js` 不能引用 `Excel.run` / `PowerPoint.run`（反之亦然）
3. `common.js` / `runtime.js` 不能引用任何宿主对象模型

**改能力表只改 `server/capabilities.json`**：`hosts` / `destructive` 由服务端注入的 `window.__ACTIONS_META__` 提供。`registry.js` 加载时会三个方向对账——有实现没登记 / 登记了没实现 / 覆盖类操作缺预览读取器——任何一条不成立就直接抛错拒绝加载，而不是静默丢掉平台校验或审批拦截。

**改了插件字节就要升 `package.json` 的 `version`**——下游按版本号判新旧，同版本不同字节不会被重新安装。

## 许可

[MIT](LICENSE)。第三方组件（pako）的许可与出处见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。

---

*非微软官方产品。本项目与 Microsoft Corporation 无关联。*
