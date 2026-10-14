# dsh-ms-office

> [English](README.en.md) · 中文

在 [DeepSeek Harness（DSH）](https://github.com/deepseek-ai/deepseek-harness) 会话里，直接操控你**已经打开的** Microsoft Office 文档——Word / Excel / PowerPoint 的读取、写入、格式化、结构操作、图表 / 公式 / 批注。

这是一个**独立的 dsh 插件**：装上它，dsh 就多了一个能读写 Office 文档的 agent 能力。除了官方 dsh，它不依赖任何其他项目。

> **「Office」指 Microsoft Office 桌面版**，不含 WPS 等兼容产品，也不含网页版。

![运行时数据流：dsh 会话 → 桥接服务 → 加载项窗格 → Office 文档](docs/architecture-runtime.svg)

**它和「让 AI 生成一个 .docx 文件」不是一回事**：加载项跑在 Office 进程里，操作的是**内存中的活文档**——你正在编辑的那一份。没有文件锁，没有「最后保存者赢」，你的编辑和 AI 的编辑由 Office 统一串行处理。

## 能做什么

- **操作活文档**——不落盘、不加锁，AI 改的就是你眼前那份
- **92 个 action**，覆盖三个应用：读选区、写内容、格式化、表格 / 图表 / 公式 / 批注、PPT 分组与对齐……
- **多文档精确路由**——每份文档一个实例，指令必须指明目标，不会打到错误的文档上
- **覆盖类操作有护栏**——写之前先读一遍当前状态，默认还要你在窗格里点「更改」才执行（见[安全](#安全)）
- **中英双语界面**，跟随你的 Office 显示语言

两类技能协同：[ms-office](skills/ms-office/SKILL.md) 负责安全操作；[基础样式技能](skills/office-document-styles/SKILL.md) 提供每种格式一份原创基础样式范本：Word覆盖文档元素，Excel固定两个Sheet，PPT展示图形、表格、空间布局与配色，并附来源记录和分层质量检查。查看[样式预览](docs/style-gallery.html)与[功能路线](docs/FEATURE-ROADMAP.md)。

中文默认字体：Word正文宋体、标题黑体；Excel宋体；PowerPoint标题黑体，正文及表格/图表标签楷体。范本修改通过实际插件执行，导出与截图保留原生Office证据。

完整操作清单见 [docs/CAPABILITIES.md](docs/CAPABILITIES.md)。

## 兼容性

|             | 状态                                           |
| ----------- | -------------------------------------------- |
| **Windows** | ✅ 完整支持（加载项自动注册走 WEF 注册表）                     |
| **macOS**   | ✅ 核心可用（Node + 加载项）；自动注册不适用，需在 Office 里手动加载清单 |
| **Linux**   | ❌ 没有 Office 桌面版加载项宿主                         |

- **Node.js ≥ 18**（桥接服务只用标准库）
- **Microsoft Office 桌面版**（Word / Excel / PowerPoint）
- **dsh**：见 `package.json` 的 `peerDependencies`

## 安装

```bash
dsh plugin --profile web add dsh-ms-office
```

**装完就完事了，不需要再做任何事。** 插件会自己完成剩下两件事：

1. **把自己加进 dsh 的加载清单**——dsh 加载的是 `dsh.profile.bundles` 这个有序清单，只装依赖不会被加载；官方安装路径会自动维护它。
2. **把自己的技能发布给 dsh 会话**——AI 据此知道「有这么个能力、什么时候该用」。

唯一例外：如果你以前**禁用**过它，去插件面板点一下「启用」即可（官方为避免误启用，不会自动把一个已存在的依赖加回来）。

> 想确认清单内容、或想手动排查？见[开发文档的「安装与加载」](docs/DEVELOPMENT.md#安装与加载)。

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
powershell -ExecutionPolicy Bypass -File server\sideload.ps1
powershell -ExecutionPolicy Bypass -File server\sideload.ps1 -Remove
```

## 安全

**默认每次覆盖类操作都要你批准。** 窗格先展示「当前 → 修改后」，预览失败禁止执行。你点「更改」后，窗格使用绑定命令的一次性内部批准入口；外部 `args.confirm` 被忽略，内部批准也不用 `confirm` 参数。审批最长 120 秒，且不超过命令到期时间。

命令默认等待期限为 180 秒。执行中超时可能进入 `execution_unknown`，继续占用实例直到迟到结果到达；先查原命令状态并核对文档，勿盲重发写入。脚本可用 `idempotencyKey` 复用同一次操作，详见[集成方文档](docs/AGENT.md)。

**只读文档受保护。** 文档处于只读 / 查看模式时，覆盖类操作直接拒绝。

**桥接服务只认本机自己人。** 不返回任何 CORS 头，带 `Origin` 的请求必须是服务自身来源，并校验 `Host` 挡 DNS 重绑定。静态面只暴露窗格需要的 5 个文件。

## 排查

| 现象                        | 处理                                                                                    |
| ------------------------- | ------------------------------------------------------------------------------------- |
| 窗格显示「桥接服务未连接」             | 检查设置面板的 `starting` / `healthy` / `unhealthy` / `stopped` 状态；服务健康后再检查窗格连接 |
| 指令报 `timeout`             | 查询原命令状态；`execution_unknown` 时先核对文档，不要盲重发 |
| 指令立即返回 `addin_offline`    | 窗格没开：打开文档 → 开始 / 开发人员 → 加载项 → 开发人员加载项 → 打开「DSH Office 执行器」                          |
| 窗格没出现                     | Windows：确认服务起过（自动注册需要它跑一次），然后**关闭并重开**文档；确认 Office 不是以管理员身份运行（回环地址例外要求普通权限）             |
| 改了端口后窗格连不上                | 关闭并重新打开 Office 文档（加载项地址写在清单里，需要重新加载）                                                   |
| 修改没生效 / 改错了位置             | 确认执行位置（光标 / 选区）正确；看窗格底部的执行日志                                                          |

## 文档

| 文档 | 内容 |
|---|---|
| [docs/CAPABILITIES.md](docs/CAPABILITIES.md) | 全部 action 清单与能力边界 |
| [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) | 架构、文件职责、分片契约、i18n、热更新、自定义 action、打包约定 |
| [docs/AGENT.md](docs/AGENT.md) | HTTP 接口、错误码、审批协议（写脚本 / 别的 agent 对接用） |

## 许可

[MIT](LICENSE)。第三方组件（pako）的许可与出处见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。

---

*非微软官方产品。本项目与 Microsoft Corporation 无关联。*

---

本项目 fork 自 <https://github.com/Mikuzjc/dsh-office-for-mso>。
