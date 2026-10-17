# dsh-ms-office

[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Node.js](https://img.shields.io/badge/node-%E2%89%A5%2018-3c873a.svg)](#兼容性)
[![CI](https://github.com/svkoh/dsh-ms-office/actions/workflows/ci.yml/badge.svg)](https://github.com/svkoh/dsh-ms-office/actions/workflows/ci.yml)

[English](README.md) · 中文

在 [DeepSeek Harness（DSH）](https://github.com/deepseek-ai/deepseek-harness) 会话里读写你已经打开的 Microsoft Office 文档：Word、Excel、PowerPoint 的内容、格式、结构、表格、图表、公式与批注。

![运行时数据流：dsh 会话 → 桥接服务 → 加载项窗格 → Office 文档](docs/architecture-runtime.svg)

加载项运行在 Office 内部，直接修改你当前打开的那份文档。改动立刻可见，不需要保存或重新打开；你和 AI 也可以同时编辑，由 Office 依次处理。

> 文中「Office」均指 Microsoft Office 桌面版，不含 WPS 等兼容产品，也不含网页版。

## 快速开始

```bash
dsh plugin --profile web add dsh-ms-office
```

`--profile web` 是 dsh 的配置名，和你平时使用的那份保持一致即可。

安装后 dsh 会加载插件，插件随即自动启动本地桥接服务，不需要手动操作。在插件面板确认服务状态为 `healthy`。

1. 打开一个 Word、Excel 或 PowerPoint 文档。
2. 打开加载项窗格：
   - Windows：服务启动时已自动注册加载项。关闭并重新打开文档，窗格即出现。
   - macOS：需要手动加载清单，见 [macOS 手动加载](#macos-手动加载)。
3. 窗格显示「已连接：等待 DSH 指令」即表示就绪。
4. 在 dsh 会话里让它读文档，例如「读一下当前 Word 的全文」。

窗格必须保持打开。Office.js 在窗格中执行，关掉会得到 `addin_offline`。窗格可以缩小或拖到角落。

### macOS 手动加载

文件 → 选项 → 自定义功能区 → 勾选「开发人员」，打开文档后依次进入 开发人员 → 加载项 → 管理「开发人员加载项」，添加 `manifest.xml`。

## 能力

- 覆盖 Word、Excel、PowerPoint 三个应用，共 92 个 action：读选区、写内容、格式化、表格、图表、公式、批注、PPT 分组与对齐等。
- 每份文档对应一个独立实例，指令必须指明目标文档，不会误伤其他文档。
- 覆盖类操作前会先读取当前状态；预览失败直接禁止执行，默认还需要你在窗格中点一次「更改」。
- 界面中英双语，跟随 Office 显示语言。

完整 action 清单见 [docs/CAPABILITIES.md](docs/CAPABILITIES.md)。

插件自带两个技能。[ms-office](skills/ms-office/SKILL.md) 负责安全操作；[office-document-styles](skills/office-document-styles/SKILL.md) 为每种宿主提供一份原创的基础样式范本——Word 覆盖文档元素，Excel 固定两个 Sheet，PowerPoint 展示图形、表格、空间布局与配色，并附带来源记录和分层质量检查。范本预览见[样式画廊](docs/style-gallery.md)，后续计划见[功能路线](docs/FEATURE-ROADMAP.md)。

中文默认字体：Word 正文宋体、标题黑体；Excel 宋体；PowerPoint 标题黑体，正文及表格、图表标签楷体。范本的所有修改都通过插件实际执行，导出与截图保留原生 Office 证据。

## 兼容性

| 平台 | 状态 | 说明 |
| --- | --- | --- |
| Windows | 完整支持 | 加载项自动注册，走 WEF 注册表 |
| macOS | 核心可用 | Node 与加载项正常；自动注册不适用，需在 Office 中手动加载清单 |
| Linux | 不支持 | 没有 Office 桌面版加载项宿主 |

- Node.js ≥ 18（桥接服务只用标准库）
- Microsoft Office 桌面版（Word / Excel / PowerPoint）
- dsh：版本要求见 `package.json` 的 `peerDependencies`

## 文档

| 文档 | 内容 |
| --- | --- |
| [docs/CAPABILITIES.md](docs/CAPABILITIES.md) | 全部 action 清单与能力边界 |
| [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) | 架构、文件职责、分片契约、i18n、热更新、自定义 action、打包约定 |
| [docs/AGENT.md](docs/AGENT.md) | HTTP 接口、错误码、审批协议，供脚本和其他 agent 对接 |

## 许可

[MIT](LICENSE)。第三方组件 pako 的许可与出处见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。

本项目 fork 自 <https://github.com/Mikuzjc/dsh-office-for-mso>。
