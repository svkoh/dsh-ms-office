---
name: office-for-mso
description: 通过本机桥接服务和 Office 加载项，操控用户当前打开的 Word/Excel/PowerPoint 文档（读/写/格式化/图表/公式/批注）。用于「读一下 Word 全文」「Excel 生成图表」「把选中的翻译成英文」等请求。
whenToUse: 用户要求操作 Word/Excel/PowerPoint 文档（读取、编辑、格式化、分析、生成图表/公式、批注等），且这些文档是用户当前已打开的。
---

# Office for MSO（DSH ↔ Microsoft Office）

本技能由 `dsh-office-for-mso` 插件提供。插件启动时拉起一个本机 HTTP 桥接服务
（`server/server.js`），并把本技能发布到 `~/.agents/skills` 供 DSH 发现。

## 架构

```
DSH 会话(AI) --POST /office/command--> 桥接服务 --轮询--> Office 加载项窗格 --Office.js--> Word/Excel/PPT
```

- **桥接服务**：本插件托管的子进程，默认监听 `127.0.0.1:3000`。
- **监听端口**：优先用插件设置里配的端口（未配置则 3000）；该端口被占用时插件会另选一个空闲端口，
  并把加载项清单里的 SourceLocation 改写成实际端口。**不要写死 3000** —— 用下面第 1 步的
  `/office/status`，或看插件设置页显示的端口。
- **加载项窗格**：Office 文档侧边的「DSH Office 执行器」必须保持开启（关闭后指令立即 `addin_offline`）。

## 使用流程

1. **查在线实例**：`GET http://127.0.0.1:<port>/office/status` → `instances`：`{instanceId, host, docUrl, docTitle}`
   - 识别文档：**docUrl 优先**；为空（未保存）时用 **docTitle**
   - 目标不在线 → 提醒用户打开文档，并在 **开始/开发人员 → 加载项** 中打开「DSH Office 执行器」
2. **查能力**（可选）：`GET /office/capabilities` —— 78 个 action，含各自支持的平台与参数说明
3. **发指令**：`POST /office/command`，body：`{ "action": "...", "instance": "<instanceId>", "args": {...} }`
   - **`instance` 必填**（多文档同时打开时，不指定实例可能让错误的文档执行指令）
   - 串行等待（超时约 90s）
4. **错误处理**（完整表见 `GET /office/errors`）：
   - `confirm_required` → 把 preview 给用户；**由窗格在用户批准后附加 `confirm:true`**。AI **不得**预先发送 `confirm:true`
   - `rejected` → 停止，不要重发
   - `addin_offline` → 提醒用户打开窗格，不要重试
   - 无匹配的 search/dryRun → `ok:true + count:0`，如实告知，不要重试

## 端口变更后

端口变了，加载项清单里的 SourceLocation 也跟着变；用户需**关闭并重开**相关 Office 文档，
窗格才会按新地址重新加载。

## 安全

- 桥接服务只接受来自本机自身来源的请求：`Origin` 不是它自己的、或 `Host` 不匹配的请求一律 403
- 覆盖类写操作默认进入窗格审批（`confirmMode=ask`）；用户可在窗格切到 `auto`
- 不要伪造 `confirm:true`
- 只操作用户本机已打开、且窗格在线的文档
