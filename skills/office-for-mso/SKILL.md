---
name: office-for-mso
description: 通过本机桥接服务和 Office 加载项，操控用户当前打开的 Microsoft Office 桌面版文档（Word / Excel / PowerPoint）——读取、写入、格式化、结构操作、图表 / 公式 / 批注。用于「读一下 Word 全文」「Excel 生成图表」「把选中的翻译成英文」这类请求。
whenToUse: 用户要求操作 Word / Excel / PowerPoint 文档（读取、编辑、格式化、分析、生成图表/公式、批注等），且这些文档是用户已经在 Microsoft Office 桌面版里打开的。
---

# Office for MSO（DSH ↔ Microsoft Office）

## 适用边界

- 只对**用户当前已打开**的 Microsoft Office **桌面版**文档生效（不含 WPS，也不含 Office 网页版）。
- 执行发生在**文档里的加载项窗格**中，所以窗格必须开着；关掉就会得到 `addin_offline`。
- **没有「打开某个文件」的能力**——文件不在 Office 里打开，这个技能就碰不到它。

## 1. 先拿到桥接地址

默认是 `http://127.0.0.1:3000`，但**端口是可配置的，被占用时服务会自己换一个**，所以不要假定 3000。
实际地址写在桥接服务自己维护的公告文件里——**读它，不要猜**：

```
$DSH_HOME/dsh-office-for-mso/bridge.json      # $DSH_HOME 未设置时是 ~/.dsh
```

```json
{ "plugin": "dsh-office-for-mso", "url": "http://127.0.0.1:3000", "port": 3000, "pid": 1234 }
```

取 `url` 字段拼后面的路径。下面出现的 `/office/...` 都相对这个地址。

- **文件不存在** → 桥接服务没在跑。让用户在插件设置里确认它是「运行中」，**不要反复重试**。
- 文件里的端口和你上次用的不一样，**以文件为准**（用户可能改过）。

## 2. 确认目标文档在线

`GET /office/status` → `instances: [{ instanceId, host, docUrl, docTitle, online, lastBeat }]`

- **按 `docUrl` 识别是哪一份文档**；`docUrl` 为空（文档未保存）时用 `docTitle`。
- 有多个实例都像目标时，**先问用户是哪一份**，不要猜。
- 目标不在线 → 提醒用户打开文档，并在 **开始 / 开发人员 → 加载项 → 开发人员加载项** 里
  打开「DSH Office 执行器」窗格。**不要反复重试。**
- 如果用户刚改过端口，清单里的地址虽然已跟着变，但 Office 只在加载文档时读一次清单——
  提醒用户**关闭并重新打开**该文档。

## 3. 查能力

`GET /office/capabilities?lang=<zh|en>` → `actions: [{ name, hosts, destructive, requirement, args }]`

- `hosts` = 能跑这个 action 的应用；`destructive: true` = 覆盖类操作（见第 5 步）。
- `args` 是每个参数的说明（含类型与是否必填）。
- 枚举值大小写不敏感（`align:"left"` 也能通过）；被拒时错误里会列出合法值，照清单改，不要猜。
- 字段名以 `args` 为准。`kind`→`type`、`width`/`height`→`w`/`h`、`fontSize`→`size` 是兼容别名；颜色只接受 `#RRGGBB`。
- `shapeType` 的合法值由 Office 决定，插件不猜枚举：任何 PascalCase 名字都会放行，执行失败时错误里会补一份「已验证可用」的清单——照清单改，别猜。
- **不要在回复里写死 action 数量**，以这次返回为准。
- `lang` 省略时按服务端环境推断；给用户看的内容请显式传用户的语言。

## 4. 发指令

`POST /office/command`：

```json
{ "action": "read_selection", "instance": "<instanceId>", "args": {} }
```

- **`instance` 必填。** 多文档同时打开时省略它，指令可能落到错误的文档上——这是硬性拒绝，不是警告。
- 同步等待，最长约 90s（超时返回 `timeout`）。
- 返回 `{ ok, commandId, host, instance, result }`，失败时是 `{ ok: false, code, error, result? }`。
- **请求体按 UTF-8 字节发。** 脚本里发中文，先把 JSON 编码成 UTF-8 再发（PowerShell：
  `[Text.Encoding]::UTF8.GetBytes($json)` 配 `Content-Type: application/json; charset=utf-8`）。
  直接 `ConvertTo-Json` 交出去会按本地代码页编码——中文标题曾被写成 `??????`。
  服务端现在会明确回 `bad_encoding`，不会静默把乱码写进文档。

## 5. 处理结果

- 覆盖类操作在 `ask` 模式下返回 `confirm_required`，并带上 `result.preview`。
  **把 preview 展示给用户，然后等窗格里的审批结果**——`confirm:true` 只能由窗格在用户批准后附加，
  **你不得预先发送它**。
- `rejected`：用户拒绝或审批超时。停下来，尊重决定，**不要重发**。
- `addin_offline`：窗格没开。提醒用户，**不要重试**。
- `bad_encoding`：请求体不是合法 UTF-8（多半是脚本按本地代码页发了中文）。改成 UTF-8 重发，**不要重发同一份字节**。
- 「没找到」不是错误：`search` / `replace_all(dryRun)` 无命中返回 `ok:true + count:0`，如实转述。
- 写入成功后结果里带 `previousState` / `afterState` / `document`——**用它核验，不要盲信**。

完整错误码与处理建议（含每个码是否可重试）：`GET /office/errors?lang=<zh|en>`。

## 安全约定

- 桥接服务只接受本机自身来源的请求：`Origin` 不是它自己的、或 `Host` 不匹配的一律 403。
- 覆盖类写操作默认走窗格审批（`confirmMode=ask`）；用户可在窗格切到自动模式。
- **不要伪造 `confirm:true`。**
- 只操作用户本机已打开、且窗格在线的文档。
