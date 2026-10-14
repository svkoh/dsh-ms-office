# 集成方文档（AI / 脚本）

> [English](AGENT.en.md) · 中文

给调用方看的 HTTP 协议。AI 的操作约定见[技能文档](../skills/ms-office/SKILL.md)。

## 先拿到桥接地址

默认 `http://127.0.0.1:3000`，但端口可配置，被占用时服务会另选端口。读取服务维护的公告文件，不要猜地址：

```text
$DSH_HOME/dsh-ms-office/bridge.json      # $DSH_HOME 未设置时是 ~/.dsh
```

```json
{ "plugin": "dsh-ms-office", "url": "http://127.0.0.1:3000", "port": 3000, "pid": 1234, "startedAt": "…" }
```

取 `url` 字段拼下文路径。文件不存在时，确认插件设置中的桥接服务状态；不要反复重试。设置面板区分 `starting`、`healthy`、`unhealthy`、`stopped`，进程存在不等于服务健康。端口变化时以公告文件为准。

## 端点

| 方法 + 路径 | 作用 |
|---|---|
| `GET /office/status` | 在线文档实例、心跳与未结束命令 |
| `GET /office/capabilities` | action 名称、平台、是否覆盖类、参数说明 |
| `POST /office/command` | 发指令并等待最终结果 |
| `GET /office/commands/:id` | 查询命令状态、到期时间和已记录结果 |
| `POST /office/commands/:id/start` | 仅窗格使用：原子领取执行权 |
| `POST /office/result` | 仅窗格使用：提交带实例标识的结果 |
| `GET /office/errors` | 错误码目录、处理建议和是否可重试 |
| `GET /office/i18n` | 窗格 UI 字典 |
| `POST /office/config` | 切审批模式（`confirmMode`: `ask` / `auto`） |
| `GET /actions.js` | 拼装后的 action 实现，供窗格加载 |

文案端点接受 `?lang=zh|en`，见[语言](#语言)。

### 确认目标实例

`GET /office/status` 返回 `instances: [{ instanceId, host, docUrl, docTitle, online, lastBeat }]`。

- 按 `docUrl` 识别文档；未保存文档的 URL 为空时用 `docTitle`。
- 多个实例都像目标时，先问用户是哪份，不要猜。
- 目标不在线时提醒用户打开文档和窗格，不要反复重试。
- `pendingCommands` 列出未结束命令；旧字段 `pending` 保留兼容。不同实例可并行，同一实例最多一条未结束指令。

### 发指令

`POST /office/command`：

```json
{
  "action": "read_selection",
  "instance": "<instanceId>",
  "args": {},
  "idempotencyKey": "<optional-key-for-this-operation>"
}
```

- `instance` 必填，缺失会被拒绝。
- 请求等待执行、拒绝或超时结果；默认 `COMMAND_TIMEOUT_MS=180000`（180 秒）。
- 可选顶层 `idempotencyKey` 为不超过 128 字符的字符串。同一实例、同一 key、相同参数复用原命令的等待或结果；同 key 但参数不同返回 HTTP 409。一次逻辑操作的网络重试应保留原 key。
- 成功：`{ ok: true, commandId, host, instance, result }`；失败：`{ ok: false, code, error, result? }`。
- 请求体必须是 UTF-8 字节。PowerShell 可用 `[Text.Encoding]::UTF8.GetBytes($json)`，配合 `Content-Type: application/json; charset=utf-8`。非法 UTF-8 返回 `bad_encoding`；开头的 UTF-8 BOM 会被剥掉。

写入成功后，按实际返回的 `previousState` / `afterState` / `document` 核验。字段取决于动作、宿主及读取是否成功；Excel 回读最多实际加载 5000 格，不能假定每次都有完整全文。写后核验失败以 `result.verificationWarnings` 提示，写入本身仍可为 `ok:true`。

### 查询、超时与幂等

`GET /office/commands/:id` 返回顶层 `commandId`、`status`、`expiresAt` 和 `response`。`response` 是已记录的结果，也可能是超时结果，不能仅凭它判断 Office 是否完成写入。

| status | 含义 |
|---|---|
| `pending` | 排队或等待窗格审批，尚未领取执行权 |
| `executing` | 窗格已原子领取，正在执行 |
| `completed` | 已收到窗格结果，以结果的 `ok` 判断成功或失败 |
| `expired` | 执行前到期，释放该实例占位 |
| `execution_unknown` | 领取后超时，执行结果尚不明确 |

`execution_unknown` 会继续占用实例，直到迟到的窗格结果到达。**不要盲目重发写入或换 key 重试**；先查原命令并核对文档。历史最多 1000 条，完成后保留最多一小时；历史及幂等信息只在内存中，服务重启不保留。

## 审批协议

覆盖类操作（`destructive: true`）在默认 `confirmMode=ask` 下由窗格审批：

1. 窗格读取当前状态并展示「当前 → 修改后」；预览失败时禁止执行。
2. 调用方保持等待，用户在窗格批准或拒绝，不需要重发命令。
3. 审批最长 120 秒，且不得超过命令 `expiresAt`；拒绝或审批超时返回 `rejected`。
4. 批准后窗格先原子领取执行权，再通过绑定 `commandId` 的一次性内部闭包执行。

**外部 `args.confirm` 会被忽略。** 窗格内部批准不使用 `confirm` 参数，JSON 参数不能代替用户审批。普通分发入口可能返回 `confirm_required`，这不是要求调用方带 `confirm:true` 重发。

所有窗格执行路径（含只读与自动模式）均先调用 `POST /office/commands/:id/start`，请求体为 `{ "instance": "<instanceId>" }`。服务端仅在 `pending` 且未过期时原子设为 `executing`；重复领取或过期返回 HTTP 409。窗格向 `POST /office/result` 提交 `{ commandId, instance, ok, result?, error?, code? }`，其中 `ok` 必须是布尔值。调用方不应调用这两个窗格端点。

用户可在窗格切换自动模式，或设置 `POST /office/config {"confirmMode":"auto"}`。自动模式仍先预览，再执行；按实际返回的状态字段核验。

## 错误码

| code | 含义 | 处理 |
|---|---|---|
| `instance_required` | 缺 `instance` | 查 status 获取目标 instanceId |
| `addin_offline` | 窗格未开启或离线 | 提醒用户开窗格，不要重试 |
| `busy` | 同实例已有未结束指令 | 查询原命令；结果未知时不要另发写入 |
| `timeout` | 命令期限内未收到结果 | 查命令状态；`execution_unknown` 不得盲重发 |
| `bad_json` / `body_too_large` / `bad_args` | 请求体或参数非法 | 修正后重发 |
| `bad_encoding` | 请求体不是合法 UTF-8 | 按 UTF-8 重发，别用本地代码页 |
| `unknown_action` | action 不存在 | 查询 capabilities |
| `unsupported_host` | 当前应用不支持该动作 | 换动作或应用 |
| `confirm_required` | 需要窗格审批 | 交给窗格，勿带批准参数重发 |
| `rejected` | 用户拒绝或审批超时 | 停止，不要重发 |
| `not_found` | 定位或查找目标不存在 | 如实告知，不重复同一查找 |
| `readonly` | 文档只读或查看模式 | 提醒用户另存可编辑副本 |
| `requirement` / `unsupported` | 当前环境缺该 API | 如实告知 |
| `execution` | 执行期异常 | 报告 error；结合命令状态判断结果 |

完整目录以 `GET /office/errors` 为准。查找或 dry-run 无命中可返回 `ok:true` 和零命中数，应如实转述。

## Excel 参数与读取限制

`apply_filter.column` 相对目标区域从 **1** 开始。值筛选可用 `criteria: {"values":["a"]}` 或 Values 数组简写 `criteria: ["a"]`；自定义筛选用 Custom 条件对象，如 `criteria: {"criterion1":">2"}`，具体筛选方式字段以 capabilities 为准。

`read_range` 和 Excel `read_document` 最多实际加载 5000 格。处理大表时分区域读取，并检查返回的截断信息，勿把回读当成完整工作表。

## 语言

`?lang=` → `Accept-Language` → `OFFICE_LANG` → 系统 locale → 中文。`en*` 选英文，`zh*` 选中文，其余回退中文。面向英文用户显式传 `?lang=en`；文案响应以 `lang` 回显实际语言。语言只影响文案，`hosts` / `destructive` / `requirement` 不随语言变化。

## 安全约定

桥接服务不返回 `Access-Control-Allow-*` 头。带 `Origin` 的请求必须匹配服务自身来源，否则返回 `403 forbidden_origin`；同时校验 `Host` 防止 DNS 重绑定。无 `Origin` 的本机脚本请求可正常使用。

把审批交给窗格中的用户；超时后先查原命令，避免重复修改文档。
