# 集成方文档（AI / 脚本）

> [English](AGENT.en.md) · 中文

给**调用方**看的：怎么写脚本或 agent 去驱动这台桥接服务。

> 如果你是在写「让 AI 自己会用这个能力」的说明，看的是[技能文档](../skills/office-for-mso/SKILL.md)——
> 它和本文是同一套协议的两种视角：技能讲**该怎么做**，本文讲**线上到底有什么**。

## 先拿到桥接地址

默认 `http://127.0.0.1:3000`，但**端口可配置，被占用时服务会自己换一个**，所以不要假定 3000。

实际地址写在服务自己维护的公告文件里——**读它，不要猜**：

```
$DSH_HOME/dsh-office-for-mso/bridge.json      # $DSH_HOME 未设置时是 ~/.dsh
```

```json
{
  "plugin": "dsh-office-for-mso",
  "url": "http://127.0.0.1:3000",
  "port": 3000,
  "pid": 1234,
  "startedAt": "2026-10-12T12:00:00.000Z"
}
```

取 `url` 字段拼后面的路径。下文所有 `/office/...` 都相对这个地址。

- **文件不存在** → 桥接服务没在跑。让用户在插件设置里确认它是「运行中」，**不要反复重试**。
- 文件里的端口和你上次用的不一样，**以文件为准**（用户可能改过端口）。

## 端点

| 方法 + 路径                 | 作用                                    |
| ----------------------- | ------------------------------------- |
| `GET /office/status`    | 在线文档实例列表（含心跳时间），用来确认目标在线              |
| `GET /office/capabilities` | 能力表：每个 action 的名称 / 平台 / 是否覆盖类 / 参数说明 |
| `POST /office/command`  | 发一条指令                                 |
| `GET /office/errors`    | 错误码目录：每个码的含义、处理建议、是否可重试               |
| `GET /office/i18n`      | 窗格 UI 字典（一般只有窗格自己用）                   |
| `POST /office/config`   | 切审批模式（`confirmMode`: `ask` / `auto`）   |
| `GET /actions.js`       | 拼装后的 action 实现（窗格加载用）                 |

所有端点都接受 `?lang=zh|en`（见[语言](#语言)）。

### `GET /office/status`

```json
{
  "instances": [
    {
      "instanceId": "…",
      "host": "Word",
      "docUrl": "file:///C:/…/report.docx",
      "docTitle": "report.docx",
      "online": true,
      "lastBeat": 1760000000000
    }
  ]
}
```

- **按 `docUrl` 识别是哪一份文档**；`docUrl` 为空（文档尚未保存）时用 `docTitle`。
- 有多个实例都像目标时，**先问用户是哪一份**，不要猜。
- 目标不在线 → 提醒用户打开文档和窗格。**不要反复重试。**

### `POST /office/command`

```json
{ "action": "read_selection", "instance": "<instanceId>", "args": {} }
```

- **`instance` 必填。** 多文档同时打开时省略它，指令可能落到错误的文档上——这是硬性拒绝，不是警告。
- 同步等待，最长约 90s（超时返回 `timeout`）。
- 成功：`{ ok: true, commandId, host, instance, result }`
- 失败：`{ ok: false, code, error, result? }`
- **请求体必须是 UTF-8 字节。** 脚本发中文时先把 JSON 按 UTF-8 编码再发出去（PowerShell：
  `[Text.Encoding]::UTF8.GetBytes($json)` 配 `Content-Type: application/json; charset=utf-8`）。
  按本地代码页发出去会被明确拒绝（`bad_encoding`），不会静默替换成乱码写进文档——静默写坏用户文档比报错更糟。
  开头的 UTF-8 BOM 会被剥掉；空 body 不算编码错误。

写入类操作成功后，`result` 里带 `previousState` / `afterState` / `document`——**用它核验，不要盲信**。

## 审批协议

覆盖类操作（`destructive: true`）在 `confirmMode=ask`（默认）下**不会直接执行**：

1. 服务返回 `confirm_required`，并带上 `result.preview`（当前状态与将要变成什么）
2. **把 preview 展示给用户**，然后等窗格里的审批结果
3. 用户在窗格点「更改」→ 指令才真正执行；点「拒绝」或 120 秒无操作 → `rejected`

**`confirm:true` 只能由窗格在用户批准后附加，调用方不得预先发送它。** 服务端会忽略调用方传的 `confirm`。

切到自动模式：窗格顶部下拉，或 `POST /office/config {"confirmMode":"auto"}`。自动模式仍会写前 re-read，结果里带 `previousState` / `afterState` 供核验。

## 错误码

| code                                       | 含义                | 处理                          |
| ------------------------------------------ | ----------------- | --------------------------- |
| `instance_required`                        | 缺 `instance`      | 查 status 拿 instanceId 重发    |
| `addin_offline`                            | 窗格未开启 / 离线        | 提醒用户开窗格，**不要重试**            |
| `busy`                                     | 上一条指令还在跑          | 稍等重发                        |
| `timeout`                                  | 90s 无结果           | 查 status；在线可重试一次            |
| `bad_json` / `body_too_large` / `bad_args` | 请求体非法 / 太大 / 参数非法 | 修正后重发                       |
| `bad_encoding`                             | 请求体不是合法 UTF-8        | 按 UTF-8 重发，别用本地代码页        |
| `unknown_action`                           | action 名不存在       | 查 capabilities 用正确名称        |
| `unsupported_host`                         | 该 action 不支持当前应用  | 换 action 或换应用               |
| `confirm_required`                         | 需要用户审批            | 交给窗格审批，不得预先发 `confirm:true` |
| `rejected`                                 | 用户拒绝 / 审批超时       | 停止，**不要重发**                 |
| `not_found`                                | 定位 / 查找目标不存在      | 如实告知，**不要重试同一查找**           |
| `readonly`                                 | 文档只读 / 查看模式       | 告知用户另存为可编辑副本                |
| `requirement` / `unsupported`              | 当前环境缺这个 API       | 如实告知，绝不假装成功                 |
| `execution`                                | 执行期异常             | 把 error 原样报告给用户             |

`GET /office/errors` 返回同一份目录（含更详细的处理建议与是否可重试）。

**「没找到」不是错误**：`search` / `replace_all(dryRun)` 无命中返回 `ok:true + count:0`，如实转述即可。

## 语言

`?lang=` → `Accept-Language` → 环境变量 `OFFICE_LANG` → 系统 locale → 中文。
以 `en` 开头取英文，以 `zh` 开头取中文，其余一律中文。

`?lang=` 的优先级最高，**如果你正在服务一个英文用户，显式传 `?lang=en`**，返回的错误含义、`aiAction` 与参数说明就都是英文。响应里会带 `lang` 字段回显实际生效的语言。

**语言只影响文案**：`hosts` / `destructive` / `requirement` 这些决定行为的数据在两种语言下逐字节相同。

## 安全约定

桥接服务只接受本机自身来源的请求：

- 不返回任何 `Access-Control-Allow-*` 头
- 带 `Origin` 的请求，其 `Origin` 必须等于服务自身来源，否则 `403 forbidden_origin`
- 校验 `Host` 头，挡 DNS 重绑定
- 没有 `Origin` 的调用（本机 curl / pwsh）照常放行——那是设计内的使用者

调用方只需要遵守一条：**不要伪造 `confirm:true`**，把审批交给用户。
