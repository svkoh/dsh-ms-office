# Integration reference (AI / scripts)

> English · [中文](AGENT.md)

The HTTP protocol for callers. For instructions an AI should follow, see the [skill document](../skills/ms-office/SKILL.en.md).

## Finding the bridge address

The default is `http://127.0.0.1:3000`, but the port is configurable and the service chooses another when it is occupied. Read the announcement file maintained by the service rather than guessing:

```text
$DSH_HOME/dsh-ms-office/bridge.json      # ~/.dsh when $DSH_HOME is unset
```

```json
{ "plugin": "dsh-ms-office", "url": "http://127.0.0.1:3000", "port": 3000, "pid": 1234, "startedAt": "…" }
```

Append the paths below to `url`. If the file is missing, check the bridge state in plugin settings; do not retry in a loop. The panel distinguishes `starting`, `healthy`, `unhealthy` and `stopped`; a process being present does not prove the service is healthy. Trust the announcement when the port changes.

## Endpoints

| Method + path | Purpose |
|---|---|
| `GET /office/status` | Online document instances, heartbeats and unfinished commands |
| `GET /office/capabilities` | Action names, hosts, destructive flags and argument descriptions |
| `POST /office/command` | Submit a command and wait for its final response |
| `GET /office/commands/:id` | Query command state, expiry and recorded response |
| `POST /office/commands/:id/start` | Pane only: atomically claim execution |
| `POST /office/result` | Pane only: submit a result with its instance |
| `GET /office/errors` | Error catalogue, handling advice and retryability |
| `GET /office/i18n` | Pane UI dictionary |
| `POST /office/config` | Switch approval mode (`confirmMode`: `ask` / `auto`) |
| `GET /actions.js` | Concatenated action implementations loaded by the pane |

Text endpoints accept `?lang=zh|en`; see [Language](#language).

### Confirming the target instance

`GET /office/status` returns `instances: [{ instanceId, host, docUrl, docTitle, online, lastBeat }]`.

- Identify the document by `docUrl`; use `docTitle` when an unsaved document has no URL.
- If several instances match, ask the user which one rather than guessing.
- If the target is offline, ask the user to open the document and pane; do not retry in a loop.
- `pendingCommands` lists unfinished commands; the legacy `pending` field remains for compatibility. Different instances can run concurrently, with at most one unfinished command per instance.

### Sending a command

`POST /office/command`:

```json
{
  "action": "read_selection",
  "instance": "<instanceId>",
  "args": {},
  "idempotencyKey": "<optional-key-for-this-operation>"
}
```

- `instance` is required; omitting it is rejected.
- The request waits for execution, rejection or timeout; the default is `COMMAND_TIMEOUT_MS=180000` (180 seconds).
- Optional top-level `idempotencyKey` is a string of at most 128 characters. The same instance, key and parameters reuse the original wait or response; the same key with different parameters returns HTTP 409. Keep the original key when retrying a network request for one logical operation.
- Success: `{ ok: true, commandId, host, instance, result }`; failure: `{ ok: false, code, error, result? }`.
- The body must be UTF-8 bytes. PowerShell can use `[Text.Encoding]::UTF8.GetBytes($json)` with `Content-Type: application/json; charset=utf-8`. Invalid UTF-8 returns `bad_encoding`; a leading UTF-8 BOM is stripped.

Verify a write against whichever `previousState` / `afterState` / `document` fields are actually returned. Availability depends on the action, host and successful reads; Excel readback loads at most 5000 cells, so do not assume a complete document is always returned. A post-write verification failure appears in `result.verificationWarnings`; the write itself can still be `ok:true`.

### Querying state, timeouts and idempotency

`GET /office/commands/:id` returns top-level `commandId`, `status`, `expiresAt` and `response`. `response` is the recorded result and may be a timeout response; it alone does not prove whether Office finished a write.

| status | Meaning |
|---|---|
| `pending` | Queued or awaiting pane approval; execution has not been claimed |
| `executing` | The pane has atomically claimed execution |
| `completed` | A pane result arrived; inspect its `ok` for success or failure |
| `expired` | Expired before execution; releases the instance slot |
| `execution_unknown` | Timed out after execution was claimed; outcome is uncertain |

`execution_unknown` keeps the instance slot until a late pane result arrives. **Do not blindly resend a write or retry with a new key**; query the original command and inspect the document first. History holds at most 1000 commands, with completed entries retained for at most one hour. History and idempotency records are in memory and do not survive a service restart.

## Approval protocol

Destructive actions (`destructive: true`) use pane approval in the default `confirmMode=ask`:

1. The pane reads the current state and shows "now → after"; a failed preview prevents execution.
2. The caller keeps waiting while the user approves or rejects in the pane; no command resend is needed.
3. Approval lasts at most 120 seconds and cannot extend beyond `expiresAt`; rejection or approval timeout returns `rejected`.
4. After approval, the pane atomically claims execution and invokes a single-use internal closure bound to `commandId`.

**External `args.confirm` is ignored.** Internal pane approval does not use a `confirm` argument; JSON parameters cannot replace user approval. The ordinary dispatcher may return `confirm_required`, but this is not an instruction to resend with `confirm:true`.

Every pane execution path, including reads and automatic mode, first calls `POST /office/commands/:id/start` with `{ "instance": "<instanceId>" }`. The server atomically changes an unexpired `pending` command to `executing`; duplicate or expired claims return HTTP 409. The pane submits `{ commandId, instance, ok, result?, error?, code? }` to `POST /office/result`, where `ok` must be a boolean. Callers should not invoke these pane endpoints.

The user can switch to automatic mode in the pane, or set `POST /office/config {"confirmMode":"auto"}`. Automatic mode still previews before executing; verify against the state fields actually returned.

## Error codes

| code | Meaning | Handling |
|---|---|---|
| `instance_required` | Missing `instance` | Look up the target instanceId via status |
| `addin_offline` | Pane closed or offline | Ask the user to open it; do not retry |
| `busy` | The instance has an unfinished command | Query that command; do not send another write while its outcome is unknown |
| `timeout` | No result before the command deadline | Query command state; do not blindly retry `execution_unknown` |
| `bad_json` / `body_too_large` / `bad_args` | Invalid body or arguments | Fix and resend |
| `bad_encoding` | Invalid UTF-8 body | Resend as UTF-8 rather than the local codepage |
| `unknown_action` | Action does not exist | Query capabilities |
| `unsupported_host` | The application does not support the action | Switch action or application |
| `confirm_required` | Pane approval needed | Leave it to the pane; do not resend approval arguments |
| `rejected` | User declined or approval timed out | Stop; do not resend |
| `not_found` | Locate or search target missing | Report honestly; do not repeat the same lookup |
| `readonly` | Document is read-only or in view mode | Ask the user to save an editable copy |
| `requirement` / `unsupported` | An API is missing in this environment | Report honestly |
| `execution` | Runtime exception | Report the error and check command state |

The complete catalogue is available from `GET /office/errors`. A search or dry run with no hits can return `ok:true` and a zero hit count; report it faithfully.

## Excel arguments and read limits

`apply_filter.column` is **1-based** within the target range. Value filters accept `criteria: {"values":["a"]}` or the Values array shorthand `criteria: ["a"]`; Custom filters use a condition object such as `criteria: {"criterion1":">2"}`. Check capabilities for the filter-mode argument.

`read_range` and Excel `read_document` load at most 5000 cells. Read large sheets in separate regions and check returned truncation information; do not treat a bounded readback as the complete sheet.

## Language

`?lang=` → `Accept-Language` → `OFFICE_LANG` → system locale → Chinese. `en*` selects English, `zh*` selects Chinese, other tags fall back to Chinese. Pass `?lang=en` for an English-speaking user; text responses echo the effective `lang`. Language affects text only: `hosts` / `destructive` / `requirement` remain unchanged.

## Security contract

The bridge returns no `Access-Control-Allow-*` headers. A request carrying `Origin` must match the service's own origin, otherwise it returns `403 forbidden_origin`; `Host` is also validated to block DNS rebinding. Local script requests without an `Origin` can use the bridge normally.

Leave approval to the user in the pane. Query the original command after a timeout to avoid duplicate edits.
