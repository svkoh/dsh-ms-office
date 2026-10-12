# Integration reference (AI / scripts)

> English · [中文](AGENT.md)

For **callers**: how to write a script or an agent that drives this bridge.

> If you are writing the "make the AI use this capability" instructions instead, that is the
> [skill document](../skills/office-for-mso/SKILL.md) — the same protocol from the other side:
> the skill says **what to do**, this document says **what is actually on the wire**.

## Finding the bridge address

`http://127.0.0.1:3000` by default, but **the port is configurable and the service picks a free one when it is taken**, so never assume 3000.

The live address is in the announcement file the service maintains — **read it, do not guess**:

```
$DSH_HOME/dsh-office-for-mso/bridge.json      # ~/.dsh when $DSH_HOME is unset
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

Take `url` and append the paths below. Every `/office/...` path is relative to it.

- **File missing** → the bridge is not running. Ask the user to confirm it is "running" in the plugin settings; **do not retry in a loop**.
- If the port differs from last time, **trust the file** (the user may have changed it).

## Endpoints

| Method + path               | Purpose                                                            |
| --------------------------- | ------------------------------------------------------------------ |
| `GET /office/status`        | Online document instances (with heartbeat time) — confirm the target is online |
| `GET /office/capabilities`  | Capability table: name / hosts / destructive / arg text per action  |
| `POST /office/command`      | Send a command                                                     |
| `GET /office/errors`        | Error catalogue: meaning, handling advice, retryability             |
| `GET /office/i18n`          | Pane UI dictionary (normally only the pane uses this)              |
| `POST /office/config`       | Switch approval mode (`confirmMode`: `ask` / `auto`)               |
| `GET /actions.js`           | The concatenated action implementations (loaded by the pane)        |

All endpoints accept `?lang=zh|en` (see [Language](#language)).

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

- **Identify the document by `docUrl`**; when it is empty (unsaved document) use `docTitle`.
- If several instances look like the target, **ask the user which one** — do not guess.
- Target offline → ask the user to open the document and the pane. **Do not retry in a loop.**

### `POST /office/command`

```json
{ "action": "read_selection", "instance": "<instanceId>", "args": {} }
```

- **`instance` is required.** Omitting it with several documents open can land the command on the wrong one — this is a hard refusal, not a warning.
- Synchronous, up to about 90s (returns `timeout` beyond that).
- Success: `{ ok: true, commandId, host, instance, result }`
- Failure: `{ ok: false, code, error, result? }`

After a successful write, `result` carries `previousState` / `afterState` / `document` — **verify with it, do not trust blindly**.

## Approval protocol

Destructive actions (`destructive: true`) do **not** execute directly while `confirmMode=ask` (the default):

1. The service returns `confirm_required` together with `result.preview` (the current state and what it would become)
2. **Show the preview to the user**, then wait for the pane's verdict
3. The user clicks "Apply" → the command runs; "Reject" or 120s of inactivity → `rejected`

**`confirm:true` is attached only by the pane after the user approves; a caller must never pre-send it.** The server ignores any `confirm` sent by a caller.

To switch to automatic mode: the dropdown at the top of the pane, or `POST /office/config {"confirmMode":"auto"}`. Automatic mode still re-reads before writing and returns `previousState` / `afterState` for verification.

## Error codes

| code                                       | Meaning                                   | Handling                                           |
| ------------------------------------------ | ----------------------------------------- | -------------------------------------------------- |
| `instance_required`                        | `instance` missing                        | Look up the instanceId via status and resend       |
| `addin_offline`                            | Pane closed / offline                     | Ask the user to open the pane, **do not retry**    |
| `busy`                                     | A command is still running                | Retry shortly                                      |
| `timeout`                                  | No result within 90s                      | Check status; one retry is reasonable              |
| `bad_json` / `body_too_large` / `bad_args` | Malformed / oversized body, bad arguments | Fix and resend                                     |
| `unknown_action`                           | No such action                            | Check capabilities for the right name              |
| `unsupported_host`                         | Action does not support this application  | Switch action or application                       |
| `confirm_required`                         | User approval needed                      | Hand it to the pane; never pre-send `confirm:true` |
| `rejected`                                 | User refused / approval timed out         | Stop, **do not resend**                            |
| `not_found`                                | Locate / search target missing            | Report honestly, **do not repeat the same lookup** |
| `readonly`                                 | Document is read-only / in view mode      | Tell the user to save an editable copy             |
| `requirement` / `unsupported`              | The API is missing in this environment    | Report honestly, never fake success                |
| `execution`                                | Runtime exception                         | Report the error verbatim                          |

`GET /office/errors` returns the same catalogue (with fuller handling advice and retryability).

**"Nothing found" is not an error**: `search` / `replace_all(dryRun)` return `ok:true + count:0`; report that honestly.

## Language

`?lang=` → `Accept-Language` → `OFFICE_LANG` env var → system locale → Chinese.
A tag starting with `en` selects English, one starting with `zh` selects Chinese, anything else falls back to Chinese.

`?lang=` has the highest priority — **if you are serving an English-speaking user, pass `?lang=en` explicitly** and the error meanings, `aiAction` text and arg descriptions all come back in English. The response echoes the effective language in its `lang` field.

**Language only affects text**: the data that decides behaviour — `hosts` / `destructive` / `requirement` — is byte-identical across both languages.

## Security contract

The bridge only accepts requests from its own origin on the local machine:

- No `Access-Control-Allow-*` headers at all
- A request carrying `Origin` must have `Origin` equal to the service's own origin, otherwise `403 forbidden_origin`
- The `Host` header is validated, which blocks DNS rebinding
- Requests without an `Origin` (local curl / pwsh) pass through — those are the intended callers

There is exactly one rule for callers: **never forge `confirm:true`** — leave approval to the user.
