---
name: ms-office
description: Drive the Microsoft Office desktop documents the user already has open (Word / Excel / PowerPoint) through a local bridge and an Office add-in — read, write, format, restructure, charts / equations / comments. Use it for requests like "read the whole Word document", "make a chart in Excel", or "translate the selection into English".
whenToUse: The user asks to work on a Word / Excel / PowerPoint document (read, edit, format, analyse, generate charts/equations, add comments, …) and that document is already open in Microsoft Office desktop.
---

# MS Office (DSH ↔ Microsoft Office)

## Scope

- It only reaches Microsoft Office **desktop** documents the user **already has open** (not WPS, not Office on the web).
- Execution happens in the **add-in task pane inside the document**, so the pane must stay open; closing it yields `addin_offline`.
- **There is no "open a file" capability** — if the file is not open in Office, this skill cannot touch it.

## 1. Get the bridge address first

It defaults to `http://127.0.0.1:3000`, but **the port is configurable and the service picks another one when
it is taken** — so do not assume 3000. The live address is written by the bridge itself into a small
announcement file. **Read it; do not guess:**

```
$DSH_HOME/dsh-ms-office/bridge.json      # ~/.dsh when $DSH_HOME is unset
```

```json
{ "plugin": "dsh-ms-office", "url": "http://127.0.0.1:3000", "port": 3000, "pid": 1234 }
```

Take the `url` field and append the paths below to it. Every `/office/...` path is relative to that address.

- **File missing** → the bridge is not running. Ask the user to confirm it reads "running" in the plugin
  settings, and **do not retry in a loop**.
- If the port differs from what you used before, **trust the file** (the user may have changed it).

## 2. Confirm the target document is online

`GET /office/status` → `instances: [{ instanceId, host, docUrl, docTitle, online, lastBeat }]`

- **Identify the document by `docUrl`**; when it is empty (unsaved document) fall back to `docTitle`.
- If several instances could be the target, **ask the user which one** — do not guess.
- Target offline → ask the user to open the document and enable the "DSH Office Runner" pane under
  **Home / Developer → Add-ins → Developer Add-ins**. **Do not retry in a loop.**
- If the user just changed the port, the manifest URL has followed, but Office only reads that manifest when a
  document loads — ask the user to **close and reopen** the document.

## 3. Look up capabilities

`GET /office/capabilities?lang=<zh|en>` → `actions: [{ name, hosts, destructive, requirement, args }]`

- `hosts` = the applications that can run the action; `destructive: true` = an overwriting action (see step 5).
- `args` documents each parameter, including its type and whether it is required.
- Enum values are case-insensitive (`align:"left"` passes as well); when one is rejected the error lists the
  valid values — fix it from the list instead of guessing.
- Field names follow `args`. `kind`→`type`, `width`/`height`→`w`/`h`, `fontSize`→`size` are accepted aliases;
  colors accept `#RRGGBB` only.
- `shapeType` values are Office's, not ours: the plugin does not guess the enum, so any PascalCase name is
  accepted, and when execution fails the error appends the list of values verified against the add-in — fix
  it from that list instead of guessing.
- **Never quote a fixed action count back to the user** — use what this call returns.
- Omit `lang` to let the server infer it; pass the user's language explicitly for anything the user will read.

## 4. Send the command

`POST /office/command`:

```json
{ "action": "read_selection", "instance": "<instanceId>", "args": {} }
```

- **`instance` is required.** With several documents open, omitting it can land the command on the wrong
  document — this is a hard rejection, not a warning.
- It waits for the final result, up to 180s by default (`timeout`); set the calling script's HTTP timeout slightly longer.
- Commands are serialized per instance; different instances execute independently. For writes, supply a top-level `idempotencyKey` (non-empty string, at most 128 characters). The same instance, key and arguments reuse the original command or result; changed arguments are rejected.
- Returns `{ ok, commandId, host, instance, result }`, or `{ ok: false, code, error, result? }` on failure.
- **Send the body as UTF-8 bytes.** When a script writes non-ASCII text, encode the JSON as UTF-8 first
  (PowerShell: `[Text.Encoding]::UTF8.GetBytes($json)` with `Content-Type: application/json; charset=utf-8`).
  Handing `ConvertTo-Json` straight to the request encodes it in the local codepage — a document title once
  came out as `??????`. The bridge now rejects such a body with `bad_encoding` instead of silently
  writing mojibake into the user's document.

## 5. Handle the result

- In `ask` mode the pane shows a preview and waits for user approval; keep waiting on the original request rather than resending. Approval lasts at most 120 seconds and cannot exceed `expiresAt`; a failed preview prevents execution.
- External `args.confirm` is ignored. Internal approval invokes a single-use closure bound to `commandId`, without a `confirm` argument. Do not resend approval arguments when seeing `confirm_required`.
- `timeout`: query `GET /office/commands/:id` first. It returns top-level `commandId`, `status`, `expiresAt` and `response`; states are `pending` / `executing` / `completed` / `expired` / `execution_unknown`. `response` may be a timeout response and does not prove a write never happened.
- `execution_unknown` keeps the instance occupied until a late result arrives; do not blindly resend writes. Query the original command and inspect the document. History holds at most 1000 commands, completed entries last at most one hour, and records do not survive a restart. `/office/status` exposes `pendingCommands` and retains legacy `pending`.
- `rejected`: the user declined, or approval timed out. Stop, respect the decision, **do not resend**.
- `addin_offline`: the pane is not open. Tell the user, **do not retry**.
- `bad_encoding`: the body was not valid UTF-8 (usually a script sent non-ASCII text in the local codepage).
  Resend it as UTF-8 — **do not resend the same bytes**.
- "Nothing found" is not an error: `search` / `replace_all(dryRun)` return `ok:true + count:0`. Report that faithfully.
- Verify against the `previousState` / `afterState` / `document` fields actually returned; fields may be absent or truncated. Do not assume complete document readback on every write. `verificationWarnings` means verification was incomplete; `ok:true` still reports a successful write.
- Excel `read_range` / `read_document` load at most 5000 cells; read large sheets in separate regions. `apply_filter.column` is 1-based. Values filters accept `criteria:{values:["a"]}` or shorthand `criteria:["a"]`; Custom uses condition objects such as `criteria:{criterion1:">2"}`. Check capabilities for the filter-mode argument.
- `POST /office/commands/:id/start` and `/office/result` are pane-only; callers must not claim execution or submit results themselves.

Full error codes and what to do about each (including whether it is retryable): `GET /office/errors?lang=<zh|en>`.

## Security contract

- The bridge only accepts requests from its own local origin: a mismatched `Origin`, or a `Host` that does not
  match, is rejected with 403.
- Overwriting writes go through pane approval by default (`confirmMode=ask`); the user can switch the pane to auto.
- Leave approval to the user in the pane; external approval parameters grant no execution rights.
- Only touch documents the user has open locally and whose pane is online.
