---
name: office-for-mso
description: Drive the Microsoft Office desktop documents the user already has open (Word / Excel / PowerPoint) through a local bridge and an Office add-in — read, write, format, restructure, charts / equations / comments. Use it for requests like "read the whole Word document", "make a chart in Excel", or "translate the selection into English".
whenToUse: The user asks to work on a Word / Excel / PowerPoint document (read, edit, format, analyse, generate charts/equations, add comments, …) and that document is already open in Microsoft Office desktop.
---

# Office for MSO (DSH ↔ Microsoft Office)

## Scope

- It only reaches Microsoft Office **desktop** documents the user **already has open** (not WPS, not Office on the web).
- Execution happens in the **add-in task pane inside the document**, so the pane must stay open; closing it yields `addin_offline`.
- **There is no "open a file" capability** — if the file is not open in Office, this skill cannot touch it.

## 1. Get the bridge address first

It defaults to `http://127.0.0.1:3000`, but **the port is configurable and the service picks another one when
it is taken** — so do not assume 3000. The live address is written by the bridge itself into a small
announcement file. **Read it; do not guess:**

```
$DSH_HOME/dsh-office-for-mso/bridge.json      # ~/.dsh when $DSH_HOME is unset
```

```json
{ "plugin": "dsh-office-for-mso", "url": "http://127.0.0.1:3000", "port": 3000, "pid": 1234 }
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
- **Never quote a fixed action count back to the user** — use what this call returns.
- Omit `lang` to let the server infer it; pass the user's language explicitly for anything the user will read.

## 4. Send the command

`POST /office/command`:

```json
{ "action": "read_selection", "instance": "<instanceId>", "args": {} }
```

- **`instance` is required.** With several documents open, omitting it can land the command on the wrong
  document — this is a hard rejection, not a warning.
- It waits synchronously, up to roughly 90s (`timeout`).
- Returns `{ ok, commandId, host, instance, result }`, or `{ ok: false, code, error, result? }` on failure.

## 5. Handle the result

- In `ask` mode an overwriting action returns `confirm_required` plus `result.preview`.
  **Show the preview to the user and wait for the pane's approval** — `confirm:true` is attached by the
  pane only after the user approves, and **you must never send it yourself**.
- `rejected`: the user declined, or approval timed out. Stop, respect the decision, **do not resend**.
- `addin_offline`: the pane is not open. Tell the user, **do not retry**.
- "Nothing found" is not an error: `search` / `replace_all(dryRun)` return `ok:true + count:0`. Report that faithfully.
- A successful write returns `previousState` / `afterState` / `document` — **verify against them rather than trusting blindly**.

Full error codes and what to do about each (including whether it is retryable): `GET /office/errors?lang=<zh|en>`.

## Security contract

- The bridge only accepts requests from its own local origin: a mismatched `Origin`, or a `Host` that does not
  match, is rejected with 403.
- Overwriting writes go through pane approval by default (`confirmMode=ask`); the user can switch the pane to auto.
- **Never forge `confirm:true`.**
- Only touch documents the user has open locally and whose pane is online.
