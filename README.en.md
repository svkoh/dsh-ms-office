# dsh-office-for-mso

> English · [中文](README.md)

Drive the Microsoft Office documents you **already have open** — Word, Excel and PowerPoint — from a [DeepSeek Harness (DSH)](https://github.com/deepseek-ai/deepseek-harness) session: reading, writing, formatting, structural edits, charts, formulas and comments.

This is a **standalone dsh plugin**. Install it and dsh gains an agent capability that reads and writes Office documents. Apart from official dsh, it depends on nothing else.

> "Office" means the **Microsoft Office desktop applications** — not WPS or other compatible suites, and not the web versions.

![Runtime data flow: dsh session → bridge → add-in task pane → Office document](docs/architecture-runtime.svg)

**This is not "have the AI generate a .docx file."** The add-in runs inside the Office process and operates on the **live, in-memory document** — the one you are editing right now. No file locking, no last-writer-wins; your edits and the AI's edits are serialised by Office itself.

## What it does

- **Works on the live document** — nothing hits disk, nothing gets locked
- **78 actions** across all three applications: read a selection, write content, format, tables / charts / formulas / comments, group and align PowerPoint shapes…
- **Precise multi-document routing** — each document gets its own instance and every command must name its target, so it cannot land on the wrong one
- **Mechanism-level guardrails on writes** — the current state is read before writing and, by default, you click "Apply" in the task pane first (see [Security](#security))
- **Bilingual UI** (Chinese / English) that follows your Office display language

The full action list is in [docs/CAPABILITIES.en.md](docs/CAPABILITIES.en.md).

## Compatibility

|             | Status                                                                                                 |
| ----------- | ------------------------------------------------------------------------------------------------------ |
| **Windows** | ✅ Fully supported (add-in auto-registration via the WEF registry)                                      |
| **macOS**   | ✅ Core works (Node + add-in); auto-registration does not apply, load the manifest manually from Office |
| **Linux**   | ❌ No Office desktop add-in host                                                                        |

- **Node.js ≥ 18** (the bridge uses only the standard library)
- **Microsoft Office desktop** (Word / Excel / PowerPoint)
- **dsh** — see `peerDependencies` in `package.json`

## Installation

```bash
dsh plugin --profile web add dsh-office-for-mso
```

**That is the whole install.** The plugin takes care of the remaining two steps itself:

1. **It adds itself to dsh's load list** — dsh loads the ordered `dsh.profile.bundles` list, and a bare dependency is not loaded; the official install path maintains that list for you.
2. **It publishes its skill to the dsh session** — which is how the AI learns that this capability exists and when to reach for it.

One exception: if you **disabled** it earlier, enable it again from the plugin panel (to avoid re-enabling things by accident, the official path will not add an already-present dependency back on its own).

> Want to inspect the manifest, or debug the install? See [Installation and loading](docs/DEVELOPMENT.en.md#installation-and-loading) in the development guide.

## First run

1. Open a Word / Excel / PowerPoint document
2. Open the add-in task pane:
   - **Windows**: the bridge registers the add-in under `HKCU\...\WEF\Developer` when it starts. **Close and reopen** your Office document and the pane appears.
   - **macOS**: File → Options → Customize Ribbon → tick "Developer" → open a document → Developer → Add-ins → manage "Developer Add-ins" → add `manifest.xml` (the same manual route works on Windows)
3. The pane shows **已连接：等待 DSH 指令** ("connected: waiting for a DSH command") — you are set
4. Ask your dsh session something like "read the whole of the Word document I have open"

The pane can be shrunk and tucked into a corner, but it **must stay open while you use it** — Office.js runs inside the pane, and closing it produces `addin_offline`.

Register or remove the add-in by hand:

```powershell
powershell -ExecutionPolicy Bypass -File server\sideload.ps1
powershell -ExecutionPolicy Bypass -File server\sideload.ps1 -Remove
```

## Security

**Every destructive operation needs your approval by default.** The current state is read first, the pane shows "now → after", and nothing runs until you click "Apply". The AI **cannot** bypass this in advance — the pane attaches `confirm:true` only after you approve, and an approval left untouched for 120s counts as a rejection.

**Read-only documents are protected.** When the document is in read-only / view mode, destructive operations are refused outright.

**The bridge only talks to its own callers.** No CORS headers at all; a request carrying `Origin` must match the service's own origin, and the `Host` header is validated to block DNS rebinding. The static surface exposes only the five files the pane needs.

## Troubleshooting

| Symptom                                     | What to do                                                                                                                                                                                                                         |
| ------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Pane shows "bridge not connected"           | Check the plugin is "running" in the settings panel                                                                                                                                                                                 |
| Commands time out                           | The add-in is offline: make sure the document is open and the pane says connected                                                                                                                                                   |
| Commands return `addin_offline` immediately | The pane is closed: open the document → Home / Developer → Add-ins → Developer Add-ins → open "DSH Office 执行器"                                                                                                                     |
| The pane never appears                      | Windows: make sure the bridge has run at least once (it does the registration), then **close and reopen** the document; also check Office is not running as administrator (loopback exemptions require a normal-privilege process)  |
| The pane cannot connect after changing port | Close and reopen your Office documents (the add-in URL lives in the manifest and has to be reloaded)                                                                                                                                |
| Edits land in the wrong place               | Check the cursor / selection is where you expect; watch the execution log at the bottom of the pane                                                                                                                                 |

## Documentation

| Document | Contents |
|---|---|
| [docs/CAPABILITIES.en.md](docs/CAPABILITIES.en.md) | Full action list and capability boundaries |
| [docs/DEVELOPMENT.en.md](docs/DEVELOPMENT.en.md) | Architecture, file map, fragment contract, i18n, hot reload, custom actions, packaging rules |
| [docs/AGENT.en.md](docs/AGENT.en.md) | HTTP endpoints, error codes, approval protocol (for scripts and other agents) |

## License

[MIT](LICENSE). The license and provenance of the one third-party component (pako) are in [THIRD\_PARTY\_NOTICES.md](THIRD_PARTY_NOTICES.md).

---

*Not an official Microsoft product. This project is not affiliated with Microsoft Corporation.*

---

This project is a fork of <https://github.com/Mikuzjc/dsh-office-for-mso>.
