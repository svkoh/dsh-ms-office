# dsh-office-for-mso

> English · [中文](README.md)

Drive the Microsoft Office documents you **already have open** — Word, Excel and PowerPoint — from a [DeepSeek Harness (DSH)](https://github.com/deepseek-ai/deepseek-harness) session: reading, writing, formatting, structural edits, charts, formulas and comments.

This is a **standalone dsh plugin**. Install it and dsh gains an agent capability that reads and writes Office documents. Apart from official dsh, it depends on nothing else.

> "Office" means the **Microsoft Office desktop applications**. WPS and other compatible suites are not supported.

```
you ──▶ dsh session (AI) ──POST /office/command──▶ bridge (local HTTP, 127.0.0.1)
                                                    │ command queued (routed by host / instance)
                            Office add-in task pane (polls the bridge, 1s)
                                                    │ executed via Office.js (78 actions)
                                                    ▼
                        result ──▶ AI picks it up ──▶ reported back to you
```

**This is not "have the AI generate a .docx file."** The add-in runs inside the Office process and operates on the **live, in-memory document** — the one you are editing right now. No file locking, no last-writer-wins; your edits and the AI's edits are serialised by Office itself.

## Highlights

- **Works on the live document** — nothing hits disk, nothing gets locked
- **78 actions** — from reading a selection to inserting native equations, from conditional formatting to grouping and aligning PowerPoint shapes
- **Precise multi-document routing** — Word, Excel and PowerPoint each run their own pane instance; every command carries an `instance` so it cannot land on the wrong document
- **Mechanism-level guardrails on writes** — destructive actions re-read the current state first and, by default, wait for you to click "Apply" in the task pane (see [Security model](#security-model))
- **Hot-reloadable actions** — edit `server/actions/*`, restart the bridge, the pane reloads itself. No need to reopen the pane

---

## Compatibility

| | Status |
|---|---|
| **Windows** | ✅ Fully supported (add-in auto-registration via the WEF registry) |
| **macOS** | ✅ Core works (Node + add-in); auto-registration does not apply, load the manifest manually from Office |
| **Linux** | ❌ No Office desktop add-in host |

- **Node.js ≥ 18** (the bridge uses only the standard library)
- **Microsoft Office desktop** (Word / Excel / PowerPoint)
- **dsh** — see `peerDependencies` in `package.json` (a range is declared for `@deepseek-ai/dsh-client-ui-slots` and checked by dsh at install time)

## Installation

### Install as a dsh plugin (recommended)

```bash
# Into a profile (web, or your own profile name)
dsh plugin --profile web add dsh-office-for-mso

# Straight from the repository
dsh plugin --profile web add github:svkoh/dsh-office-for-mso

# Point at a local checkout while developing
dsh plugin --profile web add /path/to/dsh-office-for-mso
```

> `dsh plugin --profile <name> <args>` forwards its arguments to pnpm inside the profile directory.

Afterwards, **make sure the package is listed in the profile manifest** — dsh loads the ordered `dsh.profile.bundles` list, and a bare dependency is not loaded:

```jsonc
// $DSH_HOME/profiles/web/package.json
{
  "dsh": {
    "profile": {
      "bundles": ["@deepseek-ai/dsh-base", "@deepseek-ai/dsh-web-app", "dsh-office-for-mso"]
    }
  }
}
```

Installing through dsh's plugin manager maintains that list for you; if you installed with `dsh plugin ... add` by hand, check it yourself.

### Let the AI actually use the skill

On startup the plugin publishes its bundled skill to `~/.agents/skills/office-for-mso`, which is how a dsh session learns that this capability exists and when to reach for it. If skill publishing is disabled, copy it manually:

```powershell
mkdir "$HOME\.agents\skills\office-for-mso" -Force | Out-Null
copy "skills\office-for-mso\SKILL.md" "$HOME\.agents\skills\office-for-mso\"
```

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
powershell -ExecutionPolicy Bypass -File server\sideload.ps1            # register
powershell -ExecutionPolicy Bypass -File server\sideload.ps1 -Remove    # remove
```

## Architecture

| File | Role |
|---|---|
| `lib/index.js` | Host-side entry: starts the bridge, mounts the settings routes |
| `lib/sidecar-host.mjs` | Bridge lifecycle: spawn, health check, port selection, crash restart, dispose |
| `lib/client.js` | Plugin settings panel (status dot / port / autostart switch) |
| `cordis.patch.yml` | The loader row that mounts this plugin |
| `server/server.js` | Bridge: command queue, routing, heartbeat, capability discovery, static files |
| `server/taskpane.js` | The task pane: polling, heartbeat, dispatch, actions hot-reload |
| `server/actions/` | Action implementations **split by host**: `word.js` / `excel.js` / `ppt.js` / `common.js` (host-agnostic) / `runtime.js` (shared runtime) / `registry.js` (assembly + reconciliation) |
| `server/actions/parts.json` | Fragment concatenation list (order = dependency order); read by both the server and the tests |
| `server/capabilities.json` | **Single source of truth for the capability table** (name / hosts / destructive / args) |
| `server/taskpane.html` | Pane UI (including the approval panel) |
| `server/manifest.xml` | Office add-in manifest template (port 3000; rewritten for the live port into `$DSH_HOME/dsh-office-for-mso/manifest.xml`) |
| `server/pako.min.js` | zlib implementation used to unpack OOXML (third-party — see [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)) |
| `skills/office-for-mso/SKILL.md` | The skill text the AI reads (status → command → errors → safety) |
| `test/` | 105 tests (`npm test`, no test framework dependency — plain `node --test`) |

**Multi-document model**: every pane has its own `instanceId`. `GET /office/status` returns the online documents with `docUrl` (the path, used to tell documents apart) and `docTitle`. Commands must carry `instance`; without it they are refused — that is the mechanism preventing a command from being executed by the wrong document when several are open.

**Hot reload**: before each poll the pane compares the version of `/actions.js` (the maximum mtime across all fragments) and reloads it if it changed; the heartbeat also compares the mtime of `server/taskpane.js` (`codeVersion`) and reloads the whole pane when it changed. So **restarting the bridge** is all it takes after a code change.

**Port**: `127.0.0.1:3000` by default. Change it in the plugin settings panel; if the port is taken the plugin picks a free one and rewrites the add-in manifest accordingly. After a port change, close and reopen your Office documents.

## Capabilities (78 actions)

> `server/capabilities.json` is the single source of truth; `GET /office/capabilities` serves the same data. The table below is a human-readable overview.
> **D** = destructive: the current state is read before writing, and by default you approve it in the pane.
> Hosts: W = Word, E = Excel, P = PowerPoint.

### Common

| action | hosts | D | Purpose |
|---|---|---|---|
| `read_selection` | W/E/P | | Read the selection; `withStyles` adds style details |
| `write_selection` | W/E/P | D | Replace the selection with `text` |
| `read_document` | W/E/P | | Word full text / Excel used range (`sheet` picks a sheet) / PowerPoint per-slide text |
| `get_environment` | W/E/P | | Host version, platform, requirement-set support |
| `get_document_info` | W/E/P | | Document url / title / mode (read-only detection) / settings values |
| `set_setting` | W/E/P | D | Persist a key-value pair in the document (survives sessions) |

### Word

| action | hosts | D | Purpose |
|---|---|---|---|
| `read_styles` | W/E | | Selection styles (font, size, bold, colour, highlight…) |
| `replace_all` | W/E | D | Find and replace `search` → `replace`; `dryRun` returns the hit count only |
| `append_text` | W | D | Append a paragraph at the end of the document |
| `read_tables` | W | | Structured read of every table plus its style |
| `read_table_borders` / `set_table_borders` | W | D | Read / set the six border locations (WordApi 1.3+) |
| `set_font` | W | D | Set the font for the whole document (including table text) |
| `word_list_paragraphs` / `word_delete_paragraphs` | W | D | List paragraphs with zero-based `index`, then delete explicitly via `targets:[{index,expectedText}]` |
| `insert_paragraph` | W | D | Insert a paragraph with an optional built-in style and location |
| `insert_table` | W | D | Native editable table via WordApi 1.3 (`styleBuiltIn` is language-neutral) |
| `word_insert_html` | W | D | Insert HTML into the body or the selection |
| `word_set_page_setup` | W | D | Per-section margins and header/footer distances (points) |
| `word_insert_equation` | W | D | Insert a native equation (structured nodes, via OOXML) |
| `word_read_equations` | W | | Read equation OOXML for the whole document |
| `word_read_selection_paragraph` | W | | Read the paragraph around the selection and any equation OOXML in it |
| `word_export_file` | W | | Export docx / pdf (base64) |
| `word_insert_image` | W | D | Insert an image before / after / replacing the selection |
| `word_transfer_content` | W | D | Move OOXML between selection, body, tables, search hits, headers and footers |
| `apply_style` | W | D | Apply a built-in style (selection or all) |
| `format_selection` | W/E | D | Format the selection (font, size, bold, italic, colour, highlight) |
| `set_paragraph_format` | W | D | Paragraph format (alignment, indent, line spacing, list) |
| `search` | W | | Find text, returns the hits |
| `add_comment` / `read_comments` | W/E | D | Add / read comments |
| `read_properties` | W/E | | Document properties (title, author, word count…) |
| `word_list_sections` | W | | List sections |
| `word_read_headers_footers` / `word_set_header_footer` | W | D | Read / write headers and footers |
| `word_insert_break` | W | D | Insert a page or section break |
| `word_insert_caption` | W | D | Insert a caption (static numbering verified; dynamic `SEQ` must be opted into) |
| `locate_select` | W/E | | Locate and select (bookmark, named range, text, Excel range) — side-effect free |

### Excel

| action | hosts | D | Purpose |
|---|---|---|---|
| `list_sheets` | E | | List worksheets (name, position, visibility) |
| `read_range` | E | | Read a range (values, formulas, number formats) |
| `write_range` | E | D | Batch write `values` or `formulas` (chunked above 5000 cells) |
| `format_range` | E | D | Format a range and read the resulting format back |
| `insert_chart` | E | D | Data → chart (Column / Line / Bar / Pie / Area / Scatter…) |
| `add_sheet` / `rename_sheet` / `delete_sheet` | E | D | Worksheet management (`delete_sheet` supports `dryRun`) |
| `apply_sort` | E | D | Multi-field sort |
| `apply_filter` | E | D | AutoFilter |
| `evaluate_formula` | E | | Evaluate expressions such as `SUM(A1:A5, C1:C5)` (whitelist: SUM / AVERAGE / COUNT / MAX / MIN / PRODUCT) |
| `list_tables` / `create_table` | E | D | Enumerate / create native table objects |
| `delete_table` / `convert_table_to_range` | E | D | Delete the table and its data / drop the table object but keep the data |
| `read_conditional_formats` / `add_conditional_format` / `delete_conditional_format` / `clear_conditional_formats` | E | D | Conditional-format lifecycle (ExcelApi 1.6+) |
| `read_data_validation` / `set_data_validation` / `clear_data_validation` / `validate_data` | E | D | Data-validation lifecycle (1.8+; `validate_data` details need 1.9+) |

### PowerPoint

| action | hosts | D | Purpose |
|---|---|---|---|
| `read_slides` | P | | Read the currently selected slide range (use `read_document` for the whole deck) |
| `ppt_read_notes` | P | | Speaker notes for the whole file (parses OOXML notesSlides) |
| `ppt_inspect_slides` | P | | Structure snapshot: shape ID, type, name, geometry, z-order, text |
| `ppt_list_placeholders` | P | | Placeholder types and bounds |
| `ppt_add_slides` / `ppt_add_elements` | P | D | Batch-create slides, then add editable elements by the returned `slideId` |
| `ppt_delete_slides` / `ppt_delete_shapes` | P | D | Delete by explicit ID (refuses to empty the deck) |
| `ppt_insert_image` | P | D | Insert an image (positioned, or into the selection / a placeholder) |
| `ppt_insert_table` | P | D | Native editable table via PowerPointApi 1.8+ |
| `ppt_group_shapes` / `ppt_ungroup_shape` / `ppt_set_z_order` | P | D | Group / ungroup / z-order |
| `ppt_align_shapes` / `ppt_distribute_shapes` | P | D | Align / distribute using deterministic geometry over shape bounds |
| `ppt_render_slides` | P | | Render slides to PNG |
| `ppt_export_file` | P | | Native Office export |

### Boundaries

Some things Office.js simply does not offer. The plugin returns an error code instead of pretending to succeed:

- **No API exists**: creating native charts in PowerPoint, image cropping, inserting into an arbitrary placeholder ID; cross-platform "Link to Previous" and clipboard moves in Word
- **Missing in some runtimes**: Word's `body.comments` and `paragraphFormat` do not exist on every build (the plugin returns `requirement`, it does not fake success)
- **Known API traps**: Excel has no `workbook.getRange` / `calculate` (the official replacements are used); `Range.autoFilter` must go through the worksheet

## Security model

This layer deserves its own section, because it decides whether *any web page you visit* can act through your Office documents.

### The bridge only talks to its own callers

The bridge listens on `127.0.0.1`, but **listening on loopback is not security by itself** — any page in your browser can send requests to `127.0.0.1`. So:

- **No `Access-Control-Allow-*` headers at all** (no wildcard CORS)
- **A request that carries `Origin` must have `Origin` equal to the service's own origin**, otherwise `403 forbidden_origin`
- **The `Host` header is validated**, which blocks DNS rebinding (when an attacker's domain resolves to 127.0.0.1, `Host` still carries the attacker's domain)
- **Static files are allow-listed**: only the five files the pane needs are served. `manifest.xml`, `*.ps1`, `capabilities.json` and `package.json` are not on the HTTP surface

Requests without an `Origin` (local curl, or an agent calling through pwsh) pass through — those are the intended callers.

### Guardrails on destructive operations

- **Re-read before write (mechanism level)**: every destructive action reads the current state first; it never overwrites from memory
- **Approval mode by default**: with `confirmMode=ask` (the default) the pane shows an approval panel with the current state and the resulting content, and nothing runs until you click "Apply". The AI **cannot** pre-send `confirm:true` to bypass this — the pane attaches it only after you approve. An approval left untouched for 120s counts as a rejection
- **Read-only protection**: when the document is in `readOnly` / `view` mode, destructive actions return a `readonly` error
- **Explicit delete targets**: deleting paragraphs requires `word_list_paragraphs` first, then passing the verified zero-based `index` together with `expectedText`; targets whose text changed, that contain images, or that are the final paragraph are refused
- **`dryRun` previews**: `replace_all` / `delete_sheet` can return an impact preview without executing

To switch to automatic mode, use the dropdown at the top of the pane (or `POST /office/config {"confirmMode":"auto"}`). Automatic mode still re-reads before writing and returns `previousState` / `afterState` for verification.

## Customising (user-actions.js)

To add your own actions or override built-in ones, create `server/user-actions.js` (already in `.gitignore`, so upgrades will not touch it):

```js
window.__USER_ACTIONS__ = {
  my_action: {
    hosts: ['Word'],
    destructive: false,
    impl: (host, args) => ({ ok: true, result: { text: String(args.text || '') } }),
  },
  // the same name overrides a built-in action
};
```

Restart the bridge and it takes effect (hot reload — no need to reopen the pane).

> Actions in `user-actions.js` do not go through `capabilities.json`, so they must **declare their own `hosts` / `destructive`**; add a `preview` as well if you want approval plus post-write verification.

## Troubleshooting

| Symptom | What to do |
|---|---|
| Pane shows "bridge not connected" | Check the plugin is "running" in the settings panel, or `curl http://127.0.0.1:3000/office/status` |
| Commands time out | The add-in is offline: make sure the document is open and the pane says connected |
| Commands return `addin_offline` immediately | The pane is closed: open the document → Home / Developer → Add-ins → Developer Add-ins → open "DSH Office 执行器" |
| The pane never appears | Windows: make sure the bridge has run at least once (it does the registration), then **close and reopen** the document; also check Office is not running as administrator (loopback exemptions require a normal-privilege process) |
| Edited `server/actions/*` but nothing changed | Restart the bridge (the pane hot-reloads itself) |
| Edited `taskpane.js` but nothing changed | Restart the bridge, then wait for the next heartbeat (≤5s) to reload the pane; if it still does not, reopen the pane |
| The pane cannot connect after changing the port | Close and reopen your Office documents (the add-in URL lives in the manifest and has to be reloaded) |
| Edits land in the wrong place | Check the cursor / selection is where you expect; watch the execution log at the bottom of the pane |

## Notes for AI callers

The full error table is available at `GET /office/errors` (each code comes with handling advice and whether it is retryable).

- Call `GET /office/status` first and confirm the target document's instance is online with a fresh heartbeat
- **`instance` is required** — omitting it with several documents open lets the wrong one execute the command
- On `addin_offline`, **do not retry**; tell the user to open the pane
- On `confirm_required`, show `result.preview` to the user and let the pane approve it; **never pre-send `confirm:true`**
- On `rejected`, stop and respect the user's decision; do not resend
- **"Nothing found" is not an error**: `search` / `replace_all(dryRun)` return `ok:true + count:0`; report that honestly

| code | Meaning | Handling |
|---|---|---|
| `instance_required` | `instance` missing | Look up the instanceId via status and resend |
| `addin_offline` | Pane closed / offline | Ask the user to open the pane, **do not retry** |
| `busy` | A command is still running | Retry shortly |
| `timeout` | No result within 90s | Check status; one retry is reasonable |
| `bad_json` / `body_too_large` / `bad_args` | Malformed / oversized body, bad arguments | Fix and resend |
| `unknown_action` | No such action | Check capabilities for the right name |
| `unsupported_host` | Action does not support this application | Switch action or application |
| `confirm_required` | User approval needed | Hand it to the pane; never pre-send `confirm:true` |
| `rejected` | User refused / approval timed out | Stop, **do not resend** |
| `not_found` | Locate / search target missing | Report honestly, **do not repeat the same lookup** |
| `readonly` | Document is read-only / in view mode | Tell the user to save an editable copy |
| `requirement` / `unsupported` | The API is missing in this environment | Report honestly, never fake success |
| `execution` | Runtime exception | Report the error verbatim |

## Development

```bash
npm install     # only devDependencies for lint; the plugin itself has no runtime deps
npm test        # 105 tests
npm run lint
npm run check   # lint + test

# Run the bridge on its own (for debugging)
PORT=3000 node server/server.js
```

**Restarting the bridge is enough after a code change** (the pane hot-reloads):

- `server/actions/*` → restart the service
- `taskpane.js` / `taskpane.html` → restart, then wait one heartbeat for the auto-reload
- `server.js` / `lib/*` → restart dsh

### Fragment contract (`server/actions/`)

The server concatenates the fragments in `parts.json` order into **a single script** served at `/actions.js`: the browser side still sees one `<script>`, so loading and hot-reload are unchanged — while on disk the code is separated by host, so Word / Excel / PowerPoint no longer interleave.

Each fragment is a **self-contained IIFE**; they share no lexical scope and exchange everything through `window.__OFFICE_ACTIONS__`:

| Fragment | Mount point |
|---|---|
| `runtime.js` | `A.isSetSupported` / `A.validateMatrix` / `A.sleep` / `A.notFound` |
| `word.js` | `A.hosts.Word = { actions, selectTargets, previewReaders, afterContent, environmentProbe }` |
| `excel.js` | `A.hosts.Excel = { … }` |
| `ppt.js` | `A.hosts.PowerPoint = { … }` |
| `common.js` | `A.hosts.Common = { actions, previewReaders, afterContent }` (host-agnostic, Common API only) |
| `registry.js` | Assembly: flatten into the registry, reconcile against `capabilities.json`, expose `__ACTIONS__` / `__EXECUTE__` / `__AFTER_CONTENT__` |

Three boundaries are enforced by tests (`test/actions-boundary.test.mjs`) — they are what makes the split meaningful rather than cosmetic:

1. A host fragment must contain **no** host-branching conditionals — the functions no longer take a `host` parameter at all
2. `word.js` must not reference `Excel.run` / `PowerPoint.run` (and vice versa)
3. `common.js` / `runtime.js` must not reference any host object model

**To change the capability table, edit `server/capabilities.json` only**: `hosts` / `destructive` come from the server-injected `window.__ACTIONS_META__`. On load, `registry.js` reconciles in three directions — implemented but unregistered / registered but unimplemented / destructive without a preview reader — and refuses to load if any of them fails, instead of silently dropping host validation or the approval interception.

**Bump `version` in `package.json` whenever you change the plugin's bytes** — downstream compares version numbers, and the same version with different bytes is not reinstalled.

## License

[MIT](LICENSE). The license and provenance of the one third-party component (pako) are in [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).

---

*Not an official Microsoft product. This project is not affiliated with Microsoft Corporation.*
