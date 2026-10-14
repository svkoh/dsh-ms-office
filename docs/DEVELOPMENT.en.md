# Development guide

> English · [中文](DEVELOPMENT.md)

For people changing this plugin. Usage lives in the [README](../README.en.md); the wire protocol is in the [integration reference](AGENT.en.md).

## Architecture

### Runtime data flow

![Runtime data flow: dsh session → bridge → add-in task pane → Office document](architecture-runtime.svg)

### Module structure

![Module structure: how the files on disk are served as HTTP endpoints](architecture-modules.svg)

### File responsibilities

| File                                             | Role                                                                                                                                                                                  |
| ------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `lib/index.js`                                   | Host-side entry: starts the bridge, mounts the settings routes                                                                                                                        |
| `lib/sidecar-host.mjs`                           | Bridge lifecycle: spawn, health check, port selection, crash restart, dispose; picks `SKILL.md` or `SKILL.en.md` by host language when publishing the skill                           |
| `lib/client.js`                                  | Plugin settings panel (status dot / port / autostart switch)                                                                                                                          |
| `cordis.patch.yml`                               | The loader row that mounts this plugin                                                                                                                                                |
| `server/server.js`                               | Bridge: command queue, routing, heartbeat, capability discovery, static files, per-request text assembly                                                                              |
| `server/taskpane.js`                             | The task pane: polling, heartbeat, dispatch, actions hot-reload                                                                                                                       |
| `server/actions/`                                | Action implementations **split by host**: `word.js` / `excel.js` / `ppt.js` / `common.js` (host-agnostic) / `runtime.js` (shared runtime) / `registry.js` (assembly + reconciliation) |
| `server/actions/parts.json`                      | Fragment concatenation list (order = dependency order); read by both the server and the tests                                                                                         |
| `server/capabilities.json`                       | **Structure-only source of truth** for the capability table: name / hosts / destructive / requirement. Human-readable arg text is **not** here                                        |
| `i18n/zh.json` · `i18n/en.json`                  | **Single source of truth for all text**: pane UI, server messages and error catalog, capability arg descriptions                                                                      |
| `server/taskpane.html`                           | Pane UI skeleton (Fluent 2 semantic roles, follows the system dark mode; every visible string is filled in from `/office/i18n`)                                                       |
| `server/manifest.xml`                            | Office add-in manifest template (port 3000; rewritten for the live port into `$DSH_HOME/dsh-ms-office/manifest.xml`; display name / description carry `en-US` overrides)         |
| `server/pako.min.js`                             | zlib implementation used to unpack OOXML (third-party — see [THIRD\_PARTY\_NOTICES.md](../THIRD_PARTY_NOTICES.md))                                                                    |
| `skills/ms-office/SKILL.md` · `SKILL.en.md` | The skill text the AI reads (Chinese / English variants — exactly one is published, so only one skill is registered)                                                                  |
| `docs/`                                          | Documentation and the two diagrams (the SVGs have no external font dependency and read on light and dark themes)                                                                      |
| `test/`                                          | Tests (`npm test`, no test framework dependency — plain `node --test`)                                                                                                               |

## How it runs

### Multi-document model

Every pane has its own `instanceId`. `GET /office/status` returns the online documents with `docUrl` (the path, used to tell documents apart) and `docTitle`. Commands must carry `instance`; without it they are refused — that is the mechanism preventing a command from being executed by the wrong document when several are open.

### Commands and health states

See the [integration reference](AGENT.en.md) for the command protocol. The default is `COMMAND_TIMEOUT_MS=180000`; approval lasts at most 120 seconds and cannot exceed `expiresAt`. Every pane execution uses an atomic `/start` claim, and results carry `instance` plus boolean `ok`. External `args.confirm` is ignored; internal approval uses a single-use closure bound to the command. Timeouts during execution become `execution_unknown` and retain the instance slot until a late result arrives, preventing blind retries from duplicating writes. The settings panel distinguishes `starting` / `healthy` / `unhealthy` / `stopped`; a child process being present does not prove availability. Post-write verification failures add `verificationWarnings` instead of misreporting a successful write; complete document readback is not guaranteed.

### Hot reload

Before each poll the pane compares the version of `/actions.js` (the maximum mtime across all fragments) and reloads it if it changed; the heartbeat also compares the mtime of `server/taskpane.js` and `i18n/*.json` (`codeVersion`) and reloads the whole pane when it changed.

So **restarting the bridge** is all it takes after a code change. Note that `codeVersion` covering `i18n/*.json` is required — otherwise editing a string would not trigger a reload.

### Port and bridge.json

`127.0.0.1:3000` by default; change it in the plugin settings panel. If the port is taken the plugin picks a free one and rewrites the add-in manifest accordingly. After a port change the user must close and reopen their Office documents (the add-in URL lives in the manifest and Office reads it once, when the document loads).

On startup the service writes its current address to `$DSH_HOME/dsh-ms-office/bridge.json`:

```json
{ "plugin": "dsh-ms-office", "url": "http://127.0.0.1:3000", "port": 3000, "pid": 1234, "startedAt": "…" }
```

That file is for **callers** (AI / scripts): they cannot see the plugin settings panel and the port is variable, so there has to be a machine-readable landing point. A missing file means the bridge is not running. **Deleting or corrupting it is harmless** — the next startup rewrites it.

### i18n

The UI language is resolved in this order: **`?lang=` → `Accept-Language` → `OFFICE_LANG` env var → system locale → Chinese**. A tag starting with `en` selects English, one starting with `zh` selects Chinese, anything else falls back to Chinese.

- **Task pane**: reads `Office.context.displayLanguage` (the editing language the user set in Office), falling back to `navigator.language`; it fetches its dictionary from `GET /office/i18n?lang=` at startup and **does not render any status text until the dictionary has arrived** (otherwise it would flash "connected" and then flip to an error).
- **Server**: `/office/capabilities`, `/office/errors` and `/office/i18n` all return a `lang` field and assemble their bodies in that language; startup logs go through the same dictionary.
- **Adding a language**: drop an `<lang>.json` into `i18n/`, add it to `LANGS` in `server/server.js`; `test/i18n.test.mjs` enforces that both sides have identical key sets and identical `{placeholder}` sets. Add an `<Override Locale="…">` for the manifest display name too.
- **Structure never varies with language**: `capabilities.json` holds only machine-consumed fields, so switching language only changes text. Editing a description **cannot** touch the data that decides host validation or approval interception — that is the entire point of the structure / text split.

## Installation and loading

### Profile manifest (`dsh.profile.bundles`)

dsh loads the ordered `dsh.profile.bundles` list; a bare dependency is not loaded. This plugin declares `dsh.bundle` in its own `package.json`, so `dsh plugin add` and the desktop plugin panel add it to that list for you:

```jsonc
// $DSH_HOME/profiles/web/package.json — written automatically, no hand editing
{
  "dsh": {
    "profile": {
      "bundles": ["@deepseek-ai/dsh-base", "@deepseek-ai/dsh-web-app", "dsh-ms-office"]
    }
  }
}
```

**One exception**: if the dependency already exists (the user **disabled** it earlier), the official path will not add it back, to avoid re-enabling something on purpose. In that case the user enables it from the plugin panel.

### Skill publishing

On every startup the plugin publishes the bridge and style skills to `~/.agents/skills/ms-office` and `~/.agents/skills/office-document-styles`, which is how a dsh session learns that this capability exists and when to reach for it. **This call is unconditional — there is no switch in the code.** A failure only leaves a `skill publishing skipped: …` line in the plugin log and does not affect the plugin.

The skill name is its **directory name**, so both language variants live in the **same directory** (`SKILL.md` + `SKILL.en.md`); at publish time the host language (`LC_ALL` / `LC_MESSAGES` / `LANG` starting with `en`) selects one and it is copied over `SKILL.md`. The bridge and design skill have distinct roles; language variants of the same skill must share one directory.

### Publish allow-list (`package.json#files`)

This plugin is `"sync": "npm-files"` in the parent repo's `plugins/manifest.json` ⇒ **the published package contains only what `files` lists**. When you add a directory you **must** add it to `files` too, otherwise you ship a package with missing files and **packaging reports nothing** (`test/plugin-manifest.test.mjs` catches this class statically).

## Fragment contract (`server/actions/`)

The server concatenates the fragments in `parts.json` order into **a single script** served at `/actions.js`: the browser side still sees one `<script>`, so loading and hot-reload are unchanged — while on disk the code is separated by host, so Word / Excel / PowerPoint no longer interleave.

Each fragment is a **self-contained IIFE**; they share no lexical scope and exchange everything through `window.__OFFICE_ACTIONS__`:

| Fragment      | Mount point                                                                                                                            |
| ------------- | -------------------------------------------------------------------------------------------------------------------------------------- |
| `runtime.js`  | `A.isSetSupported` / `A.validateMatrix` / `A.sleep` / `A.notFound`                                                                     |
| `word.js`     | `A.hosts.Word = { actions, selectTargets, previewReaders, afterContent, environmentProbe }`                                            |
| `excel.js`    | `A.hosts.Excel = { … }`                                                                                                                |
| `ppt.js`      | `A.hosts.PowerPoint = { … }`                                                                                                           |
| `common.js`   | `A.hosts.Common = { actions, previewReaders, afterContent }` (host-agnostic, Common API only)                                          |
| `registry.js` | Assembly: flatten into the registry, reconcile against `capabilities.json`, expose `__ACTIONS__` / `__EXECUTE__` / `__AFTER_CONTENT__` |

Three boundaries are enforced by tests (`test/actions-boundary.test.mjs`) — they are what makes the split meaningful rather than cosmetic:

1. A host fragment must contain **no** host-branching conditionals — the functions no longer take a `host` parameter at all
2. `word.js` must not reference `Excel.run` / `PowerPoint.run` (and vice versa)
3. `common.js` / `runtime.js` must not reference any host object model

**To change the capability table, edit `server/capabilities.json` only**: `hosts` / `destructive` come from the server-injected `window.__ACTIONS_META__`. On load, `registry.js` reconciles in three directions — implemented but unregistered / registered but unimplemented / destructive without a preview reader — and refuses to load if any of them fails, instead of silently dropping host validation or the approval interception.

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

## Development

```bash
npm install     # only devDependencies for lint; the plugin itself has no runtime deps
npm test
npm run lint
npm run check   # lint + test

# Run the bridge on its own (for debugging)
PORT=3000 node server/server.js
```

**Restarting the bridge is enough after a code change** (the pane hot-reloads):

- `server/actions/*` → restart the service
- `taskpane.js` / `taskpane.html` / `i18n/*` → restart, then wait one heartbeat for the auto-reload
- `server.js` / `lib/*` → restart dsh

**Bump `version` in `package.json` whenever you change the plugin's bytes** — downstream compares version numbers, and the same version with different bytes is not reinstalled.

**Adding tests**: `node --test`, no framework dependency. The helper boots the bridge on a random port (`test/helpers/server.mjs`). Test environments must set `OFFICE_SKIP_SIDELOAD=1` to prevent Office add-in registration. Node vm tests use Office stubs and do not establish real Office validation.

## License

[MIT](../LICENSE). The license and provenance of the one third-party component (pako) are in [THIRD\_PARTY\_NOTICES.md](../THIRD_PARTY_NOTICES.md).

## Document style resources

`skills/office-document-styles/SKILL.md` and `SKILL.en.md` ship presets, examples, provenance and offline planning/quality scripts. The six original recipes use `resources/presets.json` as their source of truth; browser previews are not Office screenshots. See [FEATURE-ROADMAP.en.md](FEATURE-ROADMAP.en.md) for remaining boundaries.

`lib/skill-publisher.mjs` compares the entire published tree, including resources when the main skill text is unchanged. An ownership manifest limits stale-file cleanup to previously published files, preserving unknown user files. Symlink/junction and traversal paths are rejected; the manifest updates after successful copying. Source and resource changes require installation and a DSH restart; editing this checkout does not update an already installed runtime.
