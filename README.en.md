# dsh-ms-office

[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Node.js](https://img.shields.io/badge/node-%E2%89%A5%2018-3c873a.svg)](#compatibility)
[![CI](https://github.com/svkoh/dsh-ms-office/actions/workflows/ci.yml/badge.svg)](https://github.com/svkoh/dsh-ms-office/actions/workflows/ci.yml)

English · [中文](README.zh.md)

Read and write the Microsoft Office documents you already have open — Word, Excel and PowerPoint — from a [DeepSeek Harness (DSH)](https://github.com/deepseek-ai/deepseek-harness) session: content, formatting, structure, tables, charts, formulas and comments.

![Runtime data flow: dsh session → bridge → add-in task pane → Office document](docs/architecture-runtime.svg)

The add-in runs inside Office and edits the document you currently have open. Changes are visible immediately, with no saving or reopening, and you and the AI can edit at the same time — Office applies the changes in order.

> Throughout this document, "Office" means the Microsoft Office desktop applications, not WPS or other compatible suites, and not the web versions.

## Quick start

```bash
dsh plugin --profile web add dsh-ms-office
```

`--profile web` is a dsh profile name; use whichever one you normally run.

Installing makes dsh load the plugin, which then starts the local bridge service on its own — there is nothing to launch by hand. Confirm the service state is `healthy` on the plugin panel.

1. Open a Word, Excel or PowerPoint document.
2. Open the add-in task pane:
   - Windows: the service registers the add-in when it starts. Close and reopen the document and the pane appears.
   - macOS: the manifest has to be loaded by hand — see [Loading the manifest on macOS](#loading-the-manifest-on-macos).
3. The pane shows "Connected — waiting for DSH commands" — you are ready.
4. Ask your dsh session to read the document, for example "read the whole of the Word document I have open".

The pane must stay open. Office.js runs inside it, and closing it produces `addin_offline`. The pane can be shrunk or tucked into a corner.

### Loading the manifest on macOS

File → Options → Customize Ribbon → tick "Developer", then with a document open go to Developer → Add-ins → manage "Developer Add-ins" → add `manifest.xml`.

## Capabilities

- 92 actions across Word, Excel and PowerPoint: reading selections, writing content, formatting, tables, charts, formulas, comments, PowerPoint grouping and alignment.
- One instance per document, and every command must name its target, so a command cannot land on the wrong document.
- Destructive operations read the current state first; a failed preview blocks execution, and by default you still click "Apply" in the pane.
- Bilingual UI (Chinese / English) that follows your Office display language.

The full action list is in [docs/CAPABILITIES.en.md](docs/CAPABILITIES.en.md).

Two skills ship with the plugin. [ms-office](skills/ms-office/SKILL.en.md) handles safe operations; [office-document-styles](skills/office-document-styles/SKILL.en.md) provides one original base style atlas per host — Word covers document elements, Excel is fixed to two sheets, and PowerPoint shows shapes, tables, spatial layout and colour roles, each with provenance records and layered quality checks. See the [style gallery](docs/style-gallery.en.md) and the [feature roadmap](docs/FEATURE-ROADMAP.en.md).

Chinese defaults: Word body SimSun and headings SimHei; Excel SimSun; PowerPoint titles SimHei, body and table/chart labels KaiTi. Every change to an atlas runs through the plugin itself, and exports and screenshots keep native Office evidence.

## Compatibility

| Platform | Status | Notes |
| --- | --- | --- |
| Windows | Full | Add-in auto-registration via the WEF registry |
| macOS | Core | Node and the add-in work; auto-registration does not apply, load the manifest manually from Office |
| Linux | None | No Office desktop add-in host |

- Node.js ≥ 18 (the bridge uses only the standard library)
- Microsoft Office desktop (Word / Excel / PowerPoint)
- dsh — see `peerDependencies` in `package.json`

## Documentation

| Document | Contents |
| --- | --- |
| [docs/CAPABILITIES.en.md](docs/CAPABILITIES.en.md) | Full action list and capability boundaries |
| [docs/DEVELOPMENT.en.md](docs/DEVELOPMENT.en.md) | Architecture, file map, fragment contract, i18n, hot reload, custom actions, packaging rules |
| [docs/AGENT.en.md](docs/AGENT.en.md) | HTTP endpoints, error codes, approval protocol, for scripts and other agents |

## License

[MIT](LICENSE). The license and provenance of the third-party component (pako) are in [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).

This project is a fork of <https://github.com/Mikuzjc/dsh-office-for-mso>.
