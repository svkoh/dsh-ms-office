# Changelog

## 0.3.3 — 2026-10-19

- Render the Office bridge settings on the plugin's own detail page. The Plugins page addresses an installed bundle as `{ kind: "bundle", pkg }`, while the section only accepted `{ kind: "item", id }`, so it returned null on the one page it was written for. Both shapes are now matched by name, and a row page still renders nothing.
- Verified the host half against dsh 0.2.1-rc.2: the entry mounts through `ctx.inject(["webServer"])`, registers `/api/ms-office/{describe,save,sideload}`, round-trips the port/autostart settings and disposes cleanly. The declared `engines.dsh` and peer ranges already admit rc.2, so they are unchanged.

## 0.3.2 — 2026-10-15

- Prepare the first public NanoAI plugin artifact using the existing npm package whitelist.
- Add the native `dsh-plugin.json` declaration for the existing host and web client entries.
- Normalize `dsh.bundle.patch` to `cordis.patch.yml` for the NanoAI installer archive contract.
- Align the package lock version with the package version and retain MIT and bundled pako license notices.

This release packages the existing Office bridge, 92 actions, bilingual skills, and native document style resources. It adds no Office actions. Automated checks use local HTTP and Office object stubs; they do not establish desktop Office rendering compatibility.
