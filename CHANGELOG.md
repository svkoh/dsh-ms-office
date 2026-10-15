# Changelog

## 0.3.2 — 2026-10-15

- Prepare the first public NanoAI plugin artifact using the existing npm package whitelist.
- Add the native `dsh-plugin.json` declaration for the existing host and web client entries.
- Normalize `dsh.bundle.patch` to `cordis.patch.yml` for the NanoAI installer archive contract.
- Align the package lock version with the package version and retain MIT and bundled pako license notices.

This release packages the existing Office bridge, 92 actions, bilingual skills, and native document style resources. It adds no Office actions. Automated checks use local HTTP and Office object stubs; they do not establish desktop Office rendering compatibility.
