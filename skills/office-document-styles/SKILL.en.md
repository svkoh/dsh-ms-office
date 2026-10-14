---
name: office-document-styles
description: Baseline design recipes and quality checks for Word reports, Excel tables and budgets, and PowerPoint presentations. Build a reviewable Office command plan with consistent headings, fonts, colors, spacing and slide grids, then verify content and rendered output.
whenToUse: The user wants to create, format, improve or standardize a Word, Excel or PowerPoint document, asks for a template or visual style, or needs document quality checks.
---

# Office document styles

Understand the content before choosing a layout. Use the host's default recipe when no brand is specified; existing company templates and explicit user preferences take priority. All bundled design tokens and recipes are original. Internet examples are references, with provenance and licensing recorded in [sources.md](references/sources.md). Third-party templates, images, masters and fonts are not bundled, and a color recipe is not an Office theme.

## Resources
Use one native file per host, with different pages demonstrating elements and layouts. Do not multiply similar business templates.
- [style-atlas.json](resources/style-atlas.json): element coverage, fonts, geometry, action mapping and boundaries.
- [Word atlas](resources/templates/word-base-styles.docx): cover, real TOC, headings/paragraphs/lists, table, alignment, color, embedded image/caption, header/footer/PAGE.
- [Excel atlas](resources/templates/excel-base-styles.xlsx): exactly two sheets, basic styles and data/charts.
- [PowerPoint atlas](resources/templates/ppt-base-styles.pptx): shapes, real native table, data chart, images, flow, timeline, grids and color roles.
- [sources.json](resources/sources.json), [sources.md](references/sources.md) and [quality-guide.md](references/quality-guide.md): references, licensing and verification.
Available PDFs and native PNGs are listed in each verification record; missing exports remain explicit. External document-generation scripts do not qualify as plugin acceptance. Save a separate working copy and reuse only needed elements. Historical business compilers remain compatible but are not the default catalog. Run scripts/style-atlas.mjs list, show word|excel|ppt, or check with Node18+. The check validates definitions/files, not visual quality or live Office.js support. Never replace user facts with DEMO numbers or invent material to meet page, character or row targets. USER is internal provenance and is not printed.
## Apply and verify

1. Load the `ms-office` bridge skill. Read its announcement, `/office/status`, `/office/capabilities`, and `get_environment`; identify the exact live instance. Source code availability does not prove the installed runtime supports an action. Stop with an accurate explanation if a requirement is missing.
2. Capture the baseline: Word paragraphs and document text; Excel sheet names; PPT `ppt_read_presentation_info`, `ppt_inspect_slides({withStyles:true})`, and document text. Native slide sizes are in pt; exported PNG pixel dimensions do not determine the document's page size.
3. Select the needed element/layout after reading existing content and styles. Confirm a blank document from live reads before new-document generation; preserve existing user paragraphs, tables and slides. Avoid conflicting Excel sheet names. If using a batch plan, validate preconditions and checkPlan.
4. Briefly explain the design and affected scope, then work within existing authorization. Keep the current approval mode. External `confirm:true` and private approval hooks cannot grant permission.
5. Execute steps in order with the plan's instance and each step's stable idempotency key. Resolve bindings through `resolveStep(step,results)` only after the referenced action succeeds; `results[from]` contains its `result` data. Never send unresolved nulls or placeholders, use eval, or guess IDs. Bindings are local to this plan and instance.
6. On rejection or failure, stop dependent writes and retain receipts for successful steps. The batch is not an atomic transaction. For timeout/execution_unknown, query the original command and inspect the document before any retry; keep the existing key.
7. Read back content and styles, then inspect real output. Render only newly created PPT pages with `ppt_render_slides({slideIds:[...]})`; render Excel ranges with `excel_render_range` when available; render complete Word pages including headers/footers with `word_render_pages({expectedPageCount:actualCount})` on Windows PrintLayout; export native Word PDF when available. Decode Base64 into local artifacts before inspecting them. Export success alone is not proof of correct pagination or absence of overflow.
8. Report completed work, verified checks, repairs and outstanding manual checks. Content correctness and visual correctness require separate evidence.

## Host defaults
**Word:** one multi-page atlas; SimSun12pt body,24pt first-line indent,true1.5 spacing; SimHei14pt H1,18pt document title,28pt cover. Actual global atlas body, Normal, Title, Heading1/2/3/Caption definitions. Create the true1.5 global body style with word_set_styles({styles:[{name:"范本正文",preset:"body-1.5"}]}). Apply this named style to body paragraphs; unused Normal spacing is not claimed as1.5, real TOC/PAGE fields, native table and embedded image. Direct formatting does not redefine global styles. When the live runtime supports WordApi1.5, use word_read_styles/word_set_styles for global styles and word_insert_field/header-footer field parameters for real TOC/PAGE. lineSpacing is in points; verify native360/auto before claiming true1.5 spacing. Choose alignment by purpose and keep color restrained.
**Excel:** Chinese text uses SimSun (宋体),11pt body; exactly two sheets. Basic styles show bold titles, headers/borders, alignment, numeric types, real auto-width, wrap/auto-height, native Table/filter/freeze/conditional formats. Data/charts show actual SUM/difference/IFERROR formulas and a native Chart linked to source cells. format_range accepts autoFitColumns (legacy autoFit alias), wrapText and autoFitRows; dimensions/autofit require ExcelApi1.2. Format data, auto-fit only data columns, bound long-note widths, then wrap and auto-fit rows. Fixed and automatic sizes are mutually exclusive. Configure printing with excel_set_print_layout and verify actual native PDF paper size, page count and scaling.
**PowerPoint:** Chinese titles use SimHei (黑体); body, table and chart labels use KaiTi (楷体). One deck with visibly distinct element/layout pages at960x540pt. Fonts30/18/16/14/12pt for title/body/label/table-chart/footer. One white-navy-teal palette, dark titles/KPIs, transparent text, no font shrinking. Includes a real native Table, embedded image and editable shape chart; geometry is not a native Chart. Show two-column,three-column,flow,timeline and grid layouts. Preserve user slides and check actual dimensions/capabilities, including PowerPointApi1.8 for table insertion.
## Quality boundary

Plan checks cover contrast, type sizes, geometry, density, names, host compatibility and bindings. Content checks cover evidence, units, formula results and demo labeling. Real-output checks cover page breaks, clipping, font substitution, table widths, reading order and image quality. Report unobserved checks as outstanding instead of claiming they passed.


Bridge discovery: read `url` from `$DSH_HOME/dsh-ms-office/bridge.json`; `http://127.0.0.1:3000` is only a default reference. Follow the bridge skill for actual discovery.
