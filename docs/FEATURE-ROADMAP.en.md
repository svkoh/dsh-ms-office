# Functional audit and roadmap

This review describes source version0.3.0. Installed capabilities must be checked at runtime. Object-model tests and offline plan checks do not establish real Office rendering quality.

## Implemented in this iteration

- Word: exact index/text paragraph patches with semantic styles, fonts and spacing; real Paragraph/list APIs, correct alignment enum, requirement gates, actual page-setting previews and strict table-border targets/readback.
- Excel: freeze panes, point dimensions, wrapping/alignment/autofit rows, layout readback and complete-range PNG export with a cell budget.
- PowerPoint: patches to existing shapes, optional style inspection, actual native page dimensions and document order, and rendering restricted to selected IDs.
- Design: two original recipes per host, provenance, labeled examples, reviewable command plans, safe result bindings and offline quality checks.
- Publishing: complete resource-tree synchronization, language selection, ownership-limited stale-file cleanup and path safety.

## Remaining common tasks

| Priority | Gap | Proposed completion criterion |
| --- | --- | --- |
| P1 | Word global style definitions and native updateable TOC/page numbers/cross-references | Read/update Normal, Heading and Caption definitions; verify native fields and exported pagination |
| P1 | Excel print layout | Read/write print area, orientation, scaling and repeated headings, then inspect actual printed pages |
| P1 | Editing existing charts | List/read stable IDs, patch series/axes/legend/layout/colors, render and verify values |
| P1 | Reusing PPT master/layout templates | List existing layouts and create pages from them; distinguish reuse from full theme/master editing |
| P1 | Multi-step preview, progress, cancellation and recovery | Explicit scope and authorization, per-step receipts and accurate partial completion; no universal rollback claim |
| P1 | Native visual regression set | Real report, spreadsheet and presentation examples with PDF/PNG checks for CJK fonts, overflow and pagination |
| P2 | Excel pivot tables | Ordinary pivot create/read/refresh, valid fields and free destination; no PowerPivot/DAX equivalence claim |
| P2 | Existing Word table maintenance | Stable targeting, values/width/header/row-column edits and readback; explicit merged-table boundaries |
| P2 | PPT order, native background and sections | Verified native APIs and stable IDs, preserving notes and relationships |
| P2 | Brand profiles and scenario recipes | Versioned company colors/fonts/type/numeric formats; user templates take priority |
| P2 | Accessibility and publishing checks | Semantic headings, table headers, alt text and reading order; unobservable checks remain pending |
| P3 | Notes writing and deeper theme/master editing | Verify actual platform APIs first instead of inventing setters |

Excel already supports formula matrices, native Tables, conditional formatting and validation. PPT already supports native tables/images, grouping, z-order, alignment, distribution and notes reading. Notes extraction may still include page-number placeholders, which require filtering. Word already has HTML/OOXML, equations, tables, headers/footers and export. These are not evidence that global style definitions, native TOCs or every rich-document operation are supported.

Next acceptance should use a roughly five-page report, a sheet mixing money/counts/ratios/totals, and a six-to-eight-page deck. Capture actual readbacks and native exports, verify content and visual quality separately, and turn failures into small reproducible native fixtures before expanding the catalog.

The [gallery](style-gallery.en.md) shows native Office page previews. See the [style skill](../skills/office-document-styles/SKILL.en.md), [source evidence](../skills/office-document-styles/references/sources.md), and [full Chinese audit](FEATURE-ROADMAP.md).
