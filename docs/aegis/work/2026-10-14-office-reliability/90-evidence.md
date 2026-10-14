# Verification evidence

Outcome: approved workspace implementation delivered; confidence B. Running installed plugin was not replaced or published. No user Office document content/style was changed for verification.

## Fresh checks
- `npm run check > .tmp/office-reliability-check.log 2>&1`: exit 0; lint passed; 425 tests passed, 30 suites, 0 failures/cancellations/skips. Full default isolated runner used. Log summary lines584–591.
- `git diff --check`: exit0.
- Independent advisory review found unsupported enum aliases, alias-before-exact custom style selection, and overbroad rendering diagnostic wording. All corrected. Collision regression covers custom Heading-1 and 列表段落 with real listParagraph enum shape; lower WordApi1.3 semantic style behavior retained.
- Independent single-model document-based skill dry-run covered report/table addition under deadline, typo-only locked template, spreadsheet rows/notes/filter/totals, and exact custom style patch. Chinese/English core obligations match. This is not measured cross-model reliability or live model editing.

## Live native Word acceptance
Read-only bridge status exposed one online Word instance; COM exposed one open Word document with5 tables. Read-only COM row probe captured two rows each from two tables: final two native Characters each return [13,7], contiguous native ranges. Native positions differ from Range.Text UTF-16 length, confirming why raw offset subtraction is unsafe.

Read-only installed `read_document` returned current body length6676, tabs103, paragraphMarks96. Using that body and expectedPageCount8 with the workspace renderer succeeded:8 native A4 pages (595x841pt), PNG1190x1682 at144dpi. Fingerprint bc830e5a8f06ec28be7de7848d7f1250da57c8eb3f8da96b533bfd39a6027035. All8 PNG payloads validated by renderer; page3 opened and inspected (body/page number visible). Receipt `.tmp/live-render-receipt.json`; images `.tmp/live-render-pages/`. This validates the existing live document matching/rendering path, not every document layout.

## Repair, retention, and limits
Retired flat table text splitting and one-name paragraph-style validation in place. Existing action names, result envelope, live-instance routing, approval and before/after rendering guards remain. WordApi1.3 enum path retained for hosts lacking1.5;1.5 uses exact actual style names before supported aliases. No fuzzy candidate rendering or implicit style creation.

Structural row projection fails closed for unsupported marker shapes, nested tables and some merged-row access. Broad empty/Unicode/object/overlap/bounds cases have synthetic real-PowerShell coverage; live exotic-table/multi-candidate matrix remains untested. Other style writes are verified by mocked Office contract tests, not by modifying the user's report. Theme-color host behavior is detected/reported through actual field comparison; its original incident was not independently reproduced.

Complexity control: existing owners extended with bounded lookup/verification/projection; obsolete parsing/one-name validation removed; no duplicate runtime owner or architectural split. .codegraph is a local generated index, excluded from Git and source-only provenance scanning. Pre-existing untracked atlas work preserved.

## Delivery boundary
Code, bilingual skills, capability descriptions, input/output explanations and UTF-8 examples are ready in this checkout. No runtime deployment or package publish performed; source checkout and installed plugin remain distinct. Model beautification rule is instruction-based and preserves explicit constraints/reasonable templates, with bounded local changes and honest visual-evidence limits.
