# Native Word preview reproduction

The editable source is [word-base-styles.docx](word-base-styles.docx). It retains genuine TOC and PAGE fields. Microsoft Word confirmed five pages, one TOC, one table, one inline image, and one footer PAGE field.

Run from this repository on Windows with desktop Microsoft Word installed:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/render-word-atlas.ps1
```

The standalone [renderer](../../scripts/render-word-atlas.ps1) opens a temporary source copy read-only in one dedicated Word instance. It records that instance's native window handle, process ID, start time, and zero-document baseline before opening the copy. A supervising process enforces a 30-second deadline and stops only its identified owned Word process and worker. It does not use ActiveDocument, clipboard, Application.Quit, global Word options, or any existing user document.

Outputs are [five previews](word-base-styles/) and [native evidence](word-base-styles.evidence.json). Each PNG is a native Word page body Range.EnhMetaFileBits raster, preserving the metafile's aspect ratio. These are body-range previews: page margins, headers and footers are excluded. They are not full-page screenshots or PDF page renders. The genuine source retains those features; full-page header/footer visual QA has not been completed.

The renderer checks expected native counts and verifies the original DOCX SHA256 remains unchanged before publishing. Diagnostics remain in a uniquely named temporary run directory. PDF export is not attempted. On 2026-10-14, native ExportAsFixedFormat exceeded a 30-second deadline even for a one-paragraph ASCII control; field-free, image-free and modern-layout atlas copies also timed out. No PDF is claimed, and no PDF was synthesized from these cropped images.
