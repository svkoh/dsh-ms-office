# Full action reference

> English · [中文](CAPABILITIES.md)

Every action this plugin currently supports. **The authoritative list is the live one** — fetch it any time:

```bash
curl "$(jq -r .url "$DSH_HOME/dsh-office-for-mso/bridge.json")/office/capabilities?lang=en"
```

(`bridge.json` is explained in the [integration reference](AGENT.en.md#finding-the-bridge-address).)

**How to read this**

- **D** = destructive: the current state is read before writing, and by default you approve it in the pane.
- **Hosts**: W = Word, E = Excel, P = PowerPoint.
- **Enum values are case-insensitive** (`align:"left"` and `align:"Center"` are both accepted); a rejected value lists the valid ones in the error — fix it from the list instead of guessing.
- **Field names follow the `args` descriptions.** Common misspellings are aliased: `kind`→`type`, `width`/`height`→`w`/`h`, `fontSize`→`size`; colors accept `#RRGGBB` only.

> `server/capabilities.json` is the single source of truth for the **structure** (name / hosts / destructive / requirement);
> the human-readable arg text comes from `i18n/<lang>.json`, and the endpoint merges the two.
> The table below is a human-readable overview.

## Common

| action              | hosts | D | Purpose                                                                               |
| ------------------- | ----- | - | ------------------------------------------------------------------------------------- |
| `read_selection`    | W/E/P |   | Read the selection; `withStyles` adds style details                                   |
| `write_selection`   | W/E/P | D | Replace the selection with `text`                                                     |
| `read_document`     | W/E/P |   | Word full text / Excel used range (`sheet` picks a sheet) / PowerPoint per-slide text |
| `get_environment`   | W/E/P |   | Host version, platform, requirement-set support                                       |
| `get_document_info` | W/E/P |   | Document url / title / mode (read-only detection) / settings values                   |
| `set_setting`       | W/E/P | D | Persist a key-value pair in the document (survives sessions)                          |

## Word

| action                                                 | hosts | D | Purpose                                                                                              |
| ------------------------------------------------------ | ----- | - | ---------------------------------------------------------------------------------------------------- |
| `read_styles`                                          | W/E   |   | Selection styles (font, size, bold, colour, highlight…)                                              |
| `replace_all`                                          | W/E   | D | Find and replace `search` → `replace`; `dryRun` returns the hit count only                           |
| `append_text`                                          | W     | D | Append a paragraph at the end of the document                                                        |
| `read_tables`                                          | W     |   | Structured read of every table plus its style                                                        |
| `read_table_borders` / `set_table_borders`             | W     | D | Read / set the six border locations (WordApi 1.3+)                                                   |
| `set_font`                                             | W     | D | Set the font for the whole document (including table text)                                           |
| `word_list_paragraphs` / `word_delete_paragraphs`      | W     | D | List paragraphs with zero-based `index`, then delete explicitly via `targets:[{index,expectedText}]` |
| `insert_paragraph`                                     | W     | D | Insert a paragraph with an optional built-in style and location                                      |
| `insert_table`                                         | W     | D | Native editable table via WordApi 1.3 (`styleBuiltIn` is language-neutral)                           |
| `word_insert_html`                                     | W     | D | Insert HTML into the body or the selection                                                           |
| `word_set_page_setup`                                  | W     | D | Per-section margins and header/footer distances (points)                                             |
| `word_insert_equation`                                 | W     | D | Insert a native equation (structured nodes, via OOXML)                                               |
| `word_read_equations`                                  | W     |   | Read equation OOXML for the whole document                                                           |
| `word_read_selection_paragraph`                        | W     |   | Read the paragraph around the selection and any equation OOXML in it                                 |
| `word_export_file`                                     | W     |   | Export docx / pdf (base64)                                                                           |
| `word_insert_image`                                    | W     | D | Insert an image before / after / replacing the selection                                             |
| `word_transfer_content`                                | W     | D | Move OOXML between selection, body, tables, search hits, headers and footers                         |
| `apply_style`                                          | W     | D | Apply a built-in style (selection or all)                                                            |
| `format_selection`                                     | W/E   | D | Format the selection (font, size, bold, italic, colour, highlight)                                   |
| `set_paragraph_format`                                 | W     | D | Paragraph format (alignment, indent, line spacing, list)                                             |
| `search`                                               | W     |   | Find text, returns the hits                                                                          |
| `add_comment` / `read_comments`                        | W/E   | D | Add / read comments                                                                                  |
| `read_properties`                                      | W/E   |   | Document properties (title, author, word count…)                                                     |
| `word_list_sections`                                   | W     |   | List sections                                                                                        |
| `word_read_headers_footers` / `word_set_header_footer` | W     | D | Read / write headers and footers                                                                     |
| `word_insert_break`                                    | W     | D | Insert a page or section break                                                                       |
| `word_insert_caption`                                  | W     | D | Insert a caption (static numbering verified; dynamic `SEQ` must be opted into)                       |
| `locate_select`                                        | W/E   |   | Locate and select (bookmark, named range, text, Excel range) — side-effect free                      |

## Excel

| action                                                                                                            | hosts | D | Purpose                                                                                                   |
| ----------------------------------------------------------------------------------------------------------------- | ----- | - | --------------------------------------------------------------------------------------------------------- |
| `list_sheets`                                                                                                     | E     |   | List worksheets (name, position, visibility)                                                              |
| `read_range`                                                                                                      | E     |   | Read a range (values, formulas, number formats)                                                           |
| `write_range`                                                                                                     | E     | D | Batch write `values` or `formulas` (chunked above 5000 cells)                                             |
| `format_range`                                                                                                    | E     | D | Format a range and read the resulting format back                                                         |
| `insert_chart`                                                                                                    | E     | D | Data → chart (Column / Line / Bar / Pie / Area / Scatter…)                                                |
| `add_sheet` / `rename_sheet` / `delete_sheet`                                                                     | E     | D | Worksheet management (`delete_sheet` supports `dryRun`)                                                   |
| `apply_sort`                                                                                                      | E     | D | Multi-field sort                                                                                          |
| `apply_filter`                                                                                                    | E     | D | AutoFilter                                                                                                |
| `evaluate_formula`                                                                                                | E     |   | Evaluate expressions such as `SUM(A1:A5, C1:C5)` (whitelist: SUM / AVERAGE / COUNT / MAX / MIN / PRODUCT) |
| `list_tables` / `create_table`                                                                                    | E     | D | Enumerate / create native table objects                                                                   |
| `delete_table` / `convert_table_to_range`                                                                         | E     | D | Delete the table and its data / drop the table object but keep the data                                   |
| `read_conditional_formats` / `add_conditional_format` / `delete_conditional_format` / `clear_conditional_formats` | E     | D | Conditional-format lifecycle (ExcelApi 1.6+)                                                              |
| `read_data_validation` / `set_data_validation` / `clear_data_validation` / `validate_data`                        | E     | D | Data-validation lifecycle (1.8+; `validate_data` details need 1.9+)                                       |

## PowerPoint

| action                                                       | hosts | D | Purpose                                                                          |
| ------------------------------------------------------------ | ----- | - | -------------------------------------------------------------------------------- |
| `read_slides`                                                | P     |   | Read the currently selected slide range (use `read_document` for the whole deck) |
| `ppt_read_notes`                                             | P     |   | Speaker notes for the whole file (parses OOXML notesSlides)                      |
| `ppt_inspect_slides`                                         | P     |   | Structure snapshot: shape ID, type, name, geometry, z-order, text                |
| `ppt_list_placeholders`                                      | P     |   | Placeholder types and bounds                                                     |
| `ppt_add_slides` / `ppt_add_elements`                        | P     | D | Batch-create slides, then add editable elements by the returned `slideId`        |
| `ppt_delete_slides` / `ppt_delete_shapes`                    | P     | D | Delete by explicit ID (refuses to empty the deck)                                |
| `ppt_insert_image`                                           | P     | D | Insert an image (positioned, or into the selection / a placeholder)              |
| `ppt_insert_table`                                           | P     | D | Native editable table via PowerPointApi 1.8+                                     |
| `ppt_group_shapes` / `ppt_ungroup_shape` / `ppt_set_z_order` | P     | D | Group / ungroup / z-order                                                        |
| `ppt_align_shapes` / `ppt_distribute_shapes`                 | P     | D | Align / distribute using deterministic geometry over shape bounds                |
| `ppt_render_slides`                                          | P     |   | Render slides to PNG                                                             |
| `ppt_export_file`                                            | P     |   | Native Office export                                                             |

## Boundaries

Some things Office.js simply does not offer. The plugin returns an error code instead of pretending to succeed:

- **No API exists**: creating native charts in PowerPoint, image cropping, inserting into an arbitrary placeholder ID; cross-platform "Link to Previous" and clipboard moves in Word
- **Missing in some runtimes**: Word's `body.comments` and `paragraphFormat` do not exist on every build (the plugin returns `requirement`, it does not fake success)
- **Known API traps**: Excel has no `workbook.getRange` / `calculate` (the official replacements are used); `Range.autoFilter` must go through the worksheet
- **shapeType values are Office's, not ours**: the plugin does not guess the enum and accepts any PascalCase name; when execution fails it appends the list of values verified against the add-in, so the caller does not get a bare `InvalidArgument`

## Adding your own actions

Beyond the built-in list you can add custom actions or override built-in behaviour in `server/user-actions.js` — see the [development guide](DEVELOPMENT.en.md#customising-user-actionsjs).
