# Full action reference

> English · [中文](CAPABILITIES.md)

Every action this plugin currently supports. **The authoritative list is the live one** — fetch it any time:

```bash
curl "$(jq -r .url "$DSH_HOME/dsh-ms-office/bridge.json")/office/capabilities?lang=en"
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

| action | Hosts | D | Purpose | Requirement | Arguments |
| --- | --- | --- | --- | --- | --- |
| `read_selection` | W/E/P |  |  |  | `withStyles`: boolean, optional: also return styling |
| `write_selection` | W/E/P | D |  |  | `text`: string, required: text that replaces the selection |
| `read_document` | W/E/P |  |  |  | `sheet`: string, optional (Excel): worksheet name |
| `read_styles` | W/E |  |  |  |  |
| `replace_all` | W/E | D |  |  | `search`: string, required<br>`replace`: string, required<br>`dryRun`: boolean, optional: only return the hit count, do not execute |
| `format_selection` | W/E | D |  |  | `font`: string, optional<br>`size`: number, optional<br>`bold`: boolean, optional<br>`italic`: boolean, optional<br>`color`: string, optional (#RRGGBB)<br>`highlight`: string, optional (Word) |
| `add_comment` | W/E | D |  |  | `text`: string, required<br>`selectionText`: string, optional (Word): comment on the selection containing this text<br>`cell`: string, optional (Excel): A1 cell address |
| `read_comments` | W/E |  |  |  |  |
| `read_properties` | W/E |  |  |  |  |
| `get_environment` | W/E/P |  | Read Office environment, native API support and confirmation mode |  | `syncConfig`: optional boolean; sync pane to existing user-selected server confirmation mode; no mode/approval input<br>`reloadPane`: optional boolean; reload this pane only after public command status is explicitly empty; leaves the document intact |
| `get_document_info` | W/E/P |  |  |  | `keys`: array, optional: settings keys to read (e.g. ["lang","theme"]) |
| `set_setting` | W/E/P | D |  |  | `key`: string, required: the settings key<br>`value`: any, required: any JSON value (persisted with the document, readable across sessions) |
| `locate_select` | W/E |  |  |  | `bookmark`: string, optional: Word bookmark / Excel named range<br>`anchor`: string, optional: same as bookmark<br>`text`: string, optional: locate the first matching text<br>`range`: string, optional (Excel): A1:B5 or Sheet!A1:B5<br>`address`: string, optional (Excel): same as range<br>`sheet`: string, optional (Excel): the worksheet to use<br>`blinks`: number, optional: number of blinks (0 = select and hold; 1-5 blinks)<br>`interval`: number, optional: blink interval in ms (default 300, 100-800) |

## Word

| action | Hosts | D | Purpose | Requirement | Arguments |
| --- | --- | --- | --- | --- | --- |
| `append_text` | W | D |  |  | `text`: string, required: paragraph appended at the end of the document |
| `read_tables` | W |  |  |  |  |
| `read_table_borders` | W |  |  | WordApi 1.3+ |  |
| `set_table_borders` | W | D |  | WordApi 1.3+ | `color`: string, required: #RRGGBB<br>`tableIndexes`: number[], optional: 1-based, defaults to all tables<br>`locations`: array of Top\|Bottom\|Left\|Right\|InsideHorizontal\|InsideVertical, optional<br>`enableGrid`: boolean, optional: apply TableGrid first so the full grid exists<br>`width`: Optional positive width in pt; omitted preserves existing width |
| `set_font` | W | D |  |  | `font`: string, required: font family name |
| `word_list_paragraphs` | W |  |  |  | `filter`: all\|empty\|nonEmpty, optional, defaults to all |
| `word_delete_paragraphs` | W | D |  |  | `targets`: array, required: [{index,expectedText}], index is 0-based |
| `insert_paragraph` | W | D |  |  | `text`: string, required<br>`style`: string, optional: built-in style such as Heading1<br>`location`: string, optional: end\|start\|afterSelection\|afterSelectionParagraph |
| `insert_table` | W | D |  | WordApi 1.3+ | `rows`: array, required: rectangular 2-D array, at most 1000 rows / 63 columns / 10000 cells<br>`location`: string, optional: start\|end\|beforeSelection\|afterSelection<br>`style`: string, optional: localized/custom table style name only<br>`styleBuiltIn`: string, optional: language-neutral built-in style enum, e.g. GridTable4_Accent1; cannot be combined with style<br>`headerRow`: boolean, optional: treat the first row as a header row |
| `word_insert_html` | W | D |  | WordApi 1.1+ | `html`: string, required: 8 MiB max<br>`target`: body\|selection, optional, defaults to body<br>`mode`: replace\|append\|prepend\|before\|after, optional, defaults to append |
| `word_set_page_setup` | W | D |  | WordApiDesktop 1.3+ | `sections`: all\|number[], optional: 0-based, defaults to all<br>`topMargin`: number, optional (points)<br>`bottomMargin`: number, optional (points)<br>`leftMargin`: number, optional (points)<br>`rightMargin`: number, optional (points)<br>`headerDistance`: number, optional (points)<br>`footerDistance`: number, optional (points)<br>`pageWidth`: optional positive pt<br>`pageHeight`: optional positive pt<br>`differentFirstPageHeaderFooter`: optional boolean |
| `word_insert_equation` | W | D |  | WordApi 1.1+ Range.insertOoxml | `expression`: structured equation node: text\|group\|fraction\|subscript\|superscript<br>`target`: object, optional: {kind:selection\|body\|search,query/occurrence/matchCase/matchWholeWord}; occurrence is 0-based<br>`mode`: replace\|before\|after\|prepend\|append, optional, defaults to after<br>`display`: boolean, optional: defaults to true |
| `word_read_equations` | W |  |  | WordApi 1.1+ Range.getOoxml |  |
| `word_read_selection_paragraph` | W |  |  | WordApi 1.1+ Range.getOoxml |  |
| `word_export_file` | W |  |  |  | `type`: docx\|pdf, defaults to docx |
| `word_insert_image` | W | D |  |  | `base64`: string, required: PNG/JPEG base64<br>`mode`: before\|after\|replace, optional, defaults to after<br>`width`: number, optional (px)<br>`height`: number, optional (px) |
| `word_transfer_content` | W | D |  | WordApi 1.1+ | `source`: object, required: {kind:selection\|table\|search,index/query/occurrence/matchCase}<br>`destination`: object, required: {kind:selection\|body\|header\|footer,section/type}<br>`operation`: copy\|move, optional, defaults to copy<br>`mode`: replace\|append\|prepend\|before\|after, optional, defaults to append<br>`expectedText`: string, optional: verify the source text before writing<br>`overwrite`: boolean, optional: must be true to replace a non-empty target |
| `apply_style` | W | D |  |  | `style`: string, required: Heading1 / Title / Normal, …<br>`scope`: string, optional: selection\|all |
| `set_paragraph_format` | W | D |  |  | `alignment`: string, optional: left\|center\|right\|justify<br>`indent`: number, optional (cm)<br>`lineSpacing`: Optional native numeric line spacing; Word UI divides by12 (18 =>1.5)<br>`listType`: none\|numbered\|bulleted; WordApi1.3; existing lists must be explicitly removed before conversion |
| `search` | W |  |  |  | `query`: string, required<br>`matchCase`: boolean, optional<br>`wildcard`: boolean, optional: use wildcards |
| `word_list_sections` | W |  |  | WordApi 1.1+ | `textLimit`: number, optional |
| `word_read_headers_footers` | W |  |  | WordApi 1.1+ | `section`: number\|all<br>`types`: array of primary\|firstPage\|evenPages<br>`includeFields`: optional boolean; WordApi1.5; read actual native field type/code/text |
| `word_set_header_footer` | W | D |  | WordApi 1.1+ | `kind`: header\|footer<br>`section`: number\|all<br>`type`: primary\|firstPage\|evenPages<br>`text`: optional string; required unless field is provided<br>`mode`: replace\|append\|prepend<br>`field`: optional {type:TOC\|Page\|NumPages,switches?}; mutually exclusive with text; requires WordApi1.5 |
| `word_insert_break` | W | D |  | WordApi 1.1+ | `type`: page\|line\|sectionNext\|sectionContinuous\|sectionEven\|sectionOdd<br>`location`: start\|end\|beforeSelection\|afterSelection<br>`paragraph`: optional {index:non-negative integer,expectedText:string}; exact native paragraph target |
| `word_insert_caption` | W | D | Static numbering needs WordApi 1.1+; a dynamic SEQ field needs WordApi 1.5+ (unavailable on this machine) | WordApi 1.1+ | `label`: string<br>`title`: string<br>`sequenceId`: ASCII identifier<br>`dynamic`: boolean, optional, defaults to false; true attempts a dynamic SEQ field<br>`number`: number, optional: static numbering starts at 1<br>`position`: beforeSelection\|afterSelection<br>`separator`: string, optional |
| `word_format_paragraphs` | W | D | Format paragraphs by exact index and complete text using built-in or native atlas body styles | WordApi 1.3+ | `targets`: required1..500 exact targets[{index,expectedText,style?,font?:{name,size,color,bold,italic},alignment?,spaceBefore?,spaceAfter?,lineSpacing?,leftIndent?,rightIndent?,firstLineIndent?}]; unique zero-based indexes and complete-text guards preflight every target. Supports built-in or atlas body styles. Dimensions are points; numeric lineSpacing18 is not true1.5 spacing |
| `word_read_styles` | W |  | Read actual global paragraph style definitions and script-specific fonts | WordApi 1.5+ | `names`: string[] optional; exact localized style names; omit to list paragraph styles |
| `word_set_styles` | W | D | Set actual global paragraph style definitions; numeric lineSpacing is points, not a multiple | WordApi 1.5+ | `styles`: required style array{name,font?,paragraphFormat?}; or the single reserved body-1.5 atlas preset for a global SimSun12pt/24pt-indent/true1.5 style. Its exact localized name is documented in the style skill. Numeric lineSpacing is points, not a multiple |
| `word_read_fields` | W |  | Read actual Word field codes and cached results | WordApi 1.5+ | `target`: body\|selection; default body |
| `word_insert_field` | W | D | Insert and read back real TOC or PAGE fields | WordApi 1.5+ | `type`: TOC\|Page\|NumPages<br>`switches`: optional single-line field switches<br>`target`: body\|selection; default selection<br>`mode`: replace\|prepend\|append\|before\|after; body accepts only replace/prepend/append<br>`paragraph`: TOC-only optional{index,expectedText}; zero-based body paragraph with exact complete-text guard. Uses the available native Desktop1.4 contents collection with default heading levels1–3; custom switches are rejected before mutation. Reads back actual generated entries |
| `word_create_document` | W | D | Natively create and open a blank Word document; no file/Base64 input; original content preserved | WordApi 1.3+ |  |
| `word_render_pages` | W |  | Read-only native Word full-page PNGs including headers/footers; Windows desktop Word PrintLayout required | WordApi 1.1+; Windows desktop Word native PrintLayout | `expectedPageCount`: required integer 1..10; expected native page count |

## Excel

| action | Hosts | D | Purpose | Requirement | Arguments |
| --- | --- | --- | --- | --- | --- |
| `list_sheets` | E |  | Read all worksheet names, positions and visibility |  |  |
| `read_range` | E |  |  |  | `address`: string, required: Sheet!A1:B10<br>`limit`: number, optional: positive integer maximum cells to load; default and hard cap are 5000<br>`withLayout`: Optional boolean: read actual wrap/alignment and supported point dimensions (mixed values remain null) |
| `write_range` | E | D |  |  | `address`: string, required<br>`values`: array, optional: 2-D array<br>`formulas`: array, optional: 2-D array (A1 formulas) |
| `format_range` | E | D |  | ExcelApi 1.1+; column/row dimensions and autofit require ExcelApi 1.2+ | `address`: string, required<br>`font`: string, optional<br>`size`: number, optional<br>`bold`: boolean, optional<br>`fill`: string, optional (#RRGGBB)<br>`numberFormat`: string, optional<br>`autoFit`: Optional boolean: compatible alias for autoFitColumns; ExcelApi 1.2+<br>`tableStyle`: string, optional (e.g. TableStyleMedium2)<br>`tableName`: required when tableStyle is set<br>`columnWidth`: Optional positive width in pt; mutually exclusive with autoFitColumns:true or autoFit:true<br>`rowHeight`: Optional positive height in pt; mutually exclusive with autoFitRows:true<br>`wrapText`: Optional boolean<br>`horizontalAlignment`: General\|Left\|Center\|Right\|Fill\|Justify\|CenterAcrossSelection\|Distributed<br>`verticalAlignment`: Top\|Center\|Bottom\|Justify\|Distributed<br>`autoFitRows`: Optional boolean: autofit row height; ExcelApi1.2<br>`autoFitColumns`: Optional boolean: autofit columns; ExcelApi 1.2+; cannot combine true with columnWidth. autoFit is a compatible alias.<br>`color`: optional text color #RRGGBB |
| `insert_chart` | E | D |  |  | `type`: string, required: Line\|Column\|Bar\|Pie\|Area\|Scatter, …<br>`dataRange`: string, required: Sheet!A1:B10<br>`title`: string, optional<br>`position`: Optional {startCell,endCell}: local A1 corners on the dataRange worksheet. Always creates on the source worksheet, never an unrelated active sheet<br>`font`: optional font family; Chinese default SimSun<br>`seriesColors`: optional #RRGGBB[]<br>`fill`: optional chart background #RRGGBB<br>`color`: optional chart text #RRGGBB |
| `add_sheet` | E | D |  |  | `name`: string, required |
| `rename_sheet` | E | D |  |  | `oldName`: string, required<br>`newName`: string, required |
| `delete_sheet` | E | D |  |  | `name`: string, required<br>`dryRun`: boolean, optional |
| `apply_sort` | E | D |  |  | `address`: string, required<br>`fields`: array, required: [{column:1,ascending:true}] |
| `apply_filter` | E | D |  |  | `address`: string, required<br>`columns`: array, optional: [{column:1,filterType:'Values',criteria:{values:['a']}}]; column is 1-based, Values also accepts criteria:['a']; Custom uses criteria:{criterion1:'>2'} |
| `evaluate_formula` | E |  |  |  | `formula`: string, required: e.g. SUM(A1:A10) |
| `list_tables` | E |  |  | ExcelApi 1.1+ | `sheet`: string, optional<br>`includeValues`: boolean, optional<br>`limit`: number, optional |
| `create_table` | E | D |  | ExcelApi 1.1+ | `address`: string<br>`hasHeaders`: boolean<br>`name`: string, optional<br>`style`: string, optional |
| `delete_table` | E | D |  | ExcelApi 1.1+ | `name`: string |
| `convert_table_to_range` | E | D |  | ExcelApi 1.2+ | `name`: string |
| `read_conditional_formats` | E |  |  | ExcelApi 1.6+ | `address`: string |
| `add_conditional_format` | E | D |  | ExcelApi 1.6+ | `address`: string<br>`type`: Custom\|CellValue\|ContainsText<br>`rule`: object<br>`format`: object, optional |
| `delete_conditional_format` | E | D |  | ExcelApi 1.6+ | `address`: string<br>`id`: string |
| `clear_conditional_formats` | E | D |  | ExcelApi 1.6+ | `address`: string |
| `read_data_validation` | E |  |  | ExcelApi 1.8+ | `address`: string |
| `set_data_validation` | E | D |  | ExcelApi 1.8+ | `address`: string<br>`rule`: object, a single validation type<br>`ignoreBlanks`: boolean, optional<br>`errorAlert`: object, optional<br>`prompt`: object, optional |
| `clear_data_validation` | E | D |  | ExcelApi 1.8+ | `address`: string |
| `validate_data` | E |  | details needs ExcelApi 1.9+ | ExcelApi 1.8+ | `address`: string<br>`details`: boolean, optional |
| `excel_set_freeze_panes` | E | D |  | ExcelApi 1.7+ | `sheet`: Required worksheet name<br>`mode`: Required rows\|columns\|range\|none<br>`count`: Positive integer for rows/columns only<br>`range`: For range mode only: local A1 range to freeze, e.g. A1:B2; not a cursor anchor |
| `excel_render_range` | E |  |  | ExcelApi 1.7+ | `address`: Required range address; render the whole range as PNG for visual review<br>`maxCells`: Optional integer 1..10000, default5000; reject over-budget ranges without truncating |
| `excel_export_file` | E |  | Export using the native Office File API; actual host support is required | Common File API; actual host support is checked | `type`: xlsx\|pdf; default xlsx |
| `excel_render_chart` | E |  | Native Office Chart.getImage chart preview | ExcelApi 1.9+ | `sheet`: string required<br>`name`: string required |
| `excel_set_print_layout` | E | D | Set bounded native worksheet print area and page layout with readback | ExcelApi 1.9+ | `sheet`: string required<br>`address`: sheet-local A1 rectangle; at most 10000 cells<br>`orientation`: Landscape\|Portrait; default Landscape<br>`paperSize`: A4\|Letter\|A3; default A4<br>`fitToPagesWide`: integer1..100; default1<br>`fitToPagesTall`: integer1..100; default1 |

## PowerPoint

| action | Hosts | D | Purpose | Requirement | Arguments |
| --- | --- | --- | --- | --- | --- |
| `read_slides` | P |  | Reads the currently selected slide range, not the whole deck; use read_document for the whole presentation |  |  |
| `ppt_read_notes` | P |  |  |  |  |
| `ppt_add_slides` | P | D |  | PowerPointApi 1.4+ | `count`: number, optional: 1-50, defaults to 1<br>`clearPlaceholders`: boolean, optional: defaults to false |
| `ppt_add_elements` | P | D |  | PowerPointApi 1.4+ | `slideId`: string, required: target slide ID<br>`coordinateSpace`: object, required: element design coordinates {width,height}<br>`targetSize`: object, required: real target size in points {width,height}<br>`elements`: array, required: text\|shape\|line elements; shared x/y/w/h; text supports font/size/color/bold/italic/align/valign/margins/autoFit/runs, shape supports shapeType/fill/lineColor, line supports connectorType/color/lineWidth/dashStyle |
| `ppt_delete_slides` | P | D |  | PowerPointApi 1.2+ | `slideIds`: string[], required: exact slide IDs; refuses to delete every slide |
| `ppt_delete_shapes` | P | D |  | PowerPointApi 1.4+ | `slideId`: string, required<br>`shapeIds`: string[], required: exact shape IDs; returns the remaining IDs after deletion |
| `ppt_render_slides` | P |  |  | PowerPointApi 1.8+ | `width`: number, optional: PNG width, defaults to 960<br>`height`: number, optional: PNG height, defaults to 540<br>`slideIds`: Optional unique nonempty native slide ID array; omitted means all pages. Output preserves document order and original 1-based indexes; all IDs checked before rendering |
| `ppt_export_file` | P |  |  |  | `type`: pptx\|pdf; default pptx; native Office export |
| `ppt_inspect_slides` | P |  |  | PowerPointApi 1.4+ | `slideIds`: string[], optional<br>`withStyles`: Optional boolean: read text/font/fill when supported, otherwise report stylesRead:false and reason |
| `ppt_list_placeholders` | P |  |  | PowerPointApi 1.8+ | `slideId`: string, optional |
| `ppt_insert_image` | P | D |  | ImageCoercion 1.1+ | `base64`: string<br>`mode`: positioned\|selectionOrPlaceholder<br>`left`: number<br>`top`: number<br>`width`: number, optional<br>`height`: number, optional<br>`slideId`: optional string: select exact native slide before inserting image |
| `ppt_insert_table` | P | D | style needs PowerPointApi 1.9+ | PowerPointApi 1.8+ | `slideId`: string<br>`values`: rectangular 2-D array<br>`left`: number, optional<br>`top`: number, optional<br>`width`: number, optional<br>`height`: number, optional<br>`style`: PowerPoint.TableStyle enum string, optional, e.g. MediumStyle2Accent1 |
| `ppt_group_shapes` | P | D |  | PowerPointApi 1.8+ | `slideId`: string<br>`shapeIds`: string[]<br>`name`: string, optional |
| `ppt_ungroup_shape` | P | D |  | PowerPointApi 1.8+ | `slideId`: string<br>`groupShapeId`: string |
| `ppt_set_z_order` | P | D |  | PowerPointApi 1.8+ | `slideId`: string<br>`shapeIds`: string[]<br>`operation`: BringForward\|BringToFront\|SendBackward\|SendToBack |
| `ppt_align_shapes` | P | D |  | PowerPointApi 1.4+ | `slideId`: string<br>`shapeIds`: string[]<br>`mode`: left\|center\|right\|top\|middle\|bottom |
| `ppt_distribute_shapes` | P | D |  | PowerPointApi 1.4+ | `slideId`: string<br>`shapeIds`: string[]<br>`mode`: horizontal\|vertical |
| `ppt_update_shapes` | P | D |  | PowerPointApi 1.4+ | `slideId`: Required native slide ID<br>`updates`: Required 1..500 patches: {shapeId,text?,font?:{name,size,color,bold,italic},fill?:{color,transparency},left?,top?,width?,height?}; geometry/font size in pt. Only supplied fields change. text replaces the whole text, not runs; font/text/fill limited to TextBox/GeometricShape. |
| `ppt_read_presentation_info` | P |  |  | Common API getFileAsync (Compressed) |  |

## Boundaries

Some things Office.js simply does not offer. The plugin returns an error code instead of pretending to succeed:

- **No API exists**: creating native charts in PowerPoint, image cropping, inserting into an arbitrary placeholder ID; cross-platform "Link to Previous" and clipboard moves in Word
- **Missing in some runtimes**: Word's `body.comments` is unavailable on some builds; paragraph formatting uses Paragraph properties, built-in styles need1.3 and page setup needsDesktop1.3
- **Known API traps**: Excel has no `workbook.getRange` / `calculate` (the official replacements are used); `Range.autoFilter` must go through the worksheet
- **shapeType values are Office's, not ours**: the plugin does not guess the enum and accepts any PascalCase name; when execution fails it appends the list of values verified against the add-in, so the caller does not get a bare `InvalidArgument`

## Adding your own actions

Beyond the built-in list you can add custom actions or override built-in behaviour in `server/user-actions.js` — see the [development guide](DEVELOPMENT.en.md#customising-user-actionsjs).
