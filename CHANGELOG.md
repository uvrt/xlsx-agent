# Changelog

xlsx-agent has not been released to PyPI; it is installed from source. Each version that
reaches `main` is tagged `v<version>`, from 0.1.1, the first tagged. [`ROADMAP.md`](ROADMAP.md)
has the detail of every milestone. Why each dependency is needed at its version:
[Dependencies](#dependencies), at the end.

## Unreleased

- The loaded core's online budget is now 6,800 tokens as count-tokens counts them (raised
  from 6,500 by the user on 2026-11-05, keeping `read_batch` in the core): round 32, R15
  took the core 6,431 -> 6,678 counted online (claude-sonnet-5-5). Tests only; no tool
  definition changed.
- The online budget is now 7,000 (the user, 2026-11-06: the budget is guidance, clear tools
  first), and `xlsx_derive`'s descriptions that 0.1.16's trim left unclear are restored:
  `join.on` says its first name is the key column of the outer `from` (the rows joined to)
  and the second of this join's table, `join.from` that it is the table to join;
  `pivot_values` reads as sentences again (a listed value without rows still gets its
  column; rows of a value not listed are left out and named in `pivot_unlisted`); `coerce`
  says "Opt in:" again. Also made clear: `fill`'s `from` is the fill's own column, read in
  the same row; `join` adds to each row of the outer `from`; `join.columns` are this join's
  table's; `xlsx_compare`'s, `xlsx_validate_data`'s and `xlsx_query`'s `cursor` is passed back with
  "the same call" (none has a target). Descriptions, tests and docs only; no behaviour
  changed, so no version bump (`tools/check_version.py` asks for none). `xlsx_derive`
  1,182 -> 1,245 estimated, the core 4,783 -> 4,844 (about 6,866 online), every definition
  5,754 -> 5,815 (guards 4,785 -> 4,845 and 5,755 -> 5,820); a reviewer's core 4,629 ->
  4,690.

## 0.1.26 (unreleased) -- AutoFilters, hidden rows and columns, pivot refresh on open

- **AutoFilters** (production round 44, lane A, R55): `wb.set_autofilter(target, ref=None,
  criteria={column: values}, hide=True)` and `wb.clear_autofilter(target, keep=False)`
  (`xlsx_agent.edit.filters`). A sheet's filter (one per sheet: `"Stock"` with `ref`, or
  `"Stock!A1:F8"`) or a table's own (`"t:Orders"`); value lists by header text or column
  letters, `None` for blank cells. Written as Excel for Mac 16.106 wrote the same filters
  (measured through its object model, 2026-11-08): `autoFilter` with `filterColumn`/`filters`
  (values sorted as text, `blank="1"`), `sheetPr filterMode="1"` while there are criteria,
  the hidden `_xlnm._FilterDatabase` name local to the sheet; a table's filter in its part,
  with no name nor `filterMode`. Matching as Excel's: text without regard to case, a number
  by its shown text in the workbook's locale, a formula by its cached value. **Excel does not
  apply a saved filter when it opens a file** (measured, sheet and table filters), so the rows
  a filter excludes are hidden here and the result says so (`hidden_by: "xlsx-agent"`); the
  range's other rows are shown, rows hidden by hand among them, as Excel shows them.
  Refused: a date, boolean or error cell in a filtered column (`filter_unmeasured`: Excel's
  text for them is the region's and language's, not measured), a formula column once any cell
  was written or the file asks to be recalculated (`filter_formulas`); `hide=False` writes the
  criteria and leaves hiding to Excel. Turning a filter off shows the range's rows and drops
  `autoFilter` and `filterMode` (the name stays, as Excel leaves it); `keep=True` is Excel's
  Clear. Hiding or showing rows asks for a recalculation on open when a `SUBTOTAL`,
  `AGGREGATE` or table totals row could read them.
- **Rows and columns hidden by hand**: `wb.hide_rows(sheet, rows)`, `show_rows`,
  `hide_columns(sheet, columns)`, `show_columns` (`5`, `"5:7"`, `"C:D"` or lists); a row made
  for the flag is removed again when shown, a column keeps no width hiding gave it.
- **Pivot caches refreshed on open, opt-in** (R61): `wb.refresh_pivots_on_load(True|False,
  pivots=None)` (`xlsx_agent.edit.pivots`) sets or takes away `refreshOnLoad` on the cache
  definitions (all, or those of the pivot tables or cache ids named) as a targeted edit of
  the part's bytes: every other byte is kept. Measured: Excel for Mac kept the flag but did
  **not** refresh the pivot on open (object-model open, Launch Services open, 8 s waited,
  records left out); a refresh by hand did. Excel for Windows not measured.
- New edit codes: `filter_range`, `filter_column`, `filter_table`, `filter_unmeasured`,
  `filter_formulas`, `no_pivot`, `pivot_part`. No agent tool (the core is at its budget).
  The oracle's `Session.operate` takes another script (`tools/excel_filter.applescript`).

## 0.1.25 (unreleased) -- Aptos Narrow fitted, one sheet-name rule, dates into General cells

- **Fitted widths in Aptos Narrow**, Excel's default body face (production round 44, wave
  2): a fit refused it (`unmeasured_font`). It is measured now with its own advance widths,
  which ooxml-common 0.42.0 adds (`own_metrics`, generated from the face Office installs,
  as Aptos's are; no font file anywhere). Fonts are looked up with `own_metrics`, the very
  face's widths: Calibri Light and Aptos Light, which the renderers measure with a near
  face's widths, now refuse a fit instead of being fitted 1.3% or more off. A workbook
  whose default font has no widths at all (Rockwell) refuses a column width
  (`unmeasured_font`: the width's unit, its widest digit, is unknown) where it failed with
  an `AssertionError`.
- **The fit's padding is the cell font's** (`layout.fit_padding`): 5 px, and 2 more for
  every 4 px of the font's own widest digit past the first -- 7 for Calibri, Aptos and
  Aptos Narrow at 9 to 14 points, as before; 5 at 6 to 8 points, 9 at 18 to 24 (Calibri),
  11 at 28, 13 at 36. Measured (Excel for Mac 16.106, 2026-11-09): Calibri and Aptos at
  6 to 36 points in Calibri 11 and Aptos 11 workbooks, 578 columns, all exact (were 2 to
  6 px off below 9 and above 16 points); Aptos Narrow regular at 8 to 24 points in Aptos
  Narrow 11 and 12 workbooks, and bold Aptos Narrow, Aptos and Calibri at 9 to 18 points:
  exact but for a word with accented letters (1 px) and Aptos Narrow bold 12, which Excel
  measures unrounded (3 px on "Stock level by warehouse", 14 px on 40 letters).
- **`Workbook.new()` names sheets by the measured rules** that `add_sheet` and
  `rename_sheet` apply (`edit.sheets.name_problems`, now with `taken` for a workbook not
  made yet): control characters and text reading as an `_xHHHH_` escape refused, names
  compared as Excel compares them (`Ａ1` is `A1`), a lone space and `History` taken.
  `blank.sheet_name_problems` keeps its name and calls it.
- **A date into a General cell gets the short date** (`set_cells`, and so the tool): a new
  cell, a General one, or one of a built-in format Excel replaces for a typed date --
  currency and accounting (5-8, 44), percentages (9, 10), scientific (11, 48), fractions
  (12, 13), times (18-21, 45-47) -- is given `numFmtId` 14, the rest of its format kept,
  and listed in the result's `short_date`; a number format (1-4, 37-40), text (49) or a
  custom code that shows no day is still refused, saying that the cell would show the
  date's serial number. Measured: "2026-01-05" and "1/5/2026" typed into cells of 46
  formats (Excel for Mac 16.106, 2026-11-09). `xlsx_set_cells`' description says so:
  519 -> 525 estimated tokens (deferred group; the core unchanged at 4,938; every
  definition 7,076 -> 7,082).
- Needs ooxml-common 0.42.0 (`own_metrics`, Aptos Narrow's table).

## 0.1.24 (unreleased) -- edits on large sheets pay for what they touch

- **Each edit step pays for what it touched** (production round 44, lane P: on a sheet of
  50,000 rows by 4 columns with 5,000 formulas, writing one cell took 12 s, every step
  10-36 s). Nothing changes what is written: the bytes after every step, and after undo,
  are those 0.1.23 writes (compared on the test workbooks and the large sheet).
  - The step's structure check (`check_structure(rows=)`) reads, of a sheet whose cells
    alone were written, the cells of the rows written and of its tables' header rows;
    every row is still checked as a row, the `dimension` against the cells read. A sheet
    changed otherwise (formatting, layout, rows and columns) is checked whole, from the
    tree the writer holds, in one in-line pass. `wb.validate()` stays the whole check
    (and is 3-4x faster); an issue in a row a step did not write is validate's to find.
  - The guards read merges, array formulas and validation rules from the sheet's tree
    (`Workbook._layout`), and `write_table`/`append_rows` a block's cells from it
    (`xlsx_agent.edit.content`), typing only the cells inside the block: no step streams
    the sheet to type its values. A sheet's protection is read from its tree once parsed.
  - The parts a step changed come from the history's own states (two fewer copies of
    every edited part per step); the save's read-back types only the cells written.
  - Formulas found directly when rows or columns move (`restructure._formula_cells`).
  - Measured (best of 3, CPU s, before -> after, this machine; `tools/bench_edits.py`):
    one cell 12.2 -> 1.5 (the next steps about 0.5 each), 1,000 cells 13.1 -> 1.4,
    append 1,000 rows 32.7 -> 2.6, write_table 1,000 rows 10.9 -> 1.6, format a column
    12.3 -> 3.4, insert/delete/move rows 10.1-10.7 -> 3.2-3.3, five one-cell steps
    51.5 -> 3.4, `validate()` 5.4 -> 1.6, save 7.4 -> 2.8; undo 0.14, `to_bytes` 0.3
    unchanged; peak memory unchanged (280-340 MB, the parsed sheet).
  - With ooxml-edit 0.95.1 (its PR #112: the shared strings' index no longer counts the
    table at every new string, columns looked up, the dimension measured in line); this
    version works with 0.95.0, more slowly where strings are added.

## 0.1.23 (unreleased) -- formatting and layout from code

- **`wb.format(target, font=, fill=, border=, alignment=, number_format=, protection=,
  like=, style=)`** (production round 44, R53; E7): cells, whole columns or whole rows
  formatted in one step (`xlsx_agent.edit.format`). Only what is asked changes (bold keeps
  the fill, the number format, the borders); a font, fill, border, number format or cell
  format equal to one the workbook has is reused, never copied per cell. What is written is
  what Excel writes (measured, Excel for Mac 16.106, Excel formatting cells itself): a
  font's parts in Excel's order, a solid fill with `bgColor indexed="64"`, a border's
  automatic colour `indexed="64"`, a code that is a built-in format's by its id (`0.00` is
  2), an indent making a `general` cell `left`, and each `apply*` flag `1` exactly where
  that part differs from the cell style the format is based on. `style="Heading 1"` applies
  a named cell style of the workbook as Excel does (measured on five styles: the parts the
  style includes are its, the others the cell's), a built-in one also by its English name
  (a Dutch Excel saves it as `Kop 1`); `like="Stock!D2"` gives the cells another cell's
  whole format, from any sheet. Whole columns and rows as Excel formats them (measured: a
  column's style, a row's `customFormat`, each cell held, none made).
- **`wb.set_column_widths(target, width | "fit" | None)`** (`xlsx_agent.edit.layout`): a
  width in characters stored as Excel stores one typed in (measured: 12 is
  `12.83203125`), and `"fit"`: Excel's AutoFit, measured on 258 columns -- each character's
  advance width at the font's size rounded to whole pixels, one more pixel per five widest
  digits past fifteen, 7 px of padding, over the default font's widest digit
  (`xlsx_agent.edit.fit`). Exact on every probe column in Calibri, Arial, Times New Roman,
  Courier New, Cambria and Aptos at 10 to 16 points; 2 px (a third of a character) off at
  8 and 20 points. A font without measured widths refuses the fit (`unmeasured_font`:
  Aptos Narrow measured 15% narrower than Aptos's widths); cells not measured (merged,
  wrapped, dates, booleans...) are listed apart. **`wb.set_row_heights`** (points, or
  `None`), **`wb.freeze_panes(sheet, rows=, columns=, at=)`** and **`wb.set_gridlines`**
  as Excel writes them (measured), **`wb.protect_sheet(sheet, password=)`** and
  **`wb.unprotect_sheet`**: Excel's `sheetProtection` defaults, a password hashed as Excel
  hashes it (SHA-512, a 16-byte salt, 100,000 rounds: the hash Excel wrote is the one
  computed here, and Excel unprotected a sheet protected here with its password).
- **On a protected sheet, as Excel allows** (measured, Excel for Mac 16.106, AppleScript
  on a sheet protected with Excel's defaults and on one with `formatCells`,
  `formatColumns` and `formatRows` `0`): `freeze_panes` and `set_gridlines` always;
  `format` only where `formatCells="0"` (whole columns and rows also need `formatColumns`
  / `formatRows`), widths only where `formatColumns="0"`, heights only where
  `formatRows="0"` -- with the defaults Excel refused bold and left widths and heights as
  they were. Anything else on a protected sheet is refused (`protected`).
- A column given a style on a sheet without `sheetFormatPr` gets Excel for Mac's default
  width there (`baseColWidth` 10: `10.83203125` with Calibri 11; measured on Mac, Excel
  for Windows not measured); with `sheetFormatPr`, its `defaultColWidth`, else
  `baseColWidth` (default 8).
- Each call is one step, undone to the bytes before it, refused before anything changes
  with every problem listed and what would work (`bad_format`, `bad_target`,
  `no_cell_style`, `too_many_cells`, `unmeasured_font`, `wrong_password`, `protected`,
  `not_writable`); something already so is `changed: 0` with a note. Formatting writes no
  value: the calculation status is kept.
- **`xlsx_format`**, the deferred `xlsx_format` group's one tool (the `edit` role's): the
  formats, `width` or `fit`, `height`, `freeze` and `gridlines` in one call, flat; not
  strict (the strict set stays `xlsx_set_cells`). 499 estimated; every definition 6,578 ->
  about 7,077 (guard 7,080); the core unchanged. Protection is the host's only: no password
  in a tool call.
- Oracle (Excel for Mac 16.106): the shop stock list plain, `r`-less, with inline strings
  and with a table, and a stock list with four named styles applied, each formatted with
  every option, a fitted and a given width, a row height, frozen panes, gridlines off and a
  password: 5 of 5 open without a repair, and Excel's re-save keeps every cell's format,
  the widths, the height, the panes, the gridlines and the password's hash.
- Needs ooxml-edit 0.95.0 (`Worksheet.set_style`, `set_row_height`, `set_row_style`,
  `cells_in`; `Styles.add_format`).

## 0.1.22 (unreleased) -- inserting, deleting and moving rows and columns

- `Workbook.insert_rows`, `insert_columns`, `delete_rows`, `delete_columns`, `move_rows`,
  `move_columns` (round 44, lane M; R52's `delete_rows`): Excel's own operations, one undoable
  step each, every dependent rewritten as measured from Excel for Mac -- formulas on every
  sheet (shared groups kept where they still read alike), defined names, print areas and
  titles, tables (new columns, calculated columns filled, totals rows), merges, data
  validations and conditional formats (split where their cells' formulas part, inherited by
  inserted lines), hyperlinks, comments and their notes, filters, frozen panes, page breaks,
  column widths and row heights, chart series, pivot cache sources, drawing anchors. What
  Excel refuses is refused alike, what is not measured is refused (`not_measured`); new
  `EditError` codes `bad_position`, `array_split`, `off_grid`, `table_header_row`,
  `table_move`, `table_rows`, `merge_split`, `pivot_table`, `not_measured`. The roadmap's
  non-goal is lifted (user, 2026-11-08); what was measured is a table in ROADMAP.md.
- `Workbook.rearrange_columns(sheet, order or mapping, drop=[...])` (R59): deletes and moves
  composed in one step.
- Tests: 108 operations recorded from Excel (`tests/golden/structure`) replayed offline
  (`tests/test_structure.py`); `pytest -m oracle tests/test_structure_oracle.py` asks Excel
  again and opens each result. The oracle drives Excel's own structural edits
  (`Session.operate`, `tools/excel_structure.applescript`). No tool added; host API only.

## 0.1.21 (unreleased) -- sheets, merged cells, text XML cannot carry

- **Sheets** (production round 44, R52; E5): `wb.add_sheet(name, index)`,
  `wb.rename_sheet(sheet, name)`, `wb.delete_sheet(sheet, force=False)`,
  `wb.clear_sheet(sheet, force=False)`, `wb.set_sheet_state(sheet, "visible" | "hidden" |
  "veryHidden")` and `wb.sheet_referents(sheet)` (`xlsx_agent.edit.sheets`): each one
  undoable step, refused before anything changes, as Excel does it (measured, Excel for Mac
  16.106). A rename rewrites every formula, defined name, conditional format and data
  validation (the `x14` ones too), chart series and title (`c:f`), pivot cache source and
  table formula that names the sheet (`rewritten`, by kind and address); a hyperlink's
  location, which Excel leaves, is listed (`not_rewritten`), and a sheet's name inside a
  string is no reference. A delete is refused while anything refers to the sheet
  (`referred`, `SheetError.referents` by address); `force` deletes it as Excel does: each
  reference `#REF!`, a 3-D run ending at it shrunk inward, the names local to it gone, the
  active tab moved. `clear_sheet` is Excel's Clear All on every cell -- cells, row and
  column formats, merges, rules, validations, comments, hyperlinks, tables and pivot tables
  go; widths, custom heights, views, print settings, drawings and every reference to the
  sheet stay -- the way to regenerate a sheet (delete + add breaks every reference).
  Hiding the active sheet moves the active tab right, else left; the last visible sheet is
  never hidden. Names as Excel takes them (1-31 characters, none of `: \ / ? * [ ]`, no
  `'` at an end, unique as Excel compares names); control characters refused. A locked
  structure, a protected sheet (clear) and the host's policy (the sheet covered whole)
  refuse.
- **`xlsx_manage_sheets`** (deferred, `xlsx_edit`; non-strict): `action` add, rename,
  clear, delete (refused with `details.referents`, or `force`), hide, unhide; a no-op is
  `changed: false`; withheld sheets neither reached nor named. 280 estimated tokens (about
  414 counted); the core unchanged.
- **Merged cells** (R56): `wb.merge(range, drop=False)` gives every cell the first cell's
  format, refused while other cells hold content (each listed) unless `drop`, which drops
  them as Excel does (the first value moved into an empty first cell); an overlap with a
  merge (with the union to merge), a table or an array formula is refused.
  `wb.unmerge(range, fill=None | "anchor")`: every merge the range overlaps.
- **Text XML cannot carry** (R63b): `set_cells` writes control characters, a carriage
  return and text that reads as an escape (`_x0041_`) as Excel writes them, `_xHHHH_`
  (`_x005F_x0041_`), in shared and inline strings, rich text and formulas, and names those
  cells (`escaped`); before, it raised lxml's `ValueError`. A formula's string result
  (`t="str"`) and its text are read decoded.
- `xlsx_agent.formula.sheets`: sheet names in references renamed and deleted, structured
  references to a removed table, names folded as Excel compares them.

## 0.1.20 (unreleased) -- new workbooks, tables written, rows appended

- **`Workbook.new()`** (production round 44, R51): a new workbook as Excel makes one, its
  nine parts measured from what Excel for Mac 16.106 writes for a new workbook
  (`xlsx_agent.edit.blank`); `sheets`, `author`, `title`. **`Workbook.from_tables({"Stock":
  {"columns": [...], "rows": [...], "table": True}})`**: a sheet per entry, written by
  `write_table` at A1, in one step.
- **`wb.write_table(sheet, origin, headers, rows, as_table=False|name, style=,
  overwrite=)`** (R54, E4): a block with `set_cells`' values, returning the range written;
  `as_table` makes a real Excel table (its AutoFilter, the workbook's default style banded
  by rows), its part as Excel writes it and its name checked by Excel's rules (measured).
  Refused before anything changes, every problem listed: cells holding values (unless
  `overwrite`), merges, other tables, a bad table name (with one that would do) or style.
- **`wb.append_rows(target, rows)`** (R54): to a table (`t:Orders`: it grows, calculated
  columns filled, the format above each new cell, a totals row moved below unless something
  refers to it by address), or below the last row holding a value in a sheet's or region's
  columns (a formatted empty row does not count).
- A date written by these into a cell that shows no date gets the short date (`numFmtId`
  14), as Excel formats a typed date (measured). A workbook without a shared-string table
  gets one with its first text.
- Tool **`xlsx_write_table`** in the deferred `xlsx_edit` group (`at` a cell, or `t:<table>`
  to append; `headers`, `rows` of typed cells, `from: d:<id>`, `table`, `name`, `style`,
  `overwrite`), non-strict: the core unchanged (4,938 estimated), every definition 5,909 ->
  6,298. The shared `new_document`'s xlsx handler is `xlsx_agent.tools.NEW_DOCUMENT`, which
  an application adds (in the core it would cost 225 estimated tokens, past the budget).
- New edit codes: `not_empty`, `table_overlap`, `bad_table_name`, `bad_table_style`,
  `totals_referenced`. Oracle (`tests/test_oracle_tables.py`): 8 of 8 open without a repair.

## 0.1.19 (unreleased) -- the length of a text in a page

- **`text_chars`** (production round 41, R45) on `xlsx_read` and `xlsx_query`, of a table, a
  sheet's cells or a dataset: the characters a cell's text (a value, a formula, a comment,
  a link) may have in a page before it is cut, by default 2,000 as before (a read of one
  cell: as far as the page holds, unless `text_chars` is given). `max_text_chars` and
  `text_length` are read as it, with a warning. A longer `text_chars` gives fewer cells a
  page, never a bigger one: pages keep `max_chars`/`Limits.max_result_chars` and page with
  `next_cursor` as before.
- **`cut.text` gives the call for the rest** (rule 4): besides `count` and `first`, the
  `longest` cut text's full length and `call`, the call's own arguments (the same page, its
  cursor if any) with `text_chars` that long -- and without a `max_chars` too small for it
  -- which gives every text of the page in full, paged. A text longer than a page holds is
  cut to the page in any read, and a warning says so. A cut cell keeps its full `length`
  (`f_length`, a comment's `length`, `url_length`). A dataset read's `cut.text` and
  `xlsx_derive`'s were a count: they are now this object, `xlsx_derive`'s `call` the
  `xlsx_read` of the new dataset.
- Definitions: `xlsx_read` 690 -> 713 estimated, `xlsx_query` 513 -> 537, the core 4,924 ->
  4,972, every definition 5,896 -> 5,943: counted online 7,024, past the 7,000 budget. So
  `xlsx_query`'s description no longer says "A filter on a column of another type is
  refused, naming the type." (the refusal says so itself, with the tests the type takes):
  `xlsx_query` 537 -> 518, the core 4,972 -> 4,953 (about 6,997 online at the counted
  ratio), every definition 5,943 -> 5,924; counted online 7,006. So `xlsx_read`'s `target`
  no longer says "held_by names them (null: cells read as themselves)" (the result's
  `held_by`, the first such read's warning and `cells_only` say it): `xlsx_read` 713 -> 698,
  the core 4,953 -> 4,938 (about 6,985 online), every definition 5,924 -> 5,909. **The
  online count must be checked again**
  (`test_tools_online.py`, budget 7,000). Neither tool is strict. `xlsx_trace` is unchanged
  (its values are always whole).

## 0.1.18 (unreleased) -- key sets on datasets, any of a list, loose text

- **Key sets on datasets** (production round 40, R42): `d:by_region/{South,North}/total`
  names the rows of its keys in its order, as a table's key set does, in `xlsx_read`
  (paged, `columns`), `xlsx_trace` and `DatasetTable.read`/`select` (`DatasetSelection.order`,
  `positions`). Keys are looked up as a table's are (`KeyIndex`): a missing key names its
  place in the set and the rows like it by position, a shared key the rows by position. A
  dataset's key column is its own when a derive grouped it by one column; otherwise
  `xlsx_read`'s `key` names one (`d:by_rep/{Ann,Bob}` with `key: "Rep"`), and without it the
  read is refused (`no_key_column`) with the columns that can be the key as `valid_options`
  and the call that works in the message and `details.call`. A `key` given for a dataset
  address with no row by key changes nothing and says so. `no_key_set` is no longer raised.
- **Any of a list, and loose text** (R43): the filters' `contains_any` (text holding any of
  `values`, case ignored as `contains`; `contains` given `values` is read as it, said in the
  warnings) and a filter's `loose: true`: `eq`, `in`, `contains` and `contains_any` on a
  text or list column compare both sides with every white space character dropped
  (`str.isspace`, plus U+200B and U+FEFF) and Unicode's full case folding
  (`xlsx_agent.query.loose_text`); accents and punctuation still count. The result's
  `loosened` counts, per column, the matching rows that match only loosely, with their
  first addresses; `loose` where it loosens nothing says so. A query that matches no row
  but would with `loose` says how many and which filters. In `xlsx_query` (tables and
  datasets) and `xlsx_derive`'s `where` alike (`Derived.loosened`, `Derived.notes`).
  `xlsx_query` 473 -> 513 and `xlsx_derive` 1,245 -> 1,286 estimated; the core 4,844 ->
  4,924, every definition 5,815 -> 5,896 (guards 4,925 and 5,900). Neither tool is strict.

## 0.1.17 (unreleased) -- a workbook's values for a check of numbers

- **Numbers in text no value holds** (production round 36, R28; ooxml-edit 0.84.0): the
  xlsx format lists every Excel table and every declared table as a value source
  (`DocumentFormat.value_sources`, `xlsx_agent.tools.docformat.value_sources`), so a
  deck's or document's `check` with `include: ["numbers"]` compares its text with the
  values of every open workbook's tables, read as `from: "t:..."` reads them.
- A table's dataset says each column's number format (`Dataset.formats`, from the cells
  read: the one code a column's numbers and percentages are shown in, `None` for General
  or several), so `2.649` shown `0.00` matches a text's `2.65` and `2.6`. A derived
  dataset says it for the columns that come straight from a source's cells: a `group_by`
  column, a `min`, `max`, `first` or `last` (a pivot's too), a join's or a fill's rows
  (each side's own); computed ones (`sum`, `mean`, `share`...) have none
  (`Derived.formats`, `dataset.column_formats`).
- No tool definition changed: offline tokens and the strict set as before.

## 0.1.16 (unreleased) -- a column's blanks filled from another

Production round 35, R22.

- **`xlsx_derive`'s `fill`** (`xlsx_agent.fill`; `Workbook.derive(fill=...)`):
  `[{"column": "Limit", "from": "Default limit", "coerce": "number"}]` gives each blank
  cell of `column` -- no value, or text of no characters (a formula's `""` counts as blank)
  -- the same row's `from` value,
  after the join and before `where`, the groups and the measures. A fill never overwrites;
  several apply in order (a second fill of the same column fills what the first left
  blank). From a looked-up table: join it, then fill from the joined column.
- **`coerce`** (`number`, `date`; `number_as_text` and `date_as_text` taken for them) reads
  text as the derive's `coerce` does (the reader's `number_as_text`/`date_as_text`, which a
  declared column's cleaned text is flagged as too); a `from` value of another type fills
  nothing. Without it a value is filled as it is.
- **Provenance**: a filled value's sources are the `from` cells (by key when the table has
  a unique key column), and it is marked with that column: `Provenance.filled`, and
  `"filled": "<column>"` on the value in the tool's rows and in an `xlsx_read` of the
  dataset (`DatasetValue.filled`). `Derived.filled` (`filled` in the result) gives per
  fill the cells filled, those `coerced`, and those left blank by reason (`blank`,
  `error`, `no_cached_value`, `wrong_type`), each with first addresses.
- **Intent and refusals**: one fill object is taken as a list of one; a strict call's
  placeholder entry (`column` and `from` empty) is dropped; a fill of a column from itself
  does nothing and says so in its `note` (naming a joined column renamed beside it). A
  fill of the wrong form, or naming no column, is refused (`bad_fill`, field `fill`) with
  the form that works or the table's columns.
- Without `group_by`, `measures`, `bins` or `pivot` the filled rows are the dataset (as a
  join's); `round` is refused for them.
- Definitions: `fill` about +109 estimated; the note that a join alone gives the rows moved
  to the tool's description, now for a fill too. Counted online, the core went 6,678 ->
  6,838, past the 6,800 budget, so `xlsx_derive`'s texts were shortened with their meaning
  kept (`doc` worded as `xlsx_read`'s and `xlsx_query`'s; `join`'s, `pivot_values`', `id`'s,
  `from`'s and `coerce`'s sentences tightened): `xlsx_derive` 1,116 -> 1,182 estimated, the
  core 4,717 -> 4,783 (about 6,780 online), every definition 5,688 -> 5,754 (guards 4,720
  -> 4,785 and 5,690 -> 5,755); a reviewer's core 4,563 -> 4,629. `xlsx_derive` is not
  strict (13 optional parameters): the strict set is unchanged.

## 0.1.15 (unreleased) -- many reads in one call

Production round 32, R15.

- **`read_batch`** (ooxml-edit 0.78.0) is a workbook tool, in the core: many reads --
  records by key, cells, regions, queries, traces, `describe` -- in one call, each answered
  as the call alone (its own `max_result_chars`), the reply within 96,000 characters, and
  what does not fit listed with the call that asks for it. It changes nothing and is the
  `review` role's, so a reader is given it; a change among its ops is refused with the way
  out (`batch`). The prompt's line on the read tools says several reads go in one
  `read_batch`, and ooxml-edit's shared fragment names it.
- Definitions: +184 estimated tokens (`read_batch`), the core 4,533 -> 4,717, every
  definition 5,505 -> 5,688 (guards 4,720 and 5,690); a reviewer's core 4,379 -> 4,563.
  Not strict: the strict set is unchanged.
- `ooxml-edit>=0.78.0` (was 0.72.0): the first with `shared.READ_BATCH`.

## 0.1.14 (unreleased) -- a value after its label, anywhere in the cell

Production round 32, R9. No tool definition changed.

- **A declared extract finds its label anywhere.** Before, an `extract` match had to start
  in a cell's last 256 characters (`EXTRACT_WINDOW`), so a status line followed by a
  justification of 260-350 characters read blank. Now a pattern that begins with a label
  -- literal text of 3 characters or more, or a choice of such texts
  (`(?i)(?:Review state|Next step)\s*:\s*(...)`; `clean.pattern_label`) -- is tried once
  more where that label last occurs, when nothing matched in the window. A cell that
  matched in the window reads as before; one that read blank may now have its value. A
  pattern anchored at the text's start (`^` without `(?m)`, `\A`) is not.
- **`after`**, a new `clean` rule: the label a value follows, literal text or a list of up
  to 16 (`{"after": ["Review state", "Next step"], "extract": "^\\s*:\\s*(\\S[^,.;\\n]*)"}`).
  The labels' last occurrence is found by a plain, case-insensitive string search, never a
  pattern, and `extract` is applied to the text after it as if the cell began there (`^`
  the label's end); where that finds nothing, to the text from the label on, so a pattern
  written for the whole cell works with it too. A match starts within 256 characters of
  the label and is read within 512 (`EXTRACT_REACH`), a match reaching that far tried once
  on the whole text. Refused without an `extract`, with the form that works.
- **A label without a value is counted.** A cell holding the label (`after`'s, or a
  pattern's own) where the pattern found nothing reads blank with
  `issue: "extract_unmatched"`: per column in `xlsx_profile` and `describe` (its count and
  first cell), `Derived.unmatched` and `xlsx_derive`'s `extract_unmatched`, an
  `xlsx_read`/`xlsx_query` page's `extract_unmatched`, and a deck's or document's table
  from `t:` (the dataset's `detail`), each with a warning naming the first cell and that
  `xlsx_read` with `cells_only` shows its text.
- **Bounded**: a label is found by a string search from the end (`str.rfind` on the
  case-folded text); a cell of 100,000 characters costs at most about 15 ms (the worst
  pattern taken, `([^;]*);`, after a label at its start), a label pattern's about 0.2 ms;
  a 5,000-row table of three extract columns reads as fast as before (measured on an Intel
  Core i7-6567U, Python 3.12).

## 0.1.13 (unreleased) -- keys like 4.1: found, addressed, followed in text

Production round 31, R7b. No tool definition changed.

- **A column of dotted numbers is the key** where several columns could be and the others
  hold words: a table of clauses numbered `4`, `4.1`, `4.1.2`, `A.5.1` beside their titles
  is keyed by the numbers (`xlsx_agent.read.regions.is_dotted_key`: digits joined by dots,
  after at most two letters and a dot; at least one key with a dot). So `t:Std/4.1`,
  `t:Std/{4.1,4.2}` and `t:Std/4.1/Title` address its rows in every tool taking an address
  by key (`xlsx_read`, `xlsx_trace`, `find_text`'s `range`, `xlsx_compare`'s pairing,
  `describe`'s key; a write by key still needs the key declared). Beside another column of
  codes or whole numbers (at least half its values without a space and with a digit), or
  another dotted column, nothing is chosen and the lookup is refused, as before. Before,
  `t:Std/4.1` was refused (`no_key_column`, 2 columns unique). A clause number typed as a
  number (4.1, a fraction) is no key, as before: 4.10 would be the same number.
- **`xlsx_trace` follows dotted keys in text** (`see 4.2`) of a table keyed by them
  (declared, or found), never its whole numbers (`4`) nor another table's dotted values.
  Whole only where the dotted number is: not in `4.1.2`, `1.4.1`, `12.05.2026` or `4.10`;
  yes in `4.1.` ending a sentence, `4.1-4.3`, `(4.1)`. A key of digits alone is not found
  as a quantity: after a currency sign, a sign or a comparison (`$4.2`, `-4.2`, `= 4.2`),
  or before a percent, a unit, a word of a quantity or a score (`4.2 kg`, `4.2%`, `4.2 GB`,
  `4.2 days`, `4.2 heures`, `4.2 million`, `4.2/5`, `4.2 out of 5`:
  `xlsx_agent.trace.UNITS`; `4.2 sur la base` still names 4.2).
  Elsewhere it counts as named (`version 4.1` names clause 4.1 when the workbook has
  one). A dotted key with a letter (`A.5.1`) was found before; now it too is whole only
  where the dotted number is (`A.5.1` is not in `A.5.1.2`).
- **Refusals carry the way out** (tool-design rule 2): a key no row has, and a table with
  no one key column, now say the table's data rows and the addresses by position that name
  them (`t:Std has 6 data rows: t:Std/#1 to t:Std/#6 (t:Std/#1:#6)`); the
  latter also how a key is named or declared (`key=`, a declared table's key,
  `Workbook.declare_key`). Never the key itself, as before. A row by position written
  without its table (`#4`, `#1:#8`) is still refused -- it names no table -- now saying so,
  with `t:Products/#5`, `t:Products/#9:#16` and `d:by_category/#2` as the forms.
- Dataset rows by position (`d:x/#4`, `d:x/#1:#8`) read and trace as before; now tested.
- Tokens: no definition changed (the core 4,533 estimated, every definition 5,505).

## 0.1.12 (2026-11-04) -- rows traced to the rows they rest on

Production round 31, R7, and the workbook side of R5e. One tool added, on demand.

- **`xlsx_trace`** (R7), in a deferred group of its own, `xlsx_trace`: rows of a table or
  dataset -- `t:Items/I-4`, a key set, a row range, a whole table, a sheet's cell inside a
  table, `d:by_owner/#2`, a whole dataset -- each with the rows it rests on and why:
  `source` (a dataset row's values came from it: the sources the derive kept, a key set's
  keys each a row, a range the table rows holding it, else the cells), `names` (its text
  names the row's key -- the traced row's, a source row's, or, `depth` steps deep (2
  unless given, at most 4), a named row's; `by` and `column` say whose text and where) and
  `lists` (the row's text names the traced row's key, or a source row's). Every row's
  values in full, once, in `rows` (a declared split's items as a list; blanks left out).
  A reply holds at most `max_chars` (by default and at most 24,000): rows are traced in
  order while their links fit, then the named rows' values in link order; what did not fit
  is listed, never cut -- `left_out` as key sets `xlsx_read` reads, `untraced` as rows to
  pass back (key sets, row ranges) -- and `next` says so. At most 50 rows a call, 200
  links a row (the rest counted). The rows linked are the Excel tables' and the declared
  tables', by key where the table's key is known; a key is found in text whole, case
  counting, and only a key holding a letter and a digit is looked for (numbers and words
  would be found everywhere). `Workbook.trace(rows, depth=...)` in Python
  (`xlsx_agent.trace`: `Link`, `Trace`, `TraceResult`); the index of what each row's text
  names is built once per workbook state. A dataset is traced in its own workbook only
  (`doc`), else refused naming it. Measured on a made-up workbook of 7,000 rows: about
  3.4 s for the first trace of a workbook just opened, 1.2 s after `describe`, then about
  0.2 ms a row.
- **Budget.** `xlsx_trace` is 242 estimated (about 358 counted by the proxy), deferred: the
  core is unchanged (4,533 estimated); every definition 5,274 -> 5,505, the guard 5,275 ->
  5,510. It reads, so it is not strict: the strict set is unchanged (`xlsx_set_cells`).
- **`find_text` on a workbook with a document's `stories` at its default** (R5e,
  ooxml-edit 0.77.0): a toolbox that also serves documents is sent `find_text` with
  `stories`, and a strict decoder writes it; `"body"` and `"all"` (and another kind's empty
  `slides`) are now ignored with a warning instead of refused; `layouts` or `masters` are
  still refused, naming `range`. No change here but a test (skipped on an older
  ooxml-edit); the floor stays 0.72.0.

## 0.1.11 (2026-11-03) -- a dataset's columns picked, the blank cells listed

Production round 28, X7 and X8. Two descriptions changed; no parameter added.

- **`xlsx_read` of a dataset takes `columns`** (X7), as a table's read does:
  `{"target": "d:by_region", "columns": ["total", "Region"]}` reads those columns, in the
  order given, each once, the address's own column first (`d:by_region[total]` with
  `columns: ["Region"]` reads `total`, then `Region`; a row or a value address reads its
  row's columns). A name the dataset has no column of is `not_found` (`field: "columns"`),
  its message and `valid_options` listing the dataset's columns. A cursor belongs to the
  columns asked for. Before, `columns` with a `d:` target was refused; `d:by_region[total]`
  reads as before.
- **`pivot_values: [""]` is the blank cells** (X8), in a pivot column of any type: their
  column is named `(blank)` (`(blank) total` with more measures), as a pivot without
  `pivot_values` names it, and their rows count there instead of in `pivot_unlisted`.
  Before, `""` was read as the empty text, and the column without a name was refused with
  the dataset's generic message. A clash with a text value `"(blank)"` is resolved for the
  blank cells: the blank cells' column is always `(blank)`, the text's `(blank) (text)`,
  whichever comes first (with or without `pivot_values`; before, the one met first kept
  the name). A text cell of no characters (a formula's `""`) is not blank: its column, or
  its `pivot_unlisted` entry, is named `(empty text)` (before, it had no name and the tool
  refused the dataset), and `""` does not list it.
- **A refused pivot value names its place**: `pivot_values[2] is not a number: Size is a
  number column, ...`, `pivot_values[2] repeats pivot_values[1] ...`, `pivot_values[1] is
  no value: ...` (codes and `field` as before; never the value itself).
- Tokens (offline proxy): `xlsx_read` 688 -> 690, `xlsx_derive` 1,110 -> 1,116, the core
  4,525 -> 4,533, every definition 5,266 -> 5,274. Neither tool is strict; the strict set is
  unchanged.

## 0.1.10 (2026-10-30) -- a pivot's columns, listed

Production round 27, X6. One tool parameter added (`xlsx_derive`'s `pivot_values`).

- **Listed pivot values** (`xlsx_agent.derive`; `Workbook.derive(..., pivot_values=[...])`,
  `xlsx_derive`'s `pivot_values`). One column (per measure) for each listed value, in the
  order listed, whether any row holds it or not: a report's fixed columns. A value without
  rows has what a measure of no values has: `count` and `distinct` 0, every other measure
  blank (`None`): a sum, mean, min or max of nothing is not 0, so 0 would be a false
  figure. The values are written as filter values are and read as values of the pivot
  column's type, then matched as `group_by` groups, by type and value: in a number column
  `"5"`, `5` and `5.0` are the number 5, matching cells holding 5 and never the text
  `"5"`; in a text column `"5"` is that text, exactly; a date is its day; `"true"` a bool.
  A value not of the column's type is refused (`wrong_type`), one listed twice (as the
  column reads it) `bad_pivot`, as is `pivot_values` without `pivot`, empty, or longer
  than 100.
- **Values present but not listed** are not columns (there is no option to append them:
  list them, or leave `pivot_values` out). Their rows are left out and counted in
  `left_out` under the pivot column (`not_listed`, a new reason, with the first cells), and
  each value is named in `pivot_unlisted` (`{name, v, t, rows}`, the name its column would
  have, in the order first met; at most 100 values, `left_out` counting every row). A
  blank, and a cell of another type than the column's (the text `"5"` in a number column),
  is such a value. The columns are fixed by the list and the rows by `group_by`: a group
  whose rows all hold unlisted values is still a row, its cells empty (`count` 0, the others
  blank), its rows counted as left out. (Bins keep their behaviour: a row outside them has
  no bin.)
- **A coerced pivot column** reads its text as a `group_by` column does: with
  `coerce: ["date_as_text"]` a date stored as text in a date column is its date, matching a
  listed date and counted in `coerced` under the pivot column. This applies to `pivot`
  without `pivot_values` too (before, a pivot read such text as text).
- Works on derived fields (`from_field`), joined columns and datasets, and composes with
  `sort` (by a listed value's column), `top` and `round`.
- Tokens (offline proxy): `xlsx_derive` 1,045 -> 1,110, the core 4,460 -> 4,525, every
  definition 5,201 -> 5,266. `xlsx_derive` is not strict; the strict set is unchanged.

## 0.1.9 (2026-10-30) -- a value read out of a column, beside it

Production round 26, X4 and X5. No tool definition changes: both are declarations, the
application's configuration.

- **A field derived from another** (X4, `xlsx_agent.declare`). A declared field may read
  another column's cells by rules of its own: `{"id": "part", "from_field": "notes",
  "clean": {"extract": "Part:\\s*(\\S+)$"}}` is a column `part` right after `notes`, in
  the same table, read, queried, profiled, grouped by, joined on and derived from like any
  other. Its values carry the cells they read (`Log!B2`), and so does every source a
  derive gives (by key, `t:jobs/{J-1,J-4}/notes`: the column read). `from_field` names a
  column of the table by id or header; `id` and `clean` are required; `header`, `column`
  and `hidden` are refused. No chains: a field derived from a derived field is refused
  (`bad_source`), so no cycle can be declared; so is one derived from a hidden column (it
  would show what the column hides), and a derived `key` (`derived_key`). A value by key
  in a derived column is refused (`derived_column`), naming the column to write; a write
  policy naming it is refused, and a policy's table without `columns` leaves it out. It
  is one column for hiding: nothing is declared twice, and no second table over the same
  cells (which had to hide the same columns) is needed. `describe` and `xlsx_profile`
  show `derived_from`; `Region.derived` and `Region.source_of` in Python.
- **Rules on numbers as shown** (X5, `xlsx_agent.shown`). A `clean` with `"on": "shown"`
  reads number and percent cells too, each as the text Excel shows it with its number
  format in the workbook's locale, so `{"extract": "^(\\d+)", "on": "shown"}` reads the
  text `"5.1"` and the number 5.2 alike, as `"5"` (flagged `number_as_text`). The default,
  `"text"`, is as before. The text is made only where it is measured: the forms of the 43
  number codes of `ooxml_common.numfmt.MEASURED` (Excel for Mac 16.106, 2026-10-25), each
  of the 36 rendered held to Excel's text in the tests. The seven measured codes those
  forms do not give (a fraction, a fill, two conditions, a numeral system, engineering
  notation), any other form, a negative number, zero with a zero section, a half where
  roundings differ, and a long number in `General` are refused: the cell reads `blank`
  with `issue: "unmeasured_format"`, never a guess. Dates, times, bools and errors are
  never changed.

## 0.1.8 (2026-10-30) -- datasets that match the table an agent shows

Production round 25, X0-X3.  0.1.8, not 0.2.0: the public API grows (a derive's `join`,
`Workbook.rows_holding`, `Joined`, `xlsx_read`'s `cells_only`, the `extract` rule), but
0.2 is the roadmap's edit milestone, and nothing is removed or renamed.

- **A cited cell reads its table row** (X0). `xlsx_read` of a sheet's cell or range that
  lies inside one table's data rows -- in any form a citation writes it: `'Q3 review'!K83`,
  `Review!$K$83`, `'Ann''s list'!B4`, `'Q3 review'!K80:L90` (the address parser's forms,
  its quoting measured in 0.1.6) -- reads the table rows holding it, every column or those
  `columns` names, paged as any read: `target` is their `t:` address (one row by key when
  the table's keys are unique, else by position; a range's rows a row range), `cells` the
  address asked for, `held_by` the table and rows. Declared tables and regions come first
  (one over several sheets counts its rows sheet by sheet), then Excel tables. A header or
  totals cell, one outside every table and a range across a table's edge are read as cells,
  with `"held_by": null`. `cells_only: true` reads the cells as stored (as before) and still
  names the rows. In Python: `Workbook.rows_holding(address)`, `Regions.holding`.
  *Changed*: a read of a single cell inside a table, which gave the cell, now gives its row
  (the cell among it, with its address); `cells_only` gives the cell alone. `held_by.read`
  says which was read (`rows` or `cells`), and the first such row read on a workbook warns
  once that the row was read instead of the cell and that `cells_only: true` reads the cell.
- **Joins** (X1, `xlsx_agent.join`). `xlsx_derive`'s (and `Workbook.derive`'s) `join`,
  `{from, on: [left key, right key], columns}`, adds another table's or dataset's columns to
  each row of `from`, by key, before `where`, the groups, the measures, `sort`, `top` and
  `round`. Keys match by type and value as `group_by` groups; a list key (a declared split,
  or a list in a dataset) matches by any of its items, on either side. Every row is kept,
  unmatched ones blank. When any row matches several right rows, every joined column is a
  list column (each cell the matched rows' values in order): rows are never repeated, so the
  left table's counts, sums and keys stay true and "the issues of each item" is one cell;
  a deck or document shows it one item per line (the dataset holds the lines; reading or
  deriving from it gives the list). A joined name the left side has is renamed
  `Status (Issues)`. Each joined value's sources are the joined rows' cells. Without
  `group_by`, `measures`, `bins` or `pivot` the joined rows are the result (`round`
  refused: they are cells as stored); `joined` says what was matched. At most 1,000,000
  values held (`too_many_cells`), and the result's rows as any derive's. A joined dataset
  goes stale when either side changes -- `from`'s source, the joined dataset's source or
  the joined table's workbook (ooxml-edit 0.72.0, `define_dataset(also=...)`; the floor
  rises to it) -- and undoing the change makes it current again.
- **`extract`, a declared clean rule** (X2). A pattern with exactly one capture group,
  searched for (anchored only as written) in a declared column's text, its match starting
  in the last 256 characters: the value is the captured text, or blank when not found --
  a status line ending a narrative becomes a column, each value's source its cell. Applied
  after `strip`, before `split` and `collapse_space`. Like the other rules it is the
  application's configuration, never a model's: no tool definition changes. Its cost is
  bounded: it is tried at every place of its window, so no repeat may give characters back
  (`\s*(.+)$` refused, `\s*(\S.*)$` taken); the worst forms allowed took about 1 ms a
  cell of 32,767 characters (Python 3.12, an Intel Core i7-6567U, 2026-10-30). The pattern
  check now judges a repeat before only the end of the text (`$`, `\Z`) as one at the end,
  and a repeat at the end of a group by what follows the group (so `([^;]*);` is taken).
- **Dates stored as text through a derived dataset** (X3). A derive's `text` column of ISO
  dates -- a `first` or `last` of dates stored as text, grouped without `coerce` -- keeps
  them `date_as_text` in its dataset, so a derive from it refuses `max` naming the `coerce`
  and takes `coerce: ["date_as_text"]` (a caller's own `text` label is still taken at its
  word). A `group_by` column made a date by `coerce` groups its text by the day, shown ISO
  (counted in `coerced` under its name). A `coerce` that read none of a refused column's
  values says so: `coerce ["date_as_text"] did not apply: none of its values reads as a
  date`.
- **Definitions**: `xlsx_read` 593 -> 688, `xlsx_derive` 801 -> 1,045 estimated; the core
  4,122 -> 4,460 estimated (about 6,601 by the offline proxy; counted online 5,894 ->
  6,334, claude-sonnet-5-5), which the user allows past 6,500 for a good tool:
  the offline core check is now a regression guard (4,465). Every definition 4,862 ->
  5,201, the guard 4,865 -> 5,205. The strict set is unchanged (1 tool, 13 optional).

## 0.1.7 (2026-10-29) -- the shared tools sent with a workbook's parameters only

- **The shared tools without a deck's or document's parameters** (production round 23, V3;
  ooxml-edit 0.70.0, `Tool.kind_fields`): a workbook's toolbox is sent `find_text`
  without `slides` and `stories` (its `range` described for a workbook alone: "Only within
  this sheet, range or table."), `save_document` without a document's `clean`, and
  `undo`'s `scope` described for a sheet. A call that still passes one is answered as
  before: this library's `find_text` refuses `slides` and `stories`, ooxml-edit refuses
  `clean: true`. Offline: the core 4,272 -> 4,122 (about 6,323 -> 6,101 counted), every
  definition 5,013 -> 4,862 (guard 4,865); the strict set is unchanged (1 tool, 13
  optional parameters). The ooxml-edit floor rises to 0.70.0.
- **The Excel oracle is harder to fool** (`tests/oracle.py`,
  `tools/excel_resave.applescript`; tests and tooling only, no package change; listed
  under 0.1.7, the first version after it). Found while measuring sheet-name quoting (0.1.6):
  - *What Excel saved is the input.* Excel can start with a blank workbook ("Map1") that
    stays active after the input is opened, and the script saved the active workbook: a
    blank file, reported `ok`. The script now saves the workbook of the input's file name
    (none open: it prints `wrong workbook:` and saves nothing), and the saved file must
    have the input's sheet names in order: otherwise `mismatch`.
  - *The "hangs" were repair prompts nobody could see.* With the screen locked, System
    Events lists no window of Excel's, so the repair prompt was never found: Excel waited
    on it until the AppleEvent time-out, which looked like a hang with no dialog. Measured
    (Excel for Mac 16.106, 2026-10-28, screen locked, made-up workbooks): the window
    server still lists the prompt (an untitled window, 273 x 274-288 points), and
    ScreenCaptureKit images of it showed Excel's repair prompt for each case below and the
    shared-string control. With the screen locked the oracle now finds it there (JavaScript
    for Automation, `CGWindowListCopyWindowInfo`), stops the call, quits Excel by signal
    and reports `repair` (it cannot be answered; `detail` says so) -- the oracle suite
    took 2 min 11 s where the time-outs took over 4 min. Excel not answering with no
    prompt found is `hang` (the script reports AppleScript's -1712 in any language),
    naming any known cause and that the screen was locked, never a plain `error`.
  - *Known causes, before Excel.* `repair_causes` finds what makes Excel ask to repair a
    workbook, as measured: two sheet names it treats as equal -- case (`R`/`r`), NFKC
    forms (`a¹b`/`a１b`, a no-break space and a space, halfwidth and fullwidth katakana),
    soft hyphen, BOM and zero-width space left out (`a<U+200B>b` equals `ab`) -- and rows
    or cells out of order: `B1` listed before `A1`, row 2 before row 1, a cell or a row
    listed twice. Such a file is not given to Excel (`preflight`, naming the cause) unless
    `resave(preflight=False)`.
  - *Interrupted, Excel is still quit.* The `osascript` runs in its own process group;
    Ctrl-C, `SIGTERM` or `SIGHUP` mid-run kills it, keeps the Excel it started as the
    session's and quits it; signals during the quit are held until Excel is gone; an
    exit hook recovers a session never closed. Against `kill -9`, the session keeps
    the Excel it started (pid and start time) in `excel-started.json` while it runs, and
    the next session quits that Excel only when it is still that process, its run is
    gone and it has no workbook open; any other Excel is "in use", as before. Verified
    with Excel: each of the four signals sent to a run holding Excel on a prompt left no
    Excel running (for `kill -9`, after the next session started).
  - Outcomes are `ok`, `repair`, `mismatch`, `hang`, `preflight` and `error`
    (`oracle.OUTCOMES`); only `ok` and `repair` are cached. Unit tests for the
    pre-flight, the outcome mapping, the window-server check and the interruptions run
    without Excel; one more oracle test gives Excel a known cause and checks the next
    file is saved as itself.

## 0.1.6 (2026-10-27) -- sheet-name quoting measured in Excel

- **Quoting follows Excel, measured** (round 18, follow-up C). Excel for Mac 16.106 on
  macOS, through `tests/oracle.py`: 253 made-up sheet names, each referred to from a
  formula both in a file Excel opened and saved and through Excel's object model. Excel
  read every name quoted (`'Name'!A1`), from a file and through the object model, giving
  the referred cell's value; the bare form of a name Excel quotes was refused, read as
  something else (`a-b!A1` as `a-[2]b!A1`, another workbook) or requoted. `needs_quotes`
  now follows the `<f>` text Excel saves, and the module docstring and README.md say so in
  place of "not a measurement" and "Excel accepts a quoted name wherever it accepts a bare
  one". The measured names are `tests/test_sheet_quoting.py`.
- **Now quoted** (Excel quotes them; 0.1.5 wrote them bare): a name beginning with an
  R1C1 reference that has a row or column (`R1x`, `C1x`, `RC1x`, `R1C1x`, `R2C3D`;
  `R1048576x`), and a `.` Excel quotes: twice (`a.b.c`, `a..b`), in a name starting with
  `_` (`_a.b`), or after a part holding a digit, a boolean or an R1C1 reference
  (`Q3.sales`, `A1.b`, `AB12.c`, `TRUE.x`). Bare, the first
  kind is refused (through the object model; a file holding it did not open), the second
  read and requoted.
- **Now bare** (Excel writes them bare): names of letters of any script (`Überblick`,
  `数据`, `Привет`; combining marks and digits after the first letter: `नमस्ते`, `a٣b`),
  cell-shaped names off the grid (`XFE1`, `AB0`, `A1048577`), and R1C1-shaped names out of
  range (`R0`, `C0`, `R1048577`, `C99999999`).
- **Still quoted where Excel is irregular**: Excel writes many non-ASCII symbols bare
  (`a§b`, `x²`, `a—b`, `a€b`) and quotes others (`a©b`, `a®b`, `a£b`, `a«b`, `€x`, `a，b`),
  so every character that is not a letter, a digit or a combining mark is quoted, as is a
  name starting with one; Excel reads them quoted. Where Excel's object model (`Formula`)
  reads a name back bare that its saved file quotes (`_a.b`, `a..b`, `a，b`), the file is
  followed.
- **Compatibility.** A returned address and `canonical_address` change for the names
  above: `Überblick!B4`, no longer `'Überblick'!B4`; `'R1x'!B4`, no longer `R1x!B4`.
  Parsing is unchanged (either spelling of a name that needs no quotes is read). Callers
  comparing through `canonical_address` see no difference; text kept from 0.1.5 should be
  passed through it again.
- **The formula lexer** (`xlsx_agent.formula.tokens`) reads the bare names Excel writes
  with characters from U+00A1 to U+00BF (`aµb`, `x²`, `a§b`): they raised an internal
  assertion, or `²` started a number. Tool definitions are unchanged.

## 0.1.5 (2026-10-27) -- one canonical form for every address returned

- **The canonical form, documented and in one place** (production feedback, round 18, Q3).
  Excel's own spelling is kept as the canonical one, not always-quoted, so a returned
  address pastes into a formula unchanged: a sheet name bare when it is ASCII letters,
  digits, `_` and `.`, starting with a letter or `_`, and not a cell reference, an R1C1
  reference or a boolean; else quoted, each `'` doubled; column letters upper case, no
  `$`; a range top left to bottom right, one of one cell the cell. Written out in the
  module docstring of `xlsx_agent.address` and README.md's new "Addresses". The quoting
  rule is unchanged, and is still this library's conservative one: which further names
  Excel writes bare is not measured.
- **`canonical_address(text, sheets=None)`**: any spelling parsing accepts in that form
  (`sales!$b$4` is `sales!B4`; `Überblick!a1` is `'Überblick'!A1`); with the workbook's
  sheet names, a sheet spelled as the workbook spells it (`Sales!B4`), matched case aside.
  Exported from `xlsx_agent`.
- **One formatter.** `format_sheet`, `format_cell` and `format_range` in
  `xlsx_agent.address` write every sheet-qualified address; `SheetRef.text`,
  `RangeAddress.text` and `Value.a` use them, as do the places that wrote an address by hand
  (find results, regions, table and header flags, a validation rule's `address`, a
  cell compare's edits, the edit report's and the structure check's problems, edit
  refusals, shared formula anchors). Two of them could differ from the form: the structure
  check's `cell_row` problem gave a cell's `r` as stored (lower-case letters kept), and a
  validation rule's `address` its `sqref`'s first range as stored; both are now
  canonical. `Value.array` (an array formula's range, without its sheet) is now written
  canonically too, as stored only when it does not read as a range. Tool definitions are
  unchanged.

## 0.1.4 (2026-10-26) -- error codes with ooxml-edit's typed refusals

- **The codes, tested** (review, S9). ooxml-edit 0.66.0 raises an `ArgumentError` for a
  value it refuses and a `UnitError` for one out of its range (`ooxml_edit.errors`), and its
  dispatcher says `invalid_arguments` and `unit` for them, in a workbook as in a deck or a
  document. Nothing changes here: `ERRORS` maps the library's own errors as before (none is
  an ooxml-edit `ArgumentError`), each is now tested to keep its code, and nothing mapped an
  ooxml-edit `ValueError` that the typing makes redundant. The tools check what they write
  before ooxml-edit's sheets would refuse it, so no tool's code is known to change. New
  `tests/test_tool_errors.py` (skipped on an ooxml-edit before 0.66.0; the requirement stays
  `>=0.57.0`). Tool definitions are unchanged.

## 0.1.3 (2026-10-26) -- a linter, a type checker and py.typed

- **Linted and type checked in CI** (review, S8): `ruff check .` with a small set of
  bug-finding rules (pyflakes, syntax errors, mutable or call defaults, a return in
  `finally`, a dropped comparison, pylint's errors; no formatting, no import sorting) and
  `mypy`, not strict, over the package, both configured in `pyproject.toml` and run as
  steps of CI's `build` job. The new `dev` extra is the `test` one with both, at the
  versions CI runs. The package ships `py.typed`. What ruff found: unused imports (three
  in the package, none part of its API) and placeholder-less f-strings in tests.
- **What the type checker found**, each fixed rather than silenced:
  - A formula with a name of letters past U+FFFF (`A1+𝐀Rate`) raised `AttributeError`
    from the tokenizer, which took every character from U+00C0 on as starting a name but
    matched names only to U+FFFF. Such a name is now a name.
  - A declared field written `{"header": null}` passed the declaration's checks and
    failed later with `AttributeError`; it is refused, `bad_declaration` at
    `fields[n].header`, as an empty header is.
  - `header_row: null` (and `first_row`/`last_row`) failed with `TypeError` when the
    declaration was read; null is now the row left out, as JSON's `null` usually is (the
    default header row, no first or last row).
  - Everything else was a fact the code already relied on, now written down for the
    checker: an `assert` where a value is known to be there (a region's data rows when a
    row is named in them, a key when an address is by key), a variable no longer reused
    for values of another type, annotations that said less than the code does
    (`Workbook._read`'s `target`, `check_structure`'s `sheets`, `capped`'s mapping). One
    `type: ignore` with its reason (`re._parser`, private and without a stub).

## 0.1.2 (2026-10-25) -- one reading of a number format code

- **Format codes read as the libraries read them** (review, S1): `read.styles.classify_format`
  is now `ooxml_common.numfmt.classify` (same `(kind, elapsed, lcid)`); its own tokenizer is
  gone. The shared reading is held to Excel's answers, measured in Excel for Mac 16.106
  (`numfmt.MEASURED`, 108 codes), and `tests/test_formats.py` runs that table through
  `classify_format`. What changes: a first section with `[$-F800]`/`[$-x-sysdate]` is a
  date and one with `[$-F400]`/`[$-x-systime]` a time (Excel shows the system's long date
  and time there: `[$-x-sysdate]` alone was a number), and `上午/下午` alone is a time;
  every other code Excel opens classifies as before. `set_cells` asks the same reading whether a cell's
  format, or the `format` a date is given, shows a day, instead of ooxml-edit's
  `dates.is_date_code` as well: a Thai or era date code (`[$-th-TH]ว/ด/ปปปป`, `ggge`) is
  accepted, as Excel shows a date in it, and a refused cell format whose code shows no day
  says "give one" (`0;[Red]d` said the format was not known in the workbook's language).
- The ooxml-common floor is raised from 0.15.0 to 0.29.0, the first with `numfmt`.

## 0.1.1 (2026-10-25) -- production feedback and 0.2's edit foundation

What came after 0.1.0: production feedback and 0.2's edit foundation (E1, E3). 0.2 stays
the version that edits (ROADMAP.md).

- **Versions that mean something** (review, N3): this is 0.1.1, the first version tagged
  (`v0.1.1`). The `test` extra asks for `pptx2svg>=0.2.0`, pptx2svg's first tagged
  version. Each version that reaches `main` is tagged by the new tag workflow, and CI's
  `build` job checks that `pyproject.toml`'s version is the one this file's top version
  section names (`tools/check_version.py`). The dependency notes moved from
  `pyproject.toml` to [Dependencies](#dependencies) below. The ooxml-edit floor is raised
  from 0.54.0 to 0.57.0, the first whose `DocumentFormat` takes the `touched` hook the
  workbook format passes (round 13, R2): with 0.56.0 the tools fail to build. Checked by
  running this suite against each: 0.57.0 passes, 0.56.0 does not; ooxml-common 0.15.0
  passes, so its floor stays. No code change.

- **CI runs nightly** (03:27 UTC) and on demand, so a sibling's change that breaks this
  suite shows the next morning; ooxml-edit's own CI now runs this suite on its pull
  requests too. No code change.

- **The shared `save_document` takes `clean`** (ooxml-edit 0.60.0, round 16, F11): what
  Word keeps but never draws taken out of a document's file. A workbook's handler takes
  none, so ooxml-edit refuses `clean: true` for a workbook (`invalid_arguments`, `field:
  "clean"`); no code change here. Workbooks have no rsids or web divisions; orphaned parts
  they can carry (`OpcPackage.drop_orphans`) are not taken out yet. Core 4,230 -> 4,272
  estimated (about 6,323 counted, under 6,500); every definition 4,971 -> 5,013 (guard
  4,975 -> 5,015); strict set's members unchanged.

- **What a call changed, by sheet** (round 13, R2; ooxml-edit 0.57.0): every tool result's
  `touched` (for the application, never in what the model reads) names `{"sheet": name}`
  for each sheet whose part, or a chart or drawing it relates to, changed -- in the
  workbook's order, undo and redo included -- and `{"part": name}` for the rest (the
  workbook part, shared strings, styles, the calculation chain), from the bytes that
  changed (the format's `touched`). Cells are not listed: finding them would compare every
  changed sheet's cells, reading both versions of the sheet. No token change.

- **`xlsx_query` takes a dataset** (round 10, O5; ooxml-edit 0.55.0): `table` may be
  `d:by_region`, no `doc` needed; its rows are filtered and sorted as a table's
  (`xlsx_agent.query.dataset_matches`) over the values as the dataset reports them, each
  with its `d:` address, paged as `xlsx_read` pages a dataset, `dataset` saying where it
  came from. The shared `undo`'s `scope` names a sheet (workbooks) beside a slide (decks).
  Every definition 4,953 -> 4,971 estimated (guard 4,975); strict set unchanged.
- **The contract's gaps closed** (lane X, X6; no ooxml-edit change). *Formula keys*: a
  write by key, a new row or an allocation into a table whose key column holds a formula
  in any data row is refused (`formula_key`: declare a key column of values the
  application supplies); `compare` still pairs such rows by cached keys and says
  `formula_key` on the table (and in `xlsx_compare`). *Explicit allocation*:
  `set_cells(..., allocate={"t:stock/Lamp": 15})` names the free capacity row each new key
  starts in; refused before any change, as `allocate[n]`: `no_capacity`, `capacity_full`,
  `not_capacity`, `data_row`, `not_free` (a formula returning `""` too), `allocated_twice`,
  `existing_key`, `capacity_gap` (the rows allocated continue the data rows, with the next
  free row named). `new_rows=True` (the next free row) stays a Python convenience, refused
  under a policy whose table entry lacks `new_rows` (`not_allocated`). `xlsx_set_cells`:
  each cell's `row` allocates; `new_rows` is gone from the tool. *Structure in the
  integrity compare*: `cells=True` also lists, apart from the cells
  (`CellsDiff.structure`, `StructureChange`, `STRUCTURE_KINDS`), defined names, data
  validations (standard and `x14` as one), merged areas and conditional formats added,
  removed or changed, by name or `sqref` and rule, formulas normalised; bounded (counts,
  first 50); formulas only with `formula_text`; every one outside a write policy
  (`outside`'s `structure`) and counted in `Diff.totals`. The sheet reader keeps conditional
  formats (`SheetReader.conditional_formats`) and the key index the key cells holding
  formulas (`KeyIndex.formulas`). Definitions: `xlsx_set_cells` 533 -> 529, `xlsx_compare`
  395 -> 409; the core 4,198 -> 4,212 (about 6,234 counted by the proxy); every definition
  4,942 -> 4,953 (the guard 4,950 -> 4,960); the strict tool's complexity unchanged (13
  optional parameters, 8 free strings).
- **Formula-aware integrity compare** (lane X, X5; no ooxml-edit change).
  `Workbook.compare(a, b, cells=True)` compares every cell of the sheets both workbooks
  have by address (`Diff.cells`, `xlsx_agent.integrity`, `CellsDiff`, `CellEdit`):
  `formula_changed`, `formula_to_literal`, `literal_to_formula`, `formula_removed`,
  `formula_added`, `value_changed`, and `cache_changed` (the same formula, another cached
  value: counted, never a write) -- a formula changed or turned into its own value is found
  while the cached values agree. Formulas are compared normalised
  (`xlsx_agent.formula.normalise`: names, functions, references and errors case aside,
  ECMA-376 §18.17; whitespace dropped except the intersection operator; string literals
  and structured references as written); a shared formula's member as its expanded text,
  so a group written out cell by cell is no change; an array formula with its block.
  Values as stored (shared, inline and `str` text alike), never through a redaction. With
  `policy` (or `a`'s own write policies) the changes the policy does not allow are listed
  (`outside`: an unwritable cell, a formula under `formulas: false`, a formula replaced
  under `replace_formulas: false`; a sheet added, removed or renamed). `as_dict()` is
  bounded as an edit report: counts, each kind's cells as ranges capped at 50; formula
  text only with `formula_text=True`, withheld for a redacted formula or one reaching
  hidden cells; nothing logged. Tables are compared too only when named.
- **Validation inspection**: `Workbook.validations(sheet)` (`xlsx_agent.validations`,
  `ValidationRule`): each rule `standard` or `x14`, its type, operator (`between` when
  none is written), `allowBlank`, cells, and for a list its source -- `inline`, `range`,
  `name` (followed through names that name names, local before workbook-wide), `table`
  (a table column), `indirect`, `offset`, `formula`, `external` -- whether it resolves,
  where its values are and how many; a dependent or dynamic list is reported unresolved
  by address, with the references it uses, never evaluated. Formulas and values only on
  request. The reader keeps each rule's operator, `allowBlank`, second formula and origin.
  Rules (standard and `x14`) are tested byte for byte through unrelated edits and writes
  into validated cells.
- **A written value checked** (a warning, never a refusal): `Written.invalid` names the
  cells written with a value a resolvable rule does not allow (a resolved list, case and
  surrounding spaces aside; `whole`, `decimal`, `date`, `textLength` with number bounds or
  a bound cell holding a number), `Written.unchecked` those under a rule that cannot be
  read statically; `xlsx_set_cells` returns both and warns of the first.
- **Tools**: `xlsx_compare` takes `cells` and `formula_text` (the session's write policy
  is asked: `cells.outside`); `xlsx_read` takes `validations` (the rules over the target
  as asked, before clipping to the used range, on the first page; their formulas with
  `formulas`). Definitions: `xlsx_compare` 330 -> 395, `xlsx_read` 556 -> 593 estimated;
  the core 4,096 -> 4,198 (about 6,213 counted by the proxy, under 6,500); every
  definition 4,840 -> 4,942 (the guard 4,850 -> 4,950). No strict tool changed.
- **Guarded writes** (lane X, X4; no ooxml-edit change). The host's write policy
  (`xlsx_agent.edit.policy`, `WritePolicy`, `PolicyError`): an immutable allowlist of
  `cells` (cells, ranges, whole sheets) and `tables` (a table's data cells by `columns`;
  `new_rows` its capacity too), with `formulas`/`replace_formulas` ceilings, as plain
  data; `Workbook.open(..., policy=...)` or `xlsx_agent.tools.policy(session, ...)`, only
  narrowed later (`Workbook.narrow_policy`, `tools.narrow`). Without a policy Python writes
  everywhere but protected sheets and a tool session as its `edit` role allows; an empty
  allowlist is read-only. Raw addresses, cells by key, new rows, formulas and batch ops are
  all checked, and `Editor.write` asks again. `describe` shows `writable`. Preconditions:
  any value may carry `expect` (`type` -- `empty`/`blank`, `text`, `number`, `bool`,
  `error`, `formula`; a formula returning `""` is never empty -- and a `value`, `date` or
  `shown` text),
  and `set_cells(revision=)` the workbook's version; an unmet one refuses with the cell's
  current type (`EditError.held`), never its value. A protected sheet (`sheetProtection`)
  is never written (`protected`), unlocked cells included until Excel is measured
  (ROADMAP open question 9); structure protection refuses no cell write.
- **Cells by key and new rows in a capacity** (`xlsx_agent.edit.keys`):
  `t:<table>/<key>/<column>` writes the cell of the row holding the key now, in any region
  whose key column is declared; a duplicate or missing key, a row by place, a key set or a
  wildcard refuses -- no positional fallback -- each named `cells[n]`, never by its key.
  Declared tables take `capacity` (the most data rows; one sheet; the rows below the data
  empty, or `capacity_not_empty`); `set_cells(new_rows=True)` starts each new key's row in
  the first free one (`capacity_full` past it, `no_capacity` without one), and a free
  capacity row is never written by address (`capacity_row`). The table's rows follow the
  cells after every write, rollback, undo and redo; `describe` shows `capacity` and
  `free`. `Written.new_rows` lists the new rows' key cells.
- **`xlsx_set_cells`** takes an address by key, each cell's `expect` (its type now) and
  `expect_value` (its value now as read, as text: a number's digits, `true`/`false`, a
  date `YYYY-MM-DD`; Python's `{"shown": ...}`), `new_rows` and `revision`; a refusal's
  `details` carry `held`. Flat fields: a nested `expect` object made Claude refuse the
  strict schema ("Schema is too complex"); `tests/test_tools_schema.py` now holds the
  strict tools to the complexity Claude compiled (two object levels, 13 optional
  parameters, 8 free strings). Definitions: `xlsx_set_cells` 378 -> 533 estimated, every
  definition 4,686 -> 4,840 (the guard 4,690 -> 4,850); the core unchanged. Excel oracle: rows written
  by key into a capacity, plain, `r`-less and inline-string, open without a repair.
- **Markdown rich text in cells** (lane X, X3b; ooxml-edit 0.54.0): `wb.set_cells({address:
  {"markdown": "**Total** due *now*"}})` writes a shared string of runs (`<si><r><rPr>`)
  -- `**bold**`, `*italic*`/`_italic_`, `***both***`, `~~strike~~`, backslash escapes, code
  spans as text, by ooxml-edit's `ooxml_edit.markdown` (CommonMark's delimiter rules;
  links, HTML and every other block stay text). Each run's `rPr` is the cell's whole font
  with its marks (measured in Excel for Mac 16.106: a run with only `<b/>` shows in
  Excel's default font, not the cell's; ROADMAP *X3b*). Markdown without marks is plain
  text; text values never read Markdown. Each line given is a line of the cell (not
  CommonMark's paragraph joining); one empty line kept between lines; `-`/`*`/`+` items
  show as `• `, `1.`/`1)` items as their number renumbered from the list's first, nested
  items 4 spaces further in per level (`xlsx_agent.markdown`). Text with a line break
  gives the cell its format with `wrapText` (reused or appended); a row with
  `customHeight` keeps its height and the cell is listed in `Written.fixed_height`
  (measured: Excel grows every other row to its wrapped lines on open). Every guard and
  the call's atomicity apply; the footprint is shared strings and cell formats appended.
- **Markdown reads**: `wb.read(..., markdown=True)` (and `cell`, `cells`) give every text
  value as Markdown (`Value.markdown`): rich runs as `**`, `*`, `~~`, `•` and numbered
  lines as list items, every other character escaped, so that `{"markdown": ...}` writes
  the same cell back; plain text stays plain by default; with a redaction each run is
  redacted (a term across runs: the item's redacted text, unmarked). `xlsx_read` takes
  `markdown`; `xlsx_set_cells` a `markdown` value field and says `fixed_height` with a
  warning. Definitions: `xlsx_read` 528 -> 556, the core 4,069 -> 4,096 estimated (about
  6,022 -> 6,062 counted by the proxy); `xlsx_set_cells` 347 -> 378; every definition
  4,627 -> 4,686 (the guard 4,650 -> 4,690).
- **Typed, guarded cell writes** (0.2's E3; lane X, X3; ooxml-edit 0.53.0):
  `wb.set_cells({address: value, ...}, replace_formulas=False)` (`xlsx_agent.edit.cells`),
  replacing the internal hook `Workbook._write`. Text is always literal (`"=x"` too);
  `int`/`float` finite, a whole number past 2^53 refused (write it as text); `bool`;
  `None` a blank that keeps the cell's format (empty text refused: never a 0 or a "");
  `{"formula": "=..."}` (or a `Formula`) the only formula, written with no cached value;
  `{"value": v}`; dates as `{"date": "YYYY-MM-DD"}` or a Python `date` in the workbook's
  1900 or 1904 system, a date string otherwise in the workbook locale's date order with a
  four-digit year (ambiguous ones refused), into a cell whose format shows a day (its
  own, or a new cell's row or column format, ECMA-376 18.3.1.73/18.3.1.13; ooxml-edit's
  `Styles.is_date` and the reader's language-dependent built-in ids), else refused unless
  the call gives `"format"`; a time zone, a time of day, a `time` or a `timedelta`
  refused (a midnight `datetime` is its date). Text over 32,767 characters refused,
  never cut. Guards: a formula cell written over only with `replace_formulas`; a cell of
  an array formula or a data table, a merged area's cell other than its first, an Excel
  table's header cell, a cell named twice refused. A shared formula group's master
  written over hands the group to its next member not written in the same call (its
  formula moved by the new lexer; the group's `ref` the rest), so every other member keeps
  its formula; a master whose formula cannot be moved safely is refused
  (`shared_formula`). Every refusal is checked before anything changes, every refused
  cell listed (`EditError.problems`, the first code's cells in `addresses`), and the
  writes are one step (a refusal caught inside a batch still rolls it back). Returns
  `Written`: the cells, their old types, the groups handed on, the calculation status.
- **`xlsx_set_cells`** in a deferred `xlsx_edit` group, the `edit` role's: `doc`, `cells`
  (an `address` and exactly one of `text`, `number`, `bool`, `date` (with `format`),
  `formula`, `blank` -- typed fields, since the definitions' common subset has no type
  union), `replace_formulas`; at most 5,000 cells a call; a column the application
  withholds from the model refused (`withheld`). The result: `written`, `changed`, each
  cell's `old` type (at most 50 listed), `released`, `calculation`; refusals with every
  problem by code and address in `details`; never a value. Definitions: the core
  unchanged (4,069 estimated, about 6,022 counted by the proxy), every definition 4,280
  -> 4,627 (the guard 4,300 -> 4,650). The prompt's values line says what `cache`
  `invalidated` means (one sentence).
- **Formula freshness** (P0-4): every formula value read carries `cache`:
  `cached_unverified` (as last saved, no known invalidation, never checked to be current),
  `invalidated` (the workbook was written since it was opened -- every formula, workbook
  wide -- or arrived with `fullCalcOnLoad` or manual calculation) or `absent` (no cached
  value); and `external` when it refers to another workbook (`[1]Prices!B2`, `[1]!Rate`,
  or a defined name that does). In `xlsx_read` and `xlsx_query` per value; `xlsx_compare`'s
  cells; `xlsx_derive`'s result and `xlsx_validate_data`'s report and `xlsx_profile`'s
  columns as counts by status (`cache`); `describe`'s workbook `cache: "invalidated"`,
  and `external_formula` among the problems and a column's flags (beside the existing
  `external_links` count). A dataset or a `t:` source read from an invalidated cache is
  stale from the start (ooxml-edit's `Dataset.cache`), whatever the versions say: the deck
  and Word tools place it `stale`, with the counts in their `from`; a dataset derived
  before an edit is current again once it is undone. No evaluator: nothing is
  recalculated.
- Formula references (a part of E2): `xlsx_agent.formula`, a lexer of references only --
  strings, quoted and unquoted sheets, workbook indexes, cells and ranges with `$`, whole
  columns and rows, structured references, errors, functions and names -- with `shift`
  (the relative parts moved; names, structured references and strings untouched; off the
  grid or unreadable refused) and `is_external`.
- Tests: `tests/test_set_cells.py`; the shop fixture gains `array`, `external` and `calc`
  variants. Oracle: see the ROADMAP's E3 row.

- **The edit foundation** (0.2's E1; lane X, X2; ooxml-edit 0.51.0): `xlsx_agent.edit`.
  `with wb.batch() as edit:` is one step of ooxml-edit's `History` over the workbook's
  package: an error anywhere in it rolls everything back -- cells, shared strings, styles,
  `calcPr`, the calculation chain with its relationship and content-type entry -- and keeps
  the version (datasets stay current); an error caught inside it still rolls it back
  (`EditError` `rolled_back`); `wb.savepoint()` rolls back one part (the tool batch's ops).
  `wb.undo()` / `wb.redo()`: undo to the start is the bytes opened, redo the edited bytes;
  `wb.version` (`wb.revision`). The first write drops the calculation chain and sets
  `fullCalcOnLoad`; `wb.calculation_status` is `invalidated` then (`cached_unverified`
  before), read from the package, so an undo restores it. Each step closes with the
  shared strings' counts brought up to date and leaves an `EditReport` (`edit.report`,
  `wb.last_edit`): the cells written as ranges (capped at 50), the parts changed, the
  calculation status and the structure issues; never a value or formula, nor in the
  `xlsx_agent.edit` logger. `wb.validate()`: the structure check by code and address
  (`xlsx_agent.edit.structure`). `wb.save(target, kind=None)` and `wb.to_bytes(kind)`:
  the kind's main content type; `.xlsx`/`.xltx` refused while the workbook holds a VBA
  project; the saved bytes reopened and every written cell read back first (`readback`).
  The writes themselves are E3's; until then the internal hook `Workbook._write` writes
  one cell. Imports moved from `ooxml_edit.charts.workbook` to `ooxml_edit.sheets`
  (requires ooxml-edit 0.51.0).
- The shared `save_document`, `undo` and `batch` serve workbooks, in the core (E1):
  `save_document`'s `format` is `xlsx`, `xlsm` or `xltx`, its validation report the
  structure check against the one at open (a new issue refuses the save), its `saved` the
  cells read back and the calculation status; `undo`'s `scope` a sheet's name (a step that
  changed the parts every sheet shares is taken back by a plain undo only). The format's
  `problems` (what `open_document`'s `baseline_problems` counts) are now the structure
  issues; the cell facts stay `describe`'s. Definitions: the core 3,523 -> 4,069
  estimated (about 5,214 -> 6,022 counted by the proxy), every definition 3,746 -> 4,291
  (the guard 3,750 -> 4,300).
- The Excel oracle (E1): `pytest -m oracle` (never under xdist) drives Excel for Mac
  through `tests/oracle.py` and `tools/excel_resave.applescript`: one machine-wide lock in
  the Office group container, skipped when Excel runs or the lock is held, a repair prompt
  answered No and reported, only an Excel it started quit, results cached by content.
  The shop fixture (`tests/shop.py`, made in code) in five variants, as made, edited and
  undone: 15 of 15 open in Excel for Mac 16.106 without a repair; the control (a
  shared-string index out of range) is asked to repair.

- What a name that is no workbook tool means (round 9, N5; ooxml-edit 0.47.0): "no tool
  'xlsx_pivot': use xlsx_derive", with the best few tools in `valid_options` instead of
  every one. ooxml-edit reads the tools' names; `TOOL_HINTS` (`FORMAT.tool_hints`) adds the
  names it misses or settles wrongly: listing sheets or tables is `describe` (not
  `xlsx_read`), filtering or looking up rows `xlsx_query`, pivoting, grouping or
  aggregating `xlsx_derive`, statistics `xlsx_profile`, a diff `xlsx_compare`, validating
  `xlsx_validate_data`. No workbook tool runs in a batch, so a batch op's comment (N4) does
  not arise here. Data and tests only: definitions unchanged.
- Agents sharing one session (round 8, M1; ooxml-edit 0.46.0): views of one session
  (`session.view(roles=..., context=...)`) share the workbook and its datasets, each held
  to its own roles. Tests and the tools README only: no workbook tool or definition
  changed.
- Tool roles (round 7, item 9; ooxml-edit 0.45.0): every workbook tool is `review`'s, so a
  reviewer (`Toolbox(roles=["review"])`, `toolbox.session(roles=...)`; with `resolve` or
  `accept` alike: no workbook tool has comments or tracked changes) is sent the
  default's definitions unchanged (core 3,523, every definition 3,746 estimated; no strict
  tool) and runs every tool; a role set without `review` is sent only the session's tools
  and every read is refused naming the role. A reviewer's prompt has the role line, not
  told it may comment. Tests in `tests/test_tools_roles.py`; definitions unchanged.

- The loaded core's budget is now 6,500 tokens as count-tokens counts them (raised from 6,000
  by the user on 2026-10-17); the offline proxy and the online check use it.

- **`xlsx_read` names a row by key one way** (L8): the row is the target
  (`t:Parts/BX-17`); `key_column` is now `key`, as `xlsx_compare` names it, and
  `columns` picks a `t:` target's columns in order (`["Bin", "Count"]`). `key_column` and
  `row` are still read (as `key` and `target`, with a warning). A misread `key` works
  where it can only mean one thing -- a column that does not hold the key is read as a
  column to read, the row's key itself as the row -- with a warning, else is refused with
  what to do; several key columns are refused with the ones holding the key; a key not
  found names the key column that has it, or the rows with a key like it by position
  (`KeyIndex.near`). Definitions: the core 3,496 -> 3,523 estimated, every definition
  3,718 -> 3,746 (the guard 3,720 -> 3,750); the strict set unchanged.
- From ooxml-edit 0.36.0 (round 6, R5, R6): a text fact with a suffix that is no number
  format (`{{d:x/#1/avg:0.20}}`) says so and lists the forms; find_text's `slides` (decks
  only) takes slide ids: the core 3,490 -> 3,496 estimated, every definition 3,712 ->
  3,718 (the guard 3,716 -> 3,720).

- **Row ranges** (T3): `t:Products/#9:#16` names data rows 9 to 16 by position (from 1,
  the last included, as `#5` names one), `t:Products/#9:` row 9 to the last;
  `/#9:#16/Price` their one column, `/#9:#16/*` every column; `d:by_region/#2:#5` the
  same in a dataset. `#16:#9` is the same rows, `#9:#9` is `#9`, `#9:16` reads as
  `#9:#16`. `xlsx_read`, `find_text`'s `range`, `Workbook.read`/`resolve`/`expand`/
  `dataset` and a deck or document tool's `from` take them -- a slide table of rows 1-9 of
  a long table is `from: "t:Products/#1:#9"` -- with the withholding as before. A row
  past the last is refused (`position_out_of_range`, the range there is as the option):
  nothing is cut. `TableAddress.row_range` and `DatasetAddress.row_range`. The tool
  definitions are unchanged; the prompt fragment names the form.
- **Terms redacted from every text** (T15): `{"redact": ["Acme Corp", "re:Project\\s+\\w+"],
  "with": "[redacted]"}` in the declarations replaces each match -- literal terms case and
  white space aside (`whole_words` optional), or bounded regular expressions -- in every
  text the workbook gives: cells however stored, headers, comments, links, drop-down
  lists, number formats, an Excel table's column names, sheet names (shown and addressed
  redacted; declarations still name them as stored). A formula holding a match keeps its
  value, not its text; an Excel table or defined name whose name holds one is not listed;
  `find_text` never finds a term. The model sees only the replacement; `Workbook.redact`,
  `Workbook.redaction` and `Sheet.stored_name` for a program. No definition changes.
- **A hidden column ends at the next declared section**: a declared table's hidden
  column is hidden to just above the next table declared below it on the sheet, not to
  the sheet's last row, so a later section may show that column (was `hidden_elsewhere`).
- **Whole sheets hidden from the model** (T14): `{"sheets": "Notes", "hidden": true}` in
  the declarations withholds a sheet entirely -- it is no sheet of the workbook for
  `describe`, `open_document`'s count, reads by name or id, its Excel tables, the defined
  names on it, local to it or whose formula names it, `find_text`, the problems, a
  comparison, a drop-down drawn from it, a link to it and a declared table's glob; each
  refused exactly as a sheet, table or name the workbook does not have, the name never
  offered or listed. A formula elsewhere that reaches it keeps its value, not its text.
  `Workbook.withhold_sheets` and `withheld_sheets` for a program. No definition changes.
- **A declared table's header found by its text** (T16): `"header_row": {"match": ["Ref",
  "Item"], "rows": "8:40"}` finds the one row holding every text given (case and spaces
  aside), on each sheet, within `rows` or below the tables declared before it there; its
  rows run to the first blank row or the next match. None or several rows are refused
  (`no_header`, `ambiguous_header`) by row number. No definition changes.

- **`describe` lists the session's datasets** (T11; ooxml-edit 0.30.0, which it now
  needs): every `describe` -- of a workbook, a deck or a document -- gives `datasets`, each
  with id, columns, `types`, `rows` (a count) and `from`: `d1/t:Sales` for an
  `xlsx_derive` from a workbook (`define_dataset(..., origin=)`), `d:by_rep` for one from
  a dataset. Never a value. A derive from a declared table cannot take its hidden columns,
  and `table_dataset` leaves them out, so no listed dataset names one (tested). No
  definition changes.
- **Each compared row named by its label** (T19b): `xlsx_compare`'s row and cell
  differences carry the row's `label` -- its text in the table's first text column other
  than the key (named in the table's entry as `label`), visible cells only, at most 60
  characters -- so an answer names a row by what it is, not by its key alone. A column
  hidden from the model is never the label. Golden g2 is not re-recorded: its replay adds
  the labels to the recorded result. No definition changes.
- **Clearer descriptions** (T19): `xlsx_read`'s `key_column` says it is the column holding
  the key, `sources` that it is for `d:` only (its refusal on a `t:` or cell read says
  where it works and what to do instead), and `xlsx_derive`'s `sort` that it names output
  columns; a sort by another name is refused as "sort names an output column: one of
  ...", with the measure's column for an input column it reads, else a did-you-mean.
  Every definition 3,696 → 3,712 estimated (guard 3,700 → 3,716); the loaded core 3,474 →
  3,490.

- **A sheet's dates placed in the locale** (T18b): a date column's ISO text from `t:`, a
  derive's dates and a dataset labelled `date`/`datetime`/`time` are written in a deck or
  document in the session locale's day order, as a Python date is (`1/5/2026` en-US,
  `5/1/2026` nl-BE), and a chart's categories match them written ISO or in the locale
  (ooxml-edit 0.29.0). The workbook's own results and provenance keep ISO 8601; no cell's
  display text is produced, so the locale's writing is the only one. No definition
  changes.

- **Dates in a dataset made in Python** (T18): a Python `date`, `datetime` or `time` in a
  dataset is that type with no label (it was `text`, so `min` and `max` refused it), read
  as ISO 8601 text as a cell's date is; mixed with other values, the column takes the
  majority type as before. ISO text in an unlabelled column stays text (a column of codes
  that look like dates is not read as dates on its own) and is flagged `date_as_text`, so
  a refused `min` names the way out: the `date` label or `coerce: ["date_as_text"]`. Build
  one with ooxml-edit 0.28's `Dataset.from_records(rows, types={"opened": "date"})`; the
  labels are its `TYPE_LABELS` (`number`, `percent`, `date`, `time`, `datetime`, `text`,
  `bool`, `error`), now in the docs. In a deck or document a Python date is written in the
  session locale's day order (ooxml-edit); a derived date stays ISO text. Needs ooxml-edit
  0.28.0. No definition changes.

- **Golden g2 re-recorded** on the current `xlsx_compare` result (summary first, differences
  by key): the test no longer converts the replay back to the pre-X10 shape.
  `tools/record_goldens.py` writes the day it ran as the recording date.
- **Formulas over a hidden column** (T17's gaps): a visible cell's formula that reaches a
  cell a declared table hides -- by reference (`=D2`, `=SUM(Log!D:D)`), structured
  reference (`ChecksTable[Note]`, `[@Note]`, the whole table), a defined name over it or
  reaching it, `INDIRECT` or `OFFSET` -- is read with its cached value and without its text:
  `Value.formula_withheld`, and `"formula_withheld": true` in place of `f` in `xlsx_read`
  with `formulas`. A hidden column is now hidden from its header row to the sheet's end,
  below its table's rows too (cells above the header stay visible). `open_document`'s
  `tables` counts the declared tables. No definition changes.
- **A hidden column stays hidden, every way in** (T17): a declared table over an Excel
  table replaces it -- `describe` lists the declared table once, and `t:<the Excel table>`
  (with a column, row, key set or wildcard, in every tool) is refused as `declared` with
  the declared table as `valid_options` -- so a column the declaration hides is no longer
  shown through the Excel table's entry (its name, values and sample rows) or address. A
  defined name over hidden cells is refused and not listed or searched; a validation list
  drawn from them is not read (`reason: "withheld"`); `Workbook.read` gives them blank;
  another declared table that shows them is refused (`hidden_elsewhere`).
  `Workbook.tables` leaves a replaced Excel table out (`Regions.excel_tables` has all).
  No definition changes.
- **A title is no region**: a lone cell with nothing under it (a title, a caption, a lone
  number or date) is no longer listed by `describe` as a detected region with no rows, over
  a declared table or not; it stays a cell for `find_text` and reads. Text two rows above a
  detected region's header (a blank row between) is that region's `title`
  (`Region.title`, `Region.caption`). On the end-to-end workbook `describe` drops from 694
  to 552 estimated tokens. No definition changes.
- **Comparisons by key, summarised per column** (X10): `xlsx_compare` names every cell
  changed by key, `t:Products/SKU-2/Price` (in `at`, in place of `table`, `key` and
  `column`; the values keep their sheet addresses), and every row added or removed by key,
  `t:Products/SKU-6` (its `at` was its position). Each table's entry opens the result with
  a summary on the first page: per column the cells changed, and for numbers, percents,
  dates and times how many went up and down and the largest move each way by key
  (`"max_up": "t:checks/A190/rating +2.5"`, a date's in days, `+7d`); text and lists count
  as changed only; the rows added and removed by key (ten each, the rest counted). In
  Python, `CellChange.at`, `CellChange.delta`, `RowChange.at`, `TableDiff.by_column()`
  (`ColumnSummary`). No new parameter; the description grows by 12 estimated tokens.
- **Cleaning text** (X6): a declared field takes `clean` -- `strip` (a leading label, a
  pattern matched at the start within the cell's first 256 characters), `collapse_space`
  and `split` (a literal separator: the value becomes a `list`, a new value type; a tuple
  in Python, a list in JSON). The caller's configuration, so no definition changes;
  applied once, where the table is read (reads through it, describe, query, profile,
  derive, compare, checks, a deck's or document's `from`), never to the header, a hidden
  column or the key column (`clean_key`). A changed value says `cleaned` and keeps
  `original`; `xlsx_read`/`xlsx_query` pages count them (`cleaned`), `xlsx_profile` and
  `describe` flag them, `xlsx_derive` counts them per column (`Derived.cleaned`), and a
  `t:` source warns (`CleanedWarning`, among the deck tool's warnings) and places a list
  as one cell with a line per item. Lists: `eq`, `in`, `contains` pass when any item does;
  `count` counts them; other functions, `group_by` and `pivot` refuse them; validation
  checks their items as text. Patterns are compiled when declared and refused
  (`bad_pattern`) when their backtracking could grow past the square of what they see.
- **Declared tables** (X3) and **key-value tables** (X7): the application declares its own
  tables from its configuration, as plain data -- `Workbook.declare_tables`, the format
  (`xlsx_agent.tools.format(tables=...)`) or a session (`xlsx_agent.tools.declare`); no tool
  takes them, so no definition changes. Sheets by name or glob; a table over several
  sheets with one header, read as one, every value keeping its own cell's address and keys
  found on any sheet; header row, first and last data rows, columns; a key checked unique
  across the sheets; column ids (the header still finds the column) and descriptions,
  shown by `describe` (`about`) and `xlsx_profile`; columns hidden from the model, left out
  of every tool's results and refused when named, their cells `withheld` from reads of the
  sheet and from `find_text`. The sheet is not a column (no `_sheet`). Key-value tables,
  `t:run/<field>/value`, shown compactly by `describe`; a two-column block of facts with no
  header is detected as one -- so `describe` of a workbook with no declarations changes
  where such a block was listed as `header_unclear` (`header_not_text`): it is now a
  region, with its fields. Problems met only in a hidden column's cells are not listed;
  one met there too is counted and placed at its other cells only. A declaration that does not fit is refused naming the entry
  (`DeclarationError`: code, path, valid options), before any table is kept; at
  `open_document`, `refused`. `Value.origin` and `Value.place` give a value's own cell;
  `Region.where` its table's range on the sheets.
- **Tables as sources** (X1; ooxml-edit's O3, whose 0.22.0 is now the floor): the xlsx
  format registers a `t:` resolver and `merge_sources`, so pptx-agent's `ppt_add_table` and
  `ppt_edit_table` and the shared `edit_chart` take a workbook's table as their rows
  (`from: "t:Products"`, `t:Products[Price]`, `t:Products/*/Price`, a key set's rows). Values
  typed as `xlsx_read` gives them, nothing coerced; each value's source is its cell by key
  where the table has a key column (as a derive's), else its cell address, and a placed
  column's are merged into ranges or one key set. One value is refused
  (`invalid_arguments`, `not_rows`), and a table over `Limits.max_dataset_cells` before a
  cell is read (`limit`). The workbook is only read. In the library:
  `Workbook.dataset(address)` (`table_dataset`) and `merge_sources`. No tool definition
  changes; the workbook prompt says that a slide or document table takes `t:` or `d:`.
  Tests fill a deck's table and chart from a workbook in one session (pptx-agent and
  pptx2svg join the `test` extra).
- **Dates** (X4): `xlsx_derive`'s `first` and `last` (row order, any column);
  `coerce=["date_as_text"]`; a refusal of a column of dates or numbers stored as text names
  the issue, its count and first cells, and the coercion that reads them; `date_as_text`
  is its own reason for a cell left out; `xlsx_profile` gives the range of dates stored as
  text (`date_text`).
- **Addresses by key** (X2): `*` for every row or column (`t:Products/*/Price`,
  `t:Products/SKU-0042/*`) and key sets (`t:Products/{SKU-0042,SKU-0043}/Price`) in
  `xlsx_read` and the library; `Workbook.expand` gives what an address covers one cell at
  a time, by key (a coverage selector); `Workbook.declare_key` names a table's key column.
  A derived value's sources are written by key when its table has a key column, and so
  still name the same cells after rows are inserted or sorted (ranges otherwise, as
  before); results count sources in units (a range, or one key). A key that is `*` or
  begins with `{` is now quoted: a bare one no longer parses as a key.
- **Text lengths** (X5): `xlsx_profile` gives median and max words beside characters and
  the address of the longest text; the `words_gt` filter finds text over a word limit.

Requires Python 3.11-3.15, lxml, ooxml-edit 0.57.0 and ooxml-common 0.15.0 (neither on
PyPI yet: install them from source first).

## 0.1.0 (2026-10-13) -- read and derive

The first version: workbooks as a source an agent can describe, query and derive from,
every derived figure traceable to its cells. Nothing is edited or saved yet (0.2).

- **Open** (R1): `Workbook.open` from a path, bytes, a binary file or an `OpcPackage`;
  `.xlsx`, `.xlsm`, `.xltx`, `.xltm`; `.xls`, `.xlsb` and encrypted workbooks refused with
  an `OpenError` code; ooxml-edit's zip limits checked before unpacking; `to_bytes`
  byte-identical.
- **Addresses** (R2, R4): cells, ranges, whole sheets (`Stock!`), quoted sheets, sheets by
  id (`#3!`), tables and names (`t:Products`, `t:Stock!Prices`), detected regions
  (`t:Stock!A1`), columns, values by key, rows by position, datasets (`d:by_region`).
- **Cell reader** (R3): a streaming pass with shared strings indexed once; typed values
  (number, percent, date, time, datetime, text, bool, error, blank) in both date systems;
  cached formula values and `no_cached_value`; merges, arrays, comments, hyperlinks;
  `r`-less rows; numbers and dates stored as text recognised in the caller's locale.
- **Tables, names and regions** (R4): ListObjects, defined names over ranges,
  `declare_region` for merged or multi-row headers, single-row header detection, key
  columns.
- **Profile and catalogue** (R5, R6): per-column type and mix, fill, min/max/mean,
  distinct values, text lengths, formula share, validation lists and flags; a catalogue of
  the workbook within a token target.
- **Query, find** (R8, R9): filtered, sorted, paged rows; text found by address and kind.
- **Derive and datasets** (R10, R11): count, sum, mean, median, min, max, share and
  distinct, with `where`, `group_by`, `bins`, `pivot`, `sort`, `top` and `round` (half away
  from zero); every value with its source cells and the cells left out by reason; kept as
  ooxml-edit datasets (`d:<id>`), read back by address and derived from again.
- **Compare** (R12): two workbooks by sheet id, table name, row key and cell type.
- **Validate** (R13): a declared schema -- sheets, tables, keys, column types, required,
  unique, ranges, steps, enums, validation lists, patterns -- with problems by code and
  address, never cell text.
- **Tools** (R6-R13): `xlsx_agent.tools` for ooxml-edit's `Toolbox`: the shared
  `open_document`, `close_document`, `read_blob`, `describe` and `find_text`, and
  `xlsx_read`, `xlsx_profile`, `xlsx_query`, `xlsx_derive`, `xlsx_compare`,
  `xlsx_validate_data`. The locale is the session's (`Toolbox(locale=...)`); the reader's
  `grouping` and `local_dates` are `xlsx_agent.tools.format(...)`'s.
- **Performance** (P1): a row index so a later page starts near its rows, passes that type
  only the columns they look at, a one-pass `describe`, `find_text` reading only what may
  match.
- **Release** (R14): the loaded core counts 4,958 tokens on Claude (claude-sonnet-5-5's
  count-tokens; budget 6,000), checked by `tests/test_tools_online.py` (marked
  `provider`); the tools README per provider (`src/xlsx_agent/tools/README.md`);
  `docs/api-tour.md`; golden transcripts of three read-and-derive tasks recorded on
  Claude and replayed offline (`tests/test_golden.py`).

Requires Python 3.11-3.15, lxml, ooxml-edit 0.20.0 and ooxml-common 0.15.0 (neither on
PyPI yet: install them from source first).

## Dependencies

lxml, ooxml-edit and ooxml-common, and nothing else at runtime (ROADMAP.md, "Decisions
already made", 6). Neither sibling is on PyPI yet: install them first from sibling checkouts
(`pip install -e ../ooxml-common -e ../ooxml-edit`) or from their repositories, as CI does.
`pyproject.toml` names versions, never URLs, since PyPI refuses a distribution whose
dependencies name one. Each floor is the first version with what this package uses; what
each earlier version added is in that package's CHANGELOG. `pyproject.toml`'s version is
the one this file's top version section names, and `xlsx_agent.__version__` the same
(`tools/check_version.py`, run by CI's `build` job).

- `ooxml-edit>=0.95.0`: the format-neutral editing core the siblings share (the OPC
  package with parts as their original bytes, the safe parser, the history, the tool
  layer with workbooks a kind of the shared tools, the session's locale and datasets, the
  sources resolver hook, and `ooxml_edit.sheets`, the lossless cell writer the edits and
  addresses are built on). 0.54.0 carries `ooxml_edit.markdown` and `sheets.RichText`,
  the Markdown cells (X3b); 0.57.0 the format's `touched` hook, by which a result names
  the sheets a call changed (round 13, R2); 0.70.0 the shared tools sent with a workbook's
  parameters only (round 23, V3); 0.72.0 `define_dataset(also=...)`, by which a joined
  dataset goes stale when the joined side changes (round 25, X1); 0.78.0
  `shared.READ_BATCH`, many reads in one call (round 32, R15); 0.84.0 a check of numbers
  in a deck's or document's text against the open workbooks (`value_sources`,
  `Dataset.formats`; round 36, R28); 0.95.0 a cell's format and a row's height and format
  set on their own (`Worksheet.set_style`, `set_row_height`, `set_row_style`, `cells_in`,
  `Styles.add_format`; round 44, R53) -- the floor.
- `ooxml-common>=0.42.0`: standard library only, named for what is imported from it
  directly: the package kinds (`xlsx`, `xlsm`, `xltx`, `xltm`), themes and measured text
  widths, from 0.15.0 the shared locale (`locales.Locale`), the reader's locale, from
  0.29.0 the shared reading of a number format code (`numfmt`), and from 0.42.0 a face's
  own advance widths (`text.fontmap.own_metrics`) with Aptos Narrow's -- the floor.
- `lxml>=4.9`: the XML tree every part is read and written in.
- `pptx-agent` and `pptx2svg>=0.2.0` (`test` extra only): a deck filled from a workbook's
  table in one session (`tests/test_sources_deck.py`), and a deck's text checked against a
  workbook's values (`tests/test_numbers_deck.py`, pptx-agent 0.1.17); 0.2.0 is pptx2svg's
  first tagged version.
- `ruff==0.17.2` and `mypy==2.4.0` (`dev` extra only, with the `test` requirements):
  pinned, as CI's `build` job installs them, so a new release changes nothing until both
  are raised together.

