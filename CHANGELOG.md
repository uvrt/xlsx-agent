# Changelog

xlsx-agent has not been released to PyPI; it is installed from source. [`ROADMAP.md`](ROADMAP.md)
has the detail of every milestone.

## Unreleased

- Tool roles (round 7, item 9; ooxml-edit 0.45.0): every workbook tool is `review`'s, so a
  reviewer (`Toolbox(roles=["review"])`, `toolbox.session(roles=...)`) is sent the
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
