# Changelog

xlsx-agent has not been released to PyPI; it is installed from source. [`ROADMAP.md`](ROADMAP.md)
has the detail of every milestone.

## Unreleased

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
