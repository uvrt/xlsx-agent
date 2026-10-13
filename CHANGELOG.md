# Changelog

xlsx-agent has not been released to PyPI; it is installed from source. [`ROADMAP.md`](ROADMAP.md)
has the detail of every milestone.

## Unreleased

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
