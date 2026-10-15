# Roadmap

Working plan for `xlsx-agent`: workbooks as a source an agent can describe, query and
derive from without typing a number itself; later, edits; later still, rendering. The
third library beside `pptx-agent` and `docx-agent`, on the same foundations. Written before
any code, to be picked up cold: every item says what it builds, what it reuses, and how to
know it is done.

**Effort key:** S ≈ half a day · M ≈ 1–3 days · L ≈ 1–2 weeks · XL ≈ 3+ weeks.

The loop is the siblings'. The agent never writes XML and never does arithmetic. It reads
a catalogue of the workbook, asks for rows, asks the library to derive figures, and gets
back values that each say which cells they came from. Two things make workbooks different
from decks and documents, and most of this file is about them:

1. **The numbers are the content.** A report built from a workbook is only as good as the
   trail from each figure back to its cells. The library computes; every derived value
   carries its source cells; nothing a model typed is ever presented as data.
2. **Workbooks are messy.** Merged and multi-row headers, hidden sheets, numbers stored as
   text, formulas without a cached value, cells cut at 32,767 characters. The library
   reports what it met, by address and code, and never guesses past it.

---

## Decisions already made

These are settled and are not reopened anywhere below.

1. **Header detection.** Excel tables (ListObjects), defined names and caller-declared
   regions come first. Detection is only for a simple single-row header over a block of
   data. Merged or multi-row headers are declared by the caller (`declare_region`, with
   the header names given), never guessed.
2. **Cross-format datasets.** A derived dataset lives in a session-level store in
   ooxml-edit's `Toolbox` session. `d:<id>` is one more address, usable by pptx-agent's and
   docx-agent's table and chart tools through the existing ref mechanism. It needs an
   ooxml-edit change (O2, O3 below), which is a dependency milestone.
3. **Formulas: cached values only.** A formula's value is the one Excel cached. A formula
   without a cached value is reported, never guessed. No recalculation is in scope; it may
   be revisited on demand.
4. **Locale is a caller setting** (default: `.` as the decimal separator). A number format
   with an explicit locale tag, such as `[$-409]`, is honoured for that format.
   *Shared locale (between R9 and R10, branch `locale-shared`):* the locale is the shared
   `ooxml_common.locales.Locale` (ooxml-common 0.15.0; default en-US), as in pptx-agent and
   docx-agent; in the tool layer it is the session's (`Toolbox(locale=...)`, ooxml-edit
   0.19.0). Whether grouped text numbers and non-ISO text dates are recognised are reader
   settings, not locale facts: `grouping` (on) and `local_dates` (off).
5. **Commit scope.** 0.1 is read and derive; 0.2 is edit. 0.3 (an `xlsx2svg` renderer,
   with `render` and `check`) stays on this roadmap as a later, separate project about the
   size of docx2svg. When that time comes, what Excel shows is measured from Microsoft
   Excel (installed on the development Mac); no Office data is transcribed from
   third-party sources.
6. **Dependencies.** lxml only at runtime, like ooxml-edit, plus ooxml-edit and
   ooxml-common. Python 3.11–3.15.

---

## Goals and non-goals

**Goals**

- An agent learns what a workbook holds from one bounded catalogue: sheets, tables,
  columns with inferred types, fill rates, ranges, and what is odd about them.
- An agent reads, filters and finds rows, bounded and paged, with typed values.
- An agent asks for figures (counts, sums, means, shares, groups, pivots) and gets values
  computed here, each with its source cells, the rows left out and why, and the rounding
  applied.
- Those figures become a dataset (`d:<id>`) that a deck's or document's table and chart
  tools take directly, so no number is retyped between formats.
- Two versions of a workbook can be compared by row key, and a workbook checked against a
  declared schema, with every problem given an address.
- (0.2) Lossless edits: cells, tables, formats, validation, sheets; everything not edited
  is written back as the bytes that were read, and every edit undoes to the original bytes.

**Non-goals**

- **Recalculation** (decision 3). No formula evaluator, small or external.
- **Running macros.** A `.xlsm`'s VBA project is kept byte for byte and never read or run.
- **Following links.** External workbook links, DDE links, web queries and data
  connections are reported, never opened or refreshed.
- **The binary formats.** `.xls` and `.xlsb` are refused with a clear error.
- **Inserting and deleting rows and columns** (0.2). Every reference in the workbook would
  have to move with them; see *Open questions*.
- **Authoring pivot tables, slicers, sparklines, form controls.** Kept losslessly; not
  created or edited.
- **Showing a value as Excel shows it** before 0.3: no display strings are produced
  until they can be measured against Excel (see *Values and types*).
- **Editing a workbook that is open in Excel**, co-authoring and AutoSave.
- **A Python escape hatch and MCP.** As in the siblings: no code execution, no raw XML or
  package tools, no paths in the tool layer.

---

## What is reused, and from where

Nothing below is rewritten when a sibling already has it.

| From | What | How |
| --- | --- | --- |
| `ooxml_edit.opc` | `OpcPackage`: parts as original bytes, lazy trees, content types, relationships, snapshots, `to_bytes` with entry order kept | the workbook's package, as in the siblings. `open_embedded` lets the same reader open a chart's embedded workbook from a deck or document (later; see *Open questions*) |
| `ooxml_edit.xml` | `parse_xml` (`resolve_entities=False`, `huge_tree` off), `qn`, `register_child_order`, `insert_in_order`, `remove` | every part parsed through it; the worksheet child order registered once (0.2) |
| `ooxml_edit.history` | `History`: undo, redo, nested batches, `version` | 0.2's edits; 0.1 uses `version` to key caches and datasets |
| `ooxml_edit.tools` | `Toolbox`, `Session`, `DocumentFormat`, `@tool`, parameter helpers, `Result`/`ToolError`, `page_list`, `truncate`, refs (`Call.define_ref`, `resolve_ref`), `batch`, `Limits`, `check_package` (central-directory zip-bomb guard), the adapters | the whole tool layer; xlsx-agent adds its `FORMAT` and its tools only |
| `ooxml_edit.tools.shared` | `open_document`, `close_document`, `read_blob`, `describe`, `find_text`; in 0.2 `save_document`, `undo`, `batch`, `set_properties`, `edit_chart`, `new_document` | xlsx handlers for the shared definitions (O1) instead of `xlsx_describe` and `xlsx_find` |
| `ooxml_edit.charts.workbook` | `Cell`, `Area`, `SheetRange`, `parse_formula`, `column_letters`, `column_index`, `MAX_ROW`, `MAX_COLUMN`, `number_text` | imported as they are (0.1) |
| `ooxml_edit.charts.workbook` | the lossless cell writer: cells in row and column order, `spans` and `dimension`, shared strings with `count`/`uniqueCount`, inline strings where there is no table, a replaced formula removed with the calculation chain, tables that follow their data and are named after their headers | **lifted, not copied** (O4): moved within ooxml-edit to a public `ooxml_edit.sheets`, with its history, and given what a whole workbook needs; `charts.workbook` re-exports it, so chart edits are unchanged. 0.2's `xlsx_set_cells` and `xlsx_write_table` are built on it |
| `ooxml_common.kinds` | `kind_for`, `kind_of`, `main_content_type`, `kind_mismatch` for `xlsx`, `xlsm`, `xltx`, `xltm` | the open check and, in 0.2, the save check |
| `ooxml_common` themes and text measurement | theme fonts and colours; measured advance widths | 0.2's column auto-fit and colours; 0.3's renderer |
| `pptx2svg` | the chart renderer | 0.3: charts on sheets, as pptx2svg draws them |

What is **not** reused from `charts.workbook`, and why: its reading side (`Worksheet._read`,
`Workbook.shared_string`) parses the whole sheet and finds a shared string with a
`findall` per lookup, and knows no styles, dates or types. That is right for a chart's few
cells and wrong for a sheet of 100,000 rows. xlsx-agent's reader is its own: a streaming
pass over the sheet with the shared strings indexed once (R3).

---

## Addresses

An agent holds addresses across calls. Every result names what it returns by address;
every tool takes the same forms. Sheet and table names are part of an address and may
appear in results and diagnostics; cell content may not appear in diagnostics.

| Form | Example | Means |
| --- | --- | --- |
| cell | `Stock!F2` | one cell |
| range | `Stock!A1:H40` | a rectangle; `Stock!A:A` and `Stock!2:5` are clipped to the used range |
| whole sheet | `Stock!` | every cell of the sheet, clipped to its used range; `'Q3 sales'!` and `#3!` as for ranges. `find_text`'s `range` takes it |
| quoted sheet | `'Q3 sales'!B2:B9` | Excel's quoting: `'` around a name, `''` for a quote inside |
| sheet by id | `#3!F2` | the sheet whose `sheetId` is 3: survives a rename. Every result gives both forms the first time it names a sheet |
| table | `t:Products` | an Excel table, a defined name over a range, or a declared region |
| sheet-scoped name | `t:Stock!Prices` | a defined name local to one sheet, written with its sheet as Excel writes it; `t:Prices` never reaches it |
| detected region | `t:Stock!A1` | a region found by header detection, named by its header's first cell (a table name cannot contain `!`, so it never collides) |
| column | `t:Products[Price]` | one column of a table, header row excluded |
| value by key | `t:Products/SKU-0042/Price` | the row whose key column holds `SKU-0042`, column `Price`; a key containing `/` or `'` is quoted: `t:Products/'A/7'/Price` |
| row by position | `t:Products/#5` | the fifth data row (a position, not an id: it moves if rows are sorted or inserted) |
| row range | `t:Products/#9:#16` | data rows 9 to 16 by position, the last included; `t:Products/#9:` row 9 to the last; `/#9:#16/Price` their one column. Given the other way round it is the same rows; a row past the last is refused with the range there is, never cut (round 5, T3) |
| every row | `t:Products/*/Price` | a selector: every data row's `Price` (the cells of `t:Products[Price]`); `t:Products/*/*` every value, `t:Products/*` every row, `t:Products/SKU-0042/*` each value of one row. `Workbook.expand` gives each cell by key |
| key set | `t:Products/{SKU-0042,SKU-0043}/Price` | the rows whose keys are listed, in that order: how a derived value's sources are written when the table has a key column. A key holding `, { } / '` is quoted inside |
| dataset | `d:by_category` | a derived dataset in the session; `d:by_category[average]`, `d:by_category/Hardware/average` as for tables |
| ref | `$step1` | the siblings' refs: a name a call gave to what it made |

Rules, from the siblings: **ids name things, positions name places, and the API says which
it is giving you.** A document handle is `d1` (no colon); a dataset is `d:` plus a name, so
the two never collide. A table's key column is the one the caller names, or else one
`describe` reports as a key (unique, never blank, text or whole numbers); `value by key`
without a key column is refused with the candidate columns as `valid_options`. Names are
matched as Excel matches them (case aside); an unknown name is `not_found` with the
nearest names.

A key is a cell's text, exactly, or a whole number's digits; a blank, a fraction, a date, a
time, a boolean and an error are no key. An address by key keeps naming the same cell when
rows are inserted, deleted or sorted; a position and a range do not. A key two rows share
is refused (`duplicate_key`) and a row without one cannot be named by key (`no_key`), each
with the rows by position as `valid_options`.

---

## Values and types

Every value read comes back typed, with its address:
`{"a": "Stock!F2", "v": 12.5, "t": "number"}`.

| Type | From | Notes |
| --- | --- | --- |
| `number` | `t="n"` | as stored; never reformatted |
| `percent` | a number whose format has `%` | the stored fraction (0.25), typed so it is not read as 0.25% |
| `date`, `time`, `datetime` | a number with a date or time format | ISO 8601; the workbook's date system (1900 or 1904) applied, and serials 1–60 of the 1900 system read as Excel reads them (its 29 February 1900) |
| `text` | shared, inline, rich or formula strings | rich text flattened; length reported when cut |
| `bool` | `t="b"` | |
| `error` | `t="e"` | the code (`#N/A`, `#REF!`, `#DIV/0!`...), never a number |
| `blank` | no value | a cell with only a style is blank |
| `list` | a declared column's text split by the caller's rule (X6) | a list of texts; no cell stores one |

Which number formats are dates is decided from the format code (date and time tokens
outside quotes and brackets) and, for the built-in ids that have no code in the file, from
ECMA-376's list of built-in formats; an Excel probe in E1 confirms that list on Excel
itself, and any difference wins over the standard. The format code is returned on
request (`formats: true`). **No display string is produced in 0.1 or 0.2**: how Excel shows
`General`, a locale's date order or a currency is measured from Excel in 0.3, not
reproduced from guesses. A model that needs to say a figure asks `xlsx_derive` to round it.

**Locale** (decision 4). `Workbook.open(..., locale=...)` with the shared
`ooxml_common.locales.Locale` (or any object with its five fields; en-US when none); in
the tool layer it is the session's, `Toolbox(locale=...)` or `toolbox.session(locale=...)`,
never the model's. Its separators decide how text that looks like a number is recognised
(numbers stored as text), its date order the order of text that looks like a date, and
its language, when it is one of ECMA-376's five (zh-TW, zh-CN, ja-JP, ko-KR, th-TH), the
table the language-dependent built-in formats are read with. Two reader settings are
not locale facts: `grouping` (on: `1,234.5` is a number in en-US) and `local_dates` (off:
ISO 8601 dates only), given to `Workbook.open` or `xlsx_agent.tools.format(...)`. Excel,
measured (16.106 Mac, 2026-10-13): text recognition follows the region's separators,
grouping included, and reads `3/4/2026` day first under nl-BE and month first under
en-US, so `local_dates=True` is Excel's reading. A format code's own locale tag
(`[$-409]`, `[$-413]`) overrides it for that format's cells.

**Formulas** (decision 3). `v` is the cached value. A formula with no cached value is
`{"a": ..., "t": "blank", "formula": true, "issue": "no_cached_value"}` and is counted in
every result that touched it. A workbook that asks for a full recalculation on load
(`calcPr fullCalcOnLoad="1"`) or was saved with manual calculation is flagged once: its
cached values may be stale. Volatile functions (`NOW`, `TODAY`, `RAND`, `OFFSET`,
`INDIRECT`) are flagged per column: the value is as of the last save.

---

## Tools

Defined once each in ooxml-edit's provider-neutral schema, registered on a `Toolbox` with
`formats=[xlsx_agent.tools.FORMAT]`. **Read tools are loaded (the core); edit tools are
deferred** behind tool search. The loaded core is held under **6,000 tokens as Anthropic's
count-tokens counts them** (the siblings' per-format budget), checked offline in CI by the
siblings' proxy (compact JSON ÷ 3.5 × 1.48) and online by a `provider`-marked test.

The proposal's `xlsx_describe` and `xlsx_find` are the shared `describe` and `find_text`
with an xlsx handler (O1), as in both siblings; the catalogue's drill-down is
`xlsx_profile`.

Every result is bounded by `Limits.max_result_chars` (24,000) and
`Limits.max_list_items` (50) unless a smaller `limit` is given, and says what it cut:
`{"cut": {"rows": 1240, "chars": true}, "next_cursor": "..."}`.

### 0.1 — read and derive (core, loaded)

| Tool | Inputs | Output | Limits |
| --- | --- | --- | --- |
| `open_document` (shared) | `blob` | `d1`, the summary: sheets, tables, size, problems at open | the zip limits; `.xls`/`.xlsb` refused |
| `describe` (shared) | `doc` | the catalogue: sheets (id, name, visible, hidden or very hidden, protected, used range); tables, names and detected regions with their key columns; per column: header, inferred type, fill rate, min/max/mean or distinct values when 12 or fewer, text length median and max, formula share, data-validation list, flags (mixed types, numbers as text, no cached value, cut at 32,767); up to two sample rows per table; workbook flags (recalculation, external links, macros); the session's datasets (id, columns, types, rows, from; T11) | about 3,000 tokens by default; a workbook too big for that lists every table and the first columns, and says which were left for `xlsx_profile` |
| `xlsx_profile` | `doc`, `target` (table or column), `columns`, `cursor` | the full per-column profile for one table, every column (a date column's range as dates; text lengths in characters and words with the longest cell's address; the range of dates stored as text) | one table per call; 50 columns a page, within 24,000 characters; paged |
| `xlsx_read` | `doc` (not for a dataset), `target` (range, table, column, row, dataset), `formats`, `formulas`, `comments`, `hyperlinks`, `merges`, `sources` (a dataset's only), `key_column` (the column holding a by-key target's key, `SKU` for `t:Products/SKU-0042`), `limit`, `max_chars`, `cursor` | typed values by row, with addresses; optionally format codes, formula text, comments, links, merges; the array blocks over the rows always | default 50 rows, at most 500 or 24,000 characters; paged, a row too wide for a page split across pages |
| `xlsx_query` | `doc`, `table`, `columns`, `where` (`eq`, `in`, `gte`, `lte`, `contains`, `words_gt`, `blank`, `not_blank`, combined with and), `sort`, `limit`, `cursor` | matching rows with their addresses, and the count of all matches | as `xlsx_read`; a filter on a column of the wrong type is refused naming the column's type |
| `find_text` (shared) | `doc`, `text`, `regex`, `range` (a sheet, range or table), `cursor` | matches by address and kind: cell (text cells; numbers and dates are not searched, having no display string yet), table or region name, column header, comment, defined name; hidden ones marked; formula text not searched (the shared tool has no option for it) | 50 per page |
| `xlsx_derive` | `doc`, `from` (table, region or dataset), `where`, `group_by`, `measures` (`count`, `sum`, `mean`, `median`, `min`, `max` (of dates, a date), `first`, `last` (row order, any column), `share`, `distinct`, each with `column` and `as`), `bins` (edges for a number or date column), `pivot` (a column whose values become columns), `sort` (by output columns: a `group_by` column or a measure's `as`), `top`, `round`, `coerce` (`number_as_text`, `date_as_text`), `id` | `d:<id>`: columns, rows, the rounding applied, and for every value its sources (see *Provenance*); the rows left out and why, counted by reason | at most 1,000 result rows and 50 shown; a measure on a column whose type does not allow it is refused before anything is computed |
| `xlsx_compare` | `a`, `b` (two open documents), `tables` (default every table and declared region; `table_b` when the one table's name differs in `b`), `key`, `columns`, `cursor` | first, each table's summary: per column the cells changed, and for numbers and dates up, down and the largest move each way by key (X10); sheets added, removed and renamed (by id); rows added and removed (by key, `t:Products/SKU-6`); cells changed (address by key, `t:Products/SKU-2/Price`, and in each sheet, old and new value, type: a number and the same number as text is a `type` change); each row and cell change with its row's `label` (the first text column other than the key, T19b); columns added, removed and renamed; duplicate keys reported, not paired by guess | 50 differences per page, totals always, the summary on the first page |
| `xlsx_validate_data` | `doc`, `schema`: sheets and tables (or declared region specs) required; `key`; per column type, required, unique, enum (with `case`), the column's own drop-down list (`validation_list`), `min`/`max`, `step` (e.g. 1–5 in halves), pattern | problems by address and code (`missing_sheet`, `missing_table`, `missing_column`, `no_validation_list`, `no_cached_value`, `required_blank`, `wrong_type`, `out_of_range`, `off_step`, `not_in_enum`, `not_in_list`, `pattern_mismatch`, `duplicate_key`, `duplicate_value`), counted by code; never cell text | 50 problems per page, counts always |

`xlsx_compare` and `xlsx_validate_data` are in the core if the budget holds; if the
online count says otherwise they move to a deferred `xlsx_check` group, and R14 records the
numbers. At R12 `xlsx_compare` is 329 estimated (about 487 counted by the offline
proxy) and the core with it about 4,153 counted, so it is in the core. At R13
`xlsx_validate_data` is 630 estimated (about 932 counted) and the core with it 3,425
estimated, about 5,069 counted: both stay in the core, and no `xlsx_check` group is made.
**Counted online at R14** (claude-sonnet-5-5's count-tokens, 2026-10-13, after P1, which
changed no definition): the loaded core, 9 definitions with the tool-use prompt, is
**4,958 tokens**, under 6,000 (the proxy's 5,069 over-counts by about 2%).

### 0.2 — edit (deferred groups)

| Tool | Group | Inputs | Output | Limits |
| --- | --- | --- | --- | --- |
| `save_document`, `undo`, `batch` (shared) | core | as in the siblings; `format`: `xlsx`, `xlsm`, `xltx` | name, size, validation report; never content | a `.xlsm` saved as `.xlsx` is refused while it has a VBA project |
| `xlsx_set_cells` | `xlsx_edit` | `doc`, `cells`: address with a value or a formula; `keep_types` (default true) | cells written, their old types, formula cells whose cached values are now stale | 5,000 cells per call; sets `fullCalcOnLoad` when a formula or a formula's input changed |
| `xlsx_write_table` | `xlsx_edit` | `doc`, `at` (a new sheet or an empty area), `data` or `from: "d:<id>"`, `header`, `style`, `filters`, `banding`, `name`, `ref` | a real Excel table; its address | refuses to overwrite non-empty cells; text over 32,767 characters is split across continuation rows, each listed, never cut silently |
| `xlsx_manage_sheets` | `xlsx_edit` | `doc`, `action` (`add`, `rename`, `move`, `hide`, `unhide`, `delete`), `sheet`, `name`, `position` | the sheets after | `rename` rewrites every reference to the sheet (formulas, names, charts, validation, conditional formats); `delete` is refused while anything refers to the sheet, and lists what does |
| `xlsx_validation` | `xlsx_edit` | `doc`, `target`, drop-down list (values or a range) or a value rule | the rules on the target | an inline list within the length Excel accepts (measured in E6), longer lists as a range |
| `xlsx_format` | `xlsx_format` | `doc`, `target`, number format, font, fill, borders, alignment, wrap, column widths (`auto` measured with the real font), row heights, freeze panes, print area and titles | what changed | styles reused when identical; no new style per cell |
| `xlsx_conditional_format` | `xlsx_format` | `doc`, `target`, colour scale, data bar, icon set, or a rule with a fill | the rules on the target | 20 rules per target |
| `edit_chart` (shared) | `shared_misc` | as in the siblings, on a sheet's drawing; `from: "d:<id>"` (O3) | as in the siblings | needs O5 |
| `new_document` (shared) | core | `kind: "xlsx"` | `d1` | a new workbook as Excel makes one, measured (E9) |

### 0.3 — render and check (later, separate project)

`render` (a sheet, a range or a print page to PNG through `xlsx2svg`: grid, number formats
per locale, fonts, fills, borders, merges, wrap, conditional formats, charts through
pptx2svg's chart renderer; the shared tool's name and image budget) and `check` (`#####` in
narrow columns, wrapped text cut by a fixed row height, error values, formulas without a
cached value, numbers and dates stored as text, mixed types, broken names, content in hidden
rows and sheets, cells at the 32,767 limit). A project of docx2svg's size, with Excel as
the measured oracle (decision 5); it gets its own roadmap when it starts.

---

## The Python API

The tools are thin wrappers over this; everything a tool does, a script can do.

```python
from xlsx_agent import Workbook, Locale      # Locale: ooxml_common.locales.Locale

wb = Workbook.open("inventory.xlsx", locale=Locale(decimal_separator=",",
        thousands_separator=".", date_order="dmy", language="nl-BE", list_separator=";"))
# bytes, a path, or an OpcPackage (a chart's embedded workbook); .xlsx .xlsm .xltx .xltm

wb.sheets                        # [Sheet(name="Stock", id=3, state="visible", protected=False)]
wb.problems                      # [Problem(address="Stock!F7", code="no_cached_value")]
print(wb.describe(max_tokens=3000))      # the catalogue as text; wb.catalogue() as data

products = wb.table("t:Products", key="SKU")         # a table, name or region
stock = wb.table("t:Stock!A1")                        # a detected region
wb.declare_region("returns", "Returns!A3:F200",      # a merged two-row header,
                  header_rows=2,                      # declared, never guessed
                  headers=["SKU", "Date", "Qty", "Reason", "Store", "Refund"])
wb.declare_tables([{"name": "stock", "sheets": "Store *",   # the caller's own tables (X3):
                    "header_row": 2, "key": "SKU",          # several sheets as one, ids,
                    "fields": [{"header": "Cost", "hidden": True}]}])   # hidden columns

products.column("Price").stats()     # ColumnStats(type="number", fill=0.98, min=..., ...)
products.query(["SKU", "Price"], where={"Category": {"eq": "Hardware"}}, limit=20)
wb.read("Stock!A1:H40")              # Values with addresses and types

ds = wb.derive(from_="t:Products", where={"Status": {"in": ["Active"]}},
               group_by=["Category"],
               measures=[{"fn": "mean", "column": "Price", "as": "average"},
                         {"fn": "count", "as": "items"}],
               sort=[("average", "desc")], round=2, id="by_category")
ds.rows                              # [{"Category": "Hardware", "average": 41.25, "items": 12}, ...]
ds.sources["Hardware"]["average"]    # [Area("Stock!F3:F14")] minus ds.excluded
ds.excluded                          # {"average": {"number_as_text": Count(1, ("Stock!F9",)), ...}}
ds.dataset()                         # ooxml_edit.tools.Dataset: rows as reported, sources as text
DatasetTable(ds).read("d:by_category/Hardware/average")   # [[DatasetValue(...)]], by address
wb.derive(from_=ds, measures=[{"fn": "max", "column": "average"}])   # from a dataset (R11):
                                     # its sources are the original cells, not the dataset's

Workbook.compare(old, new, tables=["t:Products"], key="SKU")   # Diff
wb.validate_data(schema)             # ValidationReport: [DataProblem(code, address)],
                                     # counted by code; never cell text
```

0.2 adds `wb.set_cells({...})`, `wb.write_table(...)`, `wb.format(...)`,
`with wb.batch():`, `wb.undo()`, `wb.redo()`, `wb.to_bytes()` and `wb.save(path)`, all
with ooxml-edit's history.

**Code layout (proposed).** `src/xlsx_agent/`: `package.py` (open, kinds, limits),
`address.py`, `read/` (`stream.py`, `strings.py`, `styles.py`, `types.py`, `names.py`,
`tables.py`, `regions.py`), `profile.py`, `query.py`, `derive.py`, `dataset.py`,
`compare.py`, `schema.py`, `catalogue.py`, `tools/` (`FORMAT`, one module per group, a
`README.md` like the siblings'), and in 0.2 `edit/` and `formula/` (the lexer).

---

## Design rules

1. **The model never does arithmetic.** Every figure an agent reports comes from
   `xlsx_derive` or a cell. The format's prompt fragment says so, and says that a value
   without a `d:` or cell address behind it is not data.
2. **Provenance of every derived value.** Each value carries its sources as compact ranges
   (`Stock!F3:F14`) plus the cells left out with their reasons. In a tool result at most
   eight ranges per value are shown, with the rest counted; `xlsx_read` of the dataset with
   `sources: true` pages through all of them. A test re-reads every source and recomputes
   every value (see *Testing*).
3. **Computation is exact and stated.** Sums with `math.fsum`; mean is the sum over the
   count of values used; median of an even count is the mean of the two middle values;
   `share` is a fraction of the stated total. Rounding is half away from zero (Excel's
   `ROUND`), applied to the reported value only; the unrounded value stays in the dataset.
4. **Nothing is coerced silently.** Blanks, errors, numbers stored as text and
   no-cached-value formulas are left out of a measure and counted by reason; a caller may
   opt in to `coerce: ["number_as_text"]`, and the result says it did.
5. **Bounded results.** Every result has a row, item and character cap and a cursor; a
   catalogue has a token target. Totals are always given, even when rows are cut.
6. **Untrusted input.** Zip size, part count and ratio limits (`check_package`) before
   opening; XML entities never resolved, `huge_tree` off; no external link, DDE link, web
   query or data connection followed; macros never run; sheets streamed, never parsed
   whole for reading; a soft deadline on every call (`Limits.soft_deadline`), with partial
   results marked. Cell text, comments and names are data, never instructions, and the
   prompt fragment says so.
7. **No cell text in diagnostics.** Warnings, errors, exception messages and call logs
   carry addresses, types and codes only: never a cell's value, a comment, or a formula's
   text (formulas hold literals). Tool results return content; diagnostics do not.
8. **Lossless** (0.2). A part never parsed is written back as its bytes: pivot caches,
   slicers, sparklines (worksheet `extLst`), the VBA project, custom XML, printer settings.
   An edited worksheet keeps every element it does not understand. Undo of anything is the
   original bytes.
9. **Facts, not verdicts.** Flags say what is there (`number_as_text: 14 cells`), not
   whether it is acceptable.

---

## Messy workbooks

What the reader meets, and what it does. Each item has a made-up fixture (*Testing*).

| Situation | Behaviour |
| --- | --- |
| merged header cells, two or more header rows | not detected; `describe` flags the region `header_unclear` and points to `declare_region` |
| a title or caption: one cell with nothing under it | no region; left out of `describe`, still a cell for `find_text` and reads; text two rows above a detected region's header (a blank row between) is its `title` |
| duplicate or blank header cells | in a table, Excel's names are used; in a detected region the column is named by its letter (`Column F`) and flagged |
| hidden and very hidden sheets, hidden rows and columns | read like any other, and every value from them is marked `hidden` |
| protected sheets and workbook | read; flagged; in 0.2 edits are refused on a protected sheet |
| numbers and dates stored as text | typed `text`, flagged per column with a count and the first addresses; recognised with the caller's locale |
| mixed types in a column | the column's type is the majority; the minority is counted by type with first addresses |
| formula without a cached value | `issue: no_cached_value`; counted; left out of measures |
| `fullCalcOnLoad`, manual calculation, volatile functions | flagged: cached values may be stale |
| shared formulas | the value is read as any other; formula text of a shared group's member is given as the master's address plus its own position (no translation in 0.1) |
| array formulas, dynamic arrays and their spills | values read; the spill range reported as one block |
| cells at 32,767 characters | flagged `at_limit`: probably cut by the producer |
| rows and cells without `r` attributes, out of order | positions computed as Excel does (next row, next column) |
| a wrong or missing `dimension` | ignored; the used range is computed from the cells |
| a sheet of one million sparse rows | streamed; the used range and fill rate from what is there; reads bounded |
| error values | typed `error` with the code; left out of measures, counted |
| 1904 date system | applied from `workbookPr date1904` |
| rich text, phonetic runs | text flattened, phonetic runs dropped |
| defined names that are formulas, not ranges | listed, not resolved; not usable as `t:` |
| tables with a totals row, or no header row | totals row excluded from data and reported; a headerless table's columns are named as Excel names them |
| external links, DDE links, data connections | listed in the catalogue's flags; their cached values read; never followed |
| strict OOXML namespaces (`purl.oclc.org`) | read like transitional (0.1); 0.2 edits refuse them until measured |
| `.xls`, `.xlsb`, encrypted workbooks | refused with the reason |
| XML entities, zip bombs | refused by the parser and `check_package` |

---

## ooxml-edit changes this plan depends on

Each is its own PR in ooxml-edit, with its version noted here when it lands. The neutrality
test applies: nothing below names a spreadsheet in the core, except in `shared`'s
descriptions and the optional `sheets` and `charts` subpackages.

| | Change | Needed by |
| --- | --- | --- |
| **O1** (0.18.0) | The shared tools learn kind `xlsx`: `shared.KINDS`; `describe`'s and `find_text`'s descriptions say what a workbook returns and that `range` takes a sheet, range or table; `save_document`'s `format` enum gains `xlsx`, `xlsm`, `xltx`; `new_document`'s `kind` gains `xlsx`. The deck and Word cores are re-counted, since the shared definitions are sent to them too | R6 (describe, find_text); E1 and E9 (save, new) |
| **O2** (0.20.0) | A session-level **dataset store**: `Session.datasets`, `Call.define_dataset(id, dataset)` and `Call.dataset(address)`; `d:<id>` resolved by the dispatcher in a tool's declared fields (as `Tool.refs` declares ref fields), refused with the known ids as `valid_options`; a dataset records its source document and that document's `version`, outlives `close_document`, and is marked `stale` (never refreshed silently) once its source changes; `Limits.max_datasets` and `Limits.max_dataset_cells`. Format-neutral: a dataset is columns, rows and sources | R11 |
| **O3** (0.22.0) | Consumers of datasets: `edit_chart`'s `add`, `set_values` and `add_series` take `from: "d:<id>"` with a column mapping, and the chart keeps the dataset's sources in its result; pptx-agent's `ppt_add_table`/`ppt_edit_table` and docx-agent's table path take the same. Changes in ooxml-edit, pptx-agent and docx-agent, each held to its own core budget | after 0.1 (it is what makes `d:` cross-format). Landed in ooxml-edit 0.22.0 (`edit_chart`'s `from`, the `DocumentFormat.datasets` resolver hook and `merge_sources`) and pptx-agent's tables; xlsx-agent's `t:` resolver is X1 below |
| **O4** | **Lift `charts/workbook.py`'s cell writer** into a public, optional `ooxml_edit.sheets` (moved with its history, `charts.workbook` a re-export), with what a whole workbook needs: a shared-string index built once, booleans, formulas, a style index, `r`-less rows; chart outputs byte-identical in both siblings' suites | E1 |
| **O5** | A chart host for a workbook: in a `.xlsx` a chart's series point at the host's own sheets, not at an embedded workbook, so `ooxml_edit.charts` writes the host's cells through O4 | E9 |
| O6 (small) | `ooxml_edit.xml.iterparse_xml`: `iterparse` with `parse_xml`'s safe options, so they live in one place | R3 (until it lands, R3 uses the same options locally, with a test that they match) |

---

## Milestones

Each item is one PR, small enough to review in a sitting, and leaves the suite green.
0.1 is built in this order; nothing later is started before the item before it is merged.

### 0.1 — read and derive

| # | Item | Effort | Done when |
| --- | --- | --- | --- |
| R0 | **Skeleton.** `pyproject.toml` (lxml, ooxml-edit ≥ 0.17, ooxml-common ≥ 0.14; Python 3.11–3.15), `src/xlsx_agent`, CI on 3.11 and 3.15, and `tests/build.py`: a small SpreadsheetML writer for tests only, building made-up workbooks from Python (sheets, cells of every type, styles, tables, names, merges, comments) | S | a built workbook opens in the test as a zip with the right content types |
| R1 | **Open.** `Workbook.open` from bytes, path or `OpcPackage`; `kinds` check; `check_package`; refusing `.xls`, `.xlsb` and encrypted files; `sheets` with id, state and protection; `to_bytes` byte-identical | S | every fixture round-trips byte for byte; the refusals carry their reason |
| R2 | **Addresses.** Parse and format every form in *Addresses*, on `Cell` and `Area` from `charts.workbook`; errors with `valid_options` | M | a table of forms parses and formats back to itself; quoting round-trips |
| R3 | **Cell reader.** A streaming pass over a sheet (`iterparse`, O6's options); shared strings indexed once; inline and rich strings; booleans, errors; `styles.xml` number formats; types as in *Values and types*, both date systems; cached formula values and `no_cached_value`; merges; `r`-less rows; the locale | M | every row of the messy table that concerns cells has a passing fixture; a 200,000-row sheet reads in bounded memory |
| R4 | **Tables, names, regions.** ListObjects (header and totals rows, columns), defined names over ranges (workbook and sheet scope), `declare_region`, single-row header detection, key-column detection | M | `t:` and `t:...[col]` and `t:.../key/col` resolve on every fixture; a merged header is flagged, never detected |
| R5 | **Column profile.** `column.stats()`: type and mix, fill rate, min/max/mean, distinct values when few, text lengths, formula share, validation lists, flags | S | the profile of every fixture column matches a hand-written expectation |
| R6 | **Tool layer and `describe`.** `xlsx_agent.tools`: `FORMAT` (open, detect, problems, errors, prompt fragment, summary), the shared `open_document`/`close_document`/`read_blob`, the `describe` handler with its token target; the **token-budget test** from here on. Needs O1 | M | the catalogue of the largest fixture is within its target; the core is under 6,000 by the offline proxy |
| R7 | **`xlsx_read`** with formats, formulas, comments, hyperlinks, merges, paging | S | bounded on a large fixture; every cut is reported |
| R8 | **`xlsx_query`** and `table.query` | S | each filter has a test, including a refused wrong-type filter |
| R9 | **`find_text`** handler: cells, table names, headers, comments, defined names | S | matches by address and kind; no match text in diagnostics |
| R10 | **Derive engine** (library only): `where`, `group_by`, every measure, `round`, provenance and exclusions; then `bins`, `pivot`, `sort`, `top` | M | the provenance test passes on every fixture; refusals for wrong types |
| R11 | **Datasets and `xlsx_derive`.** `Dataset`, `d:<id>` in `xlsx_read` and `xlsx_derive`'s `from`; stale marking. Needs O2 | S | a dataset survives closing its workbook, is stale after its source changes, and is read back by address |
| R12 | **`xlsx_compare`** and `Workbook.compare` | S | added, removed, changed and duplicate-key cases on fixtures |
| R13 | **`xlsx_validate_data`** and `wb.validate_data` | S | every problem code has a fixture; no cell text in any problem |
| R14 | **Release 0.1.0.** The online budget count (`provider`-marked, with a key), the tools README (setup per provider, as the siblings'), `docs/api-tour.md`, golden transcripts of three read-and-derive tasks replayed to identical results, CHANGELOG | S | the online core count is under 6,000 and recorded here |

*R11 as built* (ooxml-edit 0.20.0): `xlsx_derive` keeps `Derived.dataset()` (rows as
reported, each value's sources as range text, `keys` the key columns, the `Derived` as
`detail`) under `d:<id>`; a dataset's id is ooxml-edit's `DATASET_ID` (a letter or `_`,
then letters, digits, `_ . -`, at most 64), which the address parser now enforces.
`xlsx_derive`'s result shows its first rows (`Limits.max_list_items`, within
`max_result_chars`), each value typed with its address by position (`d:by_region/#2/total`)
and at most eight source ranges, the rest counted in `more`; `cut` and `next` (a target and
cursor) continue in `xlsx_read`. `xlsx_read` takes `d:` targets without a `doc` (a dataset
outlives its workbook) and `sources: true` (eight ranges a value; one value pages through
all of them). Deriving from a dataset measures a `Derived`'s unrounded values, and each
value's sources are the original cells, merged again. Until 0.2 a workbook has no history,
so the session counts its changes upwards: a stale dataset stays stale (no undo yet); the
tests make the change with a test-only changing tool.

*R12 as built*: `Workbook.compare(a, b, tables=..., key=..., columns=...)` gives a `Diff`
(`xlsx_agent.compare`): sheets matched by `sheetId` (renamed, added, removed); tables and
regions matched by name, case aside (every table and declared region by default, and those
in one workbook only listed; named ones, or pairs, otherwise); columns by name, a header
changed in place (old name not in `b`, new name not in `a`) a rename whose cells are
compared; rows by key -- named, declared, or the one key column both workbooks find, else
`no_key_column` with the candidates. A key two rows of one workbook share is reported as
`duplicate_key` with each workbook's rows by position, and its rows are left out of every
other count (never paired, never first-wins); rows without a key are counted (`unkeyed`).
Cells are compared by type, then value: `changed`, `type`, `added` (blank before),
`removed` (blank after); a formula by its cached value. `xlsx_compare` lists the
differences (sheets, tables, columns, duplicate keys, unkeyed rows, rows added and removed
with their cells, cells changed) 50 a page within 24,000 characters, `totals` always, the
comparison kept for its next pages.

*R13 as built*: `wb.validate_data(schema)` (`xlsx_agent.schema`) takes a JSON-able schema:
`sheets` (names, case aside) and `tables`, each a `t:` address or a declared region spec
(`range`, `name`, `header_rows`, `headers`; used for the check only, never kept), with
`key` (columns unique together, never blank) and `columns`: `name` and any of `type` (a
reader type; nothing coerced, so a percent is not a number), `required`, `unique` (values
the same as keys are: text exactly, whole numbers by digits), `min`/`max` (written as
filter values, read by the declared type), `step` (number or percent, from `min` or 0,
compared as shortest decimals), `enum` (`case`: `exact` by default, or `ignore`),
`validation_list` (the union of the data-validation lists over the column, as the profile
reads them) and `pattern` (Python regex, whole match, text cells). A schema that is not one
is refused before reading (`SchemaError`, `bad_schema`, its path, valid options; tool
error `invalid_arguments` on `schema`). The `ValidationReport` lists `DataProblem`s by
code and address -- a run of cells next to each other in a column with the same problem
is one problem over a range with its `cells` -- with the table, column, what was expected
(type, range, step, enum or list size, case rule, a `hint` `case_differs`/`space_differs`
without the value, `first` for a repeat, `reason`, `nearest` names) and a one-line
`message`; never the cell's text. `counts` (problem cells by code) are always complete;
at most 10,000 problems are kept, the rest counted (`cut`). `xlsx_validate_data` pages
them as `xlsx_compare` pages differences; in the strict schema `min`, `max` and `enum` are
strings and `key` a list.

*R14 as built* (0.1.0): `tests/test_tools_online.py` (marked `provider`, skipped without
`ANTHROPIC_API_KEY`) counts the core online -- 4,958 tokens, above -- and sends the
definitions as the adapter does, which Claude accepts; its strict-fields check skips while
no tool is sent strict (Claude's strict plan is for writing tools; 0.1's all read). Three
golden transcripts (`tests/golden/`, `tests/test_golden.py`) were recorded on
claude-haiku-5-5 by `tools/record_goldens.py` and replay offline to the same results, call
for call: g1 derives sales by region and reads `d:by_region` back, g2 compares a generated
inventory with a reviewed copy, g3 checks orders against a declared schema; their inputs
(`tests/golden_inputs.py`) fix every value and date. The tools README is
`src/xlsx_agent/tools/README.md`; the Python API, `docs/api-tour.md`; `CHANGELOG.md` has
0.1.0.

*X4 and X5 as built* (production feedback on 0.1.0, 2026-10-13): **X4, dates.**
`xlsx_derive` gains `first` and `last` (any column type: the value of the first and the
last row used, in the table's row order; sources as `min`'s, every cell compared, the
picked one first or last). `min` and `max` already took dates and gave a date; the tool
description now says so. What made a "date column" fail: dates stored as text make the
column `text` (its type is the majority's), and `min`/`max` refuse text. The refusal now
names the issue (`date_as_text`, or `local_date_text` for text in the locale's day order
that the workbook was opened without `local_dates` to read), its count and first cells
(`DeriveError.issue`, `.cells`; the tool error's `details`), and the recovery,
`coerce: ["date_as_text"]` -- opt-in, listed in `coerced`, never silent (design rule 4).
With a coercion the column's type is decided with the coerced cells counted as what they
read as (also for `number_as_text`, which had the same gap). A date stored as text in a
date column is left out as `date_as_text` (was `wrong_type`). A two-digit year is not
read: Excel's century rule is not measured. `min`/`max` of a coerced cell give the value
as read (a number, a date), no longer its text. `xlsx_profile` already gave a date
column's `min`/`max` as dates and flagged `date_as_text` with its count; it now also gives
`date_text`, the range of the dates stored as text. **X5, text lengths.** `xlsx_profile`'s
`text_length` adds `median_words`, `max_words` (runs between white space) and `longest`
(the address of the longest text in characters). The long-text filter is a query test,
`words_gt` (text columns; `xlsx_query` and `xlsx_derive`'s `where`): it lists or counts the
cells over a limit, and as one more enum value in the shared filter schema it was the
cheapest form (about 7 estimated tokens, against about 25 for a `max_words` parameter on
`xlsx_profile` or `xlsx_read`). Budget: the loaded core 3,452 estimated, about 5,109 by the
offline proxy (was 3,425, 5,069); every definition 3,674 (guard 3,700). `describe` is
unchanged.

*X2 as built* (production feedback, 2026-10-13): **Addresses by key, stable across edits.**
The grammar gains `*` -- every row in a row's place (`t:Products/*/Price`, `t:Products/*`),
every column in a column's (`t:Products/SKU-0042/*`) -- and key sets,
`t:Products/{SKU-0042,SKU-0043}/Price`. A selector reads as what it covers and keeps its
text; a key set resolves to its rows in its order (`Selection.positions`) and `xlsx_read`
pages them; `xlsx_derive`'s `from` takes `t:Products/*` as the table and refuses a part of
one; datasets take `*` and refuse key sets (`no_key_set`). `Workbook.expand(selector)` is
the coverage helper: one address by key per cell (per row without a column), in the
table's order, refusing a row with no key (`no_key`) or a shared one (`duplicate_key`).
`Workbook.declare_key(table, column)` is the caller's key column for a table with several
candidates (or none), shown by `describe`; X3's declared schemas will build on it. A
region's key column is read once into a `KeyIndex`, so a key set or an expansion is one
pass, not one per key. **Sources by key**: when the table's key column (declared, or the
one detected) names every row alone, each derived value's sources are one key set,
`t:Products/{SKU-0003,SKU-0007}/Price`, in row order (`Derived.source_key` names the
column); otherwise ranges, as before -- the default only where a key exists, because a
range stops being true when a row is inserted and a key set does not, and no parameter is
needed. A derive from a dataset gathers the keys of one column into one set again.
Excluded and left-out cells keep their cell addresses (they are diagnostics of the read,
not a plan's inputs); `xlsx_read`/`xlsx_query` cells keep their `a` (the cell read now).
**Result size** (offline estimator, a 200-row table keyed by code, interleaved groups):
`xlsx_derive` results show at most eight sources a value, now counted in units (a range,
or one key), so they stay bounded: ungrouped mean 207 → 231 tokens; mean by 5 groups 817 →
735; by each of 200 codes 3,252 → 3,766. The whole provenance (library, nothing cut):
ungrouped 19 → 444 (one range against 200 keys), by 5 interleaved groups 1,882 → 1,037
(single-cell ranges are longer than keys). Reading one value's 200 sources pages 50 keys
at a time (283 tokens a page, against one 149-token page for the one range). With two key
candidates and none declared, nothing changes. **Budget**: only `xlsx_read`'s target
mentions `*`: the loaded core 3,462 estimated, about 5,124 by the offline proxy (was
3,452, 5,109); every definition 3,684 (guard 3,700). The goldens are unchanged (their
tables have no key column). The strict check (`strict_problems`) is clean.

*X1 as built* (production feedback, 2026-10-13; the O3 consumer side): **Workbook tables
as sources for decks and documents.** The xlsx format registers ooxml-edit's resolver hook
for `t:` (`DocumentFormat.datasets`) and `merge_sources`, so a session holding a workbook
and a deck fills pptx-agent's `ppt_add_table`/`ppt_edit_table` and the shared `edit_chart`
straight from a table: `from: "t:Products"` (or `d2/t:Products` with several workbooks).
`table_dataset` (`Workbook.dataset`) gives the rows: the whole table (`t:Products`,
`/*`, `/*/*`), one column (`[Price]`, `/*/Price`), a key set's rows in its order, or a row;
one value is refused as `not_rows` (`invalid_arguments`, the table and its column as the
options). Values are `xlsx_read`'s: numbers, dates as ISO 8601 text, text, bools, error
codes, blanks as `None`; issues are not coerced (a number stored as text stays text, and
`edit_chart` refuses it naming its cell). A column's type label is its values' one type.
Sources follow X2's rule: by key (`t:Products/SKU-0042/Price`) when the key column --
declared, or the one detected -- names every row alone, else the cell address; the merger
X2's derive-from-a-dataset used (`dataset.merge_sources`, factored out of
`DatasetTable.sources_of`) writes a placed column's as ranges or one key set. The session's
`Limits.max_dataset_cells` is applied before any cell is read (`too_many_cells`, `limit`;
the dispatcher checks again after). Nothing is written to the workbook: no undo step, the
version and the bytes unchanged. The cross-format tests use pptx-agent (and pptx2svg, which
its layout checks measure with) as test-only dependencies, installed in CI from their
repositories. docx-agent's tables take the same `from` when they land there. **Budget**: no
tool definition changes -- the loaded core stays 3,462 estimated, about 5,124 by the
offline proxy; every definition 3,684 (guard 3,700); `strict_problems` is clean. The
workbook prompt gains one sentence (a slide or document table or chart takes `t:` or
`d:`).

*X3 and X7 as built* (production feedback, 2026-10-13): **Tables declared from the
caller's configuration** (`xlsx_agent.declare`). Declarations are the application's
configuration, never a model's -- no tool parameter, so no definition changes -- given as
plain data to `Workbook.declare_tables` (or `declare_table`), to the format
(`xlsx_agent.tools.format(tables=...)`: every workbook it opens) or to one session
(`xlsx_agent.tools.declare(session, ...)`, in place of the format's); the format's
`configure` hook declares them on a workbook joining the session, after its locale, so
ooxml-edit is unchanged. A declaration gives the table's `name`; its `sheets` (exact
names, case aside, or globs; the workbook's order); `header_row` (one row; 0 for none),
`first_row`, `last_row` (else the first blank row ends the data, and more below is flagged
`rows_below`), `columns` (`A:F`); `key`; a `description`; `fields` by header or column
letter, each with an `id`, a `description` and `hidden`; `type` (`rows`, `key_value`); and
`optional`. **A table over several sheets** is one rectangle of its own
(`read.regions.Span`): its rows each sheet's, in order, its header the first sheet's (every
sheet's must be the same text, or the declaration is refused naming both cells); a read
through it (`Workbook._read` on its own sheet, `t:<name>`, a name no sheet can have) reads
each sheet and gives every value in the table's place with its own cell kept
(`Value.origin`; `Value.a` is the cell's address). So reads, queries, samples, profiles
(`range` each sheet's, joined by `,`), a derive's sources (by key; without a key, each
sheet's ranges) and left-out cells, a dataset's sources, comparisons and checks (a run of
problem cells never crosses two sheets) all name the real cell, and a key finds its row on
any sheet; a key two rows share, on one sheet or two, or a row without one, is refused at
declaration, naming the cells, never the key. **Decided: the sheet is not a column**
(no `_sheet`): a value is a cell and a sheet's name is no cell's, and every address names
its sheet; for one sheet's figures, declare that sheet's table too. **Ids and
descriptions**: a column's `id` addresses it and is what `describe` shows; its header text
finds it too (`Region.headers`); descriptions show in `describe` (`about`, the table's and
each column's) and `xlsx_profile` (`description`, and `header` when an id names it).
**Hidden columns** are left out of the table's own rectangle, so nothing that reads a
table can show them, and an address naming one is `not_found` without it among the
options; their cells are withheld from reads of the sheet's cells (`xlsx_read` gives
`{"a": ..., "withheld": true}`), from `find_text`, and from a region detected or declared
over them (`RegionError` `declared`, pointing to the table); the blocks a declared table
covers are no longer detected as regions of their own. An Excel-hidden column stays a column
unless declared hidden. A sheet's problems (`Workbook.problems`, describe's `problems`)
leave its withheld cells out: one met only there is not listed, one met there too is
counted and placed at its other cells (the sheet is read once more for it, kept while the
declarations stay); a declared table's own flags (`blank_header`, `rows_below`) never name
a hidden cell. **X7, key-value tables**: `type: "key_value"`, columns `field` and
`value` (ids rename them; headers find them) and any more; the key is the field column;
`t:run/report_date/value`; `describe` shows the fields and their values compactly
(`fields`, at most 25, the rest counted). A two-column block with no header whose first row
holds a non-text value (so header detection fails, `header_not_text`) and whose first column
is a different text on each of two or more rows is **detected** as one
(`t:'Run facts'!A1/author/value`); wider blocks are not, since a first row of years over a
table looks the same. Hidden sheets were already read like any other and their values
flagged `hidden`; now tested. **Validation**: `DeclarationError` (`code`, `path` such as
`tables[0].fields[2].header`, `valid_options`) for a declaration of the wrong form (checked
when given to the format or session) or one the workbook does not fit, before any table is
kept; at `open_document` it is `refused` with the code and path in `details`. **Sizes**
(offline estimator, a made-up workbook of six area sheets of 20 checks with six columns,
one a long text, a summary sheet and a hidden sheet of run facts): `describe` undeclared
2,864 estimated tokens on main, 2,914 here (the run facts now detected as key-value, +50;
both cut to fit 3,000: 3,494 and 3,544 in full); declared as one table with its key 1,335
(nothing cut); with two ids and a hidden column 1,272; with a table description and five
column descriptions 1,361 (+89 for six descriptions); with the run facts declared as
key-value 1,357. **Budget**: no definition changes (a declaring format's definitions equal
the plain one's: tested); the loaded core is as before, 5,009 online; `strict_problems` is
clean; the workbook prompt is unchanged. The goldens are unchanged; describe of an
undeclared workbook changes only where a two-column key-value block was `header_unclear`.

*X6 as built* (production feedback, 2026-10-13): **Cleaning text, on the declaration.**
A declared field takes `clean` (`xlsx_agent.clean`): `strip`, a pattern matched at the start
of the text (Python's `re`, within the first 256 characters: a leading label such as
`^Lead:[^;]*;\s*`); `split`, a literal separator, each item trimmed and empty ones dropped,
the value a `list` (a new value type, last in `TYPES`: a tuple in Python, a list in JSON;
even of one item, so the column has one type); `collapse_space`, runs of white space to
one space and none at the ends (each item's, with a split); applied in that order, to text
values only. **Caller configuration, not a tool parameter**: no definition changes (tested:
a format declaring rules gives the plain one's definitions, and the `xlsx_*` definitions
are 10,155 characters of compact JSON, as on main); the loaded core is unchanged, 5,009
online; `strict_problems` clean; goldens unchanged. **Where**: a table with rules reads
through its `Span` (as one with hidden columns did), which cleans the data rows as they are
read -- reads through the table, describe's samples, query, profile, derive, compare,
checks and `t:` sources; never the header; a cell read by its sheet address is as stored.
Text cleaned to nothing is `blank`; cleaned text that reads as a number or date is flagged
`number_as_text`/`date_as_text`, so `coerce` reads it. **Never silent**: a changed value has
`cleaned` (`"cleaned": true` in a tool's cell) and `original` (Python only); a one-item list
of unchanged text is not counted. Counted with first addresses in `xlsx_read`/`xlsx_query`
pages (`cleaned`), as a `cleaned` flag in `xlsx_profile` and `describe`, per column in
`xlsx_derive` (`Derived.cleaned`, every row read), and for a `t:` source in the dataset's
`detail` and a `CleanedWarning`, a format warning the dispatcher adds to the deck or
document tool's result (counts and addresses, never text). **Keys: refused**
(`clean_key`), not cleaned before matching: an address by key names the key as its cell
holds it (what Excel and the caller see), two keys cleaned alike would name one row, and
a list is no key; a key-value table's field column is its key. Clean rules on a hidden
column are refused too (nothing to show). **Lists downstream**: `xlsx_query` and `where`
take `eq`, `in` and `contains` on a list column, passing when any item does; a list sorts
item by item. `xlsx_derive` takes a list column in `count` only (its cells, not items: an
item count would be a new function, a definition change); every other function,
`group_by` and `pivot` refuse it (`wrong_type`), and a list in a column of another type is
left out as `wrong_type`; so a `d:` dataset never holds one. `t:` sources join a list's
items by a line break (`dataset.LIST_JOIN`; deck and document tables render one cell of
several lines), type label `text`. `xlsx_compare` compares lists as values.
`xlsx_validate_data` checks a list as `text`, item by item, each code once per cell;
`unique` compares whole lists; the schema's `type` enum leaves `list` out (unchanged
definition). **Pattern cost** (decided: reject, with a window, no timer -- Python's `re`
cannot be interrupted): compiled at declaration (`DeclarationError` `bad_pattern`, path
`tables[i].fields[j].clean.strip`, the parser's message and position), at most 200
characters, and refused when a repeat holds a repeat, an alternation sits in a repeat, it
uses a backreference or conditional group, more than one repeat can give characters back
to what follows (a one-character class repeat followed by a character it cannot match,
`[^;]*;`, or ending the pattern, cannot; `.*;` and `\s*[^;]*` can), or its alternations
combine more than 16 ways. Read from Python's own parser (`re._parser`, 3.11 on;
refused where it is missing). With the 256-character window the worst form allowed
(`^\s*\s*x` on a 32,767-space cell) took 0.2 ms a cell on 3.12, 0.1 ms on 3.15 (Intel Core
i7-6567U, 2026-10-13); `^Lead:[^;]*;\s*` about a microsecond. Whether two character
classes share a character is decided on the code points up to U+3000 (and a few past
it): two classes that share only a character past those would be taken as disjoint, and
such a pattern could cost more than the square bound on text of that character.

*X10 as built* (production feedback, 2026-10-13): **Comparisons by key, summarised per
column.** Every cell changed is named by key as `b` names it, `t:Products/SKU-2/Price`
(`CellChange.at`; the values keep each workbook's sheet address), and every row added or
removed by key in its workbook, `t:Products/SKU-6` (`RowChange.at`). In `xlsx_compare`'s
differences `at` takes the place of `table`, `key` and `column` (a row's `at` was its
position, `t:Products/#5`; its cells have their sheet addresses); a key longer than
`TEXT_SHOWN` keeps the old form, cut. `TableDiff.by_column()` gives a `ColumnSummary` per
column with a cell changed, in `b`'s order: `changed` (every kind), and for numbers,
percents, dates and times changed in value, `up`, `down`, `max_up` and `max_down` as the
address by key and the move (`t:checks/A190/rating +2.5`; a number's exact in shortest
decimals, 0.09 - 0.08 is `+0.01`; a date's in days, `+7d`; a date and time's in days to
four places; a time's `+1:30`); the first met wins a tie; a type change, text, a boolean,
an error and a list count as `changed` only (`CellChange.delta` is `None`). The tool puts
it in each table's entry, first in the result, on the first page only: `columns` (at most
`max_list_items`, `more_columns`), `added` and `removed` (ten keys each, `more_added`,
`more_removed`); the tables' summaries take at most half the page, the rest counted in
`cut.summaries`. A declared hidden column is no column of its table, so it is never
compared, counted or named (tested). **No new parameter**: the first page is the summary
plus the first 50 differences, and paging is the model's choice; a `summary_only` was not
needed for the common case. **Result size** (offline estimator, a made-up 200-row checks
table of six columns, keyed by code): 15 changes (7 ratings, 3 statuses, a due date, a
note, 2 rows removed, 1 added), one page either way: 1,406 tokens on main, 1,380 here (the
summary +125, the differences -151 by naming cells by key); every second rating changed
(100 changes, two pages): first page 3,348 → 2,936, both pages 6,740 → 5,863; the summary
alone, 120 tokens, answers "how many went up, down, and the largest". **Budget**: the
description adds a sentence on the summary and drops "with addresses": `xlsx_compare` 329 →
330 estimated, the loaded core 3,462 → 3,474 estimated, about 5,124 → 5,142 by the offline
proxy (5,009 online before; about 5,027 expected, not counted online); every definition
3,684 → 3,696 (guard 3,700). `strict_problems` is clean. **Goldens**: g2's comparison now
gives the summary and names differences by key, so its replay differs from the recording;
at first not re-recorded (recording calls a model): `tests/test_golden.py` converted the
replay back to the recorded form, and a new test checked g2's summary. g1 and g3 are
unchanged. *Re-recorded* (2026-10-14, approved by the user): g2 alone, by
`tools/record_goldens.py g2-compare-review` (claude-haiku-5-5, same settings, the same fixed
clock; three requests, 9,635 tokens written or uncached, 15,383 read from the cache, 689
out). The model's path is the old one -- `describe` d1, `describe` d2, one `xlsx_compare` of
`t:Stock` by SKU, same arguments -- and the replay now checks the current shape directly
(the conversion is gone); the summary test reads the recorded comparison. Its answer
reports the same four changes, now with the price's `+0.01`, but names SKU-0005's item
wrongly ("Wall plug 6 mm", SKU-0003's; the comparison names the row by key only and the
model took the name from the wrong sample row); the answer is for the reader and not
replayed, so it is kept as recorded.

*Title cells as built* (end-to-end check, 2026-10-14): **A title is no region.** A sheet
with a title in A1, a blank row and a table from row 3 listed the title as a detected
region of its own (`t:'Area North'!A1`, no rows, its one column the title), declared table
or not. A region that is one cell and no data rows (`Region.caption`) is now left out by
`Regions.find`, so `describe` does not list it; a lone number or date (unclear before,
`header_not_text`) is left out too. The cell stays reachable by `find_text`, a read and
its own `t:Sheet!A1`. Text two rows above a detected region's header, the only cell over
its columns there and the row between blank, is that region's `title` (`Region.title`;
`describe` shows it cut as a sample is). A title right above a header is still a two-row
header (`header_unclear`, unchanged); a declared table or Excel table gets no title (its
caller's `description` is its `about`; a title per sheet would cost a field per sheet). A
one-column table with data rows is a region as before. **Tokens** (offline proxy, the
end-to-end workbook: two sheets titled over the declared `checks`, a key-value sheet):
`describe`'s data 694 → 552, the whole result 766 → 625. Definitions unchanged
(43,415 / 42,739 for the three-format box), `strict_problems` clean. **Goldens** unchanged:
none of g1-g3 has a lone cell.

*T17 as built* (production feedback, 2026-10-14): **A hidden column stays hidden, every way
in.** A declared table over an Excel table (a ListObject) left both in `describe`: the
Excel table's entry gave the hidden column's name, its values (`min`/`max`, the few-values
list) and its sample rows, and `t:<the Excel table>` reached it through `xlsx_read`,
`xlsx_query`, `xlsx_profile`, `xlsx_derive`, `xlsx_compare`, `xlsx_validate_data`, a
dataset's `from`, `expand` and `find_text` (its header, by name). A defined name over the
hidden cells did the same, a validation list drawn from them showed them in the profile,
and `Workbook.read` of the sheet's cells gave them (only `xlsx_read` withheld them). Two
choke points now: **names** in `Regions` -- an Excel table a declared table covers any
cell of is replaced by it (`Regions.replaced_by`; left out of `Regions.tables`, so out of
`describe`, `find_text`, compare's pairing and `known()`; `t:<it>` in any form refused,
`declared`, with the declared table as `valid_options`; `Regions.excel_tables` keeps every
one, so its name stays taken and duplicate names are still flagged), and a defined name
over a hidden column's cells is refused (`declared`) and neither listed nor searched;
**cells** in `Workbook._read` -- every read of a sheet's cells gives a withheld cell blank
(the read tool still marks it `withheld`), so profile, catalogue samples, detection and
anything later that reads a range never see it. A validation list drawn from withheld
cells is not read (`reason: "withheld"`), a table's flags at a withheld cell are dropped
from the problems, and a second declared table that shows a column another hides is
refused (`hidden_elsewhere`, the column letter as `valid_options`; declaring a table again
under its own name still finds its hidden header). The Excel table's name is not listed
as an alias: it would name the same cells two ways with different columns. **Tokens**
(offline proxy, the test workbook: one Excel table under the declaration, two names over
the hidden column): `describe` 1,311 → 505. Definitions unchanged (no tool text changed).
**Not covered**: formula text of visible cells that names a hidden cell or column
(`=D2`, `ChecksTable[Note]`) is given as written; cells of a hidden column outside its
table's rows are not withheld (as before).

*T18 as built* (production feedback, 2026-10-14): **dates in a dataset made in Python.**
A dataset built in Python with `datetime.date` values typed them `text`, so `xlsx_derive`'s
`min` refused the column; the `date` label worked but was not documented. Now
`DatasetTable` reads a Python `date`, `datetime` or `time` as that type without a label,
its value ISO 8601 text as a cell's (to the millisecond; an aware datetime's wall-clock
time), so derive, `xlsx_read` (plain JSON) and keys see what a sheet's date gives; a
column mixing them with other values takes the majority type (`majority_type`), blanks
aside. **ISO text is not inferred** -- decided: a column of codes that happen to be
ISO-shaped would silently change type, and the reader's rule is that nothing is coerced
silently. ISO text in an *unlabelled* column is flagged `date_as_text`, as a cell's is,
so a refused `min` says so and offers both ways: the `date` label
(`Dataset.from_records(..., types={"opened": "date"})`, ooxml-edit 0.28.0, which also
gives `TYPE_LABELS`) or `coerce: ["date_as_text"]`; a column labelled `text` is left
alone. Placing: ooxml-edit writes a Python date in the session locale's day order
(`3/1/2026` en-US, `1/3/2026` nl-BE); a derive's dates are ISO text, as a sheet's are,
and are placed as written (until T18b). Definitions unchanged.

*T18b as built* (2026-10-14): **one way to write a date in a deck or document.** A date
column's ISO text -- a sheet's date column through `t:` (`table_dataset` labels a column
of one type with it), a derive's `min`/`max`/`first`/`last` of dates, a dataset labelled
`date`, `datetime` or `time` -- is placed as a Python date is, in the session locale's
day order (ooxml-edit 0.29.0, `sources.label(..., type=)` through `Rows.label`), so one
deck no longer mixes `1/3/2026` and `2026-01-20`. A column mixing dates with other types
has no label and keeps its ISO text, as does an elapsed time (`PT25H30M0S`). **Which text
wins:** a cell's own display text from its number format is never produced (*Values and
types*: no display string until Excel's is measured), so nothing competes with the
locale's writing; the workbook's results (`xlsx_read`, `xlsx_derive`, provenance) keep
ISO 8601 -- only placing writes the locale's. A chart's categories match a date column's
rows written either way (ISO, as an older deck has them, or in the locale). Definitions
unchanged.

*Withheld formulas as built* (T17's gaps, 2026-10-14): **a formula over a hidden column
keeps its value, not its text.** A visible cell's formula that reaches a withheld cell is
read with `formula_text` `None` and `Value.formula_withheld` set; `xlsx_read` with
`formulas` gives `"formula_withheld": true` in place of `f` (only when `formulas` is asked
for: a plain read is unchanged). `xlsx_read` is the one result that gives formula text
(`find_text`, the profile, the catalogue and the rest never do), and every read of a
sheet's cells goes through `Workbook._read`, so the rule sits there beside T17's blanking.
**What reaches** (`Regions.formula_withheld`, over `xlsx_agent.read.reach`, a scan of the
text for its references, not E2's lexer): a cell, range, whole column or row over a
withheld cell, on the formula's sheet, another, or a run of sheets (`North:South!D2`); a
structured reference to the hidden column, a run of columns over it, or the whole table
(`ChecksTable[Note]`, `[@Note]`, `ChecksTable[[#This Row],[Note]]`,
`ChecksTable[[Item]:[Note]]`, `ChecksTable`, `ChecksTable[#All]`), by the table's name or,
for `[@...]`, the table the cell lies in; a defined name over a withheld cell, or whose
formula reaches one (names in names followed); `INDIRECT` and `OFFSET`, whose reach the
text does not say; a spill reference (`C2#`) is taken to run to the sheet's end. String
literals, error literals and another workbook's references (`[1]Log!D2`) reach nothing; a
name or table the workbook does not have reaches nothing (a `LET` or `LAMBDA` name). A
shared formula's group is judged whole, its relative references moved over the range its
first cell's `ref` fills, so every cell of a group gets one answer. *Why the text*: it says
what the hidden column is -- its header in a structured reference, a name over it, a
literal its cells are compared with. *Why the value stays*: a visible cell is what the
owner shows, and a count, a sum or a flag computed from hidden cells is visible data; the
reader cannot tell a copy (`=D2`) from a figure over many cells, so a column whose formulas
copy hidden content is declared hidden too (documented in `declare`). **Below the table**:
a hidden column is now hidden from its header row (the first data row, without one) to the
sheet's last row (`Span.withheld`), on each of the table's sheets: a row added after
`last_row`, a total or a note under the table is the column's, and nothing in a
declaration says where its content ends; the cells above the header (a title, a run-details
block beside the table) are not the table's and stay visible, so a key-value table above
it may still show that column. A name, region or declared table over the column below the
table is refused as one over its rows. **`open_document`'s `tables`** counts the declared
tables as well as the Excel tables shown (a replaced one once, as its declared table), so
a workbook whose every Excel table is declared over no longer says `0`. Definitions
unchanged (no tool text changed; every definition 3,696, guard 3,700). **Not covered**:
`Workbook.cells` (a program's view of the file as stored) gives every formula as written,
as it gives hidden cells; the `arrays` an `xlsx_read` lists may name a block that crosses
the hidden column (its address only); a formula over a hidden column whose text names it
through a construct the scan does not know is given as written -- the scan errs towards
withholding wherever it can place nothing, and E2's lexer is the place to replace it.

*T19 as built* (production feedback, 2026-10-15): **descriptions that say what a parameter
is.** `xlsx_read`'s `key_column` was read as "the key to look up"; it is the *column*
holding the key, and now says so: "Column holding target's key (SKU for
t:Products/SKU-0042); if describe gives no or several key_candidates." `sources` says it
is for `d:` only, and its refusal on a cell or `t:` read says where it works and what to do
instead ("sources works on a dataset (d:) only, where each value lists the cells it came
from; a t: or cell read gives each cell its own address (a), its source: read it without
sources"). `xlsx_derive` has its own `sort` description (`xlsx_query`'s sorts by input
columns, which are its output): "By output column (group_by, a measure's as), not input".
A sort by a name that is no output column is refused as "sort names an output column: one
of Region, total", with the one meant: for an input column a measure reads, that measure's
(`'Amount' is an input column: did you mean 'total', its measure?`), else the nearest
name (ooxml-edit's `close_match`); under a pivot it points to `pivot_columns` and offers
no measure. **Budget**: `xlsx_read` 492 → 501, `xlsx_derive` 794 → 801 estimated; the
loaded core 3,474 → 3,490 estimated, about 5,142 → 5,165 by the offline proxy (5,024
online before: about 5,048 expected, to be counted online); every definition 3,696 →
3,712, the guard 3,700 → 3,716. `strict_problems` is clean; the strict set is unchanged
(no schema shape changed, only descriptions).

*T19b as built* (golden g2's answer, 2026-10-15): **each compared row with a label.** The
re-recorded g2 named SKU-0005 by another row's item: `xlsx_compare` named rows added,
removed and changed by key alone, and the model looked the item up in `describe`'s
samples. Each row and cell difference now carries its row's `label` after `at` -- its
text in the table's **label column** (`TableDiff.label`, in each table's entry after
`key`): the first column other than the key, in `a`'s order, whose visible cells over
both workbooks are mostly text (more text than all other types; blanks and cells on a
hidden row or column not counted), read even when `columns` leaves it out of the cells
compared. A changed cell and an added row take `b`'s label, a removed row its own; at most
60 characters (`LABEL_CHARS`), cut with `…`; none for a blank, non-text or hidden label
cell, and no `label` key at all then (a table without a text column gives none). A column
hidden from the model is no column of its table, so it is never the label (tested: with
the first text column hidden, the next labels and no hidden word appears). Declarations
have no label or title field, so none is read from them. **Size**: g2's comparison +4
labels and the column name: 434 → 470 estimated tokens (offline proxy). **Goldens**: g2 is *not*
re-recorded (no model call): its transcript keeps the result the model saw, and
`tests/test_golden.py` applies what tools have added since (`SINCE_RECORDED`: for g2's
`xlsx_compare`, the label column `Item` and each difference's label, written out from g2's
own `describe` results) to the recorded result before comparing it, key order included,
with the replay; a new test checks that SKU-0005 is labelled "Spirit level 60 cm". g2's
answer stays as recorded (it is for the reader, not replayed). No definition changes.

*T11 as built* (production feedback, 2026-10-15): **`describe` lists the session's
datasets.** A reviewer called a correct name unsupported because nothing in `describe`
showed the dataset it came from. ooxml-edit 0.30.0 adds `datasets` to every kind's
`describe`: per dataset its id, columns (at most 20 named), `types`, `rows` (a count) and
`from`, and `stale`/`closed` when true; the latest 20 datasets. `xlsx_derive` from a
workbook passes its table or region as the origin (`from: "d1/t:Sales"`); from a dataset,
that dataset (`"d:by_rep"`). Withholding holds: a derive from a declared table refuses its
hidden columns and `table_dataset` leaves them out, so a listed dataset never names one.
No definition changes.

*T14 as built* (production feedback, 2026-10-15): **a sheet withheld from the model is no
sheet of the workbook.** A declaration with every column hidden is refused (`no_columns`),
so a sheet a model must never see was blanked in the file. Now `{"sheets": "Notes",
"hidden": true}` (`declare.HiddenSheets`: `sheets` a name, glob or list; `optional`; no
other key) sits in the same declarations as the tables -- the shape the feedback asked for,
kept rather than a separate `hidden_sheets=` option: the format's `tables=`, a session's
`declare` and `Workbook.declare_tables` already carry it, and "kept whole or not at all"
covers it (a refused table withholds nothing). **Choke points**, beside T17's: the sheet
list -- `Workbook.sheets` leaves it out (`Workbook._every_sheet` keeps it for a name's
`localSheetId`, a run of sheets and the table parts), so `sheet`, `sheet_by_id`,
`sheet_of` and `_cell_sheet` (a `Sheet` object too) refuse it as `NameNotFound`, the same
message and options as for a name no sheet has, and everything that lists or walks sheets
(the catalogue, `find_text`, the problems, compare's sheet changes, `open_document`'s count,
a table's glob, a schema's required sheets) never meets it; **names and tables** --
`Regions.excel_tables` and `Regions.names` leave out a table on it and a defined name local
to it, over its cells, or whose formula names it (its name anywhere in the text, a run of
sheets over it, its table, such a name), while `_every_table`/`_every_name` keep them for
the formula scan; **cells** -- `Regions.withheld` gives a withheld sheet's every cell, so
T17b's `formula_withheld` withholds a visible formula's text when it reaches the sheet
(reference, run, table, name, `INDIRECT`/`OFFSET`), and also when the text holds the
sheet's name anywhere (`HYPERLINK("#Notes!A1")`, a string an `INDIRECT` takes): the value
stays. A link whose location names it is left out of `Workbook.hyperlinks`; a drop-down
drawn from it reads `unknown_sheet`, as from no sheet. **Decided**: the name is not
listed at all, not even as withheld (a sheet's name may say what it holds); a hidden
sheet's declaration that matches nothing is refused by path only, never echoing the name,
since a refusal at `open_document` reaches the model; ids are not renumbered (a gap shows,
as after a deletion); a sheet stays withheld once withheld, and a table or region kept
already over it is refused (`on_hidden_sheet`, naming the table). The catalogue lists no
charts or pivot tables, so none can reveal it. **Tokens** (offline proxy, the test
workbook: a log sheet whose formulas reach a withheld sheet of an Excel table, a comment
and three names, and a spare sheet): `describe` 1,075 → 612. **Not covered**: a cell's own
text that names the sheet is content and shown; `Workbook.to_bytes` is the file as stored.
Definitions unchanged (every definition 3,712, guard 3,716; a format withholding a sheet
gives the plain one's definitions, tested).

*T16 as built* (production feedback, 2026-10-15): **a header found by its text.** Sheets
of several sections needed code to find each header row. A declaration's `header_row` may
now be `{"match": [texts], "rows": "8:40"}`: on each sheet, the one row whose text cells
(in `columns`, else the used range's) hold every text -- each a cell's whole text, case and
runs of white space aside, any order, other cells besides -- within `rows`, or without it
below the last row of the tables declared before it in the same declarations on that sheet;
each sheet of a table over several finds its own row (`Part.header`), the header text still
the same on all. The data rows run to the first blank row or the row before the next row
that matches, whichever is first, and nothing past that match is flagged `rows_below`;
`first_row` and `last_row` are refused with it (`bad_rows`). **Decided**: exactly one row in
the rows looked in -- none is `no_header`, several `ambiguous_header` with the rows by
number as `valid_options` and what to do (give `rows`, or more texts) -- rather than the
first, so a section never silently takes another's header; whole texts, so `Ref` does not
find `Reference`. **Not covered**: a hidden column of one section is hidden to the sheet's
last row (T17b's *Below the table*), so a section below that shows the same column is
refused (`hidden_elsewhere`) -- hide it there too, or put the sections in other columns.
Definitions unchanged; `describe` of the two-section test sheet 459 → 471 estimated tokens
(two declared tables in place of two detected ones). *Since changed* (with T15, below): a
hidden column now ends just above the next declared section's header.

*T15 as built* (production feedback, 2026-10-15): **terms redacted from every text.** A
name that may only appear under a condition was removed by rewriting the shared strings.
Now `{"redact": ["Acme Corp", "re:Project\\s+\\w+"], "with": "[redacted]"}` sits in the
same declarations as the tables and hidden sheets (`declare.RedactTerms`,
`xlsx_agent.redact`; `whole_words` optional, several apply in order). **Terms**: a literal
is found case aside, any run of white space for its spaces, inside words unless
`whole_words` (a term missed is a leak, an extra match an untidy cell); `re:` is a Python
regular expression, checked as X6's `strip` is and more strictly -- no repeat may give
characters back (`[^;]*;`, `\s+\w+` pass; `.*;` is refused), nor may it match empty
text. A redaction searches the whole text, so a text over 512 characters is searched
`WINDOW` (256) characters at a time in windows of 512, a match a window's end cuts
checked against the whole text: matches up to 256 characters are always found. Worst
pattern tried, `[a-z]+[0-9]` on 32,767 letters: 0.26 s (3.12), 0.24 s (3.15), once per
distinct text; two literal terms about 2 µs a cell (Intel Core i7-6567U, 2026-10-15).
**Choke points**: the shared strings (redacted as the table is read, so every reader and
`find_text`'s pre-check see them so), the sheet reader (inline strings, a formula's cached
text, a number format's code; a formula whose text holds a match keeps its value, not its
text, `formula_withheld`, a shared group's first cell so judged), `Workbook.comments`
(text and author), `Workbook.hyperlinks` (address and place), the drop-down's literal list,
an Excel table's column names (before duplicates are named by letter), and the sheet list.
Everything else -- describe, reads, query, profile, derive, compare, validate, keys, a
`t:` source in a deck or document, the problems -- reads through those. **Names,
decided**: shown redacted where an address can still hold them -- a sheet (quoted;
`Sheet.stored_name` keeps the file's, two alike made unique with ` (2)`), an Excel table's
column (escaped, `t:T['[redacted'] share]`) -- and left out where a redacted name would be
no name: an Excel table or defined name whose name holds a match is not listed (its cells
read by address; the table's found as a block of its sheet), as T14 does for a withheld
sheet's names. A defined name or drop-down whose formula names a redacted sheet by its
stored name resolves to the shown one. Keys read redacted, so two that redact alike are
`duplicate_key`. **The caller's words**: declarations name sheets as stored (name or
glob; a hidden sheet too) and find headers, keys and `match` texts by their stored text
(each redacted before it is compared); a table name, column id or description of the
caller's that holds a match is refused (`redacted_term`, path only); a refusal never
echoes a term. **Order**: set first -- refused once a table, region, key or hidden sheet
is kept (`redact_late`), unless the same rules; a failing declaration list leaves no
redaction set. `Workbook.redact()` for a program. **No note to the model** (decided): the
replacement is the only sign -- no count or flag in any result, since how often a name
matched is information about it; `Workbook.redaction.redacted` counts the distinct texts
changed, for the program. `find_text` searches the redacted text, so a term is neither
found nor said to have been there (tested: a search for one gives the same result as for
a word the workbook lacks). **Not covered**: `Workbook.to_bytes` is the file as stored;
`DefinedName.formula` and `Validation.formula1` (program fields, never in a result) are as
stored; chart titles, pivot tables, headers and footers are not read by any tool.
**Tokens** (offline proxy, the test workbook): `describe` 1,105 → 998 (an Excel table and a
name left out). Definitions unchanged (tested: a format redacting gives the plain one's
definitions; every definition 3,712, guard 3,716).

*Sections bound a hidden column* (T15's follow-up, main-session decision, 2026-10-15):
with T16 sections, a hidden column hidden to the sheet's last row refused a later section
showing that column (`hidden_elsewhere`). Now a declared table's hidden column ends just
above the header row (the first row, without one) of the next table declared below it on
the same sheet (`Part.end`; tables declared in the same list, or kept already) -- the
cells between the sections stay hidden. Only declared tables bound it: an Excel table or
a detected block below does not. A table kept already is not re-bounded by one declared
later (declare the sections together); a table over the hidden table's own rows is still
`hidden_elsewhere`.

### 0.2 — edit

| # | Item | Effort | Done when |
| --- | --- | --- | --- |
| E1 | **Edit foundation.** `History`, `batch`, `undo`, `save_document` (O1), `validate()` (structure: cell order, shared-string counts, table refs, names), the stale-cache marker; the **Excel oracle** (opt-in `pytest -m oracle`: Excel opens, re-saves, and the diff shows no repair), with a machine-wide lock as docx-agent's for Word; the built-in number-format probe. Needs O4 | M | every fixture round-trips; undo gives the original bytes; Excel opens every edited fixture unprompted |
| E2 | **Formula lexer** (references only, no evaluation): cell and range references, sheet-qualified and quoted, defined names, structured references; finds and rewrites them | M | rewriting a sheet name in a corpus of made-up formulas round-trips, and Excel agrees (oracle) |
| E3 | **`xlsx_set_cells`**: values and formulas, types kept, shared strings, calculation chain dropped, `fullCalcOnLoad` set, a shared formula group split when a member changes | M | Excel opens the results unprompted; reads after an edit mark dependent caches stale |
| E4 | **`xlsx_write_table`**: a real table on a new sheet or an empty area, from data or `d:<id>`, long text split | S | Excel opens it as a table with filters and banding |
| E5 | **`xlsx_manage_sheets`** on E2: rename rewrites references; delete refuses and lists referents | S | Excel opens a renamed fixture with every formula and name intact; a refused delete lists each referent by address |
| E6 | **`xlsx_validation`**: drop-down lists and value rules, kept beside any existing rules | S | Excel shows the drop-down and rejects an out-of-rule entry (oracle) |
| E7 | **`xlsx_format`**: styles reused, number formats, fonts, fills, borders, alignment, widths (auto-fit measured: Excel's width unit and padding probed from Excel), freeze panes, print settings | M | auto-fit widths match Excel's on the probe's columns |
| E8 | **`xlsx_conditional_format`**: colour scales, data bars, icon sets, rule fills (`dxf` styles reused) | S | Excel opens each rule kind unprompted and lists it in its rules manager (oracle re-save keeps it) |
| E9 | **`edit_chart` and `new_document`.** Needs O5; a new workbook's parts measured from what Excel writes for a new workbook | M | a chart added from a dataset opens in Excel with its data in the sheet |

---

## Testing

- **Made-up fixtures only, built in the tests.** `tests/build.py` writes every fixture from
  Python with neutral content (a product inventory, a sales table, a stock list), so the
  repository holds no workbook from anywhere else. Each row of *Messy workbooks* has one.
  From 0.2, a few are also opened and saved by Excel through a committed script, so the
  suite has Excel-authored files (the siblings' method); those are committed with the
  script that made them.
- **Lossless round trip.** Every fixture opened and saved is byte-identical per part; after
  an edit, untouched parts are byte-identical; undo after any edit set gives the original
  bytes, redo the edited ones.
- **Provenance.** For every value of every derived dataset on every fixture, the test
  re-reads its source cells, drops the listed exclusions and recomputes the measure with a
  plain reference implementation; the two must agree exactly.
- **Bounds.** Every tool on the largest fixture stays within its caps and reports the cut.
- **No content in diagnostics.** Fixtures carry marker strings in cells, comments and
  formulas; the test collects every warning, error message and log record from a full run
  and fails if a marker appears.
- **Untrusted input.** An entity-expansion part, an external-entity part, a zip bomb and an
  over-long part are refused without expanding them.
- **Token budget.** `tests/test_tools_schema.py`, as the siblings': every schema in the
  common strict subset, no parameter a path, groups as listed, the loaded core under 6,000
  by the offline proxy; `tests/test_tools_online.py` (`provider`) counts it online.
- **The Excel oracle** (0.2, opt-in, never under xdist): Excel opens and re-saves edited
  fixtures; a repair prompt or a changed part we did not expect fails. Excel is never run
  in the default suite.

---

## Open questions

1. **Inserting and deleting rows and columns.** Needs E2's lexer and a rewrite of every
   reference (formulas, names, tables, merges, validation, conditional formats, charts,
   pivot sources). After 0.2, on demand.
2. **Recalculation** (decision 3). What would trigger revisiting it: agents meeting
   `no_cached_value` often enough in practice. Record the counts when they come.
3. **Display strings before 0.3.** Whether an agent needs "as Excel shows it" for some
   formats sooner; if so, which formats, measured from Excel first.
4. **Datasets after their source changes.** Stale marking is decided; whether
   `xlsx_derive` should offer a refresh of a stale dataset by id is not.
5. **Chart workbooks inside decks and documents.** Reading a chart's embedded workbook
   through `open_embedded` costs little; whether it should be an address (`d1:chart 3`)
   or stay a library call.
6. **Pivot tables as a source.** Their cached rows could be read as a region; whether that
   is wanted, given that the data underneath is usually readable too.
7. **The Excel oracle's lock.** One lock per Office application, or one shared with the
   Word and PowerPoint oracles.
8. **Strict OOXML.** Read in 0.1; whether 0.2 edits it, which needs Excel's behaviour on
   a strict file measured first.
