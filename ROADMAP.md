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
| dataset | `d:by_category` | a derived dataset in the session; `d:by_category[average]`, `d:by_category/Hardware/average` as for tables |
| ref | `$step1` | the siblings' refs: a name a call gave to what it made |

Rules, from the siblings: **ids name things, positions name places, and the API says which
it is giving you.** A document handle is `d1` (no colon); a dataset is `d:` plus a name, so
the two never collide. A table's key column is the one the caller names, or else one
`describe` reports as a key (unique, never blank, text or whole numbers); `value by key`
without a key column is refused with the candidate columns as `valid_options`. Names are
matched as Excel matches them (case aside); an unknown name is `not_found` with the
nearest names.

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
| `describe` (shared) | `doc` | the catalogue: sheets (id, name, visible, hidden or very hidden, protected, used range); tables, names and detected regions with their key columns; per column: header, inferred type, fill rate, min/max/mean or distinct values when 12 or fewer, text length median and max, formula share, data-validation list, flags (mixed types, numbers as text, no cached value, cut at 32,767); up to two sample rows per table; workbook flags (recalculation, external links, macros) | about 3,000 tokens by default; a workbook too big for that lists every table and the first columns, and says which were left for `xlsx_profile` |
| `xlsx_profile` | `doc`, `target` (table or column), `columns`, `cursor` | the full per-column profile for one table, every column | one table per call; 50 columns a page, within 24,000 characters; paged |
| `xlsx_read` | `doc`, `target` (range, table, column, row, dataset), `formats`, `formulas`, `comments`, `hyperlinks`, `merges`, `key_column`, `limit`, `max_chars`, `cursor` | typed values by row, with addresses; optionally format codes, formula text, comments, links, merges; the array blocks over the rows always | default 50 rows, at most 500 or 24,000 characters; paged, a row too wide for a page split across pages |
| `xlsx_query` | `doc`, `table`, `columns`, `where` (`eq`, `in`, `gte`, `lte`, `contains`, `blank`, `not_blank`, combined with and), `sort`, `limit`, `cursor` | matching rows with their addresses, and the count of all matches | as `xlsx_read`; a filter on a column of the wrong type is refused naming the column's type |
| `find_text` (shared) | `doc`, `text`, `regex`, `range` (a sheet, range or table), `cursor` | matches by address and kind: cell (text cells; numbers and dates are not searched, having no display string yet), table or region name, column header, comment, defined name; hidden ones marked; formula text not searched (the shared tool has no option for it) | 50 per page |
| `xlsx_derive` | `doc`, `from` (table, region or dataset), `where`, `group_by`, `measures` (`count`, `sum`, `mean`, `median`, `min`, `max`, `share`, `distinct`, each with `column` and `as`), `bins` (edges for a number or date column), `pivot` (a column whose values become columns), `sort`, `top`, `round`, `id` | `d:<id>`: columns, rows, the rounding applied, and for every value its sources (see *Provenance*); the rows left out and why, counted by reason | at most 1,000 result rows and 50 shown; a measure on a column whose type does not allow it is refused before anything is computed |
| `xlsx_compare` | `a`, `b` (two open documents), `table` (or one per document), `key`, `columns`, `cursor` | rows added and removed (by key), cells changed (address in each, old and new value, type), columns added and removed; duplicate keys reported, not paired by guess | 50 differences per page, totals always |
| `xlsx_validate_data` | `doc`, `schema`: sheets and tables required; per column type, required, unique, enum, `min`/`max`, `step` (e.g. 1–5 in halves), pattern | problems by address and code (`missing_column`, `wrong_type`, `out_of_range`, `not_in_enum`, `duplicate_key`, `required_blank`...), counted by code | 50 problems per page, counts always |

`xlsx_compare` and `xlsx_validate_data` are in the core if the budget holds; if the
online count says otherwise they move to a deferred `xlsx_check` group, and R14 records the
numbers.

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
ds.excluded                          # {"numbers_as_text": ["Stock!F9"], "blank": [...]}

Workbook.compare(old, new, table="t:Products", key="SKU")   # Diff
wb.validate_data(schema)             # [Problem(address, code)], never cell text
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
   opt in to `coerce: ["numbers_as_text"]`, and the result says it did.
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
9. **Facts, not verdicts.** Flags say what is there (`numbers_as_text: 14 cells`), not
   whether it is acceptable.

---

## Messy workbooks

What the reader meets, and what it does. Each item has a made-up fixture (*Testing*).

| Situation | Behaviour |
| --- | --- |
| merged header cells, two or more header rows | not detected; `describe` flags the region `header_unclear` and points to `declare_region` |
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
| **O2** | A session-level **dataset store**: `Session.datasets`, `Call.define_dataset(id, dataset)` and `Call.dataset(address)`; `d:<id>` resolved by the dispatcher in a tool's declared fields (as `Tool.refs` declares ref fields), refused with the known ids as `valid_options`; a dataset records its source document and that document's `version`, outlives `close_document`, and is marked `stale` (never refreshed silently) once its source changes; `Limits.max_datasets` and `Limits.max_dataset_cells`. Format-neutral: a dataset is columns, rows and sources | R11 |
| **O3** | Consumers of datasets: `edit_chart`'s `add`, `set_values` and `add_series` take `from: "d:<id>"` with a column mapping, and the chart keeps the dataset's sources in its result; pptx-agent's `ppt_add_table`/`ppt_edit_table` and docx-agent's table path take the same. Changes in ooxml-edit, pptx-agent and docx-agent, each held to its own core budget | after 0.1 (it is what makes `d:` cross-format) |
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
