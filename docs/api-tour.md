# A tour of the API

**Status: 0.1.0, milestones R0 to R14 of [`ROADMAP.md`](../ROADMAP.md)** -- read and derive.
A workbook opens from bytes, a path or a package, is catalogued in one bounded result, and
is read, queried, profiled and searched by address; figures are derived here, never by the
caller, each with the cells it came from, and kept as datasets; two workbooks are compared
by key, and a workbook is checked against a declared schema. Nothing is edited or saved
yet (0.2). The tools are thin wrappers over what follows: see
[the tools README](../src/xlsx_agent/tools/README.md).

The examples run on the made-up workbooks of the golden transcripts
(`tests/golden_inputs.py`): a sales table `t:Sales` (Date, Region, Product, Units, Revenue;
16 rows), an inventory `t:Stock` and a reviewed copy of it, and an order list `t:Orders`.
The outputs shown are what they print, shortened.

## Opening

```python
from xlsx_agent import Workbook, Locale      # Locale: ooxml_common.locales.Locale

wb = Workbook.open("sales-july.xlsx")         # a path, bytes, a binary file or an OpcPackage
wb.sheets      # [Sheet(name='Sales', id=1, state='visible', protected=False, type='worksheet')]
wb.problems    # [] -- what the reader met, by address and code: Problem(address, code)
wb.tables      # [Region(kind='table', name='Sales', sheet='Sales', columns=('Date', ...), ...)]
wb.calculation # [] -- or 'full_calc_on_load', 'manual_calculation': caches may be stale
wb.to_bytes()  # the bytes that were read
```

`.xlsx`, `.xlsm`, `.xltx` and `.xltm` open. A `.xls`, a `.xlsb`, an encrypted workbook or
anything else is refused with an `OpenError` whose `code` says why; a package over
ooxml-edit's zip limits (`Limits`, given as `limits=`) is refused before it is unpacked.
A macro is kept and never run; an external link is listed, never followed.

**Locale and reader options.** The locale decides how text that looks like a number or a
date is recognised (and the language of the language-dependent built-in formats); en-US
when none is given, never the machine's:

```python
belgian = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
                 language="nl-BE", list_separator=";")
wb = Workbook.open(data, locale=belgian,
                   grouping=True,       # "1.234,5" is a number (the default)
                   local_dates=True)    # "3/4/2026" is 3 April, as Excel reads it here
```

## The catalogue: describe and catalogue

```python
print(wb.describe(max_tokens=3000))     # what the describe tool returns, as compact JSON
wb.catalogue()                          # the same as data: kind, date_system, sheets, regions
```

Each region (table, defined name, declared or detected region) comes with its range, row
count, key columns and, per column, the inferred type, fill rate, min/max/mean or the
distinct values when there are few, text lengths, formula share, validation lists and
flags; up to two sample rows. A workbook too big for the token target lists every table
and its first columns and names what it left for the profile.

## Reading by address

```python
wb.cell("Sales!E2").as_dict()            # {'a': 'Sales!E2', 'v': 159.6, 't': 'number'}
for row in wb.read("Sales!A1:E3"): ...   # lists of Value, typed: number, percent, date,
                                         # time, datetime, text, bool, error, blank
wb.read("Sales!A:A")                     # whole columns and rows clip to the used range
wb.resolve("t:Stock/SKU-0002/Price", key="SKU")   # Selection(range=Stock!E3, ...)
wb.comments("Sales"), wb.hyperlinks("Sales"), wb.merges("Sales"), wb.arrays("Sales")
```

A value read is the one stored: a date as ISO 8601 (`'2026-07-01'`, with its `serial`), a
formula's cached value (none is `blank` with `issue='no_cached_value'`), no display
string. Every form of address is in ROADMAP.md, "Addresses": `Sales!A1:E3`, `#1!E2` (by
sheet id), `t:Sales`, `t:Sales[Revenue]`, `t:Stock/SKU-0002/Price`, `t:Sales/#5`,
`t:Stock/*/Price`, `t:Stock/{SKU-0002,SKU-0007}/Price`, `d:by_region`.

**Addresses by key** name a row by its key column's value, so they keep naming the same
cell when a consultant inserts, deletes or sorts rows (a position, `t:Sales/#5`, or a
range does not). `*` is every row (or every column), and `expand` gives what an address
covers one cell at a time, by key -- a coverage selector:

```python
wb.declare_key("t:Stock", "SKU")         # the caller's choice, when several columns could
                                         # be the key; one detected key needs none
wb.read("t:Stock/{SKU-0007,SKU-0002}/Price")   # those rows, in that order
wb.expand("t:Stock/*/Price")             # ['t:Stock/SKU-0001/Price', 't:Stock/SKU-0002/Price', ...]
wb.expand("t:Stock/SKU-0002/*")          # each value of one row
required = set(wb.expand("t:Stock/*/Price"))
used = {a for s in sources for a in wb.expand(s)}     # e.g. a derived value's sources
missing = required - used                # cells no slide, check or plan used
```

A key is a cell's text, exactly (`007` is not the number 7), or a whole number's digits (the
text `7` and the number 7 are one key); a blank, a fraction, a date and a boolean are no
key. A row whose key another row shares is refused (`duplicate_key`), and `expand` refuses
a row without a key (`no_key`), each with the rows by position: never the first of them.

Merged or multi-row headers are declared, never guessed:

```python
wb.declare_region("Returns", "Returns!A3:F200", header_rows=2,
                  headers=["SKU", "Date", "Qty", "Reason", "Store", "Refund"])
```

## Declaring tables

An application that knows its workbooks -- a template filled in per run, say -- declares
their tables from its own configuration, as plain data, before anything is described. A
model never does: no tool takes a declaration (the tool layer takes them from the
application, `xlsx_agent.tools.format(tables=...)` or `declare(session, ...)`). This
example's workbook (made up, as in `tests/test_declare.py`) has a sheet of checks per area,
each with a title over its header on row 3, and a hidden sheet of run facts:

```python
wb.declare_tables({"tables": [
    {"name": "checks", "sheets": "Area *", "header_row": 3, "key": "Ref",
     "description": "One row per check, every area",
     "fields": [{"header": "Rating (1 to 5, in halves)", "id": "rating",
                 "description": "The reviewer's rating"},
                {"header": "Remark", "hidden": True}]},
    {"name": "run", "type": "key_value", "sheets": "Run facts"}]})

wb.region("t:checks").where       # "'Area North'!A3:D5,'Area South'!A3:D6"
wb.cell("t:checks/S-2/rating").as_dict()  # {'a': "'Area South'!C5", 'v': 5, 't': 'number'}
wb.cell("t:checks/#3/Topic").a    # "'Area South'!B4": positions count across the sheets
wb.cell("t:run/report_date/value").as_dict()
# {'a': "'Run facts'!B1", 'v': '2026-10-01', 't': 'date', 'serial': 46296.0, 'hidden': True}
ds = wb.derive("t:checks", group_by=["Topic"],
               measures=[{"fn": "mean", "column": "rating", "as": "avg"}])
ds.rows[0], ds.provenance[0]["avg"].sources   # {'Topic': 'door', 'avg': 2.75},
                                              # (t:checks/{N-1,S-1}/rating,)
wb.cell("t:checks/N-1/Remark")    # NameNotFound: its columns are Ref, Topic, rating, Note
```

What a declaration says:

- **Sheets** by exact name (case aside) or glob (`Area *`), one or a list. A table over
  several sheets is one table: their rows in the workbook's order, the header the same text
  on every sheet. Every value keeps its own cell's address -- in reads, queries, samples,
  a derive's sources and left-out cells, datasets, comparisons and checks -- and a key finds
  its row on whichever sheet holds it. The table's range is each sheet's, joined by `,`.
- **Rows**: `header_row` (1 by default; 0 for none), `first_row`, `last_row`. Without
  `last_row` the data ends at the first blank row, as a detected region's does; anything
  further down within the columns is flagged `rows_below`. `columns` (`"A:F"`) when the
  header row's extent is not the table's.
- **The key**: every row has one, and no two rows share one, across every sheet; otherwise
  the declaration is refused, naming the cells (never the key).
- **Fields** (by `header` or `column` letter): an `id` to address and show the column by
  (its header still finds it), a `description` (shown by `describe` as `about`, by
  `xlsx_profile` as `description`, with the header), and `hidden`: left out of everything a
  model sees -- describe, read, query, derive, profile, compare, validate, a deck's `from`,
  `find_text`, a region detected or declared over it -- and refused as unknown when named.
  A column Excel hides is a column like any other unless declared hidden.
- **The sheet is not a column.** A value is a cell, and the sheet's name is no cell's, so
  there is no `_sheet` to group by; every address names its sheet. For one sheet's
  figures, declare its own table too (`"sheets": "Area North"`).
- **Key-value tables** (`"type": "key_value"`): a field column, a value column and any
  more; `t:run/<field>/value`, shown compactly by `describe` as its fields and values. A
  two-column block without a header whose first row holds a date or a number (so no header
  is detected) and whose first column is a different text on every row is detected as one:
  `t:'Run facts'!A1/author/value`.

A declaration that does not fit the workbook is refused before any table is kept
(`DeclarationError`: `code`, `path` such as `tables[0].fields[2].header`,
`valid_options`): an unknown sheet, a glob matching none (unless `optional`), a header that
differs on one sheet, an unknown header, an id that would name two columns, a hidden key,
a key two rows share. Nothing is guessed past.

## Cleaning text

A declared field may also say how its text is cleaned (`xlsx_agent.clean`): a leading label
stripped, white space collapsed, a cell split into a list. Like the rest of a declaration,
the rules are the application's configuration -- no tool takes them, so no definition
changes -- and they are applied once, where the table is read. This example's workbook
(made up, as in `tests/test_clean.py`) has a sheet of tasks per team whose "Owners" cell
starts with the lead's label before the owners, `Lead: Robin;  Kim ;Ash`:

```python
wb.declare_tables([{"name": "tasks", "sheets": "Team *", "key": "Ref", "fields": [
    {"header": "Owners", "id": "owners",
     "clean": {"strip": r"^Lead:[^;]*;\s*", "split": ";"}},
    {"header": "Note", "clean": {"strip": r"^Lead:[^;]*;\s*", "collapse_space": True}}]}])

wb.cell("t:tasks/E-1/owners").as_dict()
# {'a': "'Team East'!C2", 'v': ['Kim', 'Ash'], 't': 'list', 'cleaned': True}
wb.cell("t:tasks/E-1/owners").original     # 'Lead: Robin;  Kim ;Ash': the text as stored
wb.cell("'Team East'!C2").v                # the same: a cell by its address is as stored
wb.cell("t:tasks/W-1/Note").as_dict()      # only a label: nothing left, so blank
# {'a': "'Team West'!D2", 'v': None, 't': 'blank', 'cleaned': True}
wb.column("t:tasks[owners]").stats().flags
# {'cleaned': Count(count=2, first=("'Team East'!C2", "'Team West'!C2"))}
```

- **The rules**, in this order, on text values only: `strip` (a regular expression matched
  at the start, within the cell's first 256 characters), `split` (a literal separator; each
  item trimmed, empty ones dropped; the value a `list` -- a tuple in Python, a list in
  JSON -- even of one item), `collapse_space` (runs of white space become one space, none
  at the ends; on each item of a split). Text cleaned to nothing is `blank`; cleaned text
  that reads as a number or date is flagged `number_as_text` or `date_as_text`, so a
  derive's `coerce` reads it.
- **Never silent**: a value a rule changed says `cleaned` and keeps `original`; the cells
  are counted with their first addresses in an `xlsx_read` or `xlsx_query` page
  (`cleaned`), as a `cleaned` flag in `xlsx_profile` and `describe`, per column in
  `xlsx_derive` (`Derived.cleaned`), and as a `CleanedWarning` when the table goes into a
  deck or document (the tool's result lists it among its warnings).
- **Bounded**: a pattern is compiled when declared, and refused (`bad_pattern`) when its
  backtracking could grow past the square of the 256 characters it sees -- a repeat in a
  repeat (`(a+)+`), an alternation in a repeat (`(a|ab)*`), a backreference, or more than
  one repeat that can give characters back (`.*;` can; `[^;]*;` cannot).
- **Not the key**: the key column is never cleaned (`clean_key`): an address by key names
  the key as its cell holds it, and two keys cleaned alike would name one row. Nor is a
  hidden column, nor the header.

Lists flow on as follows:

```python
tasks = wb.region("t:tasks")
[r[0].v for r in tasks.query(["Ref"], where={"owners": {"eq": "Kim"}}).rows]
# ['E-1', 'E-2']: eq, in and contains pass when any item does
wb.derive("t:tasks", measures=[{"fn": "count", "column": "owners", "as": "with owners"}]).rows
# [{'with owners': 3}]: cells holding a list (rows, not items)
wb.derive("t:tasks", group_by=["owners"])  # DeriveError wrong_type: a list is no one group
table_dataset(wb, "t:tasks[owners]").rows  # into a deck or document: a line per item
# (('Kim\nAsh',), ('Kim',), ('Ash\nLee',), (None,))
```

Every other function refuses a list column; `xlsx_compare` compares lists as values;
`xlsx_validate_data` checks a list as `text`, item by item (an `enum` holds for each).

## Query

```python
sales = wb.table("t:Sales")
sales.query(["Date", "Revenue"], where={"Region": {"eq": "North"}},
            sort=[("Revenue", "desc")], limit=2)
# QueryResult(address='t:Sales', columns=('Date', 'Revenue'),
#             rows=((Value(... row=5, v='2026-07-06', t='date'), Value(... v=177, ...)), ...),
#             matches=5, ...)
```

Filters are `eq`, `in`, `gte`, `lte`, `contains`, `words_gt`, `blank` and `not_blank`,
combined with and; a filter on a column of the wrong type is refused naming the column's
type. `{"Note": {"words_gt": 150}}` finds the text cells longer than 150 words.

## Profile

```python
wb.table("t:Sales").column("Revenue").stats()
# ColumnStats(address='t:Sales[Revenue]', range='Sales!E2:E17', rows=16, type='number',
#             fill=1.0, min=34.5, max=239.4, mean=108.83125, distinct=12, ...)
wb.profile("t:Sales", ["Units"])         # every column (or these) in one streamed pass
```

A date column's `min` and `max` are its range, as dates. Text columns give their lengths in
characters and in words (median and max) and the address of the longest text; dates stored
as text are flagged `date_as_text`, counted, and their own range is `date_text`.

```python
wb.key_columns("t:Stock")    # ['SKU', 'Item', 'Qty']: unique, never blank, text or whole
                             # numbers; with more than one, a value by key names its key=
```

## Find

```python
list(wb.find_text("Kettle", within="Sales!"))
# [Found(kind='cell', address='Sales!C2', start=0, end=6, ...), ...]
```

Text cells, table and region names, column headers, comments and defined names are
searched (a compiled pattern for a regex); numbers, dates and formula text are not.

## Derive, and datasets

```python
ds = wb.derive("t:Sales", group_by=["Region"],
               measures=[{"fn": "sum", "column": "Revenue", "as": "revenue"},
                         {"fn": "count", "as": "orders"}],
               sort=[("revenue", "desc")], round=2, id="by_region")
ds.rows
# [{'Region': 'North', 'revenue': 509.9, 'orders': 5},
#  {'Region': 'West', 'revenue': 470.9, 'orders': 3}, ...]
ds.sources["North"]["revenue"]           # Sales!E2, E5, E9, E13, E17 (RangeAddress each)
ds.excluded                              # {} -- else per measure, per reason, with addresses
```

When the table has a key column -- declared, or the one column that is unique, never blank,
text or whole numbers -- each value's sources are written by key instead, one key set in
the table's order, which still names the same cells after rows are inserted or sorted:

```python
by_category = wb.derive("t:Stock", group_by=["Category"],
                        measures=[{"fn": "sum", "column": "Qty"}])
by_category.source_key                   # 'SKU'
[str(s) for s in by_category.provenance[0]["sum Qty"].sources]
# ['t:Stock/{SKU-0001,SKU-0002,SKU-0003}/Qty']
```

The measures are `count`, `sum`, `mean`, `median`, `min`, `max`, `first`, `last`, `share`
and `distinct`, with `where`, `bins`, `pivot`, `sort` and `top`. `min` and `max` take dates
and give a date; `first` and `last` take any column and give the value of the first and
last row used, in the table's order. Sums are `math.fsum`; rounding is half away from zero
on the reported value only. Blanks, errors, numbers and dates stored as text and formulas
without a cached value are left out of a measure and counted by reason
(`coerce=["number_as_text", "date_as_text"]` opts in). A measure on a column whose type does
not take it is refused (`DeriveError`) before anything is computed; when the column is text
holding dates or numbers stored as text, the refusal says how many, names the first cell
and the `coerce` that reads them (`error.issue`, `error.cells`).

A derived result is a dataset, read back by address and derived from again -- its sources
stay the original cells:

```python
from xlsx_agent import DatasetTable

ds.dataset()        # ooxml_edit.tools.Dataset: columns, rows as reported, sources as text
DatasetTable(ds).read("d:by_region/North/revenue")   # [[DatasetValue(d:by_region/#1/revenue, ...)]]
wb.derive(ds, measures=[{"fn": "max", "column": "revenue", "as": "best"}]).rows
# [{'best': 509.9}]
```

In the tool layer `xlsx_derive` keeps the dataset in the session as `d:by_region`, which
`xlsx_read` reads (with `sources: true`) after its workbook is closed too.

## Workbook tables into slides and documents

A table's rows as an ooxml-edit `Dataset` -- what a deck or document tool places when given
`from: "t:Stock"` (pptx-agent's `ppt_add_table`, the shared `edit_chart`) -- typed as read,
nothing coerced, each value with its cell by key where the table has a key column, else its
cell address:

```python
wb.dataset("t:Stock").sources[0][3]      # ('Stock!D2',): several key candidates, none named
wb.declare_key("t:Stock", "SKU")
rows = wb.dataset("t:Stock/{SKU-0003,SKU-0001}")   # or t:Stock, t:Stock[Qty], t:Stock/*/Qty
rows.columns                             # ('SKU', 'Item', 'Category', 'Qty', 'Price')
rows.rows[0]                             # ('SKU-0003', 'Wall plug 6 mm', 'Hardware', 450, 0.05)
rows.sources[0][3]                       # ('t:Stock/SKU-0003/Qty',)

from xlsx_agent import merge_sources
merge_sources(["Stock!D2", "Stock!D3", "Stock!D4"])        # ['Stock!D2:D4']
merge_sources(["t:Stock/SKU-0001/Qty", "t:Stock/SKU-0002/Qty"])
# ['t:Stock/{SKU-0001,SKU-0002}/Qty']
```

In the tool layer the xlsx format registers this as its `t:` resolver and `merge_sources`
(ooxml-edit 0.22.0, O3): one session holding a workbook and a deck fills a slide table or a
chart from `t:Stock`, and the result names the cells each placed value came from. The
workbook is only read.

## Compare

```python
old = Workbook.open("inventory-generated.xlsx")
new = Workbook.open("inventory-reviewed.xlsx")
diff = Workbook.compare(old, new, tables=["t:Stock"], key="SKU")
# Diff(tables=(TableDiff(a='t:Stock', key='SKU', rows_a=8, rows_b=8, paired=7, unchanged=5,
#      added=(RowChange(change='added', key='SKU-0009', row='t:Stock/#8', ...),),
#      removed=(... 'SKU-0006' ...), changed=(CellChange('changed', 'SKU-0002', 'Price'),
#      CellChange('type', 'SKU-0005', 'Qty'): 9 became the text '9')), ...)
```

Sheets are matched by id (renamed, added, removed), tables by name, rows by key, cells by
type then value. A key two rows share is reported as `duplicate_key`, never paired.

## Validate

```python
report = wb.validate_data({"tables": [{"table": "t:Orders", "key": ["Order"], "columns": [
    {"name": "Customer", "type": "text", "required": True},
    {"name": "Status", "enum": ["Open", "Shipped", "Closed"]},
    {"name": "Qty", "type": "number", "min": "1"},
    {"name": "Discount", "type": "number", "min": "0", "max": "0.5"}]}]})
report.counts
# {'required_blank': 1, 'wrong_type': 1, 'out_of_range': 2, 'not_in_enum': 1, 'duplicate_key': 1}
report.problems[0]
# DataProblem(code='required_blank', address='Orders!B4', cells=1, table='t:Orders',
#             column='Customer', ...)
```

A schema names required `sheets` and `tables` (a `t:` address, or a region spec with
`range`, `name`, `header_rows`, `headers`), each with a `key` and `columns`: `type`,
`required`, `unique`, `min`, `max`, `step`, `enum` (with `case`), `validation_list` and
`pattern`. A schema that is not one is refused (`SchemaError`, with the path) before
anything is read. Problems carry addresses and codes, never a cell's text.
