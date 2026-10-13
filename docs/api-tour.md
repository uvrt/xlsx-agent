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
`d:by_region`.

Merged or multi-row headers are declared, never guessed:

```python
wb.declare_region("Returns", "Returns!A3:F200", header_rows=2,
                  headers=["SKU", "Date", "Qty", "Reason", "Store", "Refund"])
```

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
