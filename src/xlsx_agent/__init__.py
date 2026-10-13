"""xlsx-agent: workbooks as a source an agent can describe, query and derive from.

The agent never writes XML and never does arithmetic: it reads a catalogue of the workbook,
asks for rows, and asks the library for figures, each of which says which cells it came
from.  Built on lxml, ooxml-edit and ooxml-common, beside pptx-agent and docx-agent.

So far a workbook opens, lists its sheets and writes itself back unchanged, addresses
parse and format (:mod:`xlsx_agent.address`), cells are read, typed and streamed
(:mod:`xlsx_agent.read`), tables, defined names and regions resolve ``t:`` addresses
(:mod:`xlsx_agent.read.regions`), their columns are profiled
(:mod:`xlsx_agent.profile`), and the workbook is catalogued in one bounded result
(:mod:`xlsx_agent.catalogue`), which the shared ``describe`` tool returns; a region's rows
are queried (:mod:`xlsx_agent.query`); text is found (:mod:`xlsx_agent.find`); figures
are derived, each with its cells (:mod:`xlsx_agent.derive`); ``xlsx_read``,
``xlsx_profile`` and ``xlsx_query`` read cells, columns and rows, paged
(:mod:`xlsx_agent.tools`)::

    from xlsx_agent import Workbook, parse_address

    wb = Workbook.open("inventory.xlsx")      # bytes, a path, or an OpcPackage
    wb.sheets                                  # [Sheet(name='Stock', id=3, state='visible', ...)]
    address = parse_address("'Q3 sales'!B2:B9")
    address.text                               # "'Q3 sales'!B2:B9"
    wb.sheet_of("#3!F2")                       # the sheet whose sheetId is 3
    wb.cell("Stock!B2").as_dict()              # {'a': 'Stock!B2', 'v': 12.5, 't': 'number'}
    for row in wb.read("Stock!A:C"): ...       # lazily, clipped to the used range
    wb.resolve("t:Products/SKU-0042/Price")    # Selection(range=Stock!C43, ...)
    wb.declare_region("Returns", "Returns!A1:C40", header_rows=2, headers=[...])
    wb.table("t:Products").column("Price").stats()   # ColumnStats(type='number', ...)
    wb.table("t:Products").query(["SKU"], where={"Price": {"gte": 5}}, sort=["SKU"])
    wb.catalogue()                             # sheets, regions, columns, samples, flags
    wb.comments("Stock"), wb.hyperlinks("Stock"), wb.arrays("Stock")   # over the cells
    wb.find_text("Hardware", within="Stock!")  # matches by address and kind
    wb.derive("t:Products", group_by=["Category"],
              measures=[{"fn": "mean", "column": "Price", "as": "average"}], round=2)

A ``.xls``, ``.xlsb`` or encrypted file is refused with an :class:`OpenError` whose
``code`` says why.  See ROADMAP.md for the plan and what each milestone adds.
"""

from .address import (FORMS, Address, AddressError, DatasetAddress, NameNotFound,
                      RangeAddress, RefAddress, SheetRef, TableAddress, format_address,
                      is_name, needs_quotes, parse_address, quote_sheet)
from .derive import Derived, DeriveError, Provenance, Total
from .find import Found
from .package import REFUSALS, KindMismatchWarning, OpenError
from .profile import AllowedList, Column, ColumnStats, Count, Distinct
from .query import QueryError, QueryResult
from .read import (Comment, DefinedName, Hyperlink, Locale, NumberFormat, Problem, Region,
                   RegionError, Selection, SheetScan, UnclearRegion, Value)
from .workbook import Sheet, Workbook

__version__ = "0.0.1"

__all__ = ["FORMS", "REFUSALS", "Address", "AddressError", "AllowedList", "Column",
           "ColumnStats", "Comment", "Count", "DatasetAddress", "DefinedName", "DeriveError",
           "Derived", "Distinct",
           "Found", "Hyperlink", "KindMismatchWarning", "Locale", "NameNotFound", "NumberFormat",
           "OpenError", "Problem", "QueryError", "QueryResult", "RangeAddress", "RefAddress",
           "Provenance", "Region", "RegionError", "Selection", "Sheet", "SheetRef", "SheetScan",
           "TableAddress", "Total", "UnclearRegion", "Value", "Workbook", "__version__",
           "format_address", "is_name", "needs_quotes", "parse_address", "quote_sheet"]
