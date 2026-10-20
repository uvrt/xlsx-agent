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
are derived, each with its cells (:mod:`xlsx_agent.derive`), and kept as datasets that
are read by address and derived from again (:mod:`xlsx_agent.dataset`); two workbooks
are compared by key (:mod:`xlsx_agent.compare`); a workbook is checked against a declared
schema (:mod:`xlsx_agent.schema`); the caller's own tables are declared from its
configuration (:mod:`xlsx_agent.declare`), their text cleaned by its rules on request
(:mod:`xlsx_agent.clean`); ``xlsx_read``, ``xlsx_profile``, ``xlsx_query``,
``xlsx_derive``, ``xlsx_compare`` and ``xlsx_validate_data`` read cells, columns and rows,
paged, keep figures as ``d:<id>``, compare workbooks and check them
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
    wb.declare_tables([{"name": "checks", "sheets": "Area *", "header_row": 3,
                        "key": "Ref"}])    # the caller's tables, over several sheets
    wb.table("t:Products").column("Price").stats()   # ColumnStats(type='number', ...)
    wb.table("t:Products").query(["SKU"], where={"Price": {"gte": 5}}, sort=["SKU"])
    wb.catalogue()                             # sheets, regions, columns, samples, flags
    wb.comments("Stock"), wb.hyperlinks("Stock"), wb.arrays("Stock")   # over the cells
    wb.find_text("Hardware", within="Stock!")  # matches by address and kind
    ds = wb.derive("t:Products", group_by=["Category"],
                   measures=[{"fn": "mean", "column": "Price", "as": "average"}], round=2,
                   id="by_category")
    ds.dataset()                               # ooxml_edit.tools.Dataset: rows, sources
    DatasetTable(ds).read("d:by_category/Hardware/average")   # [[DatasetValue(...)]]
    wb.derive(ds, measures=[{"fn": "max", "column": "average"}])   # cells carried through
    Workbook.compare(old, new, tables=["t:Products"], key="SKU")   # Diff, by key
    Workbook.compare(before, after, cells=True, policy=policy).cells   # by address, formulas
    wb.validations("Stock")                    # rules: standard/x14, list sources resolved
    wb.validate_data({"tables": [{"table": "t:Products", "key": ["SKU"], "columns": [
        {"name": "Price", "type": "number", "min": 0}]}]})   # problems by code and address
    wb.set_cells({"Stock!D10": 12, "Stock!H10": {"date": "2026-10-20"},
                  "Stock!J10": {"formula": "=D10*2"}})   # 0.2 (E3): typed, guarded, one step
    wb.undo(); wb.save("stock.xlsx")           # the bytes opened again; a checked save

A ``.xls``, ``.xlsb`` or encrypted file is refused with an :class:`OpenError` whose
``code`` says why.  See ROADMAP.md for the plan and what each milestone adds.
"""

from .address import (FORMS, Address, AddressError, DatasetAddress, NameNotFound,
                      RangeAddress, RefAddress, SheetRef, TableAddress, format_address,
                      is_name, needs_quotes, parse_address, quote_sheet)
from .compare import (CellChange, ColumnChange, Diff, DuplicateKey, RowChange,
                      SheetChange, TableDiff)
from .clean import CleanedWarning
from .declare import DeclarationError
from .dataset import (DatasetSelection, DatasetTable, DatasetValue, as_dataset,
                      merge_sources, table_dataset)
from .derive import Derived, DeriveError, Provenance, Total
from .edit import EditError, EditReport, SaveReport
from .edit.policy import PolicyError, WritePolicy
from .edit.cells import Written
from .find import Found
from .package import REFUSALS, KindMismatchWarning, OpenError
from .profile import AllowedList, Column, ColumnStats, Count, Distinct
from .query import QueryError, QueryResult
from .schema import CheckedTable, DataProblem, SchemaError, ValidationReport
from .redact import RedactError, Redaction
from .read import (Comment, DefinedName, Hyperlink, Locale, NumberFormat, Problem, Region,
                   RegionError, Selection, SheetScan, UnclearRegion, Value)
from .integrity import CellEdit, CellsDiff
from .validations import ValidationRule
from .workbook import Sheet, Workbook

__version__ = "0.1.0"

__all__ = ["FORMS", "REFUSALS", "Address", "AddressError", "AllowedList", "CellChange",
           "CellEdit", "CellsDiff", "ValidationRule",
           "CheckedTable", "CleanedWarning", "Column", "DataProblem", "SchemaError", "ValidationReport",
           "ColumnChange", "Diff", "DuplicateKey", "RowChange", "SheetChange", "TableDiff",
           "ColumnStats", "Comment", "Count", "DatasetAddress", "DatasetSelection",
           "DatasetTable", "DatasetValue", "DeclarationError", "DefinedName", "DeriveError",
           "Derived", "Distinct", "EditError", "EditReport", "PolicyError", "SaveReport",
           "WritePolicy", "Written",
           "Found", "Hyperlink", "KindMismatchWarning", "Locale", "NameNotFound", "NumberFormat",
           "OpenError", "Problem", "QueryError", "QueryResult", "RangeAddress", "RefAddress",
           "Provenance", "RedactError", "Redaction", "Region", "RegionError", "Selection", "Sheet", "SheetRef", "SheetScan",
           "TableAddress", "Total", "UnclearRegion", "Value", "Workbook", "__version__",
           "as_dataset", "format_address", "is_name", "merge_sources", "needs_quotes",
           "parse_address", "quote_sheet", "table_dataset"]
