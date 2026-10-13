"""xlsx-agent: workbooks as a source an agent can describe, query and derive from.

The agent never writes XML and never does arithmetic: it reads a catalogue of the workbook,
asks for rows, and asks the library for figures, each of which says which cells it came
from.  Built on lxml, ooxml-edit and ooxml-common, beside pptx-agent and docx-agent.

So far a workbook opens, lists its sheets and writes itself back unchanged, addresses
parse and format (:mod:`xlsx_agent.address`), and cells are read, typed and streamed
(:mod:`xlsx_agent.read`)::

    from xlsx_agent import Workbook, parse_address

    wb = Workbook.open("inventory.xlsx")      # bytes, a path, or an OpcPackage
    wb.sheets                                  # [Sheet(name='Stock', id=3, state='visible', ...)]
    address = parse_address("'Q3 sales'!B2:B9")
    address.text                               # "'Q3 sales'!B2:B9"
    wb.sheet_of("#3!F2")                       # the sheet whose sheetId is 3
    wb.cell("Stock!B2").as_dict()              # {'a': 'Stock!B2', 'v': 12.5, 't': 'number'}
    for row in wb.read("Stock!A:C"): ...       # lazily, clipped to the used range

A ``.xls``, ``.xlsb`` or encrypted file is refused with an :class:`OpenError` whose
``code`` says why.  See ROADMAP.md for the plan and what each milestone adds.
"""

from .address import (FORMS, Address, AddressError, DatasetAddress, NameNotFound,
                      RangeAddress, RefAddress, SheetRef, TableAddress, format_address,
                      needs_quotes, parse_address, quote_sheet)
from .package import REFUSALS, KindMismatchWarning, OpenError
from .read import Locale, NumberFormat, Problem, SheetScan, Value
from .workbook import Sheet, Workbook

__version__ = "0.0.1"

__all__ = ["FORMS", "REFUSALS", "Address", "AddressError", "DatasetAddress", "KindMismatchWarning",
           "Locale", "NameNotFound", "NumberFormat", "OpenError", "Problem", "RangeAddress",
           "RefAddress", "Sheet", "SheetRef", "SheetScan", "TableAddress", "Value", "Workbook",
           "__version__", "format_address", "needs_quotes", "parse_address", "quote_sheet"]
