"""Round 18, Q3: one canonical form for every address the library returns, and
``canonical_address`` to turn any accepted spelling into it (xlsx_agent.address, "The
canonical form").

The workbooks are built with tests/build.py; every sheet name and value is made up.
"""

import pytest
from ooxml_edit.sheets import Area, Cell
from ooxml_edit.tools import Toolbox

from build import WorkbookBuilder
from test_address import CANONICAL, SHEET_NAMES
from xlsx_agent import (AddressError, RangeAddress, SheetRef, Workbook, canonical_address,
                        parse_address)
from xlsx_agent.address import format_cell, format_range, format_sheet
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS

#: (sheet name, how an address writes it).
NAMES = [
    ("Sales", "Sales"),
    ("Sales_2", "Sales_2"),
    ("data.2026", "data.2026"),
    ("v1.2", "'v1.2'"),
    ("Q3 Sales", "'Q3 Sales'"),
    ("Bob's", "'Bob''s'"),
    ("2024", "'2024'"),
    ("A1", "'A1'"),
    ("xfd9", "'xfd9'"),
    ("R1C1", "'R1C1'"),
    ("rc", "'rc'"),
    ("TRUE", "'TRUE'"),
    ("Überblick", "Überblick"),
    ("数据", "数据"),
    ("XFE1", "XFE1"),
    ("R1x", "'R1x'"),
    ("a©b", "'a©b'"),
    ("a-b", "'a-b'"),
]


@pytest.mark.parametrize("name, written", NAMES)
def test_a_sheet_is_quoted_only_when_the_rule_says(name, written):
    assert format_sheet(name) == written
    assert canonical_address(f"{written}!B4") == f"{written}!B4"
    assert canonical_address(f"{written}!$b$4:a1") == f"{written}!A1:B4"


@pytest.mark.parametrize("text, canonical", [
    ("Sales!$B$4", "Sales!B4"),
    ("sales!b4", "sales!B4"),                 # the sheet's spelling kept without the workbook
    ("'Sales'!B4", "Sales!B4"),               # quotes it needs not dropped
    ("r1x!a1", "'r1x'!A1"),                   # quotes it needs added
    ("'Überblick'!a1", "Überblick!A1"),
    ("A1!B2", "'A1'!B2"),
    ("'Bob''s'!A01", "'Bob''s'!A1"),          # an apostrophe doubled; the row's zeros dropped
    ("Sales!B4:B4", "Sales!B4"),              # a range of one cell is the cell
    ("Sales!C9:A1", "Sales!A1:C9"),           # corner to corner, top left first
    ("Sales!$c:a", "Sales!A:C"),
    ("Sales!$5:2", "Sales!2:5"),
    ("'Q3 Sales'!", "'Q3 Sales'!"),
    ("  #3!f2 ", "#3!F2"),
    ("t:sales!$b$2", "t:sales!B2"),
    ("t:'Sales'!Rate", "t:Sales!Rate"),
    ("t:Items/'K-1'/'Price'", "t:Items/K-1/Price"),
    ("d:by_group/#2:#2", "d:by_group/#2"),
    ("$step1", "$step1"),
])
def test_any_accepted_spelling_becomes_the_canonical_form(text, canonical):
    assert canonical_address(text) == canonical
    assert canonical_address(canonical) == canonical


@pytest.mark.parametrize("text", [text for text, _, _ in CANONICAL])
def test_a_canonical_address_is_its_own_canonical_form(text):
    assert canonical_address(text) == text


@pytest.mark.parametrize("name", SHEET_NAMES)
def test_the_helper_round_trips_any_sheet_name(name):
    text = RangeAddress(SheetRef(name=name), Area.parse("C3:D4")).text
    assert canonical_address(text) == text
    assert parse_address(canonical_address(text)).sheet.name == name


def test_the_workbooks_names_give_the_workbooks_spelling():
    sheets = ["Sales", "Q3 Sales", "Überblick"]
    assert canonical_address("sales!b4", sheets=sheets) == "Sales!B4"
    assert canonical_address("'q3 sales'!B4", sheets=sheets) == "'Q3 Sales'!B4"
    assert canonical_address("ÜBERBLICK!A1", sheets=sheets) == "Überblick!A1"
    assert canonical_address("t:sales!b2", sheets=sheets) == "t:Sales!B2"
    assert canonical_address("t:'q3 sales'!Rate", sheets=sheets) == "t:'Q3 Sales'!Rate"
    assert canonical_address("Other!A1", sheets=sheets) == "Other!A1"   # not there: as given
    assert canonical_address("#3!A1", sheets=sheets) == "#3!A1"
    assert canonical_address("t:Items[Price]", sheets=sheets) == "t:Items[Price]"


def test_what_is_not_an_address_is_refused():
    for text in ("", "B4", "Sales", "'Sales!B4", "#03!A1"):
        with pytest.raises(AddressError):
            canonical_address(text)


def test_the_formatters_write_what_the_address_classes_write():
    ref = SheetRef(name="Q3 Sales")
    for area, span in ((Area.parse("B4"), "cells"), (Area.parse("A1:C9"), "cells")):
        assert format_range(ref, area, span) == RangeAddress(ref, area, span).text
    columns = parse_address("'Q3 Sales'!A:C")
    rows = parse_address("'Q3 Sales'!2:5")
    assert isinstance(columns, RangeAddress) and isinstance(rows, RangeAddress)
    assert format_range(ref, columns.area, "columns") == "'Q3 Sales'!A:C"
    assert format_range(ref, rows.area, "rows") == "'Q3 Sales'!2:5"
    assert format_range("'Q3 Sales'") == "'Q3 Sales'!"
    assert format_range(ref, Cell(4, 2)) == "'Q3 Sales'!B4"
    assert format_cell(ref, 4, 2) == format_cell("'Q3 Sales'", 4, 2) == "'Q3 Sales'!B4"
    assert format_cell("Sales", 1, 16384) == "Sales!XFD1"
    assert format_sheet(None, 3) == "#3"
    with pytest.raises(ValueError):
        format_sheet(None)


# -- what the library returns ----------------------------------------------------------------


def _book() -> WorkbookBuilder:
    book = WorkbookBuilder()
    for n, (name, _) in enumerate(NAMES):
        sheet = book.sheet(name)
        sheet.rows("A1", [["Item", "Count"], [f"word {n}", n], ["other", 2]])
        sheet.comment("B2", "word in a note")
    return book


def _returned(address: str) -> None:
    assert canonical_address(address) == address


def test_every_address_a_read_returns_is_canonical():
    wb = Workbook.open(_book().to_bytes())
    names = [s.name for s in wb.sheets]
    for name, written in NAMES:
        for text in (f"{written}!a1:b3", f"{written}!", f"{written}!A:B", f"{written}!$1:$3"):
            values = [v for row in wb.read(text) for v in row]
            assert values
            for value in values:
                _returned(value.a)
                assert canonical_address(value.a, sheets=names) == value.a
                assert value.a.startswith(f"{written}!")
        assert wb.cell(f"{written}!$b$2").a == f"{written}!B2"


def test_every_address_a_search_returns_is_canonical():
    wb = Workbook.open(_book().to_bytes())
    found = list(wb.find_text("word"))
    assert {f.kind for f in found} >= {"cell", "comment"}
    assert len([f for f in found if f.kind == "cell"]) == len(NAMES)
    for match in found:
        for address in (match.address, match.cell, match.range):
            if address is not None:
                _returned(address)


def test_every_address_a_compare_returns_is_canonical():
    before = Workbook.open(_book().to_bytes())
    book = _book()
    after = Workbook.open(book.to_bytes())
    cells = {f"{written}!B3": 7 for _, written in NAMES}
    report = after.set_cells(cells)
    assert report
    diff = Workbook.compare(before, after, cells=True).cells
    assert diff is not None and len(diff.edits) == len(NAMES)
    for edit in diff.edits:
        _returned(edit.address)
    assert {edit.address for edit in diff.edits} == set(cells)
    for sheet in diff.sheets:
        _returned(f"{sheet}!")


def test_every_address_the_read_tool_returns_is_canonical():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        session = toolbox.session()
        blob = session.add_blob(_book().to_bytes(), name="book.xlsx")
        assert toolbox.dispatch(session, "open_document", {"blob": blob}).ok
        for _, written in NAMES:
            result = toolbox.dispatch(session, "xlsx_read",
                                      {"doc": "d1", "target": f"{written}!$a$1:b3"})
            assert result.ok, result
            _returned(result.data["target"])
            cells = [cell for row in result.data["rows"] for cell in row]
            assert cells
            for cell in cells:
                _returned(cell["a"])
