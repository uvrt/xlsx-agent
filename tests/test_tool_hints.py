"""Round 9 (ooxml-edit 0.47.0, N5): a name that is no workbook tool names the tool that does
what it says, with the format's own table (``TOOL_HINTS``) for what ooxml-edit's reading
of names misses.  Made-up calls."""

from __future__ import annotations

import pytest

from ooxml_edit.tools import Toolbox
from ooxml_edit.tools.hints import tool_hints

from xlsx_agent.tools import FORMAT, GROUPS, TOOLS
from xlsx_agent.tools.docformat import TOOL_HINTS

from build import WorkbookBuilder


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _book() -> bytes:
    book = WorkbookBuilder()
    sheet = book.sheet("Stock")
    sheet.rows("A2", [["A-1", 1], ["A-2", 2]])
    sheet.table("Lines", "A1:B3", ["SKU", "Qty"])
    return book.to_bytes()


@pytest.mark.parametrize("name, meant", [
    # read from the tools' names
    ("xlsx_read_range", "xlsx_read"),
    ("xlsx_get_cells", "xlsx_read"),
    ("xlsx_query_table", "xlsx_query"),
    ("xlsx_compare_workbooks", "xlsx_compare"),
    ("xlsx_describe", "describe"),
    # the format's own table
    ("xlsx_list_sheets", "describe"),           # not xlsx_read
    ("xlsx_filter", "xlsx_query"),
    ("xlsx_pivot", "xlsx_derive"),
    ("xlsx_aggregate", "xlsx_derive"),
    ("xlsx_diff", "xlsx_compare"),
    ("xlsx_validate", "xlsx_validate_data"),
])
def test_what_a_name_means(box, name, meant):
    assert [str(hint) for hint in tool_hints(name, box.tools, box.tool_hints)] == [meant]


def test_every_entry_names_a_tool_the_format_has(box):
    for name, target in TOOL_HINTS.items():
        tool, _, action = target.partition(" ")
        assert tool in box.tools, name
        assert not action, name
        assert name not in box.tools


def test_an_unknown_tool_names_the_one_to_use(box):
    session = box.session()
    session.open(_book(), "stock.xlsx")
    result = box.dispatch(session, "xlsx_pivot", {"doc": "d1"})
    assert result.error.message == "no tool 'xlsx_pivot': use xlsx_derive"
    assert result.error.valid_options[0] == "xlsx_derive"
    assert len(result.error.valid_options) <= 5
