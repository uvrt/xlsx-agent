"""Workbook tables as sources (X1, ooxml-edit's O3): :func:`~xlsx_agent.dataset.table_dataset`
(``Workbook.dataset``), the xlsx format's ``t:`` resolver and its ``merge_sources``.  The
deck side -- a slide table and a chart filled from ``t:checks`` -- is in
``test_sources_deck.py``.  Every workbook is made up here: a table of checks keyed by code,
a table of visits with no key, and a table of pairs with two columns that could be its key.
"""

from __future__ import annotations

from datetime import date

import pytest
from build import WorkbookBuilder
from ooxml_edit.tools import Dataset, Limits, ToolError

from xlsx_agent import AddressError, NameNotFound, RegionError, Workbook, merge_sources
from xlsx_agent.tools import FORMAT
from xlsx_agent.tools.docformat import configure, resolve_table

CHECKS = [["code", "area", "rating", "due", "note"],
          ["A1", "north", 3.5, date(2026, 1, 5), "door"],
          ["A2", "north", 4, date(2026, 2, 10), None],
          ["B1", "south", 2.5, date(2026, 3, 15), "roof"]]

#: No key: every column has a repeat, a fraction or a blank.  A rating stored as text.
VISITS = [["area", "rating", "note"],
          ["north", 2.5, "gate"],
          ["north", "4.5", "gate"],
          ["south", 3.5, None]]

#: Two columns that could be the key, so neither is until one is declared.
PAIRS = [["code", "seq", "rating"],
         ["P1", 1, 1.5],
         ["P2", 2, 2.5]]


def book() -> bytes:
    built = WorkbookBuilder()
    log = built.sheet("Log")
    log.rows("A1", CHECKS)
    log.table("checks", "A1:E4", CHECKS[0], write_header=False)
    plan = built.sheet("Visit plan")
    plan.rows("B2", VISITS)
    plan.table("visits", "B2:D5", VISITS[0], write_header=False)
    plan.rows("G2", PAIRS)
    plan.table("pairs", "G2:I4", PAIRS[0], write_header=False)
    return built.to_bytes()


@pytest.fixture
def wb() -> Workbook:
    return Workbook.open(book())


# -- the rows ----------------------------------------------------------------------------------


def test_a_keyed_table_typed_with_its_cells_by_key(wb):
    data = wb.dataset("t:checks")
    assert isinstance(data, Dataset)
    assert data.columns == ("code", "area", "rating", "due", "note")
    assert data.rows == (("A1", "north", 3.5, "2026-01-05", "door"),
                         ("A2", "north", 4, "2026-02-10", None),
                         ("B1", "south", 2.5, "2026-03-15", "roof"))
    assert data.sources[0][2] == ("t:checks/A1/rating",)
    assert data.sources[1][4] == ("t:checks/A2/note",)       # a blank has its cell too
    assert data.keys == 1                                    # the key column comes first
    assert data.types == ("text", "text", "number", "date", "text")


def test_without_a_key_the_sources_are_cell_addresses(wb):
    data = wb.dataset("t:visits")
    assert data.sources[0] == (("'Visit plan'!B3",), ("'Visit plan'!C3",),
                               ("'Visit plan'!D3",))
    assert data.keys == 0


def test_a_number_stored_as_text_stays_text(wb):
    data = wb.dataset("t:visits[rating]")
    assert data.rows == ((2.5,), ("4.5",), (3.5,))
    assert data.types == (None,)                             # numbers and text: no one type


def test_a_declared_key_writes_the_sources_by_it(wb):
    assert wb.dataset("t:pairs").sources[0][2] == ("'Visit plan'!I3",)
    wb.declare_key("t:pairs", "code")
    data = wb.dataset("t:pairs")
    assert data.sources[0][2] == ("t:pairs/P1/rating",) and data.keys == 1
    wb.declare_key("t:pairs", "seq")
    data = wb.dataset("t:pairs")
    assert data.sources[1][0] == ("t:pairs/2/code",) and data.keys == 0


@pytest.mark.parametrize("address, columns, codes", [
    ("t:checks/*", 5, ["A1", "A2", "B1"]),
    ("t:checks/*/*", 5, ["A1", "A2", "B1"]),
    ("t:checks/{B1,A1}", 5, ["B1", "A1"]),                   # a key set in its order
    ("t:checks/{B1,A1}/*", 5, ["B1", "A1"]),
    ("t:checks/A2", 5, ["A2"]),
    ("t:checks/#3", 5, ["B1"]),
    ("t:checks/#2:#3", 5, ["A2", "B1"]),                     # a row range
    ("t:checks/#2:", 5, ["A2", "B1"]),
    ("t:checks/#1:#2/*", 5, ["A1", "A2"]),
])
def test_rows_by_selector(wb, address, columns, codes):
    data = wb.dataset(address)
    assert len(data.columns) == columns
    assert [row[0] for row in data.rows] == codes


@pytest.mark.parametrize("address", ["t:checks[rating]", "t:checks/*/rating",
                                     "t:checks/{B1,A1}/rating"])
def test_one_column(wb, address):
    data = wb.dataset(address)
    assert data.columns == ("rating",)
    assert [s[0][0].split("/")[1] for s in data.sources] == \
        (["B1", "A1"] if "{" in address else ["A1", "A2", "B1"])


def test_one_value_is_not_rows_and_other_addresses_are_refused(wb):
    with pytest.raises(RegionError) as error:
        wb.dataset("t:checks/A1/rating")
    assert error.value.code == "not_rows"
    assert error.value.valid_options == ["t:checks", "t:checks[rating]"]
    with pytest.raises(AddressError):
        wb.dataset("Log!A1:E4")
    with pytest.raises(NameNotFound) as missing:
        wb.dataset("t:chekcs")
    assert "t:checks" in missing.value.valid_options


def test_a_row_range_keeps_each_values_source(wb):
    data = wb.dataset("t:checks/#2:#3/rating")
    assert data.rows == ((4,), (2.5,))
    assert data.sources == ((("t:checks/A2/rating",),), (("t:checks/B1/rating",),))
    assert wb.dataset("t:visits/#3:").sources[0][0] == ("'Visit plan'!B5",)


@pytest.mark.parametrize("address, option", [
    ("t:checks/#2:#4", "t:checks/#2:#3"),
    ("t:checks/#3:#9", "t:checks/#3"),
    ("t:checks/#4:", "t:checks/#3"),
])
def test_a_row_range_past_the_last_row_is_refused_not_cut(wb, address, option):
    with pytest.raises(RegionError) as error:
        wb.dataset(address)
    assert error.value.code == "position_out_of_range"
    assert "has 3 data rows" in error.value.message and "nothing is cut" in error.value.message
    assert error.value.valid_options == [option]


def test_too_many_values_are_refused_before_reading(wb):
    with pytest.raises(RegionError) as error:
        wb.dataset("t:checks", max_cells=14)
    assert error.value.code == "too_many_cells" and "15 values" in error.value.message
    assert len(wb.dataset("t:checks", max_cells=15).rows) == 3


# -- merging -----------------------------------------------------------------------------------


def test_merge_sources_joins_ranges_and_key_sets():
    assert merge_sources(["Log!C2", "Log!C3", "Log!C4", "Log!C7"]) == ["Log!C2:C4", "Log!C7"]
    assert merge_sources(["t:checks/A1/rating", "t:checks/B1/rating",
                          "t:checks/A1/rating"]) == ["t:checks/{A1,B1}/rating"]
    assert merge_sources(["t:checks/A1/rating"]) == ["t:checks/A1/rating"]
    assert merge_sources(["'Visit plan'!C3", "'Visit plan'!C4"]) == ["'Visit plan'!C3:C4"]
    assert merge_sources(["d:x/#1/a"]) == ["d:x/#1/a"]          # anything else as it is
    assert merge_sources([]) == []


# -- the format's hook -------------------------------------------------------------------------


def test_the_format_registers_the_resolver_and_the_merger():
    assert set(FORMAT.datasets) == {"t:"}
    assert FORMAT.merge_sources is merge_sources


class _Session:
    locale = None

    def __init__(self, cells: int) -> None:
        self.limits = Limits(max_dataset_cells=cells)


def test_the_resolver_gives_tool_errors_and_the_sessions_limit(wb):
    assert resolve_table(wb, "t:checks").rows[0][0] == "A1"
    with pytest.raises(ToolError) as error:
        resolve_table(wb, "t:chekcs")
    assert error.value.code == "not_found" and "t:checks" in error.value.valid_options
    with pytest.raises(ToolError) as error:
        resolve_table(wb, "t:checks/A1/rating")
    assert error.value.code == "invalid_arguments"
    configure(wb, _Session(10))
    with pytest.raises(ToolError) as error:
        resolve_table(wb, "t:checks")
    assert error.value.code == "limit"
    assert resolve_table(wb, "t:checks[rating]").columns == ("rating",)


def test_reading_rows_changes_nothing(wb):
    before = wb.to_bytes()
    wb.dataset("t:checks")
    wb.dataset("t:visits")
    assert wb.to_bytes() == before
