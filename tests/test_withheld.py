"""A column hidden from the model by a declared table stays hidden whichever way its cells
are reached: an Excel table (ListObject) the declared table covers is replaced by it --
``describe`` lists one table, and ``t:<the Excel table>`` is refused toward the declared one
-- a defined name over the hidden cells is refused, and a read of the sheet's cells gives
them blank (:mod:`xlsx_agent.declare`, "Hidden from the model").  Every workbook is made
up here: a log of checks in an Excel table, its note column declared hidden."""

from __future__ import annotations

import dataclasses
import json

import pytest
from build import WorkbookBuilder
from ooxml_edit.tools import Toolbox

from xlsx_agent import RegionError, Workbook, table_dataset
from xlsx_agent.tools import GROUPS, TOOLS, format as xlsx_format

COLUMNS = ["Ref", "Item", "Owner", "Note"]
ROWS = [["R-1", "pump", "Kim", "private alpha"],
        ["R-2", "valve", "Lee", "private beta"],
        ["R-3", "hose", "Kim", "private gamma"]]

#: The declaration: the Excel table's cells, its note column hidden from the model.
CHECKS = {"name": "checks", "sheets": "Log", "key": "Ref",
          "fields": [{"header": "Note", "hidden": True}]}

#: Words only the hidden column holds: its header and its cells.
SECRETS = ("Note", "private", "alpha", "gamma")


def book(notes: list[str] | None = None) -> WorkbookBuilder:
    built = WorkbookBuilder()
    sheet = built.sheet("Log")
    rows = [row[:3] + [note] for row, note in zip(ROWS, notes)] if notes else ROWS
    sheet.rows("A1", [COLUMNS] + rows)
    sheet.table("ChecksTable", "A1:D4", COLUMNS, write_header=False)
    sheet.validation_list("C2:C4", source="$D$2:$D$4")   # a drop-down drawn from the notes
    built.sheet("Other").rows("A1", [["Code", "Count"], ["x", 1], ["y", 2]])
    built.define_name("NoteCells", "Log!$D$2:$D$4")
    built.define_name("WholeLog", "Log!$A$1:$D$4")
    built.define_name("Items", "Log!$B$2:$B$4")
    return built


def opened(notes: list[str] | None = None) -> Workbook:
    wb = Workbook.open(book(notes).to_bytes())
    wb.declare_tables([CHECKS])
    return wb


def clean(text: str) -> None:
    for secret in SECRETS:
        assert secret not in text, secret


def shown(thing) -> str:
    if dataclasses.is_dataclass(thing):
        thing = dataclasses.asdict(thing)
    return json.dumps(thing, ensure_ascii=False, default=str)


def refused(caught: pytest.ExceptionInfo) -> None:
    assert caught.value.code == "declared"
    assert caught.value.valid_options == ["t:checks"]
    clean(str(caught.value))
    assert "t:ChecksTable lies in t:checks" in str(caught.value)


# -- the catalogue -------------------------------------------------------------------------------


def test_the_declared_table_replaces_the_excel_table_it_covers():
    wb = opened()
    data = wb.catalogue(detail="full")
    addresses = [r["address"] for r in data["regions"]]
    assert addresses.count("t:checks") == 1 and "t:ChecksTable" not in addresses
    assert [t.name for t in wb.tables] == []
    clean(shown(data))
    clean(wb.describe())
    assert [n["name"] for n in data.get("names", [])] == []   # names over hidden cells: not listed
    assert "t:Items" in addresses
    assert "t:ChecksTable" not in wb.regions.known()


def test_without_a_declaration_the_excel_table_is_as_before():
    wb = Workbook.open(book().to_bytes())
    assert [t.name for t in wb.tables] == ["ChecksTable"]
    assert "private alpha" in shown(wb.catalogue(detail="full"))


def test_a_declared_table_over_an_excel_table_with_nothing_hidden_still_replaces_it():
    wb = Workbook.open(book().to_bytes())
    wb.declare_tables([{"name": "checks", "sheets": "Log", "key": "Ref"}])
    addresses = [r["address"] for r in wb.catalogue(detail="full")["regions"]]
    assert addresses.count("t:checks") == 1 and "t:ChecksTable" not in addresses
    with pytest.raises(RegionError) as caught:
        wb.region("t:ChecksTable")
    refused(caught)


# -- t:<the Excel table>, every way in -----------------------------------------------------------


@pytest.mark.parametrize("address", [
    "t:ChecksTable", "t:checkstable", "t:ChecksTable[Note]", "t:ChecksTable/R-1/Note",
    "t:ChecksTable/*/Note", "t:ChecksTable/{R-1,R-3}/Note", "t:ChecksTable/#2"])
def test_the_excel_tables_name_is_refused_toward_the_declared_table(address):
    wb = opened()
    for attempt in (wb.resolve, wb.expand, lambda a: list(wb.read(a)), wb.dataset):
        with pytest.raises(RegionError) as caught:
            attempt(address)
        refused(caught)


def test_profile_query_derive_compare_and_validate_refuse_the_excel_tables_name():
    wb = opened()
    attempts = [
        lambda: wb.profile("t:ChecksTable"),
        lambda: wb.column("t:ChecksTable[Note]"),
        lambda: wb.table("t:ChecksTable").query(["Note"]),
        lambda: wb.derive("t:ChecksTable", group_by=["Note"],
                          measures=[{"fn": "count", "as": "n"}]),
        lambda: Workbook.compare(wb, opened(["changed"] * 3), tables="t:ChecksTable"),
        lambda: table_dataset(wb, "t:ChecksTable[Note]"),
        lambda: list(wb.find_text("o", within="t:ChecksTable")),
    ]
    for attempt in attempts:
        with pytest.raises(RegionError) as caught:
            attempt()
        refused(caught)
    report = wb.validate_data({"tables": [{"table": "t:ChecksTable",
                                           "columns": [{"name": "Note", "type": "text"}]}]})
    clean(shown(report).replace("t:ChecksTable", ""))


def test_compare_pairs_the_declared_table_only_and_never_the_hidden_column():
    diff = Workbook.compare(opened(), opened(["changed one", "changed two", "changed three"]))
    assert [t.a for t in diff.tables] == ["t:checks"]
    text = shown(diff)
    clean(text)
    assert "ChecksTable" not in text and diff.same


# -- names, cells, find and problems -------------------------------------------------------------


@pytest.mark.parametrize("address", ["t:NoteCells", "t:WholeLog"])
def test_a_defined_name_over_hidden_cells_is_refused(address):
    wb = opened()
    with pytest.raises(RegionError) as caught:
        wb.region(address)
    assert (caught.value.code, caught.value.valid_options) == ("declared", ["t:checks"])
    clean(str(caught.value).replace(address, ""))     # the model's own words aside
    assert [v.v for row in wb.read("t:Items") for v in row] == ["pump", "valve", "hose"]


def test_a_read_of_the_sheets_cells_gives_the_hidden_ones_blank():
    wb = opened()
    rows = list(wb.read("Log!A1:D4"))
    assert [row[3].t for row in rows] == ["blank"] * 4
    assert [row[0].v for row in rows] == ["Ref", "R-1", "R-2", "R-3"]
    assert wb.cell("Log!D3").v is None
    assert [v.v for v in wb.cells("Log")][:4] == COLUMNS     # the file's cells as stored


def test_find_text_finds_neither_the_hidden_header_nor_its_cells():
    wb = opened()
    for text in SECRETS:
        assert list(wb.find_text(text)) == [], text
    assert [f.kind for f in wb.find_text("Owner")] == ["cell"]


def test_a_drop_down_drawn_from_hidden_cells_shows_none_of_them():
    wb = opened()
    owner = wb.profile("t:checks", ["Owner"])[0]
    assert [(v.values, v.reason, v.source) for v in owner.validation] == [
        (None, "withheld", "Log!D2:D4")]
    clean(shown(wb.profile("t:checks")))
    clean(shown(wb.catalogue(detail="full")))


def test_problems_name_no_hidden_cell():
    wb = opened()
    clean(shown(wb.problems))


def test_a_second_declaration_of_the_same_table_still_finds_its_hidden_header():
    wb = opened()
    again = wb.declare_tables([CHECKS])[0]
    assert again.columns == ("Ref", "Item", "Owner")


def test_another_declared_table_over_the_hidden_column_must_hide_it_too():
    from xlsx_agent import DeclarationError
    wb = opened()
    with pytest.raises(DeclarationError) as caught:
        wb.declare_tables([{"name": "notes", "sheets": "Log", "columns": "C:D"}])
    assert caught.value.code == "hidden_elsewhere"
    clean(str(caught.value))
    region = wb.declare_tables([{"name": "notes", "sheets": "Log", "columns": "C:D",
                                 "fields": [{"column": "D", "hidden": True}]}])[0]
    assert region.columns == ("Owner",)
    with pytest.raises(DeclarationError) as caught:
        Workbook.open(book().to_bytes()).declare_tables(
            [CHECKS, {"name": "notes", "sheets": "Log", "columns": "C:D"}])
    assert caught.value.code == "hidden_elsewhere"


# -- the tools -----------------------------------------------------------------------------------


def test_no_tool_shows_the_hidden_column_through_the_excel_table():
    toolbox = Toolbox(TOOLS, formats=[xlsx_format(tables=[CHECKS])], groups=GROUPS)
    session = toolbox.session()
    blob = session.add_blob(book().to_bytes(), name="log.xlsx")
    calls = [
        ("open_document", {"blob": blob}),
        ("describe", {"doc": "d1"}),
        ("xlsx_read", {"doc": "d1", "target": "t:ChecksTable"}),
        ("xlsx_read", {"doc": "d1", "target": "t:ChecksTable[Note]"}),
        ("xlsx_read", {"doc": "d1", "target": "t:NoteCells"}),
        ("xlsx_read", {"doc": "d1", "target": "Log!A1:D4"}),
        ("xlsx_query", {"doc": "d1", "table": "t:ChecksTable"}),
        ("xlsx_profile", {"doc": "d1", "target": "t:ChecksTable"}),
        ("xlsx_profile", {"doc": "d1", "target": "t:checks"}),
        ("xlsx_derive", {"doc": "d1", "from": "t:ChecksTable", "group_by": ["Note"],
                         "measures": [{"fn": "count", "as": "n"}]}),
        ("find_text", {"doc": "d1", "text": "Note"}),
        ("find_text", {"doc": "d1", "text": "private"}),
        ("find_text", {"doc": "d1", "text": "e", "range": "Log!A1:D4"}),
    ]
    texts = []
    for name, arguments in calls:
        result = toolbox.dispatch(session, name, arguments)
        text = json.dumps(result.to_json(), ensure_ascii=False, default=str)
        for given in arguments.values():   # what the model wrote it may be told again
            if isinstance(given, str):
                text = text.replace(given, "")
        texts.append(text)
    every = "\n".join(texts)
    clean(every.replace('"withheld"', ""))
    assert "t:ChecksTable" not in texts[1]
