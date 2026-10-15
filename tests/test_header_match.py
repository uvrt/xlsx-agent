"""Round 5, T16: a declared table's header row found by its text (``header_row: {"match":
[...], "rows": "8:20"}``), for sheets of several sections: the one row whose text cells
hold every text given (case and runs of white space aside, any order), within the rows
given or below the tables declared before it on the sheet; its data rows run to the first
blank row or the next row that matches.  None or several rows: refused, by row number
(:mod:`xlsx_agent.declare`).  Every workbook is made up here."""

from __future__ import annotations

import pytest
from build import WorkbookBuilder

from xlsx_agent import Workbook
from xlsx_agent.declare import DeclarationError, parse_declarations


def sections() -> Workbook:
    built = WorkbookBuilder()
    built.sheet("Sections").rows("A1", [
        ["Run sheet"],                              # 1: a title
        [],
        ["Ref", "Item", "Rating"],                  # 3
        ["A-1", "pump", 4],
        ["A-2", "valve", 3],
        ["A-3", "hose", 5],
        [],
        [],
        ["  ref ", "ITEM", "Owner"],                # 9: the same texts, cased and spaced
        ["B-1", "gauge", "Kim"],
        ["B-2", "filter", "Lee"],
    ])
    return Workbook.open(built.to_bytes())


def stacked() -> Workbook:
    """Two sections with no blank row between them."""
    built = WorkbookBuilder()
    built.sheet("Stacked").rows("A1", [
        ["Ref", "Item"], ["C-1", "bolt"], ["C-2", "nut"],
        ["Ref", "Item"], ["D-1", "washer"], ["D-2", "screw"], ["D-3", "rivet"],
    ])
    return Workbook.open(built.to_bytes())


def rows_of(wb: Workbook, address: str) -> list[list]:
    return [[v.v for v in row] for row in wb.read(address)]


def test_the_header_is_found_by_some_of_its_texts():
    wb = sections()
    region, = wb.declare_tables([{"name": "ratings", "sheets": "Sections",
                                  "header_row": {"match": ["Rating", "Ref"]}, "key": "Ref"}])
    assert region.where == "Sections!A3:C6"
    assert region.columns == ("Ref", "Item", "Rating")
    assert wb.cell("t:ratings/A-2/Rating").v == 3


def test_case_and_spaces_aside():
    wb = sections()
    region, = wb.declare_tables([{"name": "owners", "sheets": "Sections",
                                  "header_row": {"match": [" REF", "owner "]}}])
    assert region.where == "Sections!A9:C11"
    assert rows_of(wb, "t:owners") == [["B-1", "gauge", "Kim"], ["B-2", "filter", "Lee"]]


def test_several_rows_are_refused_by_row_number_only():
    wb = sections()
    with pytest.raises(DeclarationError) as caught:
        wb.declare_tables([{"name": "both", "sheets": "Sections",
                            "header_row": {"match": ["Ref", "Item"]}}])
    error = caught.value
    assert (error.code, error.path, error.valid_options) == \
        ("ambiguous_header", "tables[0].header_row", ["3", "9"])
    assert "3, 9" in str(error) and "header_row.rows" in str(error)
    for content in ("pump", "Rating", "Owner", "Run sheet"):
        assert content not in str(error)


def test_no_row_is_refused():
    wb = sections()
    with pytest.raises(DeclarationError) as caught:
        wb.declare_tables([{"name": "none", "sheets": "Sections",
                            "header_row": {"match": ["Ref", "Due date"]}}])
    assert (caught.value.code, caught.value.path) == ("no_header", "tables[0].header_row")
    with pytest.raises(DeclarationError) as caught:   # a text is a cell's whole text
        wb.declare_tables([{"name": "none", "sheets": "Sections",
                            "header_row": {"match": ["Re"]}}])
    assert caught.value.code == "no_header"


def test_below_the_tables_declared_before_it_on_the_sheet():
    wb = sections()
    first, second = wb.declare_tables([
        {"name": "ratings", "sheets": "Sections", "header_row": {"match": ["Ref", "Rating"]}},
        {"name": "owners", "sheets": "Sections", "header_row": {"match": ["Ref", "Item"]}}])
    assert (first.where, second.where) == ("Sections!A3:C6", "Sections!A9:C11")


def test_within_the_rows_given():
    wb = sections()
    region, = wb.declare_tables([{"name": "owners", "sheets": "Sections",
                                  "header_row": {"match": ["Ref", "Item"], "rows": "8:20"}}])
    assert region.where == "Sections!A9:C11"
    with pytest.raises(DeclarationError) as caught:
        wb.declare_tables([{"name": "nothing", "sheets": "Sections",
                            "header_row": {"match": ["Ref"], "rows": "4:8"}}])
    assert caught.value.code == "no_header" and "rows 4:8" in str(caught.value)


def test_the_rows_stop_at_the_next_row_that_matches():
    wb = stacked()
    first, second = wb.declare_tables([
        {"name": "first", "sheets": "Stacked", "header_row": {"match": ["Ref", "Item"],
                                                              "rows": "1:1"}},
        {"name": "second", "sheets": "Stacked", "header_row": {"match": ["Ref", "Item"]}}])
    assert first.where == "Stacked!A1:B3" and first.flags == ()   # no rows_below
    assert second.where == "Stacked!A4:B7"
    assert rows_of(wb, "t:first") == [["C-1", "bolt"], ["C-2", "nut"]]
    assert rows_of(wb, "t:second")[-1] == ["D-3", "rivet"]


def test_each_sheet_finds_its_own_header_row():
    built = WorkbookBuilder()
    built.sheet("Area 1").rows("A1", [["Area one"], ["Ref", "Item"], ["E-1", "pump"]])
    built.sheet("Area 2").rows("A1", [["Area two"], [], [], ["Ref", "Item"], ["F-1", "valve"],
                                      ["F-2", "hose"]])
    wb = Workbook.open(built.to_bytes())
    region, = wb.declare_tables([{"name": "areas", "sheets": "Area *", "key": "Ref",
                                  "header_row": {"match": ["Item", "Ref"]}}])
    assert [(p.sheet, p.header, p.first, p.last) for p in region.span.parts] == [
        ("Area 1", 2, 3, 3), ("Area 2", 4, 5, 6)]
    assert wb.cell("t:areas/F-2/Item").v == "hose"
    assert wb.cell("t:areas/F-2/Item").a == "'Area 2'!B6"


def test_a_key_value_table_by_its_header():
    built = WorkbookBuilder()
    built.sheet("Run").rows("A1", [["Notes"], [], ["Field", "Value"], ["author", "Kim"],
                                   ["pages", 12]])
    wb = Workbook.open(built.to_bytes())
    wb.declare_tables([{"name": "run", "type": "key_value", "sheets": "Run",
                        "header_row": {"match": ["field", "value"]}}])
    assert wb.cell("t:run/pages/value").v == 12


@pytest.mark.parametrize("header_row, path, code", [
    ({"match": []}, "tables[0].header_row.match", "bad_declaration"),
    ({"match": ["Ref", " "]}, "tables[0].header_row.match", "bad_declaration"),
    ({"match": [3]}, "tables[0].header_row.match", "bad_declaration"),
    ({}, "tables[0].header_row.match", "bad_declaration"),
    ({"match": ["Ref"], "rows": "9:3"}, "tables[0].header_row.rows", "bad_rows"),
    ({"match": ["Ref"], "rows": 5}, "tables[0].header_row.rows", "bad_rows"),
    ({"match": ["Ref"], "row": "1:5"}, "tables[0].header_row.row", "bad_declaration"),
])
def test_the_form_is_checked(header_row, path, code):
    with pytest.raises(DeclarationError) as caught:
        parse_declarations([{"name": "t", "sheets": "S", "header_row": header_row}])
    assert (caught.value.code, caught.value.path) == (code, path)


@pytest.mark.parametrize("key", ["first_row", "last_row"])
def test_a_header_found_by_its_text_takes_no_fixed_rows(key):
    with pytest.raises(DeclarationError) as caught:
        parse_declarations([{"name": "t", "sheets": "S", key: 9,
                             "header_row": {"match": ["Ref"]}}])
    assert (caught.value.code, caught.value.path) == ("bad_rows", f"tables[0].{key}")


def test_a_single_text_is_taken_as_a_list():
    spec, = parse_declarations([{"name": "t", "sheets": "S", "header_row": {"match": "Ref"}}])
    assert spec.header_match == ("Ref",) and spec.header_row == 1
