"""Tables declared from the caller's configuration (X3) and key-value tables (X7): a table
over several sheets read as one, column ids and descriptions, columns hidden from the
model, the declarations' own checks, and the tool layer's way to take them.  Every
workbook is made up here: a sheet of checks per area, with a title over the header, a
rating, a note and a reviewer's remark; a summary sheet; a hidden sheet of run facts."""

from __future__ import annotations

import json
from datetime import date

import pytest
from build import Formula, WorkbookBuilder
from ooxml_edit.tools import Toolbox

from xlsx_agent import (DeclarationError, NameNotFound, RegionError, Workbook,
                        table_dataset)
from xlsx_agent.catalogue import estimate_tokens
from xlsx_agent.tools import GROUPS, TOOLS, declare, format as xlsx_format

HEADER = ["Ref", "Topic", "Rating (1 to 5, in halves)", "Note", "Remark"]

AREAS = {
    "Area North": [["N-1", "door", 3.5, "fine", "private one"],
                   ["N-2", "lock", 4, "good", "private two"]],
    "Area South": [["S-1", "door", 2, "worn", "private three"],
                   ["S-2", "gate", 5, "new", "private four"],
                   ["S-3", "fence", 1.5, "loose", "private five"]],
}

#: The declaration most tests use.
CHECKS = {"name": "checks", "sheets": "Area *", "header_row": 3, "key": "Ref",
          "description": "One row per check, every area",
          "fields": [{"header": "Rating (1 to 5, in halves)", "id": "rating",
                      "description": "The reviewer's rating"},
                     {"header": "Remark", "hidden": True}]}

RUN = {"name": "run", "type": "key_value", "sheets": "Run facts"}


def areas(extra: dict | None = None, *, header: list | None = None,
          south_header: list | None = None) -> WorkbookBuilder:
    book = WorkbookBuilder()
    for name, rows in {**AREAS, **(extra or {})}.items():
        sheet = book.sheet(name)
        sheet.cell("A1", f"Checks of {name}")
        top = south_header if (south_header is not None and name == "Area South") else header
        sheet.rows("A3", [top or HEADER] + rows)
    book.sheet("Summary").rows("A1", [["Area", "Count"], ["North", 2], ["South", 3]])
    facts = book.sheet("Run facts", state="hidden")
    facts.rows("A1", [["report_date", date(2026, 10, 1)], ["author", "Pat"], ["round", 3]])
    return book


def opened(book: WorkbookBuilder | None = None, *declarations) -> Workbook:
    wb = Workbook.open((book or areas()).to_bytes())
    if declarations:
        wb.declare_tables(list(declarations))
    return wb


@pytest.fixture(scope="module")
def wb() -> Workbook:
    return opened(None, CHECKS, RUN)


# -- one table over several sheets ------------------------------------------------------------


def test_rows_are_every_sheets_in_order_each_value_its_own_cell(wb):
    region = wb.region("t:checks")
    assert (region.kind, region.rows, region.columns) == (
        "declared", 5, ("Ref", "Topic", "rating", "Note"))
    assert region.where == "'Area North'!A3:D5,'Area South'!A3:D6"
    rows = list(wb.read("t:checks"))
    assert [row[0].v for row in rows] == ["N-1", "N-2", "S-1", "S-2", "S-3"]
    assert [row[0].a for row in rows] == ["'Area North'!A4", "'Area North'!A5",
                                          "'Area South'!A4", "'Area South'!A5",
                                          "'Area South'!A6"]
    assert rows[3][2].as_dict() == {"a": "'Area South'!C5", "v": 5, "t": "number"}


def test_rows_by_position_and_key_cross_the_sheets(wb):
    assert wb.cell("t:checks/#3/rating").a == "'Area South'!C4"
    assert wb.cell("t:checks/S-2/rating").as_dict() == {"a": "'Area South'!C5", "v": 5,
                                                        "t": "number"}
    assert wb.cell("t:checks/N-2/Note").v == "good"
    rows = list(wb.read("t:checks/{S-3,N-1}/rating"))
    assert [(r[0].a, r[0].v) for r in rows] == [("'Area South'!C6", 1.5),
                                                ("'Area North'!C4", 3.5)]
    assert wb.expand("t:checks/*/rating")[:3] == ["t:checks/N-1/rating", "t:checks/N-2/rating",
                                                   "t:checks/S-1/rating"]


def test_the_table_reads_as_any_other_query_profile_derive(wb):
    found = wb.table("t:checks").query(["Ref"], where={"rating": {"gte": 4}})
    assert [row[0].a for row in found.rows] == ["'Area North'!A5", "'Area South'!A5"]
    assert found.positions == (2, 4)
    stats = {s.column: s for s in wb.profile("t:checks")}
    assert stats["rating"].range == "'Area North'!C4:C5,'Area South'!C4:C6"
    assert (stats["rating"].min, stats["rating"].max) == (1.5, 5)
    assert stats["Ref"].types["text"].first[2] == "'Area South'!A4"
    ds = wb.derive("t:checks", group_by=["Topic"],
                   measures=[{"fn": "mean", "column": "rating", "as": "avg"}])
    assert ds.rows[0] == {"Topic": "door", "avg": 2.75}
    assert [str(s) for s in ds.provenance[0]["avg"].sources] == ["t:checks/{N-1,S-1}/rating"]


def test_without_a_key_sources_are_each_sheets_ranges():
    wb = opened(None, {**CHECKS, "key": None})
    ds = wb.derive("t:checks", measures=[{"fn": "sum", "column": "rating", "as": "total"}])
    assert ds.source_key is None
    assert [str(s) for s in ds.provenance[0]["total"].sources] == [
        "'Area North'!C4:C5", "'Area South'!C4:C6"]
    data = table_dataset(wb, "t:checks[rating]")
    assert [s[0] for s in data.sources] == [("'Area North'!C4",), ("'Area North'!C5",),
                                             ("'Area South'!C4",), ("'Area South'!C5",),
                                             ("'Area South'!C6",)]


def test_a_dataset_of_the_table_has_its_sources_by_key(wb):
    data = table_dataset(wb, "t:checks/{S-2,N-1}")
    assert data.columns == ("Ref", "Topic", "rating", "Note")
    assert data.rows[0][:3] == ("S-2", "gate", 5)
    assert data.sources[0][2] == ("t:checks/S-2/rating",)


def test_a_key_two_sheets_share_is_refused_naming_both_cells():
    book = areas({"Area West": [["N-2", "roof", 3, "ok", "private six"]]})
    with pytest.raises(DeclarationError) as caught:
        opened(book, CHECKS)
    error = caught.value
    assert error.code == "duplicate_key" and error.path == "tables[0].key"
    assert error.valid_options == ["'Area North'!A5", "'Area West'!A4"]
    assert "N-2" not in str(error)


def test_a_row_without_a_key_is_refused_naming_its_cell():
    book = areas({"Area West": [[None, "roof", 3, "ok", "x"], ["W-2", "roof", 3, "ok", "y"]]})
    with pytest.raises(DeclarationError) as caught:
        opened(book, CHECKS)
    assert caught.value.code == "no_key"
    assert caught.value.valid_options == ["'Area West'!A4"]


def test_a_header_that_differs_on_one_sheet_is_refused_naming_both_cells():
    other = list(HEADER)
    other[3] = "Notes"
    with pytest.raises(DeclarationError) as caught:
        opened(areas(south_header=other), CHECKS)
    assert caught.value.code == "header_mismatch"
    assert caught.value.valid_options == ["'Area South'!D3", "'Area North'!D3"]


def test_sheets_by_exact_name_list_and_glob():
    wb = opened(None, {**CHECKS, "sheets": ["area south", "Area North"]})
    assert wb.region("t:checks").where.startswith("'Area North'!")   # the workbook's order
    wb = opened(None, {**CHECKS, "sheets": "Area South"})
    region = wb.region("t:checks")
    assert region.span is not None and region.rows == 3              # a hidden column


def test_unknown_sheets_are_refused_and_optional_tables_skipped():
    with pytest.raises(DeclarationError) as caught:
        opened(None, {**CHECKS, "sheets": ["Area North", "Area Est"]})
    assert (caught.value.code, caught.value.path) == ("unknown_sheet", "tables[0].sheets[1]")
    assert "Area East" not in caught.value.valid_options  # nearest of the real ones
    with pytest.raises(DeclarationError) as caught:
        opened(None, {**CHECKS, "sheets": "Zone *"})
    assert caught.value.code == "no_sheets"
    wb = opened()
    assert wb.declare_tables([{**CHECKS, "sheets": "Zone *", "optional": True}, RUN]) \
        == [wb.region("t:run")]
    with pytest.raises(NameNotFound):
        wb.region("t:checks")


def test_declarations_are_kept_whole_or_not_at_all():
    wb = opened()
    with pytest.raises(DeclarationError):
        wb.declare_tables([RUN, {**CHECKS, "key": "Topic"}])   # door twice
    with pytest.raises(NameNotFound):
        wb.region("t:run")


# -- what the data rows are ---------------------------------------------------------------------


def test_rows_end_at_a_blank_row_and_more_below_is_flagged():
    built = WorkbookBuilder()
    sheet = built.sheet("Area North")
    sheet.rows("A1", [HEADER, ["N-1", "door", 3, "a", "b"], ["N-2", "lock", 4, "c", "d"]])
    sheet.rows("A6", [["Total", None, 7, None, None]])
    wb = Workbook.open(built.to_bytes())
    region = wb.declare_table("checks", "Area North", header_row=1)
    assert region.rows == 2
    assert [(f.code, f.address) for f in region.flags] == [("rows_below", "'Area North'!A6")]
    region = wb.declare_table("checks", "Area North", last_row=6)
    assert region.rows == 5 and not region.flags


def test_columns_and_first_row_are_declared():
    wb = opened(None, {"name": "notes", "sheets": "Area *", "header_row": 3, "columns": "B:D"})
    assert wb.region("t:notes").columns == ("Topic", "Rating (1 to 5, in halves)", "Note")
    wb = opened(None, {"name": "body", "sheets": "Area North", "header_row": 0,
                       "first_row": 4, "columns": "A:B"})
    region = wb.region("t:body")
    assert (region.columns, region.rows) == (("Column A", "Column B"), 2)


# -- ids, descriptions, hidden columns ----------------------------------------------------------


def test_an_id_and_its_header_both_name_the_column(wb):
    assert wb.cell("t:checks/N-1/rating").v == 3.5
    assert wb.cell("t:checks/N-1/Rating (1 to 5, in halves)").v == 3.5
    assert wb.resolve("t:checks[RATING]").column == "rating"
    stats = wb.profile("t:checks", ["rating"])[0]
    assert (stats.header, stats.description) == ("Rating (1 to 5, in halves)",
                                                 "The reviewer's rating")
    assert stats.as_dict()["header"] == "Rating (1 to 5, in halves)"


def test_describe_shows_ids_descriptions_and_no_hidden_column(wb):
    data = wb.catalogue()
    region = next(r for r in data["regions"] if r["address"] == "t:checks")
    assert region["about"] == "One row per check, every area"
    assert [c["name"] for c in region["columns"]] == ["Ref", "Topic", "rating", "Note"]
    assert region["columns"][2]["about"] == "The reviewer's rating"
    text = wb.describe()
    assert "Remark" not in text and "private" not in text and "'t:" not in text
    # the blocks of the declared table are not detected again
    assert not any(r["address"].endswith("!A3") for r in data["regions"])


def test_a_hidden_column_is_unknown_and_never_offered(wb):
    with pytest.raises(NameNotFound) as caught:
        wb.cell("t:checks/N-1/Remark")
    assert "Remark" not in caught.value.valid_options
    assert "Remark" not in str(caught.value).replace("'Remark'", "")
    with pytest.raises(NameNotFound):
        wb.table("t:checks").query(["Remark"])
    with pytest.raises(NameNotFound):
        wb.derive("t:checks", measures=[{"fn": "count", "column": "Remark", "as": "n"}])
    with pytest.raises(NameNotFound):
        wb.profile("t:checks", ["Remark"])


def test_hidden_cells_are_withheld_from_find_and_from_regions_over_them(wb):
    assert list(wb.find_text("private")) == []
    assert [f.address for f in wb.find_text("door", within="t:checks")] == [
        "'Area North'!B4", "'Area South'!B4"]
    with pytest.raises(RegionError) as caught:
        wb.detect_region("t:'Area North'!A3")
    assert caught.value.code == "declared" and caught.value.valid_options == ["t:checks"]
    with pytest.raises(RegionError) as caught:
        wb.declare_region("raw", "'Area North'!A3:E5")
    assert caught.value.code == "declared"


def test_a_column_excel_hides_stays_usable_unless_declared_hidden():
    built = WorkbookBuilder()
    sheet = built.sheet("Area North")
    sheet.rows("A1", [HEADER, ["N-1", "door", 3, "a", "b"]])
    sheet.hide_column("C")
    wb = Workbook.open(built.to_bytes())
    wb.declare_table("checks", "Area North", fields=[{"column": "C", "id": "rating"}])
    value = wb.cell("t:checks/#1/rating")
    assert (value.v, value.hidden) == (3, True)


def test_ids_that_would_name_two_columns_and_hidden_keys_are_refused():
    bad = [({**CHECKS, "fields": [{"header": "Note", "id": "Topic"}]}, "id_taken"),
           ({**CHECKS, "key": "Remark"}, "hidden_key"),
           ({**CHECKS, "key": "Score"}, "unknown_column"),
           ({**CHECKS, "fields": [{"header": "Notes"}]}, "unknown_column"),
           ({**CHECKS, "fields": [{"header": "Note"}, {"column": "D"}]}, "duplicate_field"),
           ({**CHECKS, "fields": [{"column": "Z"}]}, "unknown_column"),
           ({**CHECKS, "key": None, "fields": [{"column": c, "hidden": True}
                                                for c in "ABCDE"]}, "no_columns")]
    for declaration, code in bad:
        with pytest.raises(DeclarationError) as caught:
            opened(None, declaration)
        assert caught.value.code == code, declaration


def test_the_form_of_a_declaration_is_checked_with_its_path():
    bad = [({"sheets": "Area *"}, "bad_declaration", "tables[0].name"),
           ({**CHECKS, "name": "A1"}, "bad_name", "tables[0].name"),
           ({**CHECKS, "colour": "red"}, "bad_declaration", "tables[0].colour"),
           ({**CHECKS, "type": "pivot"}, "bad_declaration", "tables[0].type"),
           ({**CHECKS, "header_row": -1}, "bad_rows", "tables[0].header_row"),
           ({**CHECKS, "first_row": 2}, "bad_rows", "tables[0].first_row"),
           ({**CHECKS, "columns": "F:A"}, "bad_columns", "tables[0].columns"),
           ({**CHECKS, "fields": [{"id": "x"}]}, "bad_declaration", "tables[0].fields[0]"),
           ({**CHECKS, "fields": [{"header": "Note", "id": " x"}]}, "bad_id",
            "tables[0].fields[0].id"),
           ({**CHECKS, "description": "x" * 301}, "bad_declaration", "tables[0].description")]
    for declaration, code, path in bad:
        with pytest.raises(DeclarationError) as caught:
            opened(None, declaration)
        assert (caught.value.code, caught.value.path) == (code, path), declaration
    with pytest.raises(DeclarationError) as caught:
        opened(None, CHECKS, {**CHECKS, "name": "Checks"})
    assert (caught.value.code, caught.value.path) == ("duplicate_name", "tables[1].name")


def test_a_name_an_excel_table_has_is_refused():
    wb = Workbook.open(_with_table().to_bytes())
    with pytest.raises(DeclarationError) as caught:
        wb.declare_table("Products", "Stock")
    assert caught.value.code == "name_taken"


def _with_table() -> WorkbookBuilder:
    book = WorkbookBuilder()
    sheet = book.sheet("Stock")
    sheet.rows("A1", [["SKU", "Qty"], ["K-1", 1]])
    sheet.table("Products", "A1:B2", ["SKU", "Qty"], write_header=False)
    return book


# -- key-value tables ---------------------------------------------------------------------------


def test_a_declared_key_value_table_is_addressed_by_field(wb):
    region = wb.region("t:run")
    assert (region.shape, region.columns, region.key) == ("key_value", ("field", "value"),
                                                          "field")
    value = wb.cell("t:run/report_date/value")
    assert value.as_dict() == {"a": "'Run facts'!B1", "v": "2026-10-01", "t": "date",
                               "serial": 46296.0, "hidden": True}
    assert wb.cell("t:run/author/value").v == "Pat"


def test_describe_shows_a_key_value_table_compactly(wb):
    region = next(r for r in wb.catalogue()["regions"] if r["address"] == "t:run")
    assert region == {"address": "t:run", "kind": "declared", "range": "'Run facts'!A1:B3",
                      "rows": 3, "shape": "key_value", "columns": ["field", "value"],
                      "fields": {"report_date": "2026-10-01", "author": "Pat", "round": 3}}


def test_a_key_value_table_with_a_header_and_more_columns():
    book = WorkbookBuilder()
    book.sheet("Facts").rows("A1", [["Item", "Setting", "Confidence", "Source"],
                                    ["report_date", date(2026, 10, 1), "high", "form"],
                                    ["author", "Pat", "low", "mail"]])
    wb = Workbook.open(book.to_bytes())
    region = wb.declare_table("facts", "Facts", type="key_value", header_row=1,
                              fields=[{"header": "Confidence", "id": "confidence"}])
    assert region.columns == ("field", "value", "confidence", "Source")
    assert wb.cell("t:facts/author/confidence").v == "low"
    assert wb.cell("t:facts/author/Setting").v == "Pat"     # its header finds it too
    assert wb.cell("t:facts/author/value").a == "Facts!B3"


def test_a_two_column_block_of_facts_is_detected_as_key_value():
    wb = opened()
    region = wb.region("t:'Run facts'!A1")
    assert (region.kind, region.shape, region.header_rows, region.rows) == (
        "detected", "key_value", 0, 3)
    assert wb.cell("t:'Run facts'!A1/round/value").v == 3
    shown = next(r for r in wb.catalogue()["regions"] if r["address"] == "t:'Run facts'!A1")
    assert shown["fields"] == {"report_date": "2026-10-01", "author": "Pat", "round": 3}


def test_a_wider_block_with_numbers_in_its_first_row_is_still_unclear():
    book = WorkbookBuilder()
    book.sheet("Plan").rows("A1", [["Year", 2025, 2026], ["Bolt", 1, 2]])
    wb = Workbook.open(book.to_bytes())
    assert wb.catalogue()["unclear"] == [{"address": "t:Plan!A1", "code": "header_unclear",
                                          "reason": "header_not_text"}]


def test_hidden_sheets_are_read_and_flagged_hidden():
    wb = opened()
    facts = wb.sheet("Run facts")
    assert facts.state == "hidden"
    assert all(v.hidden for row in wb.read("'Run facts'!A1:B3") for v in row)
    sheets = {s["name"]: s for s in wb.catalogue()["sheets"]}
    assert sheets["Run facts"]["state"] == "hidden"


# -- comparing and checking ---------------------------------------------------------------------


def test_two_versions_compare_by_key_with_each_cells_address():
    old = opened(None, CHECKS)
    changed = {**AREAS, "Area South": [["S-1", "door", 2, "worn", "x"],
                                       ["S-2", "gate", 4.5, "new", "y"],
                                       ["S-3", "fence", 1.5, "loose", "z"]]}
    book = WorkbookBuilder()
    for name, rows in changed.items():
        sheet = book.sheet(name)
        sheet.rows("A3", [HEADER] + rows)
    new = Workbook.open(book.to_bytes())
    new.declare_tables([CHECKS])
    diff = Workbook.compare(old, new, tables=["t:checks"])
    table = diff.tables[0]
    assert table.range_a == "'Area North'!A3:D5,'Area South'!A3:D6"
    cells = [(c.a.a, c.b.a, c.a.v, c.b.v) for c in table.changed]
    assert cells == [("'Area South'!C5", "'Area South'!C5", 5, 4.5)]


def test_a_check_gives_each_sheets_cells_never_a_run_across_two():
    wb = opened(None, CHECKS)
    report = wb.validate_data({"tables": [{"table": "t:checks", "columns": [
        {"name": "rating", "type": "number", "min": 3}]}]})
    assert [p.address for p in report.problems] == ["'Area South'!C4", "'Area South'!C6"]
    report = wb.validate_data({"tables": [{"table": "t:checks", "columns": [
        {"name": "Rating (1 to 5, in halves)", "type": "text"}]}]})
    assert [p.address for p in report.problems] == ["'Area North'!C4:C5", "'Area South'!C4:C6"]


# -- the tool layer -----------------------------------------------------------------------------


def _session(toolbox: Toolbox, book: WorkbookBuilder | None = None):
    session = toolbox.session()
    blob = session.add_blob((book or areas()).to_bytes(), name="checks.xlsx")
    return session, blob


def _results(toolbox, session, calls):
    out = []
    for name, arguments in calls:
        result = toolbox.dispatch(session, name, arguments)
        assert result.ok, (name, arguments, result.error)
        out.append(json.dumps(result.data, ensure_ascii=False, default=str))
    return out


def test_the_format_declares_the_tables_and_no_tool_shows_the_hidden_column():
    toolbox = Toolbox(TOOLS, formats=[xlsx_format(tables={"tables": [CHECKS, RUN]})],
                      groups=GROUPS)
    session, blob = _session(toolbox)
    texts = _results(toolbox, session, [
        ("open_document", {"blob": blob}),
        ("describe", {"doc": "d1"}),
        ("xlsx_read", {"doc": "d1", "target": "t:checks"}),
        ("xlsx_read", {"doc": "d1", "target": "t:checks/S-2/rating"}),
        ("xlsx_read", {"doc": "d1", "target": "t:checks/{S-3,N-1}", "comments": True,
                       "merges": True}),
        ("xlsx_read", {"doc": "d1", "target": "'Area North'!A1:E5"}),
        ("xlsx_read", {"doc": "d1", "target": "'Area South'!"}),
        ("xlsx_read", {"doc": "d1", "target": "t:run/report_date/value"}),
        ("xlsx_query", {"doc": "d1", "table": "t:checks",
                        "where": [{"column": "rating", "op": "gte", "value": "3"}]}),
        ("xlsx_profile", {"doc": "d1", "target": "t:checks"}),
        ("xlsx_derive", {"doc": "d1", "from": "t:checks", "id": "by_topic",
                         "group_by": ["Topic"],
                         "measures": [{"fn": "mean", "column": "rating", "as": "avg"}]}),
        ("xlsx_read", {"target": "d:by_topic", "sources": True}),
        ("find_text", {"doc": "d1", "text": "private"}),
        ("find_text", {"doc": "d1", "text": "o", "range": "t:checks"}),
        ("xlsx_validate_data", {"doc": "d1", "schema": {"tables": [
            {"table": "t:checks", "key": ["Ref"],
             "columns": [{"name": "rating", "type": "number", "max": "4"}]}]}}),
    ])
    every = "\n".join(texts)
    assert "private" not in every and "Remark" not in every and "'t:" not in every
    read = json.loads(texts[2])
    assert read["range"] == "'Area North'!A4:D5,'Area South'!A4:D6" and "sheet" not in read
    assert read["rows"][2][0] == {"a": "'Area South'!A4", "v": "S-1", "t": "text"}
    raw = json.loads(texts[5])
    assert raw["rows"][3][4] == {"a": "'Area North'!E4", "withheld": True}
    assert json.loads(texts[12])["matches"] == []
    derived = json.loads(texts[10])
    assert derived["rows"][0][1]["sources"] == ["t:checks/{N-1,S-1}/rating"]
    checked = json.loads(texts[14])
    assert [p["at"] for p in checked["problems"]] == ["'Area South'!C5"]



def test_columns_pick_a_declared_tables_columns_on_every_sheet():
    toolbox = Toolbox(TOOLS, formats=[xlsx_format(tables=[CHECKS])], groups=GROUPS)
    session, blob = _session(toolbox)
    assert toolbox.dispatch(session, "open_document", {"blob": blob}).ok
    read = toolbox.dispatch(session, "xlsx_read", {"doc": "d1", "target": "t:checks/{S-3,N-1}",
                                                   "columns": ["rating", "Ref"]})
    assert read.ok, read.error
    assert [[c["a"] for c in row] for row in read.data["rows"]] == [
        ["'Area South'!C6", "'Area South'!A6"], ["'Area North'!C4", "'Area North'!A4"]]
    hidden = toolbox.dispatch(session, "xlsx_read", {"doc": "d1", "target": "t:checks",
                                                     "columns": ["Remark"]})
    assert hidden.error.code == "not_found" and "Remark" not in json.dumps(
        hidden.error.valid_options)


def test_a_session_declares_its_own_tables_in_place_of_the_formats():
    toolbox = Toolbox(TOOLS, formats=[xlsx_format(tables=[RUN])], groups=GROUPS)
    session, blob = _session(toolbox)
    declare(session, [CHECKS])
    assert toolbox.dispatch(session, "open_document", {"blob": blob}).ok
    assert toolbox.dispatch(session, "xlsx_read", {"doc": "d1", "target": "t:checks/#1"}).ok
    missing = toolbox.dispatch(session, "xlsx_read", {"doc": "d1",
                                                      "target": "t:run/author/value"})
    assert missing.error.code == "not_found"


def test_a_workbook_the_declarations_do_not_fit_is_refused_at_open():
    toolbox = Toolbox(TOOLS, formats=[xlsx_format(tables=[{**CHECKS, "sheets": "Zone *"}])],
                      groups=GROUPS)
    session, blob = _session(toolbox)
    result = toolbox.dispatch(session, "open_document", {"blob": blob})
    assert not result.ok and result.error.code == "refused"
    assert result.error.details == {"code": "no_sheets", "path": "tables[0].sheets"}
    with pytest.raises(DeclarationError):
        xlsx_format(tables=[{"name": "checks"}])        # checked for its form at once
    with pytest.raises(DeclarationError):
        declare(session, [{"name": "checks", "sheets": 3}])


def test_declaring_changes_no_tool_definition():
    plain = Toolbox(TOOLS, groups=GROUPS).definitions("anthropic", groups="core")
    declared = Toolbox(TOOLS, formats=[xlsx_format(tables=[CHECKS])],
                       groups=GROUPS).definitions("anthropic", groups="core")
    assert declared == plain


def test_describe_size_before_and_after_a_declaration():
    """The estimate this branch's notes report (ROADMAP.md, X3): a declaration replaces two
    detected blocks with one table, and adds its descriptions."""
    wb = opened()
    before = estimate_tokens(wb.catalogue())
    wb.declare_tables([CHECKS, RUN])
    after = estimate_tokens(wb.catalogue())
    assert after < before


def test_a_formula_in_a_declared_table_keeps_its_cell():
    book = WorkbookBuilder()
    for name in ("Area North", "Area South"):
        book.sheet(name).rows("A1", [["Ref", "Score"], [f"{name[5]}-1", Formula("1+1", cached=2)]])
    wb = Workbook.open(book.to_bytes())
    wb.declare_table("checks", "Area *", key="Ref")
    value = wb.cell("t:checks/S-1/Score")
    assert (value.a, value.v, value.formula) == ("'Area South'!B2", 2, True)


def test_problems_in_hidden_columns_are_not_listed_and_mixed_ones_skip_their_cells():
    built = WorkbookBuilder()
    sheet = built.sheet("Area North")
    sheet.rows("A1", [HEADER, ["N-1", "door", 3, "fine", "12"], ["N-2", "lock", 4, "7", "13"]])
    sheet.cell("G1", Formula("NOW()", cached=1.5))           # volatile, outside the table
    sheet.cell("E9", Formula("TODAY()", cached=2.5))         # volatile, under the hidden one
    wb = Workbook.open(built.to_bytes())
    before = {p.code: (p.address, p.count) for p in wb.problems}
    assert before["number_as_text"] == ("'Area North'!E2", 3)
    wb.declare_table("checks", "Area North", fields=[{"header": "Remark", "hidden": True}],
                     last_row=9)
    after = {p.code: (p.address, p.count) for p in wb.problems}
    assert after["number_as_text"] == ("'Area North'!D3", 1)
    assert after["volatile"] == ("'Area North'!G1", 1)
    shown = json.dumps(wb.catalogue(detail="full"))
    assert "'Area North'!E" not in shown and "Remark" not in shown


def test_a_tables_own_flags_never_name_a_hidden_cell():
    built = WorkbookBuilder()
    sheet = built.sheet("Area North")
    sheet.rows("A1", [HEADER[:4] + [None], ["N-1", "door", 3, "fine", "x"]])
    sheet.rows("A4", [[None, None, None, None, "below, hidden"], [None, "below", None, None, None]])
    wb = Workbook.open(built.to_bytes())
    region = wb.declare_table("checks", "Area North", columns="A:E",
                              fields=[{"column": "E", "hidden": True}])
    assert [(f.code, f.address) for f in region.flags] == [("rows_below", "'Area North'!B5")]
