"""R13: ``wb.validate_data`` and ``xlsx_validate_data``: a workbook checked against a declared
schema.  Every problem code has a fixture; no problem, summary or error holds cell text (a
marker in the cells); then through a real ooxml-edit Toolbox, bounded and paged.

Every workbook is built with tests/build.py.
"""

import datetime as dt
import json
from datetime import date

import pytest

from build import Formula, WorkbookBuilder, cell_name
from ooxml_edit.tools import Limits, Toolbox
from test_catalogue import MARKER

from xlsx_agent import Workbook
from xlsx_agent import schema as schema_module
from xlsx_agent.schema import CODES, SchemaError
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS

CLOCK = lambda: dt.datetime(2026, 10, 13, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731

COLUMNS = ("SKU", "Name", "Price", "Score", "Category", "Status")


def _book(rows, columns=COLUMNS, *, table="Products", sheet="Stock", styles=None,
          lists=None) -> WorkbookBuilder:
    book = WorkbookBuilder()
    stock = book.sheet(sheet)
    stock.rows("A1", [list(columns)])
    for r, values in enumerate(rows, start=2):
        for c, value in enumerate(values):
            if value is not None:
                stock.cell(cell_name(r, c + 1), value,
                           style=book.style(**styles[c]) if styles and styles.get(c) else 0)
    last = cell_name(max(len(rows), 1) + 1, len(columns))
    for sqref, values in (lists or {}).items():
        stock.validation_list(sqref, values)
    if table:
        stock.table(table, f"A1:{last}", list(columns), write_header=False)
    return book


def _open(book: WorkbookBuilder) -> Workbook:
    return Workbook.open(book.to_bytes())


def _check(rows, columns_schema, *, key=None, **book) -> list[dict]:
    table = {"table": "t:Products", "columns": columns_schema}
    if key is not None:
        table["key"] = key
    report = _open(_book(rows, **book)).validate_data({"tables": [table]})
    return [p.as_dict() for p in report]


GOOD = [["SKU-1", "Bolt", 1.5, 4.5, "Hardware", "Open"],
        ["SKU-2", "Nut", 0.25, 3, "Garden", "Shut"]]


# -- every problem code has a fixture ----------------------------------------------------------


def _fixture(code: str):
    """A workbook and a schema that give ``code``."""
    rows = [list(r) for r in GOOD]
    table: dict = {"table": "t:Products", "columns": [{"name": "SKU", "type": "text"}]}
    schema = {"tables": [table]}
    lists = None
    if code == "missing_sheet":
        schema["sheets"] = ["Stock", "Returns"]
    elif code == "missing_table":
        table["table"] = "t:Orders"
    elif code == "missing_column":
        table["columns"] = [{"name": "Weight", "type": "number"}]
    elif code == "no_validation_list":
        table["columns"] = [{"name": "Status", "validation_list": True}]
    elif code == "no_cached_value":
        rows[0][2] = Formula("C3*2")
        table["columns"] = [{"name": "Price", "type": "number"}]
    elif code == "required_blank":
        rows[1][1] = None
        table["columns"] = [{"name": "Name", "required": True}]
    elif code == "wrong_type":
        rows[1][2] = "0.25"
        table["columns"] = [{"name": "Price", "type": "number"}]
    elif code == "out_of_range":
        rows[0][3] = 6
        table["columns"] = [{"name": "Score", "type": "number", "min": 1, "max": 5}]
    elif code == "off_step":
        rows[0][3] = 4.25
        table["columns"] = [{"name": "Score", "type": "number", "min": 1, "step": 0.5}]
    elif code == "not_in_enum":
        table["columns"] = [{"name": "Category", "enum": ["Hardware"]}]
    elif code == "not_in_list":
        lists = {"F2:F3": ["Open", "Closed"]}
        table["columns"] = [{"name": "Status", "validation_list": True}]
    elif code == "pattern_mismatch":
        rows[1][0] = "SKU2"
        table["columns"] = [{"name": "SKU", "pattern": r"SKU-\d+"}]
    elif code == "duplicate_key":
        rows[1][0] = "SKU-1"
        table["key"] = ["SKU"]
    elif code == "duplicate_value":
        rows[1][1] = "Bolt"
        table["columns"] = [{"name": "Name", "unique": True}]
    return _book(rows, lists=lists), schema


@pytest.mark.parametrize("code", CODES)
def test_every_problem_code_has_a_fixture(code):
    book, schema = _fixture(code)
    report = _open(book).validate_data(schema)
    assert list(report.counts) == [code], report.counts
    (problem,) = report.problems
    assert problem.code == code and problem.message


def test_the_good_rows_have_no_problems():
    report = _open(_book(GOOD)).validate_data({"sheets": ["stock"], "tables": [{
        "table": "t:Products", "key": "SKU", "columns": [
            {"name": "SKU", "type": "text", "required": True, "pattern": r"SKU-\d+"},
            {"name": "price", "type": "number", "min": 0, "max": "10"},
            {"name": "Score", "type": "number", "min": 1, "max": 5, "step": 0.5},
            {"name": "Category", "enum": ["Hardware", "Garden"]},
            {"name": "Name", "unique": True}]}]})
    assert report.ok and report.counts == {} and len(report) == 0
    assert [t.as_dict() for t in report.tables] == [
        {"table": "t:Products", "range": "Stock!A1:F3", "rows": 2, "columns": 5}]


def test_problems_say_what_was_expected():
    rows = [["SKU-1", "Bolt", "12", 6, "hardware", "Open"],
            ["SKU-1", "Nut", -1, 2.2, " Garden", "Shut"]]
    problems = _check(rows, [{"name": "Price", "type": "number", "min": "0"},
                             {"name": "Score", "type": "number", "min": 1, "max": 5,
                              "step": 0.5},
                             {"name": "Category", "enum": ["Hardware", "Garden"]}],
                      key=["SKU"])
    assert problems == [
        {"code": "wrong_type", "at": "Stock!C2", "table": "t:Products", "column": "Price",
         "expected": "number", "found": "text", "issue": "number_as_text",
         "message": "Stock!C2 (Price): not number (found text, number_as_text)"},
        {"code": "out_of_range", "at": "Stock!D2", "table": "t:Products", "column": "Score",
         "min": 1, "max": 5, "message": "Stock!D2 (Score): outside 1 to 5"},
        {"code": "not_in_enum", "at": "Stock!E2", "table": "t:Products", "column": "Category",
         "enum_size": 2, "case": "exact", "hint": "case_differs",
         "message": "Stock!E2 (Category): not one of the 2 allowed values (case counts); "
                    "one matches but for case"},
        {"code": "duplicate_key", "at": "Stock!A3", "table": "t:Products", "column": "SKU",
         "first": "Stock!A2",
         "message": "Stock!A3 (SKU): the key repeats that of the row at Stock!A2"},
        {"code": "out_of_range", "at": "Stock!C3", "table": "t:Products", "column": "Price",
         "min": "0", "message": "Stock!C3 (Price): less than 0"},
        {"code": "off_step", "at": "Stock!D3", "table": "t:Products", "column": "Score",
         "step": 0.5, "base": 1,
         "message": "Stock!D3 (Score): not 1 plus a whole number of steps of 0.5"},
        {"code": "not_in_enum", "at": "Stock!E3", "table": "t:Products", "column": "Category",
         "enum_size": 2, "case": "exact", "hint": "space_differs",
         "message": "Stock!E3 (Category): not one of the 2 allowed values (case counts); "
                    "one matches but for space at the ends"}]


def test_cells_next_to_each_other_become_one_problem():
    rows = [[f"S{r}", "x", "1" if r in (2, 3, 4, 7) else 1, 1, "a", "b"] for r in range(1, 9)]
    problems = _check(rows, [{"name": "Price", "type": "number"}])
    assert [(p["at"], p.get("cells", 1)) for p in problems] == [("Stock!C3:C5", 3),
                                                                 ("Stock!C8", 1)]
    assert problems[0]["message"].startswith("Stock!C3:C5 (Price), 3 cells: not number")
    report = _open(_book(rows)).validate_data(
        {"tables": [{"table": "t:Products", "columns": [{"name": "Price", "type": "number"}]}]})
    assert report.counts == {"wrong_type": 4}


def test_enum_case_and_types():
    rows = [["A", "x", 2, 1, "HARDWARE", "b"], ["B", "x", 3, 1, "Garden", "b"]]
    assert _check(rows, [{"name": "Category", "enum": ["hardware", "garden"],
                          "case": "ignore"},
                         {"name": "Price", "enum": ["2", "3.0"]}]) == []
    problems = _check(rows, [{"name": "Price", "type": "number", "enum": ["2"]}])
    assert [(p["at"], p["enum_size"]) for p in problems] == [("Stock!C3", 1)]


def test_dates_and_percents():
    styles = {2: {"num_fmt": "0%"}, 3: {"num_fmt": "yyyy-mm-dd"}}
    rows = [["A", "x", 0.25, date(2026, 1, 5), "a", "b"],
            ["B", "x", 0.33, date(2025, 12, 31), "a", "b"]]
    problems = _check(rows, [{"name": "Price", "type": "percent", "min": "0", "max": "1",
                              "step": 0.05},
                             {"name": "Score", "type": "date", "min": "2026-01-01"}],
                      styles=styles)
    assert [(p["code"], p["at"]) for p in problems] == [("off_step", "Stock!C3"),
                                                         ("out_of_range", "Stock!D3")]


def test_unique_values_compare_as_keys_do():
    rows = [["A", 42, 1, 1, "a", "b"], ["B", "42", 1, 1, "a", "b"], ["C", 4.5, 1, 1, "a", "b"],
            ["D", 4.5, 1, 1, "a", "b"]]
    problems = _check(rows, [{"name": "Name", "unique": True}])
    assert [(p["code"], p["at"], p["first"]) for p in problems] == [
        ("duplicate_value", "Stock!B3", "Stock!B2"), ("duplicate_value", "Stock!B5", "Stock!B4")]


def test_a_key_of_two_columns():
    rows = [["A", "x", 1, 1, "a", "b"], ["A", "y", 1, 1, "a", "b"], ["A", "x", 1, 1, "a", "b"],
            [None, "x", 1, 1, "a", "b"]]
    report = _open(_book(rows)).validate_data(
        {"tables": [{"table": "t:Products", "key": ["SKU", "Name"]}]})
    problems = [p.as_dict() for p in report]
    assert [(p["code"], p["at"]) for p in problems] == [("duplicate_key", "Stock!A4"),
                                                         ("required_blank", "Stock!A5")]
    assert problems[0]["key"] == ["SKU", "Name"] and problems[0]["first"] == "Stock!A2"


def test_validation_lists_inline_and_from_a_range():
    book = _book([["A", "x", 1, 1, "a", "Open"], ["B", "x", 1, 1, "a", "open"]],
                 lists={"F2:F3": ["Open", "Closed"]})
    book.sheets[0].rows("H1", [["Hardware"], ["Garden"]])
    book.sheets[0].validation_list("E2:E3", source="$H$1:$H$2")
    report = _open(book).validate_data({"tables": [{"table": "t:Products", "columns": [
        {"name": "Status", "validation_list": True},
        {"name": "Category", "validation_list": True, "case": "ignore"}]}]})
    problems = [p.as_dict() for p in report]
    assert [(p["code"], p["at"], p["list_size"], p.get("hint")) for p in problems] == [
        ("not_in_list", "Stock!E2:E3", 2, None), ("not_in_list", "Stock!F3", 2, "case_differs")]


def test_a_list_from_a_formula_is_not_readable():
    book = _book(GOOD)
    book.define_name("Choices", "OFFSET(Stock!$H$1,0,0,2,1)")
    book.sheets[0].validation("F2:F3", "list", "Choices")
    wb = _open(book)
    report = wb.validate_data({"tables": [{"table": "t:Products", "columns": [
        {"name": "Status", "validation_list": True}]}]})
    (problem,) = report.problems
    (allowed,) = wb.column("t:Products[Status]").stats().validation    # the profile's reading
    assert allowed.values is None and allowed.reason == "formula"
    assert problem.code == "no_validation_list" and problem.detail["reason"] == "formula"
    assert problem.address == "t:Products" and problem.column == "Status"


def test_a_declared_region_spec_is_checked_and_not_kept():
    book = WorkbookBuilder()
    sheet = book.sheet("Returns")
    sheet.rows("A1", [["Return", None], ["SKU", "Qty"], ["S1", 2], ["S2", "two"]])
    sheet.merge("A1:B1")
    wb = _open(book)
    report = wb.validate_data({"tables": [{"range": "Returns!A1:B4", "name": "Returns",
                                           "header_rows": 2, "headers": ["SKU", "Qty"],
                                           "columns": [{"name": "Qty", "type": "number"}]}]})
    assert [(p.code, p.address, p.table) for p in report] == [
        ("wrong_type", "Returns!B4", "t:Returns")]
    assert report.tables[0].as_dict() == {"table": "t:Returns", "range": "Returns!A1:B4",
                                          "rows": 2, "columns": 1}
    assert wb.regions.declared == []


@pytest.mark.parametrize("table,reason", [
    ({"table": "t:Prices"}, "not_a_range"),
    ({"table": "t:Orders"}, "not_found"),
    ({"range": "Old!A1:B4", "name": "Old"}, "sheet_not_found"),
    ({"range": "Stock!M:N", "name": "Far"}, "no_cells"),
], ids=["formula-name", "unknown", "no-sheet", "no-cells"])
def test_a_missing_table_says_why(table, reason):
    book = _book(GOOD)
    book.define_name("Prices", "SUM(Stock!$C$2:$C$3)")
    report = _open(book).validate_data({"tables": [{**table, "key": ["SKU"]}]})
    (problem,) = report.problems
    assert problem.code == "missing_table" and problem.detail["reason"] == reason


def test_a_missing_column_offers_the_nearest():
    (problem,) = _check(GOOD, [{"name": "Prise", "type": "number"}])
    assert problem["nearest"] == ["Price"]
    assert problem["message"] == "t:Products: no column Prise; nearest: Price"


def test_past_the_most_problems_cells_are_counted(monkeypatch):
    monkeypatch.setattr(schema_module, "MAX_PROBLEMS", 3)
    rows = [[f"S{r}", "x", "1" if r % 2 else 1, 1, "a", "b"] for r in range(20)]
    report = _open(_book(rows)).validate_data(
        {"tables": [{"table": "t:Products", "columns": [{"name": "Price", "type": "number"}]}]})
    assert len(report) == 3 and report.counts == {"wrong_type": 10} and report.cut == 7


# -- schema errors -----------------------------------------------------------------------------


@pytest.mark.parametrize("schema,path", [
    ([], ""),
    ({}, ""),
    ({"tabels": []}, ""),
    ({"sheets": "Stock", "tables": 3}, "tables"),
    ({"tables": [{"columns": [{"name": "A", "type": "text"}]}]}, "tables[0]"),
    ({"tables": [{"table": "t:Products", "range": "Stock!A1:B2"}]}, "tables[0]"),
    ({"tables": [{"table": "Stock!A1", "key": ["SKU"]}]}, "tables[0].table"),
    ({"tables": [{"table": "t:Products[SKU]", "key": ["SKU"]}]}, "tables[0].table"),
    ({"tables": [{"table": "t:Products", "name": "X", "key": ["SKU"]}]}, "tables[0].name"),
    ({"tables": [{"range": "Stock!A1:B2", "key": ["SKU"]}]}, "tables[0].name"),
    ({"tables": [{"range": "Stock!A1:B2", "name": "Products", "key": ["SKU"]}]}, "tables[0]"),
    ({"tables": [{"table": "t:Products"}]}, "tables[0]"),
    ({"tables": [{"table": "t:Products", "key": ["SKU", "sku"]}]}, "tables[0].key"),
    ({"tables": [{"table": "t:Products", "columns": [{"name": "A"}]}]},
     "tables[0].columns[0]"),
    ({"tables": [{"table": "t:Products", "columns": [{"name": "A", "type": "integer"}]}]},
     "tables[0].columns[0].type"),
    ({"tables": [{"table": "t:Products", "columns": [{"name": "A", "min": "1"}]}]},
     "tables[0].columns[0].min"),
    ({"tables": [{"table": "t:Products", "columns": [{"name": "A", "type": "number",
                                                     "min": "five"}]}]},
     "tables[0].columns[0].min"),
    ({"tables": [{"table": "t:Products", "columns": [{"name": "A", "type": "number",
                                                     "min": 5, "max": 1}]}]},
     "tables[0].columns[0].max"),
    ({"tables": [{"table": "t:Products", "columns": [{"name": "A", "type": "date",
                                                     "step": 1}]}]},
     "tables[0].columns[0].step"),
    ({"tables": [{"table": "t:Products", "columns": [{"name": "A", "type": "number",
                                                     "step": 0}]}]},
     "tables[0].columns[0].step"),
    ({"tables": [{"table": "t:Products", "columns": [{"name": "A", "type": "number",
                                                     "enum": [f"x {MARKER}"]}]}]},
     "tables[0].columns[0].enum[0]"),
    ({"tables": [{"table": "t:Products", "columns": [{"name": "A", "enum": ["x"],
                                                     "validation_list": True}]}]},
     "tables[0].columns[0].enum"),
    ({"tables": [{"table": "t:Products", "columns": [{"name": "A", "type": "text",
                                                     "case": "ignore"}]}]},
     "tables[0].columns[0].case"),
    ({"tables": [{"table": "t:Products", "columns": [{"name": "A", "pattern": f"({MARKER}"}]}]},
     "tables[0].columns[0].pattern"),
    ({"tables": [{"table": "t:Products", "columns": [{"name": "A", "type": "number",
                                                     "pattern": "x"}]}]},
     "tables[0].columns[0].pattern"),
    ({"tables": [{"table": "t:Products", "columns": [{"name": "A", "type": "text"},
                                                    {"name": "a", "type": "text"}]}]},
     "tables[0].columns"),
    ({"tables": [{"table": "t:Products", "columns": [{"name": "A", "required": "yes"}]}]},
     "tables[0].columns[0].required"),
], ids=lambda value: None)
def test_a_schema_that_is_not_one_is_refused_saying_where(schema, path):
    wb = _open(_book(GOOD))
    with pytest.raises(SchemaError) as caught:
        wb.validate_data(schema)
    assert caught.value.path == path and caught.value.code == "bad_schema"
    assert MARKER not in str(caught.value)


def test_a_schema_error_names_the_valid_options():
    with pytest.raises(SchemaError) as caught:
        _open(_book(GOOD)).validate_data({"tables": [{"table": "t:Products", "columns": [
            {"name": "A", "type": "integer"}]}]})
    assert caught.value.valid_options == ["number", "percent", "date", "time", "datetime",
                                          "text", "bool", "error"]


# -- no cell text in any problem ---------------------------------------------------------------


def _marked_book() -> WorkbookBuilder:
    """Every cell that gives a problem holds the marker; so do the keys."""
    rows = [[f"K1 {MARKER}", f"{MARKER}", f"1 {MARKER}", 9, f"{MARKER} x", f"{MARKER}"],
            [f"K1 {MARKER}", f"{MARKER}", 2, 2.2, f"Hardware {MARKER}", f" {MARKER} "],
            [None, f"{MARKER}", Formula(f'"{MARKER}"'), 3, f"{MARKER.lower()}", "Open"]]
    return _book(rows, lists={"F2:F4": ["Open", MARKER.lower()]})


MARKED_SCHEMA = {"sheets": ["Stock", "Nope"], "tables": [
    {"table": "t:Products", "key": ["SKU"], "columns": [
        {"name": "Name", "unique": True, "pattern": "[a-z]+"},
        {"name": "Price", "type": "number", "required": True},
        {"name": "Score", "type": "number", "min": "1", "max": "5", "step": 0.5},
        {"name": "Category", "enum": [MARKER, "Hardware"]},
        {"name": "Status", "validation_list": True},
        {"name": "Weight", "type": "number"}]},
    {"table": "t:Orders", "key": ["Id"]}]}


def test_no_cell_text_in_any_problem():
    report = _open(_marked_book()).validate_data(MARKED_SCHEMA)
    codes = set(report.counts)
    assert codes >= {"missing_sheet", "missing_table", "missing_column", "no_cached_value",
                     "required_blank", "wrong_type", "out_of_range", "off_step", "not_in_enum",
                     "not_in_list", "pattern_mismatch", "duplicate_key", "duplicate_value"}
    text = json.dumps([p.as_dict() for p in report]) + repr(report) + \
        " ".join(p.message for p in report)
    assert MARKER not in text and MARKER.lower() not in text


# -- through a real Toolbox --------------------------------------------------------------------


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _session(toolbox, *books):
    session = toolbox.session(clock=CLOCK)
    for index, book in enumerate(books):
        data = book if isinstance(book, bytes) else book.to_bytes()
        blob = session.add_blob(data, name=f"book{index}.xlsx")
        assert toolbox.dispatch(session, "open_document", {"blob": blob}).ok
    return session


def _validate(toolbox, session, schema, **arguments):
    return toolbox.dispatch(session, "xlsx_validate_data", {"doc": "d1", "schema": schema,
                                                            **arguments})


def test_xlsx_validate_data_end_to_end(box):
    rows = [["SKU-1", "Bolt", "12", 6, "hardware", "Open"],
            ["SKU-1", "Nut", -1, 2.5, "Garden", "Shut"]]
    session = _session(box, _book(rows, lists={"F2:F3": ["Open", "Closed"]}))
    # strict decoding sends every optional field, null when unset
    column = dict.fromkeys(("type", "required", "unique", "min", "max", "step", "enum", "case",
                            "validation_list", "pattern"))
    table = dict.fromkeys(("range", "name", "header_rows", "headers"))
    schema = {"sheets": ["Stock"], "tables": [{**table, "table": "t:Products", "key": ["SKU"],
                                               "columns": [
        {**column, "name": "Price", "type": "number", "min": "0"},
        {**column, "name": "Score", "type": "number", "min": "1", "max": "5", "step": 0.5},
        {**column, "name": "Category", "enum": ["Hardware", "Garden"]},
        {**column, "name": "Status", "validation_list": True}]}]}
    result = _validate(box, session, schema)
    assert result.ok, result.error
    data = result.data
    assert data["ok"] is False
    assert data["counts"] == {"wrong_type": 1, "out_of_range": 2, "not_in_enum": 1,
                              "not_in_list": 1, "duplicate_key": 1}
    assert data["tables"] == [{"table": "t:Products", "range": "Stock!A1:F3", "rows": 2,
                               "columns": 5}]
    assert [(p["code"], p["at"]) for p in data["problems"]] == [
        ("wrong_type", "Stock!C2"), ("out_of_range", "Stock!D2"), ("not_in_enum", "Stock!E2"),
        ("duplicate_key", "Stock!A3"), ("out_of_range", "Stock!C3"),
        ("not_in_list", "Stock!F3")]
    assert result.total == 6 and result.next_cursor is None and "cut" not in data
    assert result.summary == ("d1: 1 tables checked; problem cells: 1 wrong type, 2 out of "
                              "range, 1 not in enum, 1 not in list, 1 duplicate key; "
                              "problems 1-6 of 6")


def test_a_good_workbook_has_no_problems(box):
    session = _session(box, _book(GOOD))
    result = _validate(box, session, {"tables": [{"table": "t:Products", "key": ["SKU"]}]})
    assert result.ok and result.data["ok"] is True and result.data["problems"] == []
    assert result.summary == "d1: 1 tables checked; no problems"


@pytest.mark.parametrize("schema,path", [
    ({"tables": [{"table": "t:Products", "columns": [{"name": "Price", "type": "number",
                                                     "min": "cheap"}]}]},
     "tables[0].columns[0].min"),
    ({"tables": [{"table": "t:Products"}]}, "tables[0]"),
    ({}, ""),
], ids=["min", "nothing", "empty"])
def test_xlsx_validate_data_refuses_a_bad_schema(box, schema, path):
    result = _validate(box, _session(box, _book(GOOD)), schema)
    assert not result.ok and result.error.code == "invalid_arguments"
    assert result.error.field == "schema"
    assert result.error.details == {"code": "bad_schema", "path": path}
    assert "cheap" not in result.error.message


def test_xlsx_validate_data_errors(box):
    session = _session(box, _book(GOOD))
    schema = {"tables": [{"table": "t:Products", "key": ["SKU"]}]}
    result = _validate(box, session, schema, cursor="3.0.00000000")
    assert result.error.code == "invalid_arguments" and result.error.field == "cursor"
    result = box.dispatch(session, "xlsx_validate_data", {"doc": "d9", "schema": schema})
    assert result.error.code == "not_found"


LONG_ROWS = 6_000


def test_a_large_check_is_bounded_and_paged(box):
    rows = [[f"S{r:05d}", "x", "1" if r % 3 == 0 else r, 1, "a", "b"]
            for r in range(LONG_ROWS)]
    session = _session(box, _book(rows))
    schema = {"tables": [{"table": "t:Products", "key": ["SKU"],
                          "columns": [{"name": "Price", "type": "number"}]}]}
    expected = LONG_ROWS // 3
    seen, cursor, pages = [], None, 0
    while True:
        result = _validate(box, session, schema, **({"cursor": cursor} if cursor else {}))
        assert result.ok, result.error
        assert len(result.to_text()) <= box.limits.max_result_chars
        assert len(result.data["problems"]) <= box.limits.max_list_items
        assert result.total == expected and result.data["counts"] == {"wrong_type": expected}
        seen.extend(result.data["problems"])
        pages += 1
        cursor = result.next_cursor
        if cursor is None:
            break
        assert result.data["cut"]["problems"] == expected - len(seen)
    assert len(seen) == expected and pages == expected // box.limits.max_list_items
    assert len({p["at"] for p in seen}) == expected


def test_a_page_is_held_to_its_characters():
    rows = [[f"S{r}", "x", "1", 1, "a", "b"] if r % 2 else [f"S{r}", "x", 1, 1, "a", "b"]
            for r in range(200)]
    small = Limits(max_result_chars=3000)
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0, limits=small) as toolbox:
        session = _session(toolbox, _book(rows))
        result = _validate(toolbox, session, {"tables": [{"table": "t:Products", "columns": [
            {"name": "Price", "type": "number"}]}]})
    assert result.ok, result.error
    assert result.data["cut"]["chars"] is True and result.next_cursor is not None
    assert len(result.to_text()) <= small.max_result_chars


def test_no_cell_text_in_results_summaries_warnings_or_errors(box):
    session = _session(box, _marked_book())
    bad = {"tables": [{"table": "t:Products", "columns": [
        {"name": "Score", "type": "number", "enum": [MARKER]}]}]}
    results = [_validate(box, session, MARKED_SCHEMA), _validate(box, session, bad)]
    assert results[0].ok and not results[1].ok
    assert results[0].data["counts"]
    assert MARKER not in results[0].to_text() and MARKER.lower() not in results[0].to_text()
    assert MARKER not in results[1].error.message
    for result in results:
        assert MARKER not in (result.summary or "")
        assert MARKER not in " ".join(result.warnings)
    for record in session.log:
        assert MARKER not in json.dumps(getattr(record, "summary", "") or "")
