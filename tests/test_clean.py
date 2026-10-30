"""Cleaning a declared column's text (X6): a leading label stripped, white space collapsed, a
cell split into a list -- the caller's rules on the declaration, applied where the table is
read, every change counted, the pattern's cost bounded, and the tool definitions as they
were.  Every workbook is made up here: a sheet of tasks per team, whose "Owners" cell
starts with a lead's label (``Lead: Robin;``) before the owners, separated by ``;``."""

from __future__ import annotations

import json
import time
import warnings

import pytest
from build import Formula, WorkbookBuilder
from ooxml_edit.tools import Toolbox

from xlsx_agent import (CleanedWarning, DeclarationError, DeriveError, QueryError, Workbook,
                        table_dataset)
from xlsx_agent.clean import (EXTRACT_WINDOW, LABEL_WINDOW, MAX_PATTERN, Clean, RuleError,
                              compile_pattern, parse_clean)
from xlsx_agent.tools import GROUPS, TOOLS, format as xlsx_format

HEADER = ["Ref", "Task", "Owners", "Note", "Score"]

TEAMS = {
    "Team East": [["E-1", "paint", "Lead: Robin;  Kim ;Ash", "  two   spaces ", 3],
                  ["E-2", "sand", "Kim", "fine", "Lead: Sam; 4"]],
    "Team West": [["W-1", "seal", "Lead: Sam;Ash;;Lee", "Lead: Sam;", 5],
                  ["W-2", "wax", None, "dry\tenough", 2]],
}

LABEL = r"^Lead:[^;]*;\s*"

#: The declaration most tests use.
TASKS = {"name": "tasks", "sheets": "Team *", "key": "Ref",
         "fields": [{"header": "Owners", "id": "owners",
                     "clean": {"strip": LABEL, "split": ";"}},
                    {"header": "Note", "clean": {"strip": LABEL, "collapse_space": True}},
                    {"header": "Score", "clean": {"strip": LABEL}}]}


def teams(extra: dict | None = None) -> WorkbookBuilder:
    book = WorkbookBuilder()
    for name, rows in {**TEAMS, **(extra or {})}.items():
        book.sheet(name).rows("A1", [HEADER] + rows)
    return book


def opened(*declarations, book: WorkbookBuilder | None = None) -> Workbook:
    wb = Workbook.open((book or teams()).to_bytes())
    if declarations:
        wb.declare_tables(list(declarations))
    return wb


@pytest.fixture(scope="module")
def wb() -> Workbook:
    return opened(TASKS)


def cells(wb: Workbook, address: str) -> list[list[dict]]:
    return [[v.as_dict() for v in row] for row in wb.read(address)]


# -- the rules ---------------------------------------------------------------------------------


def test_strip_split_and_collapse_in_that_order():
    rule = parse_clean({"strip": LABEL, "split": ";", "collapse_space": True})
    assert rule.apply("Lead: Robin;  Kim   Lee ;Ash;;") == (("Kim Lee", "Ash"), True)
    assert rule.apply("Kim") == (("Kim",), False)          # a list of one, the text as it was
    assert rule.apply(" Kim ") == (("Kim",), True)          # trimmed: changed
    assert rule.apply("Lead: Robin;") == (None, True)       # nothing left: blank
    collapse = parse_clean({"collapse_space": True})
    assert collapse.apply(" a \t b\n c ") == ("a b c", True)
    assert collapse.apply("a b") == ("a b", False)
    strip = parse_clean({"strip": "Lead: "})                 # matched at the start only
    assert strip.apply("x Lead: y") == ("x Lead: y", False)
    assert rule.spec == {"strip": LABEL, "split": ";", "collapse_space": True}


def test_a_label_is_looked_for_in_the_first_characters_only():
    rule = parse_clean({"strip": r"^[^;]*;"})
    near = "x" * (LABEL_WINDOW - 1) + ";rest"
    far = "x" * LABEL_WINDOW + ";rest"
    assert rule.apply(near) == ("rest", True)
    assert rule.apply(far) == (far, False)


def test_read_through_the_table_cleaned_each_value_with_its_own_cell(wb):
    rows = cells(wb, "t:tasks")
    assert rows[0][2] == {"a": "'Team East'!C2", "v": ["Kim", "Ash"], "t": "list",
                          "cleaned": True}
    assert rows[0][3] == {"a": "'Team East'!D2", "v": "two spaces", "t": "text",
                          "cleaned": True}
    assert rows[1][2] == {"a": "'Team East'!C3", "v": ["Kim"], "t": "list"}   # not changed
    # a cleaned text that reads as a number is flagged as the reader flags one
    assert rows[1][4] == {"a": "'Team East'!E3", "v": "4", "t": "text",
                          "issue": "number_as_text", "cleaned": True}
    assert rows[2][2]["v"] == ["Ash", "Lee"]
    assert rows[2][3] == {"a": "'Team West'!D2", "v": None, "t": "blank", "cleaned": True}
    assert rows[3][2] == {"a": "'Team West'!C3", "v": None, "t": "blank"}
    assert rows[3][3]["v"] == "dry enough"
    assert rows[0][4] == {"a": "'Team East'!E2", "v": 3, "t": "number"}    # not text: as is


def test_the_text_as_stored_stays_reachable(wb):
    value = wb.cell("t:tasks/E-1/owners")
    assert (value.v, value.cleaned, value.original) == (("Kim", "Ash"), True,
                                                         "Lead: Robin;  Kim ;Ash")
    raw = wb.cell("'Team East'!C2")                     # the cell by its address: as stored
    assert (raw.v, raw.t, raw.cleaned) == ("Lead: Robin;  Kim ;Ash", "text", False)
    assert [v.v for v in next(iter(wb.read("'Team East'!A1:E1")))] == HEADER


def test_the_header_is_never_cleaned_and_a_single_sheet_table_reads_alike():
    book = WorkbookBuilder()
    book.sheet("Tasks").rows("A1", [["Ref", "Lead: x; Owners"], ["T-1", "Lead: a; b;c"]])
    wb = Workbook.open(book.to_bytes())
    region = wb.declare_table("tasks", "Tasks", key="Ref", fields=[
        {"column": "B", "clean": {"strip": LABEL, "split": ";"}}])
    assert region.columns == ("Ref", "Lead: x; Owners")
    assert region.where == "Tasks!A1:B2"
    assert cells(wb, "t:tasks") == [[{"a": "Tasks!A2", "v": "T-1", "t": "text"},
                                     {"a": "Tasks!B2", "v": ["b", "c"], "t": "list",
                                      "cleaned": True}]]


def test_a_formulas_cached_text_is_cleaned_too():
    book = WorkbookBuilder()
    book.sheet("Tasks").rows("A1", [["Ref", "Owners"],
                                    ["T-1", Formula('"Lead: a;b"', cached="Lead: a;b")]])
    wb = Workbook.open(book.to_bytes())
    wb.declare_table("tasks", "Tasks", key="Ref",
                     fields=[{"header": "Owners", "clean": {"strip": LABEL}}])
    value = wb.cell("t:tasks/T-1/Owners")
    assert (value.v, value.formula, value.cleaned) == ("b", True, True)


# -- the declaration ---------------------------------------------------------------------------


@pytest.mark.parametrize("clean, code, path", [
    ({"strip": "(a+)+"}, "bad_pattern", "strip"),
    ({"strip": "(a|ab)*"}, "bad_pattern", "strip"),
    ({"strip": r"(\w)\1"}, "bad_pattern", "strip"),
    ({"strip": r"^Lead:\s*.*;\s*"}, "bad_pattern", "strip"),       # two repeats give back
    ({"strip": "a?a?a?aaa"}, "bad_pattern", "strip"),
    ({"strip": "(ab|cd)(ab|cd)(ab|cd)(ab|cd)(ab|cd)"}, "bad_pattern", "strip"),
    ({"strip": "x" * (MAX_PATTERN + 1)}, "bad_pattern", "strip"),
    ({"strip": "("}, "bad_pattern", "strip"),
    ({"strip": 3}, "bad_pattern", "strip"),
    ({"split": ""}, "bad_declaration", "split"),
    ({"split": ["a"]}, "bad_declaration", "split"),
    ({"collapse_space": "yes"}, "bad_declaration", "collapse_space"),
    ({"trim": True}, "bad_declaration", "trim"),
    ({}, "bad_declaration", None),
    ({"collapse_space": False}, "bad_declaration", None),
    ({"extract": r"Status:\s*.+$"}, "bad_pattern", "extract"),         # no capture group
    ({"extract": r"(\w+): (\w+)"}, "bad_pattern", "extract"),          # two
    ({"extract": r"Status:\s*(.+)$"}, "bad_pattern", "extract"),        # \s* gives back
    ({"extract": r"(.*);"}, "bad_pattern", "extract"),
    ({"extract": ""}, "bad_pattern", "extract"),
])
def test_a_rule_that_is_not_one_is_refused_where_it_is(clean, code, path):
    declared = {**TASKS, "fields": [{"header": "Note", "clean": clean}]}
    with pytest.raises(DeclarationError) as raised:
        opened(declared)
    error = raised.value
    assert error.code == code
    assert error.path == "tables[0].fields[0].clean" + (f".{path}" if path else "")


@pytest.mark.parametrize("pattern", [
    LABEL, r"^\s*Lead\s*:\s*[^;]*;\s*", r"^Lead:.*?;\s*", r"(?i)^lead:[^;]*;",
    r"^(?:Lead|Owner):[^;]*;\s*", r"^\[[^\]]*\]\s*", r"^\s*\s*x"])
def test_patterns_that_cannot_give_characters_back_are_taken(pattern):
    assert compile_pattern(pattern).pattern == pattern


def test_the_worst_pattern_taken_stays_cheap():
    """The cost bound ROADMAP.md (X6) reports: the worst form allowed -- one repeat that can
    give characters back, before another -- on a cell as long as Excel holds."""
    rule = Clean(strip=compile_pattern(r"^\s*\s*x"))
    text = " " * 32767
    start = time.perf_counter()
    for _ in range(100):
        assert rule.apply(text) == (text, False)
    assert (time.perf_counter() - start) / 100 < 0.01


def test_the_key_column_and_hidden_columns_are_not_cleaned():
    with pytest.raises(DeclarationError) as raised:
        opened({**TASKS, "fields": [{"header": "Ref", "clean": {"collapse_space": True}}]})
    assert (raised.value.code, raised.value.path) == ("clean_key", "tables[0].fields[0].clean")
    with pytest.raises(DeclarationError) as raised:
        opened({**TASKS, "fields": [{"header": "Note", "hidden": True,
                                     "clean": {"collapse_space": True}}]})
    assert raised.value.path == "tables[0].fields[0].clean"
    book = WorkbookBuilder()
    book.sheet("Facts").rows("A1", [["author", "Lead: x; Pat"], ["round", 3]])
    with pytest.raises(DeclarationError) as raised:      # a key-value table's key: its field
        opened({"name": "run", "type": "key_value", "sheets": "Facts",
                "fields": [{"column": "A", "clean": {"strip": LABEL}}]}, book=book)
    assert raised.value.code == "clean_key"
    wb = opened({"name": "run", "type": "key_value", "sheets": "Facts",
                 "fields": [{"column": "B", "clean": {"strip": LABEL}}]}, book=book)
    assert wb.cell("t:run/author/value").v == "Pat"


def test_rule_errors_name_the_rule_never_a_cell():
    with pytest.raises(RuleError) as raised:
        parse_clean({"strip": r"(\d+)+"})
    assert raised.value.key == "strip" and "repeat inside a repeat" in raised.value.message


# -- never silent ------------------------------------------------------------------------------


def test_profile_and_describe_flag_the_cleaned_values(wb):
    owners, note = wb.profile("t:tasks", ["owners", "Note"])
    assert owners.type == "list"
    assert owners.flags["cleaned"].as_dict() == {"count": 2,
                                                 "first": ["'Team East'!C2", "'Team West'!C2"]}
    assert note.flags["cleaned"].count == 3
    assert owners.text_length_max is None
    shown = json.dumps(wb.catalogue(), default=str)
    assert '"cleaned": 2' in shown and '"list"' in shown


def test_queries_test_each_item_of_a_list(wb):
    region = wb.region("t:tasks")
    found = region.query(["Ref"], where={"owners": {"eq": "Kim"}})
    assert [row[0].v for row in found.rows] == ["E-1", "E-2"]
    found = region.query(["Ref"], where={"owners": {"in": ["Lee", "Zed"]}})
    assert [row[0].v for row in found.rows] == ["W-1"]
    found = region.query(["Ref"], where={"owners": {"contains": "as"}})
    assert [row[0].v for row in found.rows] == ["E-1", "W-1"]
    found = region.query(["Ref"], sort=["owners"])
    assert [row[0].v for row in found.rows] == ["W-1", "E-2", "E-1", "W-2"]   # item by item
    with pytest.raises(QueryError) as raised:
        region.query(["Ref"], where={"owners": {"gte": 3}})
    assert raised.value.column_type == "list"


def test_derive_counts_list_cells_refuses_grouping_by_them_and_reports_cleaning(wb):
    derived = wb.derive("t:tasks", measures=[{"fn": "count", "column": "owners"}],
                        where={"owners": {"eq": "Kim"}})
    assert derived.rows == [{"count owners": 2}]
    assert {k: c.as_dict() for k, c in derived.cleaned.items()} == {
        "owners": {"count": 2, "first": ["'Team East'!C2", "'Team West'!C2"]}}
    assert derived.as_dict()["cleaned"]["owners"]["count"] == 2
    with pytest.raises(DeriveError) as raised:
        wb.derive("t:tasks", group_by=["owners"])
    assert (raised.value.code, raised.value.column_type) == ("wrong_type", "list")
    with pytest.raises(DeriveError) as raised:
        wb.derive("t:tasks", measures=[{"fn": "first", "column": "owners"}])
    assert raised.value.valid_options == ["count"]
    note = wb.derive("t:tasks", group_by=["Note"], measures=[{"fn": "count"}])
    assert [row["Note"] for row in note.rows] == ["two spaces", "fine", None, "dry enough"]


def test_a_table_into_a_deck_joins_a_lists_items_by_line_breaks_and_warns(wb):
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        found = table_dataset(wb, "t:tasks[owners]")
    assert [row[0] for row in found.rows] == ["Kim\nAsh", "Kim", "Ash\nLee", None]
    assert found.types == ("text",)
    assert found.detail["cleaned"].as_dict() == {"count": 2, "first": ["'Team East'!C2",
                                                                       "'Team West'!C2"]}
    assert [w.category for w in caught] == [CleanedWarning]
    message = str(caught[0].message)
    assert "t:tasks: 2 values" in message and "Robin" not in message and "Kim" not in message
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        assert table_dataset(wb, "t:tasks[Task]").detail is None      # nothing cleaned


def test_a_list_column_is_checked_item_by_item(wb):
    report = wb.validate_data({"tables": [{"table": "t:tasks", "columns": [
        {"name": "owners", "type": "text", "enum": ["Kim", "Ash"]}]}]})
    assert [(p.code, p.address) for p in report] == [("not_in_enum", "'Team West'!C2")]
    report = wb.validate_data({"tables": [{"table": "t:tasks", "columns": [
        {"name": "owners", "type": "number"}]}]})
    assert report.counts == {"wrong_type": 3}


def test_two_workbooks_compare_their_cleaned_values():
    old = opened(TASKS)
    rows = {**TEAMS, "Team West": [["W-1", "seal", "Lead: Kim;Ash;;Lee", "Lead: Sam;", 5],
                                   ["W-2", "wax", "Lee", "dry\tenough", 2]]}
    book = WorkbookBuilder()
    for name, data in rows.items():
        book.sheet(name).rows("A1", [HEADER] + data)
    new = opened(TASKS, book=book)
    diff = Workbook.compare(old, new, tables=["t:tasks"])
    changes = [(c.key, c.column, c.a.v, c.b.v) for c in diff.tables[0].changed]
    assert changes == [("W-2", "owners", None, ("Lee",))]


# -- the tools ---------------------------------------------------------------------------------


@pytest.fixture(scope="module")
def tools():
    toolbox = Toolbox(TOOLS, formats=[xlsx_format(tables=[TASKS])], groups=GROUPS)
    session = toolbox.session()
    blob = session.add_blob(teams().to_bytes(), name="tasks.xlsx")
    assert toolbox.dispatch(session, "open_document", {"blob": blob}).ok
    return toolbox, session


def dispatch(tools, name, **arguments):
    toolbox, session = tools
    result = toolbox.dispatch(session, name, arguments)
    assert result.ok, result.error
    return result.data


def test_the_read_tools_mark_and_count_cleaned_values(tools):
    read = dispatch(tools, "xlsx_read", doc="d1", target="t:tasks[owners]")
    assert read["cleaned"] == {"count": 2, "first": ["'Team East'!C2", "'Team West'!C2"]}
    assert read["rows"][0][0] == {"a": "'Team East'!C2", "v": ["Kim", "Ash"], "t": "list",
                                  "cleaned": True}
    raw = dispatch(tools, "xlsx_read", doc="d1", target="'Team East'!C2", cells_only=True)
    assert raw["rows"][0][0]["v"] == "Lead: Robin;  Kim ;Ash" and "cleaned" not in raw
    assert raw["held_by"] == {"table": "t:tasks", "rows": "t:tasks/E-1"}
    row = dispatch(tools, "xlsx_read", doc="d1", target="'Team East'!C2")   # its table row
    assert row["target"] == "t:tasks/E-1" and row["rows"][0][2]["v"] == ["Kim", "Ash"]
    query = dispatch(tools, "xlsx_query", doc="d1", table="t:tasks", columns=["Ref", "Note"],
                     where=[{"column": "owners", "op": "contains", "value": "lee"}])
    assert query["rows"] == [[{"a": "'Team West'!A2", "v": "W-1", "t": "text"},
                              {"a": "'Team West'!D2", "v": None, "t": "blank",
                               "cleaned": True}]]
    assert query["cleaned"] == {"count": 1, "first": ["'Team West'!D2"]}
    profile = dispatch(tools, "xlsx_profile", doc="d1", target="t:tasks[owners]")
    assert profile["columns"][0]["flags"]["cleaned"]["count"] == 2
    derived = dispatch(tools, "xlsx_derive", doc="d1", **{"from": "t:tasks"},
                       measures=[{"fn": "count", "column": "Note"}], id="notes")
    assert derived["cleaned"]["Note"]["count"] == 3


def test_cleaning_changes_no_tool_definition():
    plain = Toolbox(TOOLS, groups=GROUPS).definitions("anthropic", defer=False)
    cleaning = Toolbox(TOOLS, formats=[xlsx_format(tables=[TASKS])],
                       groups=GROUPS).definitions("anthropic", defer=False)
    assert cleaning == plain
    # this repository's own tools, character for character as main had them before X6
    # (X10's, T19's, L8's, X3b's and X5's additions aside); the read tools
    # only (E3's xlsx_set_cells is new)
    own = [d for d in plain if d["name"].startswith("xlsx_") and d["name"] != "xlsx_set_cells"]
    assert len(json.dumps(own, separators=(",", ":"), ensure_ascii=False)) == XLSX_CHARS


#: The compact JSON length of the ``xlsx_*`` definitions before and after X6 (10,155),
#: since X10 described xlsx_compare's summary (+41), since T19 said what key_column,
#: sources and xlsx_derive's sort name (+55), and since L8 renamed xlsx_read's key_column
#: key and added its columns (+97), and since X3b added xlsx_read's markdown (+97), and
#: since X5 added xlsx_compare's cells and formula_text and xlsx_read's validations (+356),
#: and since lane X's X6 named the structure xlsx_compare's cells compares (+49), and
#: since round 10's O5 let xlsx_query take a dataset (+59: ``doc`` optional), and since
#: round 25 gave xlsx_read a cited cell's table row (X0) and xlsx_derive its join (X1)
#: (+1,184; X2's extract is a declared rule, and changes no definition).
XLSX_CHARS = 12_093


def test_a_deck_table_from_a_cleaned_table_warns_in_the_tools_result():
    pptx_agent = pytest.importorskip("pptx_agent")
    pptx_tools = pytest.importorskip("pptx_agent.tools")
    with Toolbox([*pptx_tools.TOOLS, *TOOLS],
                 formats=[pptx_tools.FORMAT, xlsx_format(tables=[TASKS])],
                 groups=[*pptx_tools.GROUPS, *GROUPS], workers=0) as toolbox:
        session = toolbox.session()
        deck = pptx_agent.Document.new()
        deck.add_slide("Blank")
        session.open(deck.to_bytes(), "deck.pptx")                       # d1
        session.open(teams().to_bytes(), "tasks.xlsx")                    # d2
        result = toolbox.dispatch(session, "ppt_add_table", {
            "doc": "d1", "slide": "s:256", "box": {"x": 40, "y": 60, "w": 600, "h": 300},
            "from": "t:tasks", "from_columns": ["Ref", "owners"]})
        assert result.ok, result.error
        assert result.data["cells"][1] == ["E-1", "Kim\nAsh"]
        assert any("t:tasks: 6 values were cleaned" in w for w in result.warnings), result.warnings


# -- round 25, X2: extract, a value captured from a text cell ---------------------------------

STATUS = r"Status:\s*(\S.*)$"


def test_extract_searches_and_gives_the_captured_text_or_blank():
    rule = parse_clean({"extract": STATUS})
    assert rule.apply("Long notes.\nStatus: Partly done") == ("Partly done", True)
    assert rule.apply("Status:Done") == ("Done", True)
    assert rule.apply("No status line") == (None, True)         # not found: blank
    assert rule.apply("Status: done\nmore") == (None, True)     # $ is the text's end
    assert parse_clean({"extract": r"(?m)^Status: (.*)$"}).apply(
        "a\nStatus: open\nb") == ("open", True)                 # (?m): a line's end
    assert parse_clean({"extract": r"^(\w+)"}).apply("word rest") == ("word", True)
    assert parse_clean({"extract": r"x(y)?"}).apply("x") == (None, True)  # the group took no part
    assert parse_clean({"extract": r"(\d+)"}).apply("12") == ("12", False)
    assert rule.spec == {"extract": STATUS}


def test_extract_comes_after_strip_and_before_split_and_collapse():
    rule = parse_clean({"strip": r"^Lead:[^;]*;\s*", "extract": r"Owners:\s*(\S.*)$",
                        "split": ",", "collapse_space": True})
    assert rule.spec == {"strip": r"^Lead:[^;]*;\s*", "extract": r"Owners:\s*(\S.*)$",
                         "split": ",", "collapse_space": True}
    assert rule.apply("Lead: Robin; notes Owners: Kim  Lee, Ash") == (("Kim Lee", "Ash"), True)


def test_an_extract_match_starts_in_the_last_characters_only():
    rule = parse_clean({"extract": STATUS})
    tail = "Status: open"
    assert rule.apply("x" * 40000 + "\n" + tail) == ("open", True)
    far = tail + "\n" + "x" * EXTRACT_WINDOW
    assert rule.apply(far) == (None, True)
    anchored = parse_clean({"extract": r"^(\w+)"})
    assert anchored.apply("word " + " " * EXTRACT_WINDOW) == (None, True)


@pytest.mark.parametrize("pattern", [STATUS, r"([^;]*);", r"(\S+)$", r"(?m)^Status: (.*)$",
                                     r"\[(\w+)\]", r"( *)\S"])
def test_extract_patterns_that_give_nothing_back_are_taken(pattern):
    assert parse_clean({"extract": pattern}).extract.pattern == pattern


def test_the_worst_extract_taken_stays_cheap():
    """The bound clean.py reports: the worst forms allowed, tried at every place of the
    window, on cells as long as Excel holds."""
    for pattern in (r"([^;]*);", r"( *)\S", r"(\s*)$"):
        rule = parse_clean({"extract": pattern})
        for text in (" " * 32767, "a" * 32767, "x:" * 16383):
            start = time.perf_counter()
            for _ in range(10):
                rule.apply(text)
            assert (time.perf_counter() - start) / 10 < 0.02, pattern


def test_a_declared_extract_reads_through_the_table_with_its_cells():
    book = WorkbookBuilder()
    book.sheet("Items").rows("A1", [["Ref", "Review"],
                                    ["I-1", "Checked twice.\nStatus: Partly done"],
                                    ["I-2", "Not looked at"], ["I-3", "Status:  2026-03-01"]])
    wb = Workbook.open(book.to_bytes())
    wb.declare_tables([{"name": "items", "sheets": "Items", "key": "Ref",
                        "fields": [{"header": "Review", "id": "status",
                                    "clean": {"extract": STATUS}}]}])
    assert cells(wb, "t:items[status]") == [
        [{"a": "Items!B2", "v": "Partly done", "t": "text", "cleaned": True}],
        [{"a": "Items!B3", "v": None, "t": "blank", "cleaned": True}],
        [{"a": "Items!B4", "v": "2026-03-01", "t": "text", "issue": "date_as_text",
          "cleaned": True}]]
    assert cells(wb, "Items!B2") == [[{"a": "Items!B2", "v": "Checked twice.\nStatus: Partly done",
                                       "t": "text"}]]                       # as stored
    grouped = wb.derive("t:items", group_by=["status"])
    assert [r["status"] for r in grouped.rows] == ["Partly done", None, "2026-03-01"]
    assert [str(s) for s in grouped.provenance[0]["status"].sources] == ["t:items/I-1/status"]
    assert grouped.cleaned["status"].count == 3
