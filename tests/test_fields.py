"""Fields derived from another (round 26, X4) and rules on numbers as shown (X5).

A declared field with ``from_field`` reads another column's cells by rules of its own --
a value extracted from a narrative, beside the narrative, in one table -- and a ``clean``
with ``"on": "shown"`` reads a number cell as the text Excel shows it.  Every workbook is
made up here: a log of jobs whose "Notes" end with the part they used (``Part: P-2``),
whose "Size" holds text ("5.1") and numbers (5.2), and a list of parts."""

from __future__ import annotations

import json
from datetime import date

import pytest
from build import WorkbookBuilder
from ooxml_common import numfmt
from ooxml_common.locales import EN_US, Locale
from ooxml_edit.tools import Toolbox

from xlsx_agent import DeclarationError, EditError, Workbook
from xlsx_agent.clean import RuleError, parse_clean
from xlsx_agent.shown import GENERAL_WIDTH, shown_text
from xlsx_agent.tools import GROUPS, TOOLS, format as xlsx_format

PART = r"Part:\s*(\S+)$"
LEADING = r"^(\d+)"

#: The declaration most tests use: the part beside the notes, the leading whole number of
#: the size beside the size.
JOBS = {"name": "jobs", "sheets": "Log", "key": "Ref",
        "fields": [{"header": "Notes", "id": "notes", "description": "What was done"},
                   {"id": "part", "from_field": "notes", "description": "The part used",
                    "clean": {"extract": PART}},
                   {"id": "whole", "from_field": "Size",
                    "clean": {"extract": LEADING, "on": "shown"}}]}

PARTS = {"name": "parts", "sheets": "Parts", "key": "Part"}

DUTCH = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
               language="nl-BE", list_separator=";")


def book(extra=None) -> WorkbookBuilder:
    """The workbook; ``extra(out, log)`` adds to it."""
    out = WorkbookBuilder()
    one = out.style(num_fmt="0.0")
    log = out.sheet("Log")
    log.rows("A1", [["Ref", "Notes", "Size", "Hours"],
                    ["J-1", "Fitted the hinge. Part: P-2", "5.1", 2],
                    ["J-2", "Checked the seal.\nPart: P-1", None, 1],
                    ["J-3", "Nothing used", None, 4],
                    ["J-4", "Replaced it. Part: P-2", 12, 3]])
    log.cell("C3", 5.2, style=one)
    parts = out.sheet("Parts")
    parts.rows("A1", [["Part", "Name"], ["P-1", "seal"], ["P-2", "hinge"]])
    if extra is not None:
        extra(out, log)
    return out


def opened(*declarations, locale=None, policy=None) -> Workbook:
    wb = Workbook.open(book().to_bytes(), locale=locale, policy=policy)
    wb.declare_tables(list(declarations) or [JOBS, PARTS])
    return wb


@pytest.fixture(scope="module")
def wb() -> Workbook:
    return opened()


def cells(wb: Workbook, address: str) -> list[list[dict]]:
    return [[v.as_dict() for v in row] for row in wb.read(address)]


# -- X4: a field derived from another ----------------------------------------------------------


def test_a_derived_field_stands_beside_its_column_and_reads_its_cells(wb):
    table = wb.table("t:jobs")
    assert table.columns == ("Ref", "notes", "part", "Size", "whole", "Hours")
    assert table.derived == (None, None, "notes", None, "Size", None)
    assert table.source_of("part") == "notes" and table.source_of("notes") is None
    rows = cells(wb, "t:jobs")
    assert rows[0][1] == {"a": "Log!B2", "v": "Fitted the hinge. Part: P-2", "t": "text"}
    assert rows[0][2] == {"a": "Log!B2", "v": "P-2", "t": "text", "cleaned": True}
    assert rows[1][2] == {"a": "Log!B3", "v": "P-1", "t": "text", "cleaned": True}
    assert rows[2][2] == {"a": "Log!B4", "v": None, "t": "blank", "cleaned": True}
    assert wb.cell("t:jobs/J-4/part").v == "P-2"
    # its cells are the column's: once in the table's range, the column's in its own
    assert table.where == "Log!A1:D5"
    assert table.located(wb.resolve("t:jobs[part]").range.area) == "Log!B2:B5"


def test_the_source_column_keeps_its_own_rules():
    wb = opened({**JOBS, "fields": [
        {"header": "Notes", "id": "notes", "clean": {"collapse_space": True}},
        {"id": "part", "from_field": "Notes", "clean": {"extract": PART}}]})   # by header
    row = next(iter(wb.read("t:jobs/J-2")))
    assert (row[1].v, row[2].v) == ("Checked the seal. Part: P-1", "P-1")
    assert row[1].original == row[2].original == "Checked the seal.\nPart: P-1"


def test_query_profile_and_describe(wb):
    result = wb.table("t:jobs").query(["Ref", "part"], where={"part": {"eq": "P-2"}})
    assert [[v.v for v in row] for row in result.rows] == [["J-1", "P-2"], ["J-4", "P-2"]]
    stats = wb.profile("t:jobs", ["part"])[0]
    assert (stats.derived_from, stats.description, stats.range) == ("notes", "The part used",
                                                                    "Log!B2:B5")
    assert stats.as_dict()["derived_from"] == "notes"
    columns = {c["name"]: c for c in wb.catalogue()["regions"][0]["columns"]}
    assert columns["part"]["derived_from"] == "notes"
    assert columns["whole"]["derived_from"] == "Size"
    assert "derived_from" not in columns["notes"]


def test_derive_groups_by_the_derived_field_with_the_source_cells(wb):
    derived = wb.derive("t:jobs", group_by=["part"],
                        measures=[{"fn": "sum", "column": "Hours", "as": "hours"}])
    assert derived.rows == [{"part": "P-2", "hours": 5}, {"part": "P-1", "hours": 1},
                            {"part": None, "hours": 4}]
    sources = [str(s) for s in derived.provenance[0]["part"].sources]
    assert sources == ["t:jobs/{J-1,J-4}/notes"]          # by key, the column read
    unkeyed = opened({**JOBS, "key": None}, PARTS)
    again = unkeyed.derive("t:jobs", group_by=["part"],
                           measures=[{"fn": "count", "column": "Ref", "as": "n"}])
    assert [str(s) for s in again.provenance[0]["part"].sources] == ["Log!B2", "Log!B5"]


def test_a_join_on_the_derived_field(wb):
    joined = wb.derive("t:jobs", join={"from": "t:parts", "on": ["part", "Part"],
                                       "columns": ["Name"]})
    assert [(r["Ref"], r["Name"]) for r in joined.rows] == [
        ("J-1", "hinge"), ("J-2", "seal"), ("J-3", None), ("J-4", "hinge")]
    assert [str(s) for s in joined.provenance[1]["Name"].sources] == ["t:parts/P-1/Name"]


def test_a_pivot_on_the_derived_field_with_listed_values(wb):
    """Round 27, X6: listed pivot values on a derived field, its sources the cells read."""
    ds = wb.derive("t:jobs", pivot="part", pivot_values=["P-1", "P-2", "P-3"],
                   measures=[{"fn": "sum", "column": "Hours", "as": "hours"}])
    assert ds.rows == [{"P-1": 1, "P-2": 5, "P-3": None}]
    assert ds.pivot_unlisted == [{"name": "(blank)", "v": None, "t": "blank", "rows": 1}]
    assert ds.left_out["part"]["not_listed"].first == ("Log!B4",)     # the cell read
    assert [str(s) for s in ds.provenance[0]["P-2"].sources] == ["t:jobs/{J-1,J-4}/Hours"]


def test_a_value_by_key_in_a_derived_column_is_refused_naming_the_source():
    wb = opened()
    with pytest.raises(EditError) as refused:
        wb.set_cells({"t:jobs/J-1/part": "P-3"})
    assert refused.value.code == "derived_column"
    assert "t:jobs/<key>/notes" in str(refused.value)
    wb.set_cells({"t:jobs/J-1/notes": "Fitted it. Part: P-3"})
    assert wb.cell("t:jobs/J-1/part").v == "P-3"


def test_a_write_policy_never_names_a_derived_column():
    named = opened(policy={"tables": [{"table": "t:jobs", "columns": ["part"]}]})
    with pytest.raises(EditError, match="'part' is read from column 'notes'") as refused:
        named.set_cells({"t:jobs/J-1/notes": "Part: P-1"})
    assert refused.value.code == "bad_policy"
    wb = opened(policy={"tables": [{"table": "t:jobs"}]})        # every column: not it
    wb.set_cells({"t:jobs/J-1/notes": "Part: P-1"})
    with pytest.raises(EditError) as again:
        wb.set_cells({"t:jobs/J-1/part": "P-1"})
    assert again.value.code == "derived_column"


def test_hiding_a_column_beside_a_derived_field_needs_nothing_more():
    wb = opened({**JOBS, "fields": [*JOBS["fields"], {"header": "Hours", "hidden": True}]},
                PARTS)
    assert wb.table("t:jobs").columns == ("Ref", "notes", "part", "Size", "whole")
    # the hidden column is withheld as before; the derived ones read their own cells
    assert cells(wb, "Log!D2")[0][0]["t"] == "blank"
    assert wb.cell("t:jobs/J-1/part").a == "Log!B2"


@pytest.mark.parametrize("field, code, path", [
    ({"id": "x", "from_field": "Nope", "clean": {"extract": PART}}, "unknown_column",
     "from_field"),
    ({"id": "x", "from_field": "part", "clean": {"extract": PART}}, "bad_source",
     "from_field"),
    ({"id": "x", "from_field": "Hours", "clean": {"extract": PART}}, "bad_source",
     "from_field"),
    ({"from_field": "notes", "clean": {"extract": PART}}, "bad_declaration", "id"),
    ({"id": "x", "from_field": "notes"}, "bad_declaration", "clean"),
    ({"id": "x", "from_field": "notes", "header": "Notes", "clean": {"extract": PART}},
     "bad_declaration", "header"),
    ({"id": "x", "from_field": "notes", "hidden": True, "clean": {"extract": PART}},
     "bad_declaration", "hidden"),
    ({"id": "x", "from_field": "", "clean": {"extract": PART}}, "bad_declaration",
     "from_field"),
    ({"id": "Size", "from_field": "notes", "clean": {"extract": PART}}, "id_taken", "id"),
    ({"id": "part", "from_field": "notes", "clean": {"extract": PART}}, "id_taken", "id"),
])
def test_a_derived_field_refused(field, code, path):
    fields = [*JOBS["fields"][:2], {"header": "Hours", "hidden": True}, field]
    with pytest.raises(DeclarationError) as refused:
        opened({**JOBS, "fields": fields})
    assert refused.value.code == code
    assert refused.value.path == f"tables[0].fields[3].{path}"


def test_a_derived_field_is_never_the_key():
    with pytest.raises(DeclarationError) as refused:
        opened({**JOBS, "key": "part"})
    assert (refused.value.code, refused.value.path) == ("derived_key", "tables[0].key")


def test_a_column_declared_twice_is_still_refused():
    with pytest.raises(DeclarationError) as refused:
        opened({**JOBS, "fields": [{"header": "Notes"}, {"column": "B", "id": "x"}]})
    assert refused.value.code == "duplicate_field"


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[xlsx_format(tables=[JOBS, PARTS])], groups=GROUPS,
                 workers=0) as toolbox:
        yield toolbox


def test_tools_read_derive_and_staleness(box):
    session = box.session()
    blob = session.add_blob(book().to_bytes(), name="log.xlsx")
    assert box.dispatch(session, "open_document", {"blob": blob}).ok
    read = box.dispatch(session, "xlsx_read", {"doc": "d1", "target": "t:jobs/J-1/part"})
    assert read.ok, read.error
    assert json.loads(json.dumps(read.data))["rows"][0][0]["a"] == "Log!B2"
    derived = box.dispatch(session, "xlsx_derive", {
        "doc": "d1", "from": "t:jobs", "id": "by_part", "group_by": ["part"],
        "measures": [{"fn": "count", "column": "Ref", "as": "jobs"}]})
    assert derived.ok, derived.error
    entry = session.dataset("d:by_part")
    assert not entry.stale
    refused = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "t:jobs/J-3/part", "text": "P-1"}]})
    assert not refused.ok and "t:jobs/<key>/notes" in str(refused.error)
    assert not entry.stale
    written = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "t:jobs/J-3/notes", "text": "Used one. Part: P-1"}]})
    assert written.ok, written.error
    assert entry.stale                                    # the cell it read changed


# -- X5: rules on numbers as shown -------------------------------------------------------------


def test_the_leading_number_of_a_text_and_of_a_number_alike(wb):
    rows = cells(wb, "t:jobs[whole]")
    assert rows[0][0] == {"a": "Log!C2", "v": "5", "t": "text", "issue": "number_as_text",
                          "cleaned": True}
    assert rows[1][0] == {"a": "Log!C3", "v": "5", "t": "text", "issue": "number_as_text",
                          "cleaned": True}          # 5.2 shown as "5.2" in 0.0
    assert rows[3][0]["v"] == "12"                 # General
    value = wb.cell("t:jobs/J-2/whole")
    assert (value.original, value.format) == ("5.2", None)    # the code is not shown unasked
    assert wb.cell("t:jobs/J-2/Size").v == 5.2                # the column read keeps its number
    # coerce reads the extracted text as a number
    derived = wb.derive("t:jobs", measures=[{"fn": "sum", "column": "whole", "as": "s"}],
                        coerce=["number_as_text"])
    assert derived.rows == [{"s": 22}]


def test_the_shown_text_follows_the_workbooks_locale():
    wb = opened({**JOBS, "fields": [{"header": "Size", "clean": {"on": "shown"}}]},
                locale=DUTCH)
    assert [row[0].v for row in wb.read("t:jobs[Size]")] == ["5.1", "5,2", None, "12"]
    assert wb.cell("t:jobs/J-2/Size").cleaned


def test_a_number_whose_shown_text_is_not_measured_is_refused_never_guessed():
    def extra(out, log):
        log.cell("C4", 0.5, style=out.style(num_fmt="# ?/?")).cell("C5", -3)
    wb = Workbook.open(book(extra).to_bytes())
    wb.declare_tables([JOBS])
    rows = cells(wb, "t:jobs[whole]")
    assert rows[2][0] == {"a": "Log!C4", "v": None, "t": "blank", "issue": "unmeasured_format",
                          "cleaned": True}
    assert rows[3][0]["issue"] == "unmeasured_format"            # a negative number


def test_dates_and_text_cells_are_read_as_before():
    wb = Workbook.open(book(lambda out, log: log.cell("C4", date(2026, 1, 5))).to_bytes())
    wb.declare_tables([{**JOBS, "fields": [{"header": "Size", "clean": {"on": "shown"}}]}])
    value = wb.cell("t:jobs/J-3/Size")
    assert (value.t, value.v, value.cleaned) == ("date", "2026-01-05", False)
    assert wb.cell("t:jobs/J-1/Size").as_dict() == {"a": "Log!C2", "v": "5.1", "t": "text",
                                                    "issue": "number_as_text"}


def test_on_is_checked():
    assert parse_clean({"on": "shown"}).spec == {"on": "shown"}
    assert parse_clean({"extract": PART, "on": "text"}).spec == {"extract": PART}
    with pytest.raises(RuleError) as refused:
        parse_clean({"extract": PART, "on": "numbers"})
    assert (refused.value.code, refused.value.key) == ("bad_declaration", "on")
    with pytest.raises(RuleError):
        parse_clean({"on": "text"})                       # no rule at all


OPEN_CODES = {"##0.0E+0", "# ?/?", "## ??/16", "0*d", "[<100]0;0.00",
              "[DBNum1][$-804]General", "[Blue][<0]0;0"}


@pytest.mark.parametrize("code, kind, text",
                         [m for m in numfmt.MEASURED if m[1] not in numfmt.DATE_KINDS])
def test_every_measured_code_shows_as_excel_showed_it_or_is_refused(code, kind, text):
    got = shown_text(46027.5, code, DUTCH)
    if code in OPEN_CODES:
        assert got is None
    else:
        assert got == text


def test_seven_measured_codes_are_open():
    measured = {m[0] for m in numfmt.MEASURED}
    assert OPEN_CODES <= measured and len(OPEN_CODES) == 7


@pytest.mark.parametrize("number, code, text", [
    (5.2, "General", "5.2"), (0.25, "0%", "25%"), (1234.5, "#,##0.00", "1,234.50"),
    (5, "0.0#", "5.0"), (0.375, "0.00", "0.38"), (0.001, "0.00E+00", "1.00E-03"),
    (999.96, "0.0E+00", "1.0E+03"), (3, "@", "3"), (1234567, "General", "1234567"),
])
def test_shown_in_en_us(number, code, text):
    assert shown_text(number, code, EN_US) == text


@pytest.mark.parametrize("number, code", [
    (-1, "0"), (0, "0;-0;\"zero\""),                 # no negative or zero section measured
    (0.125, "0.00"), (46026.5, "0"), (2.675, "0.00"),  # halves the roundings disagree on
    (12345678, "General"), (0.1 + 0.2, "General"),     # wider than General was measured
    (0.5, "#.00"), (5, "0.##"), (5, "0E-0"), (12345678901234567, "0"),
    (5, None), (5, "yyyy-mm-dd"), (True, "0"), (float("nan"), "0"),
])
def test_not_measured(number, code):
    assert shown_text(number, code, EN_US) is None


def test_general_width_is_the_measured_text():
    assert GENERAL_WIDTH == len("46027,5")
