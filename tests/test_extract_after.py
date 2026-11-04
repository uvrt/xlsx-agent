"""A declared extract finds its label anywhere in a cell (round 32, R9): a status line far
from the cell's end, ``after`` (labels found by a plain string search, the value pattern
applied from there), a pattern's own label found the same way, every cost bounded, and a
cell that holds the label but no value the pattern finds counted and warned about.  Every
workbook is made up here: a log of notes, each ending in a ``Review state:`` line and a
justification of a few hundred characters."""

from __future__ import annotations

import time
import warnings

import pytest
from build import WorkbookBuilder
from ooxml_edit.tools import Toolbox

from xlsx_agent import CleanedWarning, DeclarationError, Workbook, table_dataset
from xlsx_agent.clean import (EXTRACT_REACH, EXTRACT_WINDOW, MAX_LABEL, MAX_LABELS, UNMATCHED,
                              Labels, parse_clean, pattern_label)
from xlsx_agent.tools import GROUPS, TOOLS, format as xlsx_format

#: A pattern written for the whole cell, label and all, in three languages' words.
WHOLE = r"(?i)(?:Review state|Stand der Dinge|Etat du suivi)\s*[:：]\s*(\S[^-–—,.;:(\n]*)"

#: The same as ``after`` and a value pattern from the label's end.
AFTER = {"after": ["Review state", "Stand der Dinge", "Etat du suivi"],
         "extract": r"^\s*[:：]\s*(\S[^-–—,.;:(\n]*)"}


def note(state: str, because: int, label: str = "Review state") -> str:
    """A note: a first line, then the label, the state and a justification ``because``
    characters long."""
    return f"Looked at in week 12.\n{label}: {state}, because " + "x" * because


#: The justifications: the first two leave the label within the last 256 characters.
LENGTHS = [10, 200, 250, 260, 300, 350]


def log(rows: list[list] | None = None) -> WorkbookBuilder:
    book = WorkbookBuilder()
    data = rows or [[f"N-{n}", note("Mostly done", length)] for n, length in enumerate(LENGTHS)]
    book.sheet("Log").rows("A1", [["Ref", "Notes"]] + data)
    return book


def table(clean: dict, book: WorkbookBuilder | None = None) -> dict:
    return {"name": "log", "sheets": "Log", "key": "Ref",
            "fields": [{"header": "Notes", "id": "state", "clean": clean}]}


def opened(clean: dict, book: WorkbookBuilder | None = None) -> Workbook:
    wb = Workbook.open((book or log()).to_bytes())
    wb.declare_tables([table(clean)])
    return wb


def states(wb: Workbook) -> list:
    return [row[0].v for row in wb.read("t:log[state]")]


# -- the label anywhere ------------------------------------------------------------------------


def test_the_label_is_found_however_far_from_the_cells_end():
    """The report's case: a label line starting 257 to 351 characters from the end read
    blank; a pattern that begins with its label now finds it anywhere."""
    for length in LENGTHS:
        text = note("Mostly done", length)
        assert len(text) - text.index("Review") > 40
    wb = opened({"extract": WHOLE})
    assert states(wb) == ["Mostly done"] * len(LENGTHS)
    assert opened(AFTER) and states(opened(AFTER)) == ["Mostly done"] * len(LENGTHS)


def test_a_match_in_the_window_reads_as_it_did():
    rule = parse_clean({"extract": r"Status:\s*(\S.*)$"})
    assert rule.apply("Status: old\nmore\nStatus: new") == ("new", True)      # the window's
    assert rule.apply("Status: first " + "x" * 300 + "\nStatus: last") == ("last", True)
    far = "Status: open\n" + "x" * EXTRACT_WINDOW          # $ is the text's end: no match
    assert rule.read(far) == (None, True, True)
    assert rule.read("no label") == (None, True, False)


def test_the_labels_last_occurrence_is_the_one_tried():
    rule = parse_clean({"extract": WHOLE})
    text = "Review state: Early, " + "x" * 400 + "\nReview state: Late, " + "y" * 400
    assert rule.apply(text) == ("Late", True)
    after = parse_clean(AFTER)
    assert after.apply(text) == ("Late", True)
    mixed = "Etat du suivi: Ouvert, " + "x" * 400 + "\nReview state: Closed. " + "y" * 300
    assert after.apply(mixed) == ("Closed", True)


def test_a_pattern_anchored_at_the_start_is_not_tried_elsewhere():
    assert pattern_label(parse_clean({"extract": r"^Status: (\w+)"}).extract) is None
    assert pattern_label(parse_clean({"extract": r"\AStatus: (\w+)"}).extract) is None
    multi = pattern_label(parse_clean({"extract": r"(?m)^Status: (.*)$"}).extract)
    assert multi is not None and multi.texts == ("Status: ",)
    rule = parse_clean({"extract": r"(?m)^Status: (\w+)"})
    assert rule.apply("a\nStatus: open\n" + "x" * 400) == ("open", True)


@pytest.mark.parametrize("pattern, texts, folds", [
    (r"Status:\s*(\S.*)$", ("Status:",), False),
    (WHOLE, ("Review state", "Stand der Dinge", "Etat du suivi"), True),
    (r"(?:Review|Revision) state: (\w+)", ("Review state: ", "Revision state: "), False),
    (r"\bState(?:s|d): (\w+)", ("State",), False),
    (r"(?i:Next) step (\w+)", ("Next step ",), True),
])
def test_a_patterns_own_label_is_read_from_its_literal_start(pattern, texts, folds):
    found = pattern_label(parse_clean({"extract": pattern}).extract)
    assert found is not None and set(found.texts) == set(texts) and found.folds is folds


@pytest.mark.parametrize("pattern", [r"(\d+)", r"x(y)?", r"\s*Status: (\w+)",
                                     r"(?:ab|Review) (\w+)", r"[:：]\s*(\w+)", r"No(\w+)"])
def test_a_pattern_without_a_label_of_three_characters_has_none(pattern):
    """A digit, a class or a short text first is no label: the window alone, as before."""
    found = pattern_label(parse_clean({"extract": pattern}).extract)
    assert found is None or all(len(t) >= 3 for t in found.texts)
    if pattern in (r"(\d+)", r"x(y)?", r"\s*Status: (\w+)", r"[:：]\s*(\w+)", r"No(\w+)"):
        assert found is None


# -- after -------------------------------------------------------------------------------------


def test_after_reads_from_the_labels_end_as_if_the_cell_began_there():
    rule = parse_clean({"after": "Review state", "extract": r"^\s*:\s*(\S[^,]*)"})
    assert rule.spec == {"extract": r"^\s*:\s*(\S[^,]*)", "after": ["Review state"]}
    assert rule.apply("x " * 300 + "REVIEW STATE: Done, x") == ("Done", True)   # any case
    assert rule.apply("Review state: Done") == ("Done", True)
    assert rule.read("No label at all: Done") == (None, True, False)          # blank, no issue
    assert rule.read("Review state - Done") == (None, True, True)             # the label, no value


def test_after_takes_a_pattern_written_with_the_label_too():
    """Rule 1, intent: the pattern the table had for the whole cell, label and all, works
    with after as it is."""
    rule = parse_clean({"after": ["Review state"], "extract": WHOLE})
    assert rule.apply("x" * 1000 + "\nReview state: Partly, y") == ("Partly", True)


def test_after_a_value_running_past_the_reach_is_read_whole():
    """A match reaching EXTRACT_REACH characters is tried once more on the whole text: a
    $ there is no end, and a long value is not cut."""
    rule = parse_clean({"after": "Review state", "extract": r"^:\s*(\S.*)$"})
    long_value = "v" * (EXTRACT_REACH * 2)
    assert rule.apply("Review state: " + long_value) == (long_value, True)
    assert rule.read("Review state: " + long_value + "\nmore") == (None, True, True)
    comma = parse_clean({"after": "Review state", "extract": r"^:\s*(\S[^,]*),"})
    assert comma.apply("Review state: " + long_value + ", rest") == (long_value, True)


def test_after_a_value_starts_within_the_window_after_the_label():
    rule = parse_clean({"after": "Review state", "extract": r"(\d+)"})
    assert rule.apply("Review state" + " " * (EXTRACT_WINDOW - 1) + "42") == ("42", True)
    assert rule.read("Review state" + " " * (EXTRACT_WINDOW + 20) + "42") == (None, True, True)


@pytest.mark.parametrize("clean, message", [
    ({"after": "Review state"}, "add the value's pattern"),
    ({"after": "", "extract": r"(\w+)"}, "not a pattern"),
    ({"after": [], "extract": r"(\w+)"}, "a list of 1 to"),
    ({"after": ["ok", 3], "extract": r"(\w+)"}, "a list of 1 to"),
    ({"after": ["x" * (MAX_LABEL + 1)], "extract": r"(\w+)"}, "a list of 1 to"),
    ({"after": [f"L{n}" for n in range(MAX_LABELS + 1)], "extract": r"(\w+)"}, "a list of 1 to"),
])
def test_after_refusals_say_what_would_work(clean, message):
    """Rule 2: the refusal names the rule, its shape and a form that works."""
    with pytest.raises(DeclarationError) as caught:
        opened(clean)
    assert caught.value.code == "bad_declaration"
    assert caught.value.path.endswith("clean.after")
    assert message in str(caught.value)
    assert "Review state" in str(caught.value)      # the example


def test_labels_found_in_any_case_and_the_longest_where_two_end_alike():
    labels = Labels(("state", "Review state"))
    assert labels.texts == ("state", "Review state")                   # as declared
    assert labels.last("a REVIEW State: x") == (2, 14)
    assert labels.last("Straße Review state ß") == (7, 19)             # ß: lower-cased
    assert Labels(("İx",)).last("aİx İX") == (4, 6)                    # neither keeps length
    assert Labels(("Status",), folds=False).last("status Status") == (7, 13)
    assert Labels(("Status",), folds=False).last("status") is None


# -- a label without a value -------------------------------------------------------------------


def unmatched_book() -> WorkbookBuilder:
    return log([["N-1", note("Mostly done", 300)],
                ["N-2", "Review state - Open, the value after a dash"],
                ["N-3", "No label in this one"],
                ["N-4", "Intro " + "z" * 400 + "\nReview state: "]])


def test_a_cell_with_the_label_but_no_value_is_blank_with_an_issue():
    wb = opened({"extract": WHOLE}, unmatched_book())
    got = [(v.a, v.v, v.issue) for (v,) in wb.read("t:log[state]")]
    assert got == [("Log!B2", "Mostly done", None), ("Log!B3", None, UNMATCHED),
                   ("Log!B4", None, None), ("Log!B5", None, UNMATCHED)]
    after = opened(AFTER, unmatched_book())
    assert [v.issue for (v,) in after.read("t:log[state]")] == [None, UNMATCHED, None,
                                                                 UNMATCHED]


def test_profile_and_describe_count_them_with_the_first_cell():
    """Rule 7: what reads wrong is said where the agent looks first."""
    wb = opened(AFTER, unmatched_book())
    column = next(c for c in wb.profile("t:log") if c.column == "state")
    assert column.flags[UNMATCHED].as_dict() == {"count": 2, "first": ["Log!B3", "Log!B5"]}
    catalogue = wb.catalogue()
    region = next(r for r in catalogue["regions"] if r["address"] == "t:log")
    state = next(c for c in region["columns"] if c["name"] == "state")
    assert state["flags"][UNMATCHED] == {"count": 2, "first": "Log!B3"}


def test_derive_counts_them_per_column():
    wb = opened(AFTER, unmatched_book())
    derived = wb.derive("t:log", group_by=["state"])
    assert derived.unmatched["state"].as_dict() == {"count": 2, "first": ["Log!B3", "Log!B5"]}
    assert derived.as_dict()["unmatched"] == {"state": {"count": 2,
                                                        "first": ["Log!B3", "Log!B5"]}}
    measured = wb.derive("t:log", measures=[{"fn": "count", "column": "state", "as": "n"}])
    assert measured.unmatched["state"].count == 2
    assert not opened(AFTER).derive("t:log", group_by=["state"]).unmatched


def test_a_deck_or_document_table_from_them_warns():
    wb = opened(AFTER, unmatched_book())
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        found = table_dataset(wb, "t:log[state]")
    assert found.detail[UNMATCHED].as_dict() == {"count": 2, "first": ["Log!B3", "Log!B5"]}
    messages = [str(w.message) for w in caught if w.category is CleanedWarning]
    assert any("2 cells hold the label" in m and "Log!B3" in m for m in messages), messages
    assert not any("Open" in m for m in messages)          # never the cell's text


@pytest.fixture(scope="module")
def tools():
    toolbox = Toolbox(TOOLS, formats=[xlsx_format(tables=[table(AFTER)])], groups=GROUPS)
    session = toolbox.session()
    blob = session.add_blob(unmatched_book().to_bytes(), name="log.xlsx")
    assert toolbox.dispatch(session, "open_document", {"blob": blob}).ok
    return toolbox, session


def call(tools, name, **arguments):
    toolbox, session = tools
    result = toolbox.dispatch(session, name, arguments)
    assert result.ok, result.error
    return result


def test_the_tools_count_and_warn(tools):
    read = call(tools, "xlsx_read", doc="d1", target="t:log")
    assert read.data[UNMATCHED] == {"count": 2, "first": ["Log!B3", "Log!B5"]}
    assert any("2 cells hold the label" in w and "first Log!B3" in w
               and "cells_only" in w for w in read.warnings), read.warnings
    query = call(tools, "xlsx_query", doc="d1", table="t:log", where=[])
    assert query.data[UNMATCHED]["count"] == 2 and query.warnings
    derived = call(tools, "xlsx_derive", doc="d1", **{"from": "t:log", "group_by": ["state"],
                                            "id": "by_state"})
    assert derived.data[UNMATCHED] == {"state": {"count": 2, "first": ["Log!B3", "Log!B5"]}}
    assert any("of column state" in w for w in derived.warnings), derived.warnings
    described = call(tools, "describe", doc="d1")
    assert '"extract_unmatched":{"count":2,"first":"Log!B3"}' in described.data["text"] \
        if "text" in described.data else UNMATCHED in str(described.data)
    stored = call(tools, "xlsx_read", doc="d1", target="Log!B3", cells_only=True)
    assert stored.data["rows"][0][0]["v"].startswith("Review state - Open")
    assert not stored.warnings


def test_after_changes_no_tool_definition():
    plain = Toolbox(TOOLS, groups=GROUPS).definitions("anthropic", defer=False)
    declared = Toolbox(TOOLS, formats=[xlsx_format(tables=[table(AFTER)])],
                       groups=GROUPS).definitions("anthropic", defer=False)
    assert plain == declared


# -- the cost ----------------------------------------------------------------------------------


@pytest.mark.parametrize("clean", [
    {"extract": WHOLE},
    AFTER,
    {"after": "Review state", "extract": r"([^;]*);"},
    {"after": "Review state", "extract": r"( *)\S"},
    {"after": "Review state", "extract": r"^:\s*(\S.*)$"},
])
def test_the_cost_stays_bounded_on_long_cells(clean):
    """Cells of 100,000 characters (past Excel's 32,767, as another writer may store),
    the label at the start, nowhere, many times, or before the worst text for each rule:
    each well under the bound clean.py reports (measured about 10 ms at worst)."""
    rule = parse_clean(clean)
    n = 100_000
    for text in ("Review state: done, " + "word " * (n // 5), "word " * (n // 5),
                 "Review state: " * (n // 14), "Review state" + "a" * n,
                 "Review state" + " " * n, "x" * (n - 300) + "Review state: done, " + "y" * 280):
        rule.read(text)
        start = time.perf_counter()
        for _ in range(3):
            rule.read(text)
        assert (time.perf_counter() - start) / 3 < 0.1, (clean, text[:20])
