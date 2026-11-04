"""Round 31, R7b: tables keyed by dotted numbers (``4.1``, ``4.1.2``, ``A.5.1``) -- found as
the key, addressed by it in every tool that takes ``t:Name/<key>``, and followed in text by
``xlsx_trace`` -- and rows by position (``#4``, ``#1:#8``) of a table and of a dataset.

Every workbook is made up here: a table of clauses of a made-up handbook, numbered, with a
title and a note that may name other clauses; a table of parts beside it whose notes hold
quantities.
"""

from __future__ import annotations

import json

import pytest
from build import WorkbookBuilder
from ooxml_edit.tools import Toolbox

from xlsx_agent import AddressError, NameNotFound, RegionError, Workbook
from xlsx_agent.address import parse_address
from xlsx_agent.read.regions import is_dotted_key
from xlsx_agent.tools import GROUPS, TOOLS, FORMAT
from xlsx_agent.trace import _Finder

CLAUSES = [["Clause", "Title", "Note"],
           ["4", "Scope", "the whole of 4.1 to 4.3"],
           ["4.1", "Context", "see 4.2"],
           ["4.2", "Parties", "rests on 4.1.2; a bolt weighs 4.3 kg"],
           ["4.1.2", "Limits", "4.10 is not a clause; nor is 1.4.1"],
           ["4.3", "Boundaries", None],
           ["A.5.1", "Annex rule", "as 4.2 says."]]


def book(clauses=CLAUSES, *, parts: bool = True) -> bytes:
    out = WorkbookBuilder()
    sheet = out.sheet("Handbook")
    sheet.rows("A1", clauses)
    sheet.table("Clauses", f"A1:{'ABC'[len(clauses[0]) - 1]}{len(clauses)}", clauses[0])
    if parts:
        other = out.sheet("Parts")
        other.rows("A1", [["Part", "Note"], ["P-1", "4.2 kg, 4.1% off, -4.3, $4.2"],
                          ["P-2", "per 4.2 and 4.3."]])
        other.table("Parts", "A1:B3", ["Part", "Note"])
    return out.to_bytes()


@pytest.fixture
def wb() -> Workbook:
    return Workbook.open(book())


# -- what a dotted key is, and the key found ----------------------------------------------


@pytest.mark.parametrize("key, dotted", [
    ("4.1", True), ("4.1.2", True), ("A.5.1", True), ("4", True), ("AB.1", True),
    ("4.", False), (".4", False), ("4..1", False), ("4.1a", False), ("I-1", False),
    ("ABC.1", False), ("4 1", False)])
def test_a_dotted_key_is_digits_joined_by_dots_after_at_most_two_letters(key, dotted):
    assert is_dotted_key(key) is dotted


def test_a_column_of_dotted_numbers_beside_words_is_the_key(wb):
    assert wb.key_columns("t:Clauses") == ["Clause"]
    assert wb.resolve("t:Clauses/4.1.2/Title").range.text == "Handbook!B5"
    assert [[v.v for v in row] for row in wb.read("t:Clauses/A.5.1/Title")] == [["Annex rule"]]
    # a whole number of the column is a key too: 4 is the row holding "4"
    assert wb.resolve("t:Clauses/4").row == 1
    # a key set of dotted keys
    assert wb.resolve("t:Clauses/{4.3,4.1}").positions == (5, 2)


def test_beside_codes_or_whole_numbers_it_is_refused_with_the_way_out():
    codes = [["Clause", "Ref", "Title"], ["4.1", "R-1", "Context"], ["4.2", "R-2", "Parties"]]
    opened = Workbook.open(book(codes, parts=False))
    assert opened.key_columns("t:Clauses") == ["Clause", "Ref", "Title"]
    with pytest.raises(RegionError) as refused:
        opened.resolve("t:Clauses/4.1")
    error = refused.value
    assert error.code == "no_key_column"
    assert error.valid_options == ["Clause", "Ref", "Title"]
    # the way out: how to name the key, the rows there are, addresses by position
    assert "key=" in error.message and "Workbook.declare_key" in error.message
    assert "2 data rows: t:Clauses/#1 to t:Clauses/#2 (t:Clauses/#1:#2)" \
        in error.message
    assert opened.resolve("t:Clauses/4.1", key="Clause").row == 1
    numbered = [["No", "Clause", "Title"], [1, "4.1", "Context"], [2, "4.2", "Parties"]]
    opened = Workbook.open(book(numbered, parts=False))
    assert opened.key_columns("t:Clauses") == ["No", "Clause", "Title"]
    # two columns of dotted numbers: neither
    twice = [["Clause", "Was", "Title"], ["4.1", "3.1", "Context"], ["4.2", "3.2", "Parties"]]
    assert Workbook.open(book(twice, parts=False)).key_columns("t:Clauses") == [
        "Clause", "Was", "Title"]
    # a declared key settles it
    opened.declare_key("t:Clauses", "Clause")
    assert opened.resolve("t:Clauses/4.2").row == 2


def test_a_column_of_numbers_with_fractions_is_no_key():
    # 4.1 typed as a number is the number 4.1 (4.10 would be it too): no key
    numbers = [["Clause", "Title"], [4.1, "Context"], [4.2, "Parties"]]
    opened = Workbook.open(book(numbers, parts=False))
    assert opened.key_columns("t:Clauses") == ["Title"]


def test_a_key_not_found_says_the_key_column_and_the_rows_by_position(wb):
    with pytest.raises(NameNotFound) as missing:
        wb.resolve("t:Clauses/9.9")
    message = missing.value.message
    assert "column 'Clause'" in message
    assert "6 data rows: t:Clauses/#1 to t:Clauses/#6" in message
    assert "9.9" not in message          # a key is cell content: never in a message


def test_a_row_by_position_without_its_table_says_how():
    with pytest.raises(AddressError) as refused:
        parse_address("#4")
    assert "t:Name/#4" in refused.value.message
    assert refused.value.valid_options == ["t:Products/#5", "t:Products/#9:#16",
                                           "d:by_category/#2"]
    with pytest.raises(AddressError) as refused:
        parse_address("#1:#8")
    assert "t:Name/#1:#8" in refused.value.message


# -- found in text ------------------------------------------------------------------------


@pytest.mark.parametrize("text, found", [
    ("see 4.2", ["4.2"]),
    ("as 4.2 says.", ["4.2"]),                # a full stop ends the sentence, not the key
    ("4.1-4.3", ["4.1", "4.3"]),              # a range of clauses
    ("4.1/4.2, 4.3", ["4.1", "4.2", "4.3"]),
    ("(4.1)", ["4.1"]),
    ("4.2 s'applique", ["4.2"]),              # an apostrophe after: a word, not a unit
    ("rests on 4.1.2", ["4.1.2"]),            # the longer key
    ("4.1.3", []),                            # inside a longer dotted number
    ("1.4.1", []),
    ("12.4.2026", []),
    ("4.10", []),
    ("4.2 kg", []), ("4.2kg", []), ("4.2%", []), ("4.2 GB", []),   # quantities
    ("-4.2", []), ("$4.2", []), ("€ 4.2", []), ("= 4.2", []), ("< 4.2", []),
    ("4.2 days", []), ("4.2 heures", []), ("4.2 million", []), ("4.2/5", []),
    ("4.2 out of 5", []), ("4.2 sur 5", []), ("4.2 sur la base", ["4.2"]),
    ("4.1 m'a dit", ["4.1"]),
    ("see A.5.1.", ["A.5.1"]),
    ("A.5.1.2", []),
    ("A.5.1 kg", ["A.5.1"]),                 # with a letter it is no quantity
])
def test_a_dotted_key_is_found_whole_and_never_as_a_quantity(text, found):
    finder = _Finder(["4.1", "4.2", "4.3", "4.1.2", "A.5.1"])
    assert finder.keys_in(text) == found


def test_trace_follows_dotted_keys_and_never_whole_numbers_or_quantities(wb):
    (trace,) = wb.trace("t:Clauses/4.2", depth=1).traces
    assert [(link.row, link.why, link.column) for link in trace.links] == [
        ("t:Clauses/4.1.2", "names", "Note"),       # "4.3 kg" is a weight
        ("t:Clauses/4.1", "lists", "Note"),
        ("t:Clauses/A.5.1", "lists", "Note"),       # "as 4.2 says."
        ("t:Parts/#2", "lists", "Note")]            # "4.2 kg" in t:Parts/#1 is a weight
    # the whole number 4 is never looked for; "4.1 to 4.3" names 4.1 and 4.3
    (scope,) = wb.trace("t:Clauses/4", depth=1).traces
    assert [link.row for link in scope.links if link.why == "names"] == [
        "t:Clauses/4.1", "t:Clauses/4.3"]
    assert not [link for link in scope.links if link.why == "lists"]
    # the part whose note holds only quantities names none
    (part,) = wb.trace("t:Parts/#1", depth=1).traces
    assert part.links == []


def test_dotted_keys_of_a_table_not_keyed_by_them_are_never_looked_for():
    # a key column of codes: its other columns' dotted numbers are no keys
    codes = [["Ref", "Version", "Note"], ["R-1", "4.1", "see 4.2"], ["R-2", "4.2", "R-1"]]
    opened = Workbook.open(book(codes, parts=False))
    opened.declare_key("t:Clauses", "Ref")
    (trace,) = opened.trace("t:Clauses/R-1", depth=1).traces
    assert [(link.row, link.why) for link in trace.links] == [("t:Clauses/R-2", "lists")]


# -- every tool, and dataset rows by position ----------------------------------------------


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _session(box):
    session = box.session()
    session.open(book(), "handbook.xlsx")
    return session


def test_every_tool_taking_a_row_by_key_takes_a_dotted_key(box):
    session = _session(box)
    read = box.dispatch(session, "xlsx_read", {"doc": "d1", "target": "t:Clauses/4.1.2/Title"})
    assert read.ok, read.to_json()
    assert "Limits" in json.dumps(read.to_json())
    many = box.dispatch(session, "xlsx_read", {"doc": "d1", "target": "t:Clauses/{4.1,A.5.1}"})
    assert many.ok and many.total == 2
    traced = box.dispatch(session, "xlsx_trace", {"doc": "d1", "rows": ["t:Clauses/4.1"],
                                                  "depth": 1})
    assert traced.ok, traced.to_json()
    assert traced.data["traced"][0]["links"][0] == {
        "row": "t:Clauses/4.2", "why": "names", "column": "Note", "depth": 1}
    found = box.dispatch(session, "find_text", {"doc": "d1", "text": "Context",
                                                "range": "t:Clauses/4.1"})
    assert found.ok, found.to_json()
    assert "Handbook!B3" in json.dumps(found.to_json())
    described = box.dispatch(session, "describe", {"doc": "d1"})
    assert '"Clause"' in json.dumps(described.to_json())
    # a write by key needs the key declared, as for any key (xlsx_agent.edit.keys)
    refused = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "t:Clauses/4.3/Note", "text": "see 4.1"}]})
    assert refused.error.code == "refused"
    session.entry("d1").document.declare_key("t:Clauses", "Clause")
    edited = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "t:Clauses/4.3/Note", "text": "see 4.1"}]})
    assert edited.ok, edited.to_json()
    again = box.dispatch(session, "xlsx_read", {"doc": "d1", "target": "t:Clauses/4.3/Note"})
    assert "see 4.1" in json.dumps(again.to_json())


def test_dataset_rows_by_position_and_range(box):
    session = _session(box)
    made = box.dispatch(session, "xlsx_derive", {
        "doc": "d1", "from": "t:Clauses", "id": "clauses", "group_by": ["Clause", "Title"]})
    assert made.ok, made.to_json()
    one = box.dispatch(session, "xlsx_read", {"target": "d:clauses/#4"})
    assert one.ok and one.total == 1 and "Limits" in json.dumps(one.to_json())
    some = box.dispatch(session, "xlsx_read", {"target": "d:clauses/#1:#3"})
    assert some.ok and some.total == 3
    traced = box.dispatch(session, "xlsx_trace", {"doc": "d1", "rows": ["d:clauses/#2:#3"]})
    assert traced.ok, traced.to_json()
    assert [t["row"] for t in traced.data["traced"]] == ["d:clauses/#2", "d:clauses/#3"]
    assert traced.data["traced"][0]["links"][0]["why"] == "source"


def test_compare_pairs_rows_by_a_dotted_key(box):
    session = box.session()
    session.open(book(), "old.xlsx")
    changed = [row[:] for row in CLAUSES]
    changed[3][1] = "Interested parties"
    session.open(book([changed[0], changed[3], *changed[1:3], *changed[4:]]), "new.xlsx")
    result = box.dispatch(session, "xlsx_compare", {"a": "d1", "b": "d2",
                                                    "tables": ["t:Clauses"]})
    assert result.ok, result.to_json()
    text = json.dumps(result.to_json())
    assert '"key": "Clause"' in text and "Interested parties" in text
