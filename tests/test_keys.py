"""Addresses by row key (X2): selectors (``t:checks/*/rating``), key sets
(``t:checks/{A1,B4}/rating``), :meth:`Workbook.expand`, derived sources written by key, and
what stays true when a consultant inserts or reorders rows.  Every workbook is made up
here: a table of checks keyed by code, with a rating and a note."""

from __future__ import annotations

from datetime import date
from fractions import Fraction

import pytest
from build import WorkbookBuilder
from ooxml_edit.tools import Toolbox

from xlsx_agent import (AddressError, DatasetTable, DeriveError, NameNotFound, RegionError,
                        Workbook)
from xlsx_agent.address import TableAddress, parse_address
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS

HEADER = ["code", "area", "rating", "note"]

#: The checks, in the order the first version of the workbook holds them.
CHECKS = [
    ("A1", "north", 3.5, "door"), ("A2", "north", 4, "door"), ("B1", "south", 2.5, "roof"),
    ("A3", "north", 5, "wall"), ("B2", "south", 1.5, "roof"), ("C1", "east", 3, "wall"),
    ("B3", "south", 4.5, "door"), ("C2", "east", 2, "roof"),
]

#: Two checks a consultant adds later.
ADDED = [("D1", "west", 1, "gate"), ("A4", "north", 2.5, "door")]


def book(rows, *, name: str = "checks", at: str = "A1") -> bytes:
    built = WorkbookBuilder()
    sheet = built.sheet("Log")
    sheet.rows(at, [HEADER] + [list(row) for row in rows])
    top, left = int(at[1:]), at[0]
    right = chr(ord(left) + len(HEADER) - 1)
    sheet.table(name, f"{at}:{right}{top + len(rows)}", HEADER, write_header=False)
    return built.to_bytes()


def edited() -> list:
    """The checks after a consultant's edit: two rows inserted (one at the top, one in the
    middle), and the rest in another order."""
    rows = list(reversed(CHECKS))
    return [ADDED[0]] + rows[:4] + [ADDED[1]] + rows[4:]


@pytest.fixture(scope="module")
def wb() -> Workbook:
    return Workbook.open(book(CHECKS))


# -- the forms -------------------------------------------------------------------------------


@pytest.mark.parametrize("text, kind", [
    ("t:checks/*/rating", "column"),
    ("t:checks/*", "table"),
    ("t:checks/*/*", "table"),
    ("t:checks/A1/*", "row"),
    ("t:checks/#2/*", "row"),
    ("t:checks/{A1,B2}/rating", "values"),
    ("t:checks/{A1,B2}", "rows"),
    ("t:checks/{A1,B2}/*", "rows"),
    ("t:checks/{'a,b','#3','*','x/y',B2}/rating", "values"),
    ("t:checks/'*'/rating", "value"),
    ("t:checks/'{x'/rating", "value"),
    ("t:checks/A1/'*'", "value"),
    ("t:Log!A1/*/rating", "column"),
    ("t:Log!Prices/{A1,B2}/rating", "values"),
    ("d:by_area/*/mean", "column"),
    ("d:by_area/#2/*", "row"),
])
def test_selectors_and_key_sets_format_back_to_themselves(text, kind):
    address = parse_address(text)
    assert (address.text, address.kind) == (text, kind)
    assert parse_address(address.text) == address


@pytest.mark.parametrize("text, canonical", [
    ("t:checks/{A1}/rating", "t:checks/A1/rating"),          # a set of one key is that key
    ("t:checks/{A1,B2,A1}/rating", "t:checks/{A1,B2}/rating"),  # a key listed once
    ("t:checks/{'A1',B2}", "t:checks/{A1,B2}"),
])
def test_a_key_set_is_normalised(text, canonical):
    assert parse_address(text).text == canonical


@pytest.mark.parametrize("text, words", [
    ("t:checks/{}/rating", "lists keys between { and }"),
    ("t:checks/{A1,*}", "* is every row"),
    ("t:checks/{#3,A1}", "not positions"),
    ("t:checks/{A1,B2", "no closing }"),
    ("t:checks/{A1,B2}x", "followed by / and a column"),
    ("t:checks/{A/1,B2}", "is quoted"),
    ("t:checks/*/*/x", "at most a row and a column"),
])
def test_a_wrong_key_set_or_selector_is_refused_without_its_keys(text, words):
    with pytest.raises(AddressError) as caught:
        parse_address(text)
    assert words in str(caught.value) and "A1" not in caught.value.message.replace(
        "{A1,B2}", "")


def test_a_key_that_is_a_wildcard_or_a_brace_is_quoted():
    assert TableAddress(name="checks", key="*", column="rating").text == "t:checks/'*'/rating"
    assert TableAddress(name="checks", key="{x", column="*").text == "t:checks/'{x'/'*'"
    assert TableAddress(name="checks", keys=("a,b", "c")).text == "t:checks/{'a,b',c}"
    with pytest.raises(ValueError):
        TableAddress(name="checks", keys=("a",))                 # one key is a key
    with pytest.raises(ValueError):
        TableAddress(name="checks", key="a", any_row=True)        # one row form
    with pytest.raises(ValueError):
        TableAddress(name="checks", any_column=True)              # * follows a row


# -- resolving and reading -------------------------------------------------------------------


def _values(rows) -> list:
    return [[v.v for v in row] for row in rows]


def test_a_selector_reads_as_the_cells_it_covers_and_keeps_its_text(wb):
    assert _values(wb.read("t:checks/*/rating")) == _values(wb.read("t:checks[rating]"))
    assert _values(wb.read("t:checks/*")) == _values(wb.read("t:checks/*/*")) \
        == _values(wb.read("t:checks"))
    assert _values(wb.read("t:checks/B1/*")) == [["B1", "south", 2.5, "roof"]]
    assert wb.resolve("t:CHECKS/*/Rating").address == "t:checks/*/rating"
    assert wb.resolve("t:checks/b1/*".replace("b1", "B1")).kind == "row"


def test_a_key_set_reads_its_rows_in_its_order(wb):
    selection = wb.resolve("t:checks/{C2,A1,B3}/rating")
    assert (selection.kind, selection.positions, selection.range) == ("values", (8, 1, 7), None)
    assert [a.text for a in selection.areas] == ["Log!C9", "Log!C2", "Log!C8"]
    assert _values(wb.read("t:checks/{C2,A1,B3}/rating")) == [[2], [3.5], [4.5]]
    assert _values(wb.read("t:checks/{C2,A1}")) == [["C2", "east", 2, "roof"],
                                                   ["A1", "north", 3.5, "door"]]


def test_one_unknown_key_refuses_the_whole_set_naming_its_place_not_its_key(wb):
    with pytest.raises(NameNotFound) as caught:
        wb.resolve("t:checks/{A1,Z9}/rating")
    assert "key 2 of the set" in str(caught.value) and "Z9" not in str(caught.value)


# -- expand: the coverage selector -----------------------------------------------------------


def test_expand_gives_each_cell_by_key_in_the_tables_order(wb):
    assert wb.expand("t:checks/*/rating") == [f"t:checks/{c[0]}/rating" for c in CHECKS]
    assert wb.expand("t:checks[rating]") == wb.expand("t:checks/*/rating")
    assert wb.expand("t:checks/*") == wb.expand("t:checks") == [f"t:checks/{c[0]}"
                                                                for c in CHECKS]
    assert wb.expand("t:checks/A2/*") == [f"t:checks/A2/{h}" for h in HEADER]
    assert len(wb.expand("t:checks/*/*")) == len(CHECKS) * len(HEADER)
    assert wb.expand("t:checks/{C2,A1}/rating") == ["t:checks/C2/rating",
                                                    "t:checks/A1/rating"]
    assert wb.expand("t:checks/#3/rating") == ["t:checks/B1/rating"]   # a position, by key
    assert wb.expand("t:checks/#3") == ["t:checks/B1"]


def test_expand_refuses_rows_it_cannot_name_alone():
    rows = [("A1", 1.5), ("A2", 2.5), (None, 3.5), ("A2", 4.5), (date(2026, 1, 5), 5.5)]
    built = WorkbookBuilder()
    sheet = built.sheet("Log")
    sheet.rows("A1", [["code", "rating"]] + [list(r) for r in rows])
    sheet.table("checks", "A1:B6", ["code", "rating"], write_header=False)
    workbook = Workbook.open(built.to_bytes())
    with pytest.raises(RegionError) as no_column:
        workbook.expand("t:checks/*/rating")              # no column is a key column
    assert no_column.value.code == "no_key_column"
    workbook.declare_key("t:checks", "code")
    with pytest.raises(RegionError) as blank:
        workbook.expand("t:checks/*/rating")
    assert blank.value.code == "no_key"
    assert blank.value.valid_options == ["t:checks/#3", "t:checks/#5"]   # blank, a date
    with pytest.raises(RegionError) as shared:
        workbook.expand("t:checks/{A1,A2}/rating")
    assert shared.value.code == "duplicate_key"
    assert shared.value.valid_options == ["t:checks/#2", "t:checks/#4"]
    assert workbook.expand("t:checks/A1/rating") == ["t:checks/A1/rating"]
    with pytest.raises(RegionError) as positioned:
        workbook.expand("t:checks/#3/rating")
    assert positioned.value.code == "no_key"


@pytest.mark.parametrize("cells, key, found", [
    (["007", 7], "007", 1),        # text is its text: 007 is not the number 7
    (["007", 7], "7", 2),          # a whole number is its digits
    ([7.0, 8], "7", 1),
    (["2026-01-05", "x"], "2026-01-05", 1),   # text that looks like a date is text
])
def test_what_a_key_is(cells, key, found):
    built = WorkbookBuilder()
    sheet = built.sheet("Log")
    sheet.rows("A1", [["code", "rating"]] + [[c, n + 0.5] for n, c in enumerate(cells)])
    sheet.table("checks", "A1:B3", ["code", "rating"], write_header=False)
    workbook = Workbook.open(built.to_bytes())
    workbook.declare_key("t:checks", "code")
    assert workbook.resolve(f"t:checks/{key}/rating").row == found


@pytest.mark.parametrize("cells, code", [
    (["7", 7], "duplicate_key"),           # the text 7 and the number 7 are one key
    ([date(2026, 1, 5), "x"], None),       # a date is never matched by text
    ([7.5, "x"], None),                     # nor a fraction
])
def test_keys_that_look_alike_are_refused_never_guessed(cells, code):
    built = WorkbookBuilder()
    sheet = built.sheet("Log")
    sheet.rows("A1", [["code", "rating"]] + [[c, n + 0.5] for n, c in enumerate(cells)])
    sheet.table("checks", "A1:B3", ["code", "rating"], write_header=False)
    workbook = Workbook.open(built.to_bytes())
    assert workbook.key_columns("t:checks") == []          # no column is a key by itself
    workbook.declare_key("t:checks", "code")
    wanted = {"duplicate_key": "7", None: "2026-01-05"}.get(code, "7.5")
    if code is None:
        with pytest.raises(NameNotFound):
            workbook.resolve(f"t:checks/{wanted}/rating")
    else:
        with pytest.raises(RegionError) as caught:
            workbook.resolve(f"t:checks/{wanted}/rating")
        assert caught.value.code == code


def test_a_key_with_the_addresses_own_characters_is_quoted_and_found():
    keys = ["a/b", "x,y", "*", "{z", "#4", " pad", "it's"]
    built = WorkbookBuilder()
    sheet = built.sheet("Log")
    sheet.rows("A1", [["code", "rating"]] + [[k, n + 0.5] for n, k in enumerate(keys)])
    sheet.table("checks", f"A1:B{len(keys) + 1}", ["code", "rating"],
                             write_header=False)
    workbook = Workbook.open(built.to_bytes())
    every = workbook.expand("t:checks/*/rating")
    assert [workbook.cell(a).v for a in every] == [n + 0.5 for n in range(len(keys))]
    together = TableAddress(name="checks", keys=tuple(keys), column="rating").text
    assert [row[0].v for row in workbook.read(together)] == [n + 0.5 for n in range(len(keys))]
    assert workbook.expand(together) == every


# -- derived sources by key ------------------------------------------------------------------


def _area_means(workbook: Workbook):
    return workbook.derive("t:checks", group_by=["area"],
                           measures=[{"fn": "mean", "column": "rating", "as": "mean"},
                                     {"fn": "count"}], id="by_area")


def test_a_derived_values_sources_are_one_key_set_in_row_order(wb):
    ds = _area_means(wb)
    assert ds.source_key == "code"
    north = ds.provenance[0]
    assert [str(s) for s in north["mean"].sources] == ["t:checks/{A1,A2,A3}/rating"]
    assert [str(s) for s in north["area"].sources] == ["t:checks/{A1,A2,A3}/area"]
    assert [str(s) for s in north["count"].sources] == ["t:checks/{A1,A2,A3}/area"]
    assert ds.as_dict()["source_key"] == "code"
    data = ds.dataset()
    assert data.sources[0][1] == ("t:checks/{A1,A2,A3}/rating",)
    one = wb.derive("t:checks", where=[{"column": "code", "op": "eq", "value": "C1"}],
                    measures=[{"fn": "sum", "column": "rating"}])
    assert [str(s) for s in one.provenance[0]["sum rating"].sources] == ["t:checks/C1/rating"]


def test_without_a_key_column_the_sources_stay_ranges_and_a_declared_one_makes_them_keys():
    rows = [(f"K{r}", "north" if r % 2 else "south", r / 2, f"n{r}") for r in range(1, 7)]
    workbook = Workbook.open(book(rows))
    assert workbook.key_columns("t:checks") == ["code", "note"]     # two candidates
    ds = _area_means(workbook)
    assert ds.source_key is None
    assert [str(s) for s in ds.provenance[0]["mean"].sources] == ["Log!C2", "Log!C4",
                                                                  "Log!C6"]
    workbook.declare_key("t:checks", "code")
    keyed = _area_means(workbook)
    assert keyed.source_key == "code"
    assert [str(s) for s in keyed.provenance[0]["mean"].sources] == ["t:checks/{K1,K3,K5}/rating"]
    assert workbook.catalogue()["regions"][0]["key"] == "code"


def test_a_declared_key_column_with_a_blank_keeps_ranges():
    rows = [("K1", "north", 1, "a"), (None, "north", 2, "a"), ("K3", "south", 3, "a")]
    workbook = Workbook.open(book(rows))
    workbook.declare_key("t:checks", "code")
    ds = _area_means(workbook)
    assert ds.source_key is None
    assert [str(s) for s in ds.provenance[0]["mean"].sources] == ["Log!C2:C3"]


def test_a_dataset_derived_from_a_keyed_one_gathers_the_keys_again(wb):
    ds = wb.derive("t:checks", group_by=["area", "note"],
                   measures=[{"fn": "sum", "column": "rating", "as": "total"}])
    again = wb.derive(ds, group_by=["area"], measures=[{"fn": "sum", "column": "total"}])
    north = again.provenance[0]["sum total"]
    assert [str(s) for s in north.sources] == ["t:checks/{A1,A2,A3}/rating"]
    table = DatasetTable(again)
    assert table.cell_sources(1, 1) == ("t:checks/{A1,A2,A3}/rating",)


def test_derive_takes_a_selector_of_the_whole_table_and_refuses_a_part(wb):
    assert wb.derive("t:checks/*").rows == wb.derive("t:checks/*/*").rows \
        == wb.derive("t:checks").rows
    for part in ("t:checks/*/rating", "t:checks/{A1,A2}"):
        with pytest.raises(DeriveError) as caught:
            wb.derive(part)
        assert caught.value.code == "bad_source"


def test_a_dataset_takes_selectors_and_refuses_a_key_set(wb):
    table = DatasetTable(_area_means(wb))
    assert [[v.v for v in row] for row in table.read("d:by_area/*/mean")] == \
        [[v.v for v in row] for row in table.read("d:by_area[mean]")]
    assert table.select("d:by_area/#2/*").text == "d:by_area/#2/*"
    with pytest.raises(RegionError) as caught:
        table.select("d:by_area/{north,south}/mean")
    assert caught.value.code == "no_key_set"


# -- stable across edits ---------------------------------------------------------------------


def test_addresses_by_key_name_the_same_cells_after_rows_are_inserted_and_reordered():
    before, after = Workbook.open(book(CHECKS)), Workbook.open(book(edited()))
    # a plan made on the first version
    plan = ["t:checks/B2/rating", "t:checks/{C1,A2}/note", "t:checks/A3/*"]
    for address in plan:
        assert _values(before.read(address)) == _values(after.read(address))
    # a position is a place: it names another row now
    assert _values(before.read("t:checks/#2")) != _values(after.read("t:checks/#2"))

    ds = _area_means(before)
    for row, where in zip(ds.unrounded, ds.provenance):
        sources = [str(s) for s in where["mean"].sources]
        cells = [v for s in sources for r in after.read(s) for v in r]
        assert [c.v for c in cells] == [v for s in sources for r in before.read(s) for v in r
                                        for v in [v.v]]
        assert float(sum(Fraction(c.v) for c in cells) / len(cells)) == row["mean"]
    # the same sources as ranges name other cells after the edit
    ranges = [f"Log!C{r}" for r in (2, 3, 5)]                    # north's, before
    assert [before.cell(r).v for r in ranges] != [after.cell(r).v for r in ranges]


def test_a_coverage_rule_made_before_an_edit_finds_exactly_the_rows_added():
    before, after = Workbook.open(book(CHECKS)), Workbook.open(book(edited()))
    required = set(after.expand("t:checks/*/rating"))
    ds = _area_means(before)
    used = {a for where in ds.provenance for s in where["mean"].sources
            for a in after.expand(str(s))}
    assert required - used == {f"t:checks/{c[0]}/rating" for c in ADDED}
    assert set(before.expand("t:checks/*/rating")) == used


def test_the_same_key_addresses_resolve_in_a_table_moved_on_its_sheet():
    before, after = Workbook.open(book(CHECKS)), Workbook.open(book(edited(), at="C4"))
    for address in before.expand("t:checks/*/*"):
        assert before.cell(address).v == after.cell(address).v


# -- the tools -------------------------------------------------------------------------------


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _session(toolbox, data: bytes):
    session = toolbox.session()
    blob = session.add_blob(data, name="book.xlsx")
    assert toolbox.dispatch(session, "open_document", {"blob": blob}).ok
    return session


def test_xlsx_read_pages_a_key_set_in_its_order_and_names_no_key_in_its_summary(box):
    session = _session(box, book(CHECKS))
    target = "t:checks/{C2,A1,B3}/rating"
    first = box.dispatch(session, "xlsx_read", {"doc": "d1", "target": target, "limit": 2})
    assert first.ok, first.error
    assert first.data["target"] == target and first.total == 3
    assert [[c["a"] for c in row] for row in first.data["rows"]] == [["Log!C9"], ["Log!C2"]]
    assert first.summary == "3 values by key of t:checks: rows 1-2 of 3; more with next_cursor"
    second = box.dispatch(session, "xlsx_read", {"doc": "d1", "target": target,
                                                 "cursor": first.next_cursor})
    assert [[c["v"] for c in row] for row in second.data["rows"]] == [[4.5]]
    column = box.dispatch(session, "xlsx_read", {"doc": "d1", "target": "t:checks/*/rating"})
    assert column.data["target"] == "t:checks/*/rating" and column.total == len(CHECKS)


def test_xlsx_derive_shows_sources_by_key(box):
    session = _session(box, book(CHECKS))
    result = box.dispatch(session, "xlsx_derive", {
        "doc": "d1", "from": "t:checks", "id": "by_area", "group_by": ["area"],
        "measures": [{"fn": "mean", "column": "rating", "as": "mean"}]})
    assert result.ok, result.error
    assert result.data["rows"][0][1]["sources"] == ["t:checks/{A1,A2,A3}/rating"]
    whole = box.dispatch(session, "xlsx_derive", {"doc": "d1", "from": "t:checks/*",
                                                  "id": "all"})
    assert whole.ok, whole.error
    refused = box.dispatch(session, "xlsx_read", {"target": "d:by_area/{north,east}/mean"})
    assert refused.to_json()["error"]["code"] == "invalid_arguments"
