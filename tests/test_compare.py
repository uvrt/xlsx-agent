"""R12: ``Workbook.compare`` and ``xlsx_compare``: sheets by id, tables by name, rows by key,
cells by type and value; duplicate keys reported, never paired; the library first, then
through a real ooxml-edit Toolbox, bounded and paged on a large pair.

Every workbook is built with tests/build.py.
"""

import datetime as dt
import json
from datetime import date

import pytest

from build import Formula, WorkbookBuilder, cell_name
from ooxml_edit.tools import Limits, Toolbox
from test_catalogue import MARKER

from xlsx_agent import NameNotFound, RegionError, Workbook
from xlsx_agent.compare import CELL_CHANGES
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS
from xlsx_agent.tools.read import TEXT_SHOWN

CLOCK = lambda: dt.datetime(2026, 10, 13, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731


def _book(rows, columns=("SKU", "Name", "Price", "Qty"), *, sheet="Stock", sheet_id=None,
          table="Products", extra_sheets=(), styles=None) -> WorkbookBuilder:
    book = WorkbookBuilder()
    for name, ident in extra_sheets:
        book.sheet(name, sheet_id=ident).rows("A1", [["x"], [1]])
    stock = book.sheet(sheet, sheet_id=sheet_id)
    stock.rows("A1", [list(columns)])
    for r, values in enumerate(rows, start=2):
        for c, value in enumerate(values):
            if value is not None:
                stock.cell(cell_name(r, c + 1), value,
                           style=book.style(**styles[c]) if styles and styles.get(c) else 0)
    last = cell_name(len(rows) + 1, len(columns))
    if table:
        stock.table(table, f"A1:{last}", list(columns), write_header=False)
    return book


def _open(book: WorkbookBuilder) -> Workbook:
    return Workbook.open(book.to_bytes())


OLD_ROWS = [["SKU-1", "Bolt", 1.5, 10],
            ["SKU-2", "Nut", 0.25, 40],
            ["SKU-3", "Washer", 0.1, 100],
            ["SKU-4", "Hinge", 4.0, None],
            ["SKU-5", "Bracket", 2.75, 8]]

NEW_ROWS = [["SKU-1", "Bolt", 1.5, 10],          # unchanged
            ["SKU-2", "Nut", 0.3, "40"],          # changed; type (a number now text)
            ["SKU-4", "Hinge", None, 3],          # removed; added
            ["SKU-5", "bracket", 2.75, 8],        # changed (case counts)
            ["SKU-6", "Clamp", 6.5, 2]]           # added row; SKU-3 removed


@pytest.fixture(scope="module")
def pair():
    old = _open(_book(OLD_ROWS, sheet_id=3, extra_sheets=[("Notes", 1), ("Old", 2)]))
    new = _open(_book(NEW_ROWS, columns=("SKU", "Name", "Price", "Count"), sheet="Stock 2026",
                      sheet_id=3, extra_sheets=[("Notes", 1), ("Fresh", 4)]))
    return old, new


# -- the library -------------------------------------------------------------------------------


def test_sheets_are_matched_by_id(pair):
    diff = Workbook.compare(*pair, key="SKU")
    assert [s.as_dict() for s in diff.sheets] == [
        {"change": "removed", "id": 2, "a": "Old"},
        {"change": "renamed", "id": 3, "a": "Stock", "b": "Stock 2026"},
        {"change": "added", "id": 4, "b": "Fresh"}]


def test_added_removed_and_changed_rows_and_cells(pair):
    diff = Workbook.compare(*pair, key="SKU")
    (table,) = diff.tables
    assert (table.a, table.b, table.key) == ("t:Products", "t:Products", "SKU")
    assert (table.range_a, table.range_b) == ("Stock!A1:D6", "'Stock 2026'!A1:D6")
    assert (table.rows_a, table.rows_b, table.paired, table.unchanged) == (5, 5, 4, 1)
    assert [(c.change, c.a, c.b) for c in table.columns] == [("renamed", "Qty", "Count")]
    assert [(r.key, r.row) for r in table.added] == [("SKU-6", "t:Products/#5")]
    assert [v.as_dict() for v in table.added[0].cells] == [
        {"a": "'Stock 2026'!A6", "v": "SKU-6", "t": "text"},
        {"a": "'Stock 2026'!B6", "v": "Clamp", "t": "text"},
        {"a": "'Stock 2026'!C6", "v": 6.5, "t": "number"},
        {"a": "'Stock 2026'!D6", "v": 2, "t": "number"}]
    assert [(r.key, r.row) for r in table.removed] == [("SKU-3", "t:Products/#3")]
    assert table.removed[0].cells[1].as_dict() == {"a": "Stock!B4", "v": "Washer", "t": "text"}
    changes = [(c.change, c.key, c.column, c.a.as_dict(), c.b.as_dict()) for c in table.changed]
    assert changes == [
        ("changed", "SKU-2", "Price", {"a": "Stock!C3", "v": 0.25, "t": "number"},
         {"a": "'Stock 2026'!C3", "v": 0.3, "t": "number"}),
        ("type", "SKU-2", "Count", {"a": "Stock!D3", "v": 40, "t": "number"},
         {"a": "'Stock 2026'!D3", "v": "40", "t": "text", "issue": "number_as_text"}),
        ("removed", "SKU-4", "Price", {"a": "Stock!C5", "v": 4.0, "t": "number"},
         {"a": "'Stock 2026'!C4", "v": None, "t": "blank"}),
        ("added", "SKU-4", "Count", {"a": "Stock!D5", "v": None, "t": "blank"},
         {"a": "'Stock 2026'!D4", "v": 3, "t": "number"}),
        ("changed", "SKU-5", "Name", {"a": "Stock!B6", "v": "Bracket", "t": "text"},
         {"a": "'Stock 2026'!B5", "v": "bracket", "t": "text"})]
    assert {c.change for c in table.changed} == set(CELL_CHANGES)
    assert diff.totals() == {"sheets": 3, "columns": 1, "rows_added": 1, "rows_removed": 1,
                             "cells_changed": 5}
    assert not diff.same


def test_a_workbook_is_the_same_as_itself(pair):
    old, _ = pair
    diff = Workbook.compare(old, old, key="SKU")
    assert diff.same and diff.totals() == {} and diff.tables[0].unchanged == 5


def test_types_are_compared_before_values():
    pct = {"num_fmt": "0%"}
    old = _open(_book([["A", True, 0.5, date(2026, 1, 5)], ["B", 1, 2, date(2026, 1, 6)]],
                      columns=("Id", "Flag", "Rate", "Day"), styles={2: pct}))
    new = _open(_book([["A", 1, 0.5, date(2026, 1, 7)], ["B", 1, 2, date(2026, 1, 6)]],
                      columns=("Id", "Flag", "Rate", "Day")))
    (table,) = Workbook.compare(old, new, key="Id").tables
    assert [(c.change, c.column, c.a.t, c.b.t) for c in table.changed] == [
        ("type", "Flag", "bool", "number"),           # true is not 1
        ("type", "Rate", "percent", "number"),        # the format changed what it is
        ("changed", "Day", "date", "date"),
        ("type", "Rate", "percent", "number")]
    assert table.changed[2].b.v == "2026-01-07"


def test_a_formula_is_compared_by_its_cached_value():
    old = _open(_book([["A", 2, 4]], columns=("Id", "X", "Y")))
    new = _open(_book([["A", 2, Formula("B2*2", cached=4)]], columns=("Id", "X", "Y")))
    (table,) = Workbook.compare(old, new, key="Id").tables
    assert table.changed == () and table.unchanged == 1


def test_whole_number_keys_match_as_digits():
    old = _open(_book([[42, "a"], [7, "b"]], columns=("Id", "V")))
    new = _open(_book([[42.0, "a"], [7, "c"]], columns=("Id", "V")))
    (table,) = Workbook.compare(old, new, key="Id").tables
    assert table.paired == 2 and [c.key for c in table.changed] == ["7"]


def test_duplicate_keys_are_reported_and_never_paired():
    old = _open(_book([["K1", 1], ["K2", 2], ["K2", 3], ["K3", 4]], columns=("Key", "V")))
    new = _open(_book([["K1", 9], ["K2", 2], ["K3", 4], ["K3", 5], ["K4", 6]],
                      columns=("Key", "V")))
    (table,) = Workbook.compare(old, new, key="Key").tables
    assert [(d.key, d.a, d.b) for d in table.duplicates] == [
        ("K2", ("t:Products/#2", "t:Products/#3"), ("t:Products/#2",)),
        ("K3", ("t:Products/#4",), ("t:Products/#3", "t:Products/#4"))]
    # neither K2 nor K3 is paired, added or removed: the first of them never wins
    assert table.paired == 1 and [c.key for c in table.changed] == ["K1"]
    assert [r.key for r in table.added] == ["K4"] and table.removed == ()
    assert table.totals() == {"duplicate_key": 2, "rows_added": 1, "cells_changed": 1}


def test_rows_without_a_key_are_counted_not_paired():
    old = _open(_book([["K1", 1], [None, 2], [1.5, 3]], columns=("Key", "V")))
    new = _open(_book([["K1", 1]], columns=("Key", "V")))
    (table,) = Workbook.compare(old, new, key="Key").tables
    assert table.unkeyed_a.as_dict() == {"count": 2, "first": ["t:Products/#2", "t:Products/#3"]}
    assert table.unkeyed_b.count == 0 and table.removed == () and table.unchanged == 1


def test_the_key_column_is_detected_when_only_one_is_a_key_in_both():
    old = _open(_book([["K1", 1.5], ["K2", 2.5]], columns=("Key", "V")))
    new = _open(_book([["K1", 1.5], ["K3", 2.5]], columns=("Key", "V")))
    (table,) = Workbook.compare(old, new).tables
    assert table.key == "Key" and [r.key for r in table.added] == ["K3"]


def test_no_key_column_is_refused_with_the_candidates():
    old = _open(_book([["K1", 1], ["K2", 2]], columns=("Key", "N")))
    new = _open(_book([["K1", 1], ["K2", 3]], columns=("Key", "N")))
    with pytest.raises(RegionError) as several:
        Workbook.compare(old, new)
    assert several.value.code == "no_key_column"
    assert several.value.valid_options == ["Key", "N"]
    dup = _open(_book([["K1", 1.5], ["K1", 2.5]], columns=("Key", "V")))
    with pytest.raises(RegionError) as none:
        Workbook.compare(dup, dup)
    assert none.value.code == "no_key_column" and none.value.valid_options == []
    with pytest.raises(RegionError) as missing:
        Workbook.compare(old, _open(_book([["K1", 1]], columns=("N", "Id"))), key="Key")
    assert missing.value.code == "no_key_column"


def test_tables_by_name_and_pairs_and_regions():
    old = _open(_book([["K1", 1]], columns=("Key", "V"), table="Products"))
    new = _open(_book([["K1", 2]], columns=("Key", "V"), table="Items"))
    default = Workbook.compare(old, new, key="Key")
    assert default.tables == () and default.tables_added == ("t:Items",)
    assert default.tables_removed == ("t:Products",)
    assert default.totals() == {"unmatched_tables": 2}
    (table,) = Workbook.compare(old, new, tables=[("t:Products", "t:Items")], key="Key").tables
    assert (table.a, table.b, len(table.changed)) == ("t:Products", "t:Items", 1)
    with pytest.raises(NameNotFound):
        Workbook.compare(old, new, tables=["t:Products"], key="Key")
    with pytest.raises(ValueError):
        Workbook.compare(old, new, tables=["t:Products[Key]"], key="Key")
    # a detected region, by its header's first cell, and a key per table
    plain_old = _open(_book([["K1", 1]], columns=("Key", "V"), table=None))
    plain_new = _open(_book([["K1", 5]], columns=("Key", "V"), table=None))
    (region,) = Workbook.compare(plain_old, plain_new, tables="t:Stock!A1",
                                 key={"t:Stock!A1": "Key"}).tables
    assert region.a == "t:Stock!A1" and region.changed[0].b.as_dict()["a"] == "Stock!B2"


def test_columns_limit_the_cells_compared(pair):
    diff = Workbook.compare(*pair, key="SKU", columns=["price", "Count"])
    (table,) = diff.tables
    assert {c.column for c in table.changed} == {"Price", "Count"}
    assert len(table.added[0].cells) == 3            # the key and the two columns
    assert table.columns[0].change == "renamed"       # column changes listed regardless
    assert Workbook.compare(*pair, key="SKU", columns=["Qty"]).tables[0].changed[0].column \
        == "Count"
    with pytest.raises(NameNotFound):
        Workbook.compare(*pair, key="SKU", columns=["Nope"])


def test_added_and_removed_columns_are_not_compared():
    old = _open(_book([["K1", 1, 2]], columns=("Key", "Old", "Same")))
    new = _open(_book([["K1", 3, 9, 4]], columns=("Key", "Same", "New", "Extra")))
    (table,) = Workbook.compare(old, new, key="Key").tables
    assert [(c.change, c.a, c.b) for c in table.columns] == [
        ("added", None, "New"), ("added", None, "Extra"), ("removed", "Old", None)]
    assert [(c.column, c.a.v, c.b.v) for c in table.changed] == [("Same", 2, 3)]


# -- xlsx_compare through a Toolbox ------------------------------------------------------------


@pytest.fixture
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


def _compare(toolbox, session, **arguments):
    return toolbox.dispatch(session, "xlsx_compare", {"a": "d1", "b": "d2", **arguments})


def _pair_books():
    return (_book(OLD_ROWS, sheet_id=3, extra_sheets=[("Notes", 1), ("Old", 2)]),
            _book(NEW_ROWS, columns=("SKU", "Name", "Price", "Count"), sheet="Stock 2026",
                  sheet_id=3, extra_sheets=[("Notes", 1), ("Fresh", 4)]))


def test_xlsx_compare_end_to_end(box):
    session = _session(box, *_pair_books())
    result = _compare(box, session, key="SKU")
    assert result.ok, result.error
    data = result.data
    assert data["tables"] == [{"a": "t:Products", "range_a": "Stock!A1:D6",
                               "range_b": "'Stock 2026'!A1:D6", "key": "SKU", "rows": [5, 5],
                               "paired": 4, "unchanged": 1,
                               "totals": {"columns": 1, "rows_added": 1, "rows_removed": 1,
                                          "cells_changed": 5}}]
    assert data["totals"] == {"sheets": 3, "columns": 1, "rows_added": 1, "rows_removed": 1,
                              "cells_changed": 5}
    differences = data["differences"]
    assert result.total == len(differences) == 11 and result.next_cursor is None
    assert differences[:4] == [
        {"sheet": "removed", "id": 2, "a": "Old"},
        {"sheet": "renamed", "id": 3, "a": "Stock", "b": "Stock 2026"},
        {"sheet": "added", "id": 4, "b": "Fresh"},
        {"table": "t:Products", "column": "renamed", "a": "Qty", "b": "Count"}]
    assert differences[4]["row"] == "added" and differences[4]["key"] == "SKU-6"
    assert differences[4]["at"] == "t:Products/#5" and len(differences[4]["cells"]) == 4
    assert differences[5]["row"] == "removed" and differences[5]["key"] == "SKU-3"
    assert differences[7] == {
        "table": "t:Products", "cell": "type", "key": "SKU-2", "column": "Count",
        "a": {"a": "Stock!D3", "v": 40, "t": "number"},
        "b": {"a": "'Stock 2026'!D3", "v": "40", "t": "text", "issue": "number_as_text"}}
    assert differences[-1]["cell"] == "changed" and differences[-1]["key"] == "SKU-5"
    assert "cut" not in data
    assert result.summary == ("d1 and d2: 1 tables compared; 3 sheets, 1 columns, 1 rows added, "
                              "1 rows removed, 5 cells changed; differences 1-11 of 11")


def test_xlsx_compare_reports_duplicate_keys(box):
    old = _book([["K1", 1], ["K2", 2], ["K2", 3]], columns=("Key", "V"))
    new = _book([["K1", 1], ["K2", 2]], columns=("Key", "V"))
    result = _compare(box, _session(box, old, new), tables=["t:Products"], key="Key")
    assert result.ok, result.error
    assert result.data["differences"] == [
        {"table": "t:Products", "duplicate_key": "K2",
         "a": ["t:Products/#2", "t:Products/#3"], "b": ["t:Products/#2"]}]
    assert result.data["totals"] == {"duplicate_key": 1}


def test_xlsx_compare_with_table_b_and_columns(box):
    old = _book([["K1", 1, 1]], columns=("Key", "V", "W"), table="Products")
    new = _book([["K1", 2, 2]], columns=("Key", "V", "W"), table="Items")
    result = _compare(box, _session(box, old, new), tables=["t:Products"], table_b="t:Items",
                      columns=["W"], key="Key")
    assert result.ok, result.error
    assert result.data["tables"][0]["b"] == "t:Items"
    assert [d["column"] for d in result.data["differences"]] == ["W"]


def test_a_document_compared_with_itself_has_no_differences(box):
    session = _session(box, _pair_books()[0])
    result = box.dispatch(session, "xlsx_compare", {"a": "d1", "b": "d1", "key": "SKU"})
    assert result.ok and result.total == 0 and result.data["differences"] == []
    assert result.summary == "d1 and d1: 1 tables compared; no differences"


@pytest.mark.parametrize("arguments,code,detail", [
    ({}, "refused", "no_key_column"),
    ({"tables": ["t:Nope"], "key": "SKU"}, "not_found", None),
    ({"tables": ["t:Products[SKU]"], "key": "SKU"}, "invalid_arguments", None),
    ({"table_b": "t:Products", "key": "SKU"}, "invalid_arguments", None),
    ({"key": "Nope"}, "not_found", None),
    ({"key": "SKU", "columns": ["Nope"]}, "not_found", None),
    ({"key": "SKU", "cursor": "3.0.00000000"}, "invalid_arguments", None),
    ({"b": "d9", "key": "SKU"}, "not_found", None),
], ids=["no-key", "table", "column-address", "table_b", "key", "columns", "cursor", "doc"])
def test_xlsx_compare_errors(box, arguments, code, detail):
    old = _book([["K1", 1], ["K2", 2]], columns=("SKU", "N"))
    new = _book([["K1", 1], ["K2", 3]], columns=("SKU", "N"))
    result = _compare(box, _session(box, old, new), **arguments)
    assert not result.ok and result.error.code == code, result.error
    if detail is not None:
        assert result.error.details["code"] == detail
        assert result.error.valid_options == ["SKU", "N"]


# -- a large pair: bounded and paged -----------------------------------------------------------

LONG_ROWS = 20_000


@pytest.fixture(scope="module")
def large_pair() -> tuple[bytes, bytes]:
    """Two versions of a ledger of :data:`LONG_ROWS` rows: in the newer, every 50th row's
    amount changed, every 400th row removed, 100 rows added, and one long text."""
    old_rows, new_rows = [], []
    for r in range(1, LONG_ROWS + 1):
        row = [f"L-{r:05d}", ["North", "South"][r % 2], (r * 37) % 1000]
        old_rows.append(row)
        if r % 400 == 0:
            continue
        new_rows.append([row[0], row[1], row[2] + 1 if r % 50 == 0 else row[2]])
    new_rows[0][1] = "x" * 10_000
    new_rows.extend([f"N-{r:03d}", "East", r] for r in range(100))
    columns = ("Id", "Region", "Amount")
    return (_book(old_rows, columns=columns, table="Ledger").to_bytes(),
            _book(new_rows, columns=columns, table="Ledger").to_bytes())


def test_a_large_pair_is_bounded_and_paged(box, large_pair):
    session = _session(box, *large_pair)
    changed = sum(1 for r in range(1, LONG_ROWS + 1) if r % 50 == 0 and r % 400) + 1
    removed = LONG_ROWS // 400
    expected = 100 + removed + changed
    first = _compare(box, session, key="Id")
    assert first.ok, first.error
    assert first.total == expected
    assert first.data["totals"] == {"rows_added": 100, "rows_removed": removed,
                                    "cells_changed": changed}
    assert len(first.data["differences"]) <= box.limits.max_list_items
    assert len(first.to_text()) <= box.limits.max_result_chars
    seen, cursor, pages, texts = [], None, 0, 0
    while True:
        result = _compare(box, session, key="Id", **({"cursor": cursor} if cursor else {}))
        assert result.ok, result.error
        assert len(result.to_text()) <= box.limits.max_result_chars
        assert len(result.data["differences"]) <= box.limits.max_list_items
        assert result.total == expected
        seen.extend(result.data["differences"])
        texts += result.data.get("cut", {}).get("text", 0)
        pages += 1
        cursor = result.next_cursor
        if cursor is None:
            break
        assert result.data["cut"]["differences"] == expected - len(seen)
    assert len(seen) == expected and pages > 1
    assert [d["row"] for d in seen[:150]] == ["added"] * 100 + ["removed"] * removed
    assert len({json.dumps(d, sort_keys=True) for d in seen}) == expected   # each once
    long = next(d for d in seen if d.get("key") == "L-00001")
    assert long["cell"] == "changed" and long["b"]["length"] == 10_000
    assert len(long["b"]["v"]) == TEXT_SHOWN and texts == 1
    # a cursor belongs to its comparison
    other = _compare(box, session, key="Id", columns=["Amount"], cursor=first.next_cursor)
    assert other.error.code == "invalid_arguments" and other.error.field == "cursor"


def test_a_difference_too_big_for_a_page_is_cut_to_fit():
    columns = tuple(f"C{i:03d}" for i in range(1, 201))
    old = _book([["K0"] + [1] * 199], columns=columns)
    new = _book([["K0"] + [1] * 199, ["K1"] + ["y" * 1900] * 3 + [2] * 196], columns=columns)
    small = Limits(max_result_chars=4000)
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0, limits=small) as toolbox:
        result = _compare(toolbox, _session(toolbox, old, new), key="C001")
    assert result.ok, result.error
    (row,) = result.data["differences"]
    assert row["row"] == "added" and row["more_cells"] > 0
    assert len(row["cells"]) + row["more_cells"] == 200
    assert result.data["cut"]["text"] >= 1
    assert len(result.to_text()) <= small.max_result_chars


def test_no_cell_content_in_summaries_warnings_or_errors(box):
    old = _book([[f"K1 {MARKER}", "a"], [f"K2 {MARKER}", "b"], [f"K2 {MARKER}", "c"]],
                columns=("Key", "V"))
    new = _book([[f"K1 {MARKER}", f"z {MARKER}"], [f"K3 {MARKER}", "d"]],
                columns=("Key", "V"))
    session = _session(box, old, new)
    results = [_compare(box, session, key="Key"),
               _compare(box, session, key="Key", columns=["Nope"]),
               _compare(box, session, key="Key", cursor=f"0.0.{MARKER}")]
    assert results[0].ok and MARKER in json.dumps(results[0].data)    # content in results
    assert not any(r.ok for r in results[1:])
    for result in results:
        if result.ok:
            assert MARKER not in result.summary
        else:
            assert MARKER not in result.error.message
        assert MARKER not in " ".join(result.warnings)
    for record in session.log:
        assert MARKER not in json.dumps(getattr(record, "summary", "") or "")
