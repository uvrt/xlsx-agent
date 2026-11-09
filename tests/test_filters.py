"""AutoFilters, hidden rows and columns (R55) and pivot refresh on open (R61), round 44.

On made-up workbooks (``tests/filter_book.py``).  What Excel for Mac 16.106 wrote for the
same filters is the module's notes (:mod:`xlsx_agent.edit.filters`); the oracle compares
the two (``tests/test_oracle_filters.py``).  Tool design rules: 1 (intent: a range as the
target, a scalar for a list, a column by letters, a table by its name), 2 (refusals list
the columns, the call that works), 3 (no-ops are ``changed`` false), 4 (lists capped and
counted), 5 (many columns and rows in one step), 6 (what Excel writes).
"""

from __future__ import annotations

import io

import pytest
from lxml import etree

from xlsx_agent import Workbook
from xlsx_agent.edit import EditError
from xlsx_agent.edit.filters import FilterChange

import filter_book
from referents import unzipped

X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
STOCK = "xl/worksheets/sheet1.xml"
ORDERS = "xl/worksheets/sheet2.xml"


def _part(wb: Workbook, part: str = STOCK):
    return etree.fromstring(unzipped(wb.to_bytes())[part])


def _hidden(wb: Workbook, part: str = STOCK) -> list[int]:
    return [int(r.get("r")) for r in _part(wb, part).iter(X + "row") if r.get("hidden") == "1"]


def _filter(wb: Workbook, part: str = STOCK) -> list:
    node = _part(wb, part).find(".//" + X + "autoFilter")
    if node is None:
        return []
    return [node.get("ref")] + [
        (c.get("colId"), c.find(X + "filters").get("blank"),
         [f.get("val") for f in c.iter(X + "filter")]) for c in node.iter(X + "filterColumn")]


def _names(wb: Workbook) -> list:
    book = _part(wb, "xl/workbook.xml")
    return [(n.get("name"), n.get("localSheetId"), n.get("hidden"), n.text)
            for n in book.iter(X + "definedName")]


def _filter_mode(wb: Workbook, part: str = STOCK) -> str | None:
    node = _part(wb, part).find(X + "sheetPr")
    return None if node is None else node.get("filterMode")


# -- a sheet's filter, as Excel writes it (rule 6) --------------------------------------------


def test_a_value_list_writes_what_excel_wrote_and_hides_the_rows_it_excludes():
    wb = Workbook.open(filter_book.stock_book())
    opened = wb.to_bytes()
    change = wb.set_autofilter("Stock", "A1:F8", {"Item": ["Lamp", "Desk"]})
    # Excel (measured): filter val sorted as text, rows 4 and 6 hidden, "lamp" kept
    assert _filter(wb) == ["A1:F8", ("0", None, ["Desk", "Lamp"])]
    assert _hidden(wb) == [4, 6]
    assert _filter_mode(wb) == "1"
    assert _names(wb) == [("_xlnm._FilterDatabase", "0", "1", "Stock!$A$1:$F$8")]
    assert change.hidden == ("4", "6") and change.hidden_by == "xlsx-agent"
    assert any("does not apply a saved filter" in note for note in change.notes)
    assert change.calculation == "cached_unverified"     # no value changed
    assert wb.validate() == []
    wb.save(io.BytesIO())
    wb.undo()
    assert wb.to_bytes() == opened


def test_blanks_alone_and_with_values_are_the_blank_attribute():
    wb = Workbook.open(filter_book.stock_book())
    wb.set_autofilter("Stock", "A1:F8", {"Shelf": [None]})
    assert _filter(wb) == ["A1:F8", ("2", "1", [])]
    assert _hidden(wb) == [2, 3, 5, 6, 7]                 # as Excel hid them
    wb.set_autofilter("Stock", "A1:F8", {"Shelf": ["A", ""]})
    assert _filter(wb) == ["A1:F8", ("2", "1", ["A"])]
    assert _hidden(wb) == [3, 6, 7]


def test_two_columns_keep_the_rows_that_pass_both_and_numbers_as_shown():
    wb = Workbook.open(filter_book.stock_book())
    change = wb.set_autofilter("Stock", "A1:F8", {"Item": ["Lamp", "Desk"], "Qty": [3, 12]})
    assert _filter(wb) == ["A1:F8", ("0", None, ["Desk", "Lamp"]), ("1", None, ["12", "3"])]
    assert _hidden(wb) == [4, 5, 6, 7, 8]
    assert change.criteria == {"Item": ["Desk", "Lamp"], "Qty": ["12", "3"]}


def test_a_number_is_matched_as_its_format_shows_it_in_the_locale():
    wb = Workbook.open(filter_book.stock_book())
    change = wb.set_autofilter("Stock", "A1:F8", {"Price": [1.5]})
    assert _filter(wb) == ["A1:F8", ("3", None, ["1.50"])]   # 0.00 in en-US
    assert _hidden(wb) == [3, 4, 5, 7, 8]
    wb.set_autofilter("Stock", "A1:F8", {"Price": ["1.50"]})
    assert _hidden(wb) == [3, 4, 5, 7, 8]
    assert change.unmatched == {}


def test_a_formula_column_is_read_by_its_cached_values_in_an_untouched_file():
    wb = Workbook.open(filter_book.stock_book())
    wb.set_autofilter("Stock", "A1:F8", {"Value": [4.5, 288]})
    assert _filter(wb) == ["A1:F8", ("4", None, ["288", "4.5"])]
    assert _hidden(wb) == [4, 5, 6, 7, 8]


def test_a_filter_shows_rows_hidden_by_hand_inside_its_range_only():
    wb = Workbook.open(filter_book.stock_book(hidden=(2, 3, 10)))
    change = wb.set_autofilter("Stock", "A1:F8", {"Item": ["Lamp"]})
    # Excel (measured): row 2 (Lamp) shown, row 3 kept hidden by the filter, row 10 below
    assert _hidden(wb) == [3, 4, 6, 8, 10]
    assert change.shown == ("2",)
    wb2 = Workbook.open(filter_book.stock_book(hidden=(2, 3, 10)))
    wb2.set_autofilter("Stock", "A1:F8")                  # the buttons alone show them too
    assert _hidden(wb2) == [10]
    assert _filter_mode(wb2) is None and _filter(wb2) == ["A1:F8"]


def test_turning_a_filter_off_and_clearing_it_as_excel_does():
    wb = Workbook.open(filter_book.stock_book())
    wb.set_autofilter("Stock", "A1:F8", {"Item": ["Lamp"]})
    cleared = wb.clear_autofilter("Stock", keep=True)      # Excel's Clear (show all data)
    assert _filter(wb) == ["A1:F8"] and _hidden(wb) == [] and _filter_mode(wb) is None
    assert cleared.shown == ("3:4", "6", "8")
    wb.set_autofilter("Stock", criteria={"Item": "Lamp"})  # the sheet's own range again
    off = wb.clear_autofilter("Stock")
    assert _filter(wb) == [] and _hidden(wb) == [] and _filter_mode(wb) is None
    # the hidden name stays, as Excel left it
    assert _names(wb) == [("_xlnm._FilterDatabase", "0", "1", "Stock!$A$1:$F$8")]
    assert off.changed and wb.validate() == []


def test_a_table_filter_is_in_the_table_part_with_no_name_nor_filter_mode():
    wb = Workbook.open(filter_book.stock_book())
    change = wb.set_autofilter("t:Orders", criteria={"Item": ["Lamp"]})
    table = etree.fromstring(unzipped(wb.to_bytes())["xl/tables/table1.xml"])
    node = table.find(X + "autoFilter")
    assert node.get("ref") == "A1:C6"
    assert [(c.get("colId"), [f.get("val") for f in c.iter(X + "filter")])
            for c in node.iter(X + "filterColumn")] == [("1", ["Lamp"])]
    assert _hidden(wb, ORDERS) == [3, 5, 6]
    assert _names(wb) == [] and _filter_mode(wb, ORDERS) is None
    assert change.target == "t:Orders" and change.range == "Orders!A1:C6"
    wb.clear_autofilter("Orders")                          # rule 1: the table's name alone
    assert _hidden(wb, ORDERS) == []
    assert etree.fromstring(unzipped(wb.to_bytes())["xl/tables/table1.xml"]).find(
        X + "autoFilter").get("ref") == "A1:C6"


# -- intent, refusals, no-ops, budgets (rules 1-5) -----------------------------------------------


def test_intent_a_range_target_a_scalar_value_and_a_column_by_letters():
    wb = Workbook.open(filter_book.stock_book())
    wb.set_autofilter("Stock!A1:F8", criteria={"a": "lamp", "C": ["A"]})
    assert _filter(wb) == ["A1:F8", ("0", None, ["Lamp"]), ("2", None, ["A"])]
    assert _hidden(wb) == [3, 4, 6, 7, 8]


def test_an_empty_criteria_list_or_mapping_sets_no_criterion():
    wb = Workbook.open(filter_book.stock_book())
    change = wb.set_autofilter("Stock", "A1:F8", {"Item": []})
    assert _filter(wb) == ["A1:F8"] and any("empty list" in n for n in change.notes)
    again = wb.set_autofilter("Stock", "A1:F8", {})
    assert not again.changed and "already so" in again.notes[-1]


def test_an_unknown_column_is_refused_listing_the_filter_s_columns():
    wb = Workbook.open(filter_book.stock_book())
    before = wb.to_bytes()
    with pytest.raises(EditError) as raised:
        wb.set_autofilter("Stock", "A1:F8", {"Colour": ["red"], 2: ["x"]})
    assert raised.value.code == "filter_column"
    assert "Item (A)" in str(raised.value) and "Arrived (F)" in str(raised.value)
    assert "0- or 1-based" in str(raised.value)
    assert wb.to_bytes() == before


def test_a_sheet_without_a_filter_needs_its_range_and_is_told_the_call():
    wb = Workbook.open(filter_book.stock_book())
    with pytest.raises(EditError) as raised:
        wb.set_autofilter("Stock", criteria={"Item": "Lamp"})
    assert raised.value.code == "filter_range"
    assert "wb.set_autofilter('Stock', \"A1:F20\")" in str(raised.value)


def test_a_range_over_a_table_points_to_the_table_s_own_filter():
    wb = Workbook.open(filter_book.stock_book())
    with pytest.raises(EditError) as raised:
        wb.set_autofilter("Orders", "A1:C6", {"Item": "Lamp"})
    assert raised.value.code == "filter_table"
    assert "wb.set_autofilter('t:Orders', ...)" in str(raised.value)


def test_dates_are_refused_unless_hiding_is_left_to_excel():
    wb = Workbook.open(filter_book.stock_book())
    with pytest.raises(EditError) as raised:
        wb.set_autofilter("Stock", "A1:F8", {"Arrived": ["1/5/2026"]})
    assert raised.value.code == "filter_unmeasured"
    assert "a date or time" in str(raised.value) and "hide=False" in str(raised.value)
    change = wb.set_autofilter("Stock", "A1:F8", {"Arrived": ["1/5/2026"]}, hide=False)
    assert _filter(wb) == ["A1:F8", ("5", None, ["1/5/2026"])]
    assert _hidden(wb) == [] and change.hidden_by is None
    assert any("hide=False" in note for note in change.notes)


def test_a_formula_column_after_a_write_is_refused_or_left_to_excel():
    wb = Workbook.open(filter_book.stock_book())
    wb.set_cells({"Stock!B2": 30})
    assert wb.calculation_status == "invalidated"
    with pytest.raises(EditError) as raised:
        wb.set_autofilter("Stock", "A1:F8", {"Value": [4.5]})
    assert raised.value.code == "filter_formulas"
    assert "written since it was opened" in str(raised.value)
    assert "hide=False" in str(raised.value)
    # a column of values is filtered all the same
    wb.set_autofilter("Stock", "A1:F8", {"Item": "Desk"})
    assert _hidden(wb) == [2, 4, 5, 6, 7]
    wb.set_autofilter("Stock", "A1:F8", {"Value": [4.5]}, hide=False)
    assert _filter(wb) == ["A1:F8", ("4", None, ["4.5"])]
    assert _hidden(wb) == [2, 4, 5, 6, 7]                  # left as they were


def test_a_file_asking_to_be_recalculated_refuses_a_formula_column():
    book = filter_book.stock_builder()
    book.calc = {"fullCalcOnLoad": "1"}
    wb = Workbook.open(book.to_bytes())
    with pytest.raises(EditError) as raised:
        wb.set_autofilter("Stock", "A1:F8", {"Value": [4.5]})
    assert "asks to be recalculated" in str(raised.value)


def test_values_no_cell_matches_are_written_and_listed():
    wb = Workbook.open(filter_book.stock_book())
    change = wb.set_autofilter("Stock", "A1:F8", {"Item": ["Lamp", "Sofa"], "Qty": [99]})
    assert change.unmatched == {"Item": ["Sofa"], "Qty": [99]}
    assert _filter(wb)[1:] == [("0", None, ["Lamp", "Sofa"]), ("1", None, ["99"])]
    assert _hidden(wb) == list(range(2, 9))


def test_setting_the_same_filter_twice_is_a_no_op():
    wb = Workbook.open(filter_book.stock_book())
    wb.set_autofilter("Stock", "A1:F8", {"Item": ["Lamp"]})
    version = wb.version
    again = wb.set_autofilter("Stock", "A1:F8", {"Item": ["lamp"]})
    assert not again.changed and wb.version == version
    none = Workbook.open(filter_book.stock_book()).clear_autofilter("Stock")
    assert not none.changed and none.notes == ("the sheet has no filter",)


def test_another_range_replaces_the_sheet_s_filter():
    wb = Workbook.open(filter_book.stock_book())
    wb.set_autofilter("Stock", "A1:F8", {"Item": ["Lamp"]})
    change = wb.set_autofilter("Stock", "A1:F5", {"Item": ["Desk"]})
    assert _filter(wb) == ["A1:F5", ("0", None, ["Desk"])]
    assert _hidden(wb) == [2, 4, 5]
    assert change.shown == ("3", "6", "8")
    assert any("replaced" in note for note in change.notes)
    assert _names(wb) == [("_xlnm._FilterDatabase", "0", "1", "Stock!$A$1:$F$5")]


def test_a_result_s_lists_are_capped_and_counted():
    change = FilterChange("set_autofilter", "Stock", "Stock!A1:A200",
                          hidden=tuple(str(n) for n in range(2, 200, 2)))
    shown = change.as_dict(limit=10)
    assert len(shown["hidden"]) == 10 and shown["hidden_cut"] == 89


def test_a_protected_sheet_filters_only_when_its_protection_allows_it():
    book = filter_book.stock_builder()
    book.sheets[0].protected = True
    wb = Workbook.open(book.to_bytes())
    with pytest.raises(EditError) as raised:
        wb.set_autofilter("Stock", "A1:F8", {"Item": "Lamp"})
    assert raised.value.code == "protected" and "autoFilter" in str(raised.value)


def test_hidden_rows_ask_for_recalculation_when_a_subtotal_reads_them():
    book = filter_book.stock_builder()
    book.sheets[0].cell("B10", filter_book.Formula("SUBTOTAL(109,B2:B8)", 160))
    wb = Workbook.open(book.to_bytes())
    change = wb.set_autofilter("Stock", "A1:F8", {"Item": "Lamp"})
    assert change.calculation == "invalidated"
    wb2 = Workbook.open(filter_book.stock_book())
    assert wb2.hide_rows("Stock", 3).calculation == "cached_unverified"


# -- rows and columns by hand ----------------------------------------------------------------


def test_rows_hidden_and_shown_by_hand_round_trip_to_the_bytes_opened():
    wb = Workbook.open(filter_book.stock_book())
    opened = wb.to_bytes()
    change = wb.hide_rows("Stock", [3, "5:6", 12])
    assert _hidden(wb) == [3, 5, 6, 12] and change.changed == ("3", "5:6", "12")
    again = wb.hide_rows("Stock", "5")
    assert again.changed == () and again.already == ("5",) and again.notes   # rule 3
    wb.show_rows("Stock", ["3:6", 12])
    assert _hidden(wb) == [] and wb.to_bytes() == opened   # row 12 made and taken away
    assert wb.validate() == []


def test_columns_hidden_with_the_default_width_and_shown_again():
    wb = Workbook.open(filter_book.stock_book())
    opened = wb.to_bytes()
    wb.hide_columns("Stock", ["B", "D:E"])
    cols = [(c.get("min"), c.get("max"), c.get("hidden"), c.get("width") is not None)
            for c in _part(wb).iter(X + "col")]
    assert cols == [("2", "2", "1", True), ("4", "5", "1", True)]
    wb.show_columns("Stock", "B:E")
    assert wb.to_bytes() == opened
    with pytest.raises(EditError) as raised:
        wb.hide_rows("Stock", ["x", 0])
    assert raised.value.code == "bad_position" and len(raised.value.problems) == 2


def test_hiding_on_a_protected_sheet_follows_its_format_options():
    book = filter_book.stock_builder()
    book.sheets[0].protected = True
    wb = Workbook.open(book.to_bytes())
    with pytest.raises(EditError) as raised:
        wb.hide_rows("Stock", 3)
    assert raised.value.code == "protected" and "formatRows" in str(raised.value)


def test_one_step_for_a_filter_and_rows_in_a_batch():
    wb = Workbook.open(filter_book.stock_book())
    opened = wb.to_bytes()
    with wb.batch():
        wb.set_autofilter("Stock", "A1:F8", {"Item": "Lamp"})
        wb.hide_columns("Stock", "F")
    wb.undo()
    assert wb.to_bytes() == opened


# -- pivot caches refreshed on open (R61) ------------------------------------------------------


PIVOT = "xl/pivotCache/pivotCacheDefinition1.xml"


def test_refresh_on_load_is_one_attribute_every_other_byte_kept():
    wb = Workbook.open(filter_book.pivot_book())
    opened = unzipped(wb.to_bytes())
    change = wb.refresh_pivots_on_load(True)
    after = unzipped(wb.to_bytes())
    assert [c["pivots"] for c in change.caches] == [[{"name": "Totals", "sheet": "Summary"}]]
    old, new = opened[PIVOT], after[PIVOT]
    assert new.replace(b' refreshOnLoad="1"', b"", 1) == old
    assert {k for k in after if after[k] != opened.get(k)} == {PIVOT}
    assert wb.validate() == []
    wb.refresh_pivots_on_load(False)
    assert unzipped(wb.to_bytes())[PIVOT] == old           # taken away as Excel writes it


def test_refresh_on_load_by_pivot_name_no_ops_and_refusals():
    wb = Workbook.open(filter_book.pivot_book(refresh=True))
    again = wb.refresh_pivots_on_load(True, pivots=["totals"])
    assert not again.changed and again.unchanged[0]["cache_id"] == "1"
    off = wb.refresh_pivots_on_load(False, pivots="1")
    assert off.changed and b"refreshOnLoad" not in unzipped(wb.to_bytes())[PIVOT]
    zero = Workbook.open(filter_book.pivot_book(refresh=False))
    assert not zero.refresh_pivots_on_load(False).changed
    turned = zero.refresh_pivots_on_load(True)
    assert b'refreshOnLoad="1"' in unzipped(zero.to_bytes())[PIVOT] and turned.changed
    with pytest.raises(EditError) as raised:
        wb.refresh_pivots_on_load(True, pivots=["Sales"])
    assert raised.value.code == "no_pivot" and "['Totals']" in str(raised.value)
    with pytest.raises(EditError) as none:
        Workbook.open(filter_book.stock_book()).refresh_pivots_on_load(True)
    assert none.value.code == "no_pivot"
