"""Guards on writes (0.2, X4): the host's write policy, expectations, the revision guard,
protected sheets, cells by key and new rows in a declared table's capacity.

On the corner shop's stock list (``tests/shop.py``, made in code).  Every refusal is
checked to change nothing -- the bytes, the version, the calculation status -- and to list
every problem of the call; no refusal carries a cell's value or a key.
"""

from __future__ import annotations

import io
import json
import zipfile

import pytest

from ooxml_edit.sheets import Cell, Formula
from ooxml_edit.tools import Toolbox

from xlsx_agent import DeclarationError, PolicyError, Workbook, WritePolicy
from xlsx_agent.edit import EditError
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS, declare, narrow, policy

from shop import FIRST_INPUT, Shop

STOCK = {"name": "stock", "sheets": "Stock", "header_row": 9, "key": "Item",
         "capacity": 8}


def _open(policy=None, *, tables=(STOCK,), **options) -> tuple[bytes, Workbook]:
    data = Shop(**options).to_bytes()
    wb = Workbook.open(data, policy=policy)
    if tables:
        wb.declare_tables(list(tables))
    return data, wb


def _refused(wb: Workbook, cells, **options) -> EditError:
    before = wb.to_bytes()
    version = wb.version
    status = wb.calculation_status
    with pytest.raises(EditError) as raised:
        wb.set_cells(cells, **options)
    assert wb.to_bytes() == before and wb.version == version
    assert wb.calculation_status == status
    return raised.value


def _zip(parts: dict[str, bytes]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return buffer.getvalue()


# -- the policy's form ---------------------------------------------------------------------------


def test_a_policy_is_plain_data_checked_for_its_form():
    made = WritePolicy.from_config({
        "cells": ["Stock!B4", " Stock!D10:E30 ", "Stock!", "#3!B1"],
        "tables": ["stock", {"table": "t:stock", "columns": ["Quantity"], "new_rows": True}],
        "formulas": False})
    assert made.cells == ("Stock!B4", "Stock!D10:E30", "Stock!", "#3!B1")
    assert made.tables[0].table == "t:stock" and made.tables[0].columns is None
    assert made.tables[1].new_rows and not made.formulas and made.replace_formulas
    assert WritePolicy.from_config(["Stock!B4"]).cells == ("Stock!B4",)
    assert WritePolicy.from_config(made) is made
    assert WritePolicy.read_only().is_read_only
    assert WritePolicy.read_only().as_dict() == "nothing"
    assert made.as_dict()["tables"][1] == {"table": "t:stock", "columns": ["Quantity"],
                                          "new_rows": True}
    with pytest.raises(AttributeError):
        made.cells = ()                   # immutable
    for bad, path in [({"cell": []}, "cell"), ({"cells": "Stock!B4"}, "cells"),
                      ({"cells": ["B4"]}, "cells[0]"), ({"cells": ["t:stock"]}, "cells[0]"),
                      ({"tables": [{"table": "stock", "colmns": []}]}, "tables[0].colmns"),
                      ({"tables": ["t:stock[Quantity]"]}, "tables[0].table"),
                      ({"tables": [{"table": "stock", "columns": []}]}, "tables[0].columns"),
                      ({"tables": [{"table": "stock", "new_rows": 1}]}, "tables[0].new_rows"),
                      ({"formulas": "no"}, "formulas"), ("Stock!B4", "")]:
        with pytest.raises(PolicyError) as raised:
            WritePolicy.from_config(bad)
        assert raised.value.code == "bad_policy" and raised.value.path == path, bad
    with pytest.raises(PolicyError):
        Workbook.open(Shop().to_bytes(), policy={"cells": ["B4"]})


# -- the policy against the workbook --------------------------------------------------------------


def test_without_a_policy_everything_but_a_protected_sheet_is_writable():
    _, wb = _open(tables=())
    assert wb.policies == () and wb.writable is None
    assert wb.set_cells({"Stock!B4": "Ann", "Lists!C1": 1, "Totals!C1": 2}).changed == 3
    assert "writable" not in wb.catalogue()


def test_cells_outside_the_allowlist_are_refused_together_and_nothing_changes():
    _, wb = _open({"cells": ["Stock!B4", "Stock!D10:D14"]})
    error = _refused(wb, {"Stock!B4": "Ann", "Stock!E10": 1, "Stock!D11": 2,
                          "Totals!C1": 3, "Stock!A10": {"formula": "=1"}})
    assert error.code == "not_writable"
    assert error.addresses == ("Stock!E10", "Totals!C1", "Stock!A10")
    assert wb.set_cells({"Stock!B4": "Ann", "Stock!D11": 2}).changed == 2
    assert wb.writable == {"cells": ["Stock!B4", "Stock!D10:D14"]}
    assert wb.catalogue()["writable"] == {"cells": ["Stock!B4", "Stock!D10:D14"]}


def test_a_table_entry_is_its_columns_data_cells_by_name_id_or_header():
    _, wb = _open({"tables": [{"table": "stock", "columns": ["quantity", "In stock"]}]})
    assert wb.set_cells({"Stock!D10": 1, "Stock!G14": True}).changed == 2
    error = _refused(wb, {"Stock!D9": "Count", "Stock!E10": 1, "Stock!D20": 1,
                          "t:stock/Mug/Unit price": 2})
    assert error.code == "not_writable"
    # the header, another column, a row outside the table, by key too
    assert error.addresses == ("Stock!D9", "Stock!E10", "Stock!D20", "Stock!E13")
    whole = _open({"tables": ["stock"]})[1]
    assert whole.set_cells({"Stock!C11": "Kitchen"}).changed == 1
    assert _refused(whole, {"Stock!C9": "Kind"}).code == "not_writable"


def test_a_read_only_policy_writes_nothing():
    for given in ({"cells": []}, [], WritePolicy.read_only()):
        _, wb = _open(given)
        assert _refused(wb, {"Stock!B4": "Ann"}).code == "not_writable"
        assert wb.writable == "nothing"


def test_the_policy_refuses_formulas_and_writes_over_them_whatever_the_call_says():
    _, wb = _open({"cells": ["Stock!"], "formulas": False, "replace_formulas": False})
    error = _refused(wb, {"Stock!J10": {"formula": "=D10*2"}, "Stock!A10": "P-9",
                          "Stock!I11": 0.5}, replace_formulas=True)
    assert error.code == "formulas_not_allowed"
    assert error.addresses == ("Stock!J10", "Stock!A10", "Stock!I11")
    assert wb.set_cells({"Stock!J10": "=D10*2"}).changed == 1    # text stays text


def test_narrowing_adds_a_policy_every_write_must_pass_and_never_widens():
    _, wb = _open({"cells": ["Stock!B4", "Stock!D10:D14"]})
    wb.narrow_policy({"cells": ["Stock!D10:D14", "Stock!E10:E14"]})
    assert len(wb.policies) == 2
    assert _refused(wb, {"Stock!B4": "Ann"}).code == "not_writable"     # narrowed away
    assert _refused(wb, {"Stock!E10": 1}).code == "not_writable"        # never widened
    assert wb.set_cells({"Stock!D10": 3}).changed == 1
    assert isinstance(wb.writable, list) and len(wb.writable) == 2


def test_a_policy_naming_what_the_workbook_lacks_refuses_every_write():
    _, wb = _open({"tables": ["shelves"]})
    error = _refused(wb, {"Stock!B4": "Ann"})
    assert error.code == "bad_policy" and "tables[0]" in str(error)
    _, wb = _open({"cells": ["Cellar!A1"]})
    assert _refused(wb, {"Stock!B4": "Ann"}).code == "bad_policy"
    _, wb = _open({"tables": [{"table": "stock", "columns": ["Colour"]}]})
    with pytest.raises(PolicyError) as raised:
        wb.check_policy()
    assert raised.value.path == "tables[0].columns[0]"


def test_neither_a_batch_nor_the_editor_nor_a_key_gets_past_the_policy():
    data, wb = _open({"cells": ["Stock!B4"]})
    with pytest.raises(EditError):
        with wb.batch():
            wb.set_cells({"Stock!B4": "Ann"})
            wb.set_cells({"Stock!D10": 1})
    assert wb.to_bytes() == data and wb.version == 0
    sheet = wb.sheet("Stock")
    with pytest.raises(EditError) as raised:
        wb.editor.write(sheet, Cell(10, 4), 1)
    assert raised.value.code == "not_writable"
    _, formulas = _open({"cells": ["Stock!"], "formulas": False})
    with pytest.raises(EditError) as raised:
        formulas.editor.write(formulas.sheet("Stock"), Cell(10, 10), Formula("D10*2"))
    assert raised.value.code == "formulas_not_allowed"
    assert _refused(wb, {"t:stock/Kettle/Quantity": 1}).code == "not_writable"


# -- expectations and the revision ------------------------------------------------------------


def _empty_text_formula() -> bytes:
    """The shop with ``Stock!B4`` a formula whose cached result is empty text."""
    parts = Shop().parts()
    sheet = parts["xl/worksheets/sheet1.xml"].decode()
    assert '<c r="B4" s="5"/>' in sheet
    parts["xl/worksheets/sheet1.xml"] = sheet.replace(
        '<c r="B4" s="5"/>', '<c r="B4" s="5" t="str"><f>""</f><v></v></c>').encode()
    return _zip(parts)


def test_an_expectation_is_met_or_the_call_is_refused_by_type_never_value():
    _, wb = _open(tables=())
    ok = wb.set_cells({
        "Stock!B15": {"value": "Lamp", "expect": {"type": "empty"}},
        "Stock!D10": {"value": 100, "expect": {"type": "number", "value": 120}},
        "Stock!B11": {"value": "Pad", "expect": {"value": "Notebook"}},
        "Stock!G12": {"value": True, "expect": {"value": False}},
        "Stock!H10": {"date": "2026-02-01", "expect": {"date": "2026-01-01"}},
        "Stock!I10": {"value": 0.1, "expect": {"type": "formula", "value": 0.06}},
        "Stock!K1": {"value": 1, "expect": {"value": None}}}, replace_formulas=True)
    assert ok.changed == 7
    error = _refused(wb, {
        "Stock!B12": {"value": "Pot", "expect": {"value": "Kettle"}},     # met
        "Stock!B13": {"value": "Cup", "expect": {"type": "empty"}},       # holds text
        "Stock!D13": {"value": 1, "expect": {"value": "18"}},             # a number, not text
        "Stock!A13": {"value": "P-4", "expect": {"type": "text"}},        # a formula's
        "Stock!B20": {"value": "Lamp", "expect": {"type": "text"}}},
        replace_formulas=True)
    assert error.code == "not_expected"
    assert error.addresses == ("Stock!B13", "Stock!D13", "Stock!A13", "Stock!B20")
    assert error.held == {"Stock!B13": "text", "Stock!D13": "number",
                          "Stock!A13": "formula", "Stock!B20": "blank"}
    assert "Mug" not in str(error) and "18" not in str(error)


def test_a_formula_with_an_empty_result_is_never_empty():
    wb = Workbook.open(_empty_text_formula())
    assert wb.editor.cells.sheet("Stock").content(Cell(4, 2)).value == ""
    error = _refused(wb, {"Stock!B4": {"value": "Ann", "expect": {"type": "empty"}}},
                     replace_formulas=True)
    assert error.code == "not_expected" and error.held == {"Stock!B4": "formula"}
    assert wb.set_cells({"Stock!B4": {"value": "Ann", "expect": {"type": "formula",
                                                                  "value": ""}}},
                        replace_formulas=True).changed == 1


def test_an_expectation_that_is_none_is_refused_as_a_value():
    _, wb = _open(tables=())
    for expect in ({}, {"type": "colour"}, {"value": [1]}, {"value": 1, "date": "2026-01-01"},
                   {"type": "text", "value": None}, "number"):
        assert _refused(wb, {"Stock!B4": {"value": "Ann", "expect": expect}}).code \
            == "bad_value", expect
    assert _refused(wb, {"Stock!B4": {"expect": {"type": "empty"}}}).code == "bad_value"


def test_the_revision_guard_refuses_a_call_made_for_another_version():
    _, wb = _open(tables=())
    assert wb.set_cells({"Stock!B4": "Ann"}, revision=0).changed == 1
    error = _refused(wb, {"Stock!B4": "Bo"}, revision=0)
    assert error.code == "revision" and "version 1" in str(error)
    assert _refused(wb, {"Stock!B4": "Bo"}, revision=True).code == "bad_value"
    assert wb.set_cells({"Stock!B4": "Bo"}, revision=wb.revision).changed == 1
    wb.undo()
    assert wb.set_cells({"Stock!B4": "Cy"}, revision=1).changed == 1


# -- protected sheets ------------------------------------------------------------------------


def test_a_protected_sheet_is_never_written_and_says_so():
    _, wb = _open(tables=(), protected=True)
    assert wb.sheet("Totals").protected and not wb.sheet("Stock").protected
    error = _refused(wb, {"Totals!C1": 1, "Totals!B1": {"formula": "=1"}, "Stock!B4": "x"},
                     replace_formulas=True)
    assert error.code == "protected" and error.addresses == ("Totals!C1", "Totals!B1")
    assert "sheetProtection" in str(error)
    _, ruled = _open({"cells": ["Totals!"]}, tables=(), protected=True)
    assert _refused(ruled, {"Totals!C1": 1}).code == "protected"    # no policy lifts it
    with pytest.raises(EditError) as raised:
        wb.editor.write(wb.sheet("Totals"), Cell(1, 3), 1)
    assert raised.value.code == "protected"


def test_a_locked_structure_leaves_cells_writable():
    _, wb = _open(tables=(), locked=True)
    assert wb.protected
    assert wb.set_cells({"Stock!B4": "Ann", "Totals!C1": 1}).changed == 2


# -- cells by key -----------------------------------------------------------------------------


def test_a_value_by_key_is_the_row_holding_the_key_wherever_it_is():
    _, wb = _open()
    assert wb.set_cells({"t:stock/Kettle/Quantity": 4}).cells == ("Stock!D12",)
    wb.set_cells({"Stock!B12": "Mug", "Stock!B13": "Kettle"})     # rows reordered
    written = wb.set_cells({"t:stock/Kettle/Quantity": 5, "t:stock/Mug/Unit price": 5})
    assert written.cells == ("Stock!D13", "Stock!E12")
    assert wb.cell("Stock!D13").v == 5 and wb.cell("Stock!D12").v == 4


def test_keys_that_name_no_row_or_two_rows_or_a_place_are_refused():
    _, wb = _open()
    wb.set_cells({"Stock!B11": "Pencil"})                     # two rows share a key now
    error = _refused(wb, {"t:stock/Pencil/Quantity": 1, "t:stock/Lamp/Quantity": 1,
                          "t:stock/#2/Quantity": 1, "t:stock/{Mug,Torch}/Quantity": 1,
                          "t:stock/Mug/Colour": 1, "t:stock[Quantity]": 1,
                          "t:stock/*/Quantity": 1})
    assert [code for code, _ in error.problems] == [
        "duplicate_key", "missing_key", "positional", "positional", "bad_address",
        "bad_address", "positional"]
    assert error.addresses == ("cells[0]",)
    assert "Pencil" not in str(error) and "Lamp" not in str(error)


def test_a_table_without_a_declared_key_column_is_not_written_by_key():
    data = Shop(table=True).to_bytes()
    wb = Workbook.open(data)
    assert _refused(wb, {"t:Rates/Kitchen/Tax": 0.2}).code == "no_key_column"
    wb.declare_key("t:Rates", "Category")
    assert wb.set_cells({"t:Rates/Kitchen/Tax": 0.2}).cells == ("Lists!B3",)


# -- new rows in a table's capacity ---------------------------------------------------------


def _region(wb: Workbook) -> dict:
    return next(r for r in wb.catalogue()["regions"] if r["address"] == "t:stock")


def test_new_rows_go_into_the_first_free_capacity_rows_in_the_order_named():
    data, wb = _open()
    assert _region(wb)["capacity"] == 8
    assert _refused(wb, {"t:stock/Lamp/Quantity": 2}).code == "missing_key"
    written = wb.set_cells({"t:stock/Lamp/Quantity": 2, "t:stock/Fan/Category": "Hardware",
                            "t:stock/Lamp/Category": "Hardware",
                            "t:stock/Kettle/Quantity": 1}, new_rows=True)
    assert written.cells == ("Stock!B15", "Stock!D15", "Stock!B16", "Stock!C16",
                             "Stock!C15", "Stock!D12")
    assert written.new_rows == ("Stock!B15", "Stock!B16")
    assert wb.cell("Stock!B15").v == "Lamp" and wb.cell("Stock!B15").t == "text"
    table = wb.table("t:stock")
    assert table.rows == 7
    assert _region(wb)["free"] == 1
    assert wb.set_cells({"t:stock/Fan/Quantity": 6}).cells == ("Stock!D16",)
    assert wb.set_cells({"Stock!E16": 3.5}).changed == 1      # a data row now
    wb.undo()
    wb.undo()
    wb.undo()
    assert wb.table("t:stock").rows == 5 and wb.to_bytes() == data
    wb.redo()
    assert wb.table("t:stock").rows == 7


def test_new_rows_past_the_capacity_or_without_one_are_refused_before_any_write():
    _, wb = _open(tables=[{**STOCK, "capacity": 6}])
    error = _refused(wb, {"t:stock/Lamp/Quantity": 2, "t:stock/Fan/Quantity": 3,
                          "t:stock/Fan/Category": "Hardware"}, new_rows=True)
    assert error.code == "capacity_full" and error.addresses == ("cells[1]", "cells[2]")
    assert "Fan" not in str(error)
    _, plain = _open(tables=[{k: v for k, v in STOCK.items() if k != "capacity"}])
    assert _refused(plain, {"t:stock/Lamp/Quantity": 2}, new_rows=True).code == "no_capacity"


def test_a_free_capacity_row_is_written_only_by_key():
    _, wb = _open()
    error = _refused(wb, {"Stock!B15": "Lamp", "Stock!J15": 1})
    assert error.code == "capacity_row" and error.addresses == ("Stock!B15",)
    assert wb.set_cells({"Stock!J16": 1}).changed == 1    # outside the table's columns


def test_a_new_row_names_its_key_cell_once_and_takes_a_number_key_as_a_number():
    _, wb = _open()
    error = _refused(wb, {"t:stock/Lamp/Item": "Lamp", "t:stock/Lamp/Quantity": 1},
                     new_rows=True)
    assert error.code == "duplicate_cell" and error.addresses == ("Stock!B15",)
    _, counted = _open(tables=[{**STOCK, "key": "Quantity"}])
    written = counted.set_cells({"t:stock/7/Item": "Lamp", "t:stock/007/Item": "Fan"},
                                new_rows=True)
    assert counted.cell("Stock!D15").v == 7 and counted.cell("Stock!D15").t == "number"
    assert counted.cell("Stock!D16").v == "007" and written.new_rows == ("Stock!D15",
                                                                         "Stock!D16")


def test_new_rows_follow_the_policy():
    allowed = {"tables": [{"table": "stock", "columns": ["Quantity"], "new_rows": True}]}
    _, wb = _open(allowed)
    assert wb.set_cells({"t:stock/Lamp/Quantity": 2}, new_rows=True).changed == 2
    assert _refused(wb, {"t:stock/Fan/Category": "Hardware"}, new_rows=True).code \
        == "not_writable"
    _, closed = _open({"tables": [{"table": "stock", "columns": ["Quantity"]}]})
    assert _refused(closed, {"t:stock/Lamp/Quantity": 2}, new_rows=True).code \
        == "not_allocated"          # the next free row only where the policy says new_rows
    assert _refused(closed, {"t:stock/Lamp/Quantity": 2},
                    allocate={"t:stock/Lamp": 15}).code == "not_writable"
    _, bare = _open({"tables": [{"table": "stock", "new_rows": True}]},
                    tables=[{k: v for k, v in STOCK.items() if k != "capacity"}])
    assert _refused(bare, {"Stock!B4": "Ann"}).code == "bad_policy"


def test_a_capacity_is_declared_only_over_empty_rows_on_one_sheet():
    for given, code in [({"capacity": 3}, "bad_capacity"),
                        ({"capacity": 0}, "bad_capacity"),
                        ({"capacity": 20, "last_row": 14}, "bad_capacity"),
                        ({"capacity": 20, "sheets": ["Stock", "Totals"]}, "bad_capacity")]:
        _, wb = _open(tables=())
        with pytest.raises(DeclarationError) as raised:
            wb.declare_tables([{**STOCK, **given}])
        assert raised.value.code == code and raised.value.path == "tables[0].capacity", given
    _, wb = _open(tables=())
    with pytest.raises(DeclarationError) as raised:
        wb.declare_tables([{"name": "run", "type": "key_value", "sheets": "Totals",
                            "capacity": 5}])
    assert raised.value.code == "bad_capacity"
    for value in ("Lamp", {"formula": '=""'}):
        _, wb = _open(tables=())
        wb.set_cells({"Stock!C18": value})
        with pytest.raises(DeclarationError) as raised:
            wb.declare_tables([{**STOCK, "capacity": 20}])
        assert raised.value.code == "capacity_not_empty" and "C18" in str(raised.value)
        assert "Lamp" not in str(raised.value)


def test_a_formula_returning_empty_text_is_no_free_row():
    parts = Shop().parts()
    sheet = parts["xl/worksheets/sheet1.xml"].decode()
    sheet = sheet.replace("</sheetData>", '<row r="16"><c r="C16" t="str"><f>""</f><v></v>'
                          '</c></row></sheetData>')
    parts["xl/worksheets/sheet1.xml"] = sheet.encode()
    wb = Workbook.open(_zip(parts))
    with pytest.raises(DeclarationError) as raised:
        wb.declare_tables([STOCK])
    assert raised.value.code == "capacity_not_empty"
    wb.declare_tables([{**STOCK, "capacity": 6}])    # row 15 free, 16 not the table's
    assert wb.set_cells({"t:stock/Lamp/Quantity": 1}, new_rows=True).changed == 2
    assert _refused(wb, {"t:stock/Fan/Quantity": 1}, new_rows=True).code == "capacity_full"


# -- the contract's keys and rows (X6) ------------------------------------------------------


def test_a_key_column_of_formulas_is_refused_as_a_key():
    _, wb = _open(tables=[{**STOCK, "key": "Id"}])     # ids made from the row's number
    assert wb.cell("Stock!A11").formula and wb.cell("Stock!A11").v == "P-002"
    for options in ({}, {"new_rows": True}):
        error = _refused(wb, {"t:stock/P-002/Quantity": 1}, **options)
        assert error.code == "formula_key" and error.addresses == ("cells[0]",)
        assert "declare a key column of values" in str(error) and "P-002" not in str(error)
    error = _refused(wb, {"t:stock/P-009/Quantity": 1}, allocate={"t:stock/P-009": 15})
    assert error.code == "formula_key" and error.addresses == ("allocate[0]", "cells[0]")
    _, values = _open()                                # the same table keyed by its items
    assert values.set_cells({"t:stock/Mug/Quantity": 1}).cells == ("Stock!D13",)


def test_new_rows_go_into_the_rows_the_host_allocates_in_any_order():
    data, wb = _open()
    written = wb.set_cells({"t:stock/Lamp/Quantity": 2, "t:stock/Fan/Category": "Hardware"},
                           allocate={"t:stock/Fan": 15, "t:stock/Lamp": 16})
    assert written.cells == ("Stock!B16", "Stock!D16", "Stock!B15", "Stock!C15")
    assert written.new_rows == ("Stock!B16", "Stock!B15")
    assert wb.cell("Stock!B15").v == "Fan" and wb.table("t:stock").rows == 7
    assert _region(wb)["free"] == 1
    alone = wb.set_cells({}, allocate=[("t:stock/Stool", 17)])   # a row of its key alone
    assert alone.cells == ("Stock!B17",) and alone.new_rows == ("Stock!B17",)
    error = _refused(wb, {"t:stock/Shelf/Quantity": 1}, allocate={"t:stock/Shelf": 17})
    assert error.code == "capacity_full" and "8 rows at most" in str(error)
    for _ in range(2):
        wb.undo()
    assert wb.to_bytes() == data and _region(wb)["free"] == 3
    mixed = wb.set_cells({"t:stock/Lamp/Quantity": 2, "t:stock/Fan/Quantity": 1},
                         allocate={"t:stock/Fan": 15}, new_rows=True)
    assert mixed.new_rows == ("Stock!B16", "Stock!B15")     # Lamp the next free after Fan's


def test_an_allocation_is_refused_off_the_capacity_twice_with_a_gap_or_for_a_known_key():
    _, wb = _open()                                    # data rows 10-14, capacity 10-17
    cases = [
        ({"t:stock/Lamp": 14}, "data_row"), ({"t:stock/Lamp": 18}, "not_capacity"),
        ({"t:stock/Lamp": 9}, "not_capacity"), ({"t:stock/Lamp": 16}, "capacity_gap"),
        ({"t:stock/Lamp": 15, "t:stock/Fan": 17}, "capacity_gap"),
        ([("t:stock/Lamp", 15), ("t:stock/Fan", 15)], "allocated_twice"),
        ([("t:stock/Lamp", 15), ("t:stock/Lamp", 16)], "allocated_twice"),
        ({"t:stock/Kettle": 15}, "existing_key"), ({"t:stock/Lamp": 0}, "bad_value"),
        ({"t:stock/Lamp": True}, "bad_value"), ({"t:stock/Lamp/Item": 15}, "bad_address"),
        ({"t:stock/#3": 15}, "bad_address"), ({"t:nothing/Lamp": 15}, "bad_address")]
    for allocate, code in cases:
        error = _refused(wb, {"t:stock/Lamp/Quantity": 2}, allocate=allocate)
        assert error.code == code, (allocate, error.code)
        assert error.addresses[0].startswith("allocate["), allocate
        assert "Lamp" not in str(error) and "Kettle" not in str(error)
    assert _refused(wb, {"t:stock/Lamp/Quantity": 2}).code == "missing_key"
    assert "allocate" in str(_refused(wb, {"t:stock/Lamp/Quantity": 2}))
    _, full = _open(tables=[{**STOCK, "capacity": 5}])
    assert _refused(full, {}, allocate={"t:stock/Lamp": 15}).code == "capacity_full"
    _, plain = _open(tables=[{k: v for k, v in STOCK.items() if k != "capacity"}])
    assert _refused(plain, {}, allocate={"t:stock/Lamp": 15}).code == "no_capacity"


def test_a_row_holding_a_formula_with_an_empty_result_is_never_allocated():
    _, wb = _open()
    wb.editor.write(wb.sheet("Stock"), Cell(16, 3), Formula('=""'))   # below a free row
    assert _region(wb)["free"] == 3           # row 15 is still the next free one
    error = _refused(wb, {"t:stock/Lamp/Quantity": 2, "t:stock/Fan/Quantity": 1},
                     allocate={"t:stock/Lamp": 15, "t:stock/Fan": 16})
    assert error.code == "not_free" and error.addresses == ("allocate[1]",)
    assert wb.set_cells({"t:stock/Lamp/Quantity": 2},
                        allocate={"t:stock/Lamp": 15}).new_rows == ("Stock!B15",)


def test_allocated_rows_follow_the_policy_over_the_capacity_rows():
    opened = {"tables": [{"table": "stock", "columns": ["Quantity"]}],
              "cells": ["Stock!B15:I15"]}          # the host opens row 15 alone
    _, wb = _open(opened)
    assert _refused(wb, {"t:stock/Lamp/Quantity": 2}, new_rows=True).code == "not_allocated"
    assert _refused(wb, {"t:stock/Lamp/Quantity": 2, "t:stock/Fan/Quantity": 1},
                    allocate={"t:stock/Lamp": 15, "t:stock/Fan": 16}).code == "not_writable"
    assert wb.set_cells({"t:stock/Lamp/Quantity": 2},
                        allocate={"t:stock/Lamp": 15}).changed == 2


# -- the tool ----------------------------------------------------------------------------------


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _session(box, given=None, **options):
    session = box.session(**options)
    declare(session, [STOCK])
    if given is not None:
        policy(session, given)
    blob = session.add_blob(Shop().to_bytes(), name="stock.xlsx")
    opened = box.dispatch(session, "open_document", {"blob": blob})
    return session, opened


def test_the_session_s_policy_is_shown_by_describe_and_refuses_outside_it(box):
    given = {"tables": [{"table": "stock", "columns": ["Quantity"], "new_rows": True}],
             "formulas": False}
    session, opened = _session(box, given)
    assert opened.ok
    described = box.dispatch(session, "describe", {"doc": "d1"}).data
    assert described["writable"] == WritePolicy.from_config(given).as_dict()
    region = next(r for r in described["regions"] if r["address"] == "t:stock")
    assert region["capacity"] == 8 and region["free"] == 3
    refused = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "Stock!D10", "number": 1},
        {"address": "Stock!E10", "number": 1},
        {"address": "t:stock/Mug/Unit price", "number": 1}]})
    assert refused.error.code == "refused" and refused.error.details["code"] == "not_writable"
    assert refused.error.details["addresses"] == ["Stock!E10", "Stock!E13"]
    assert session.entry("d1").version == 0
    written = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "t:stock/Lamp/Quantity", "number": 2, "row": 15}]})
    assert written.ok and written.data["new_rows"] == ["Stock!B15"]
    with pytest.raises(ValueError):
        policy(session, {"cells": []})


def test_the_tool_allocates_a_new_key_s_row_on_its_cells(box):
    session, _ = _session(box)
    refused = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "t:stock/Lamp/Quantity", "number": 2, "row": 14}]})
    assert refused.error.details["code"] == "data_row"
    assert refused.error.details["addresses"] == ["allocate[0]"]
    missing = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "t:stock/Lamp/Quantity", "number": 2}]})
    assert missing.error.details["code"] == "missing_key"
    wrong = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "Stock!D15", "number": 2, "row": 15}]})
    assert wrong.error.code == "invalid_arguments" and wrong.error.field == "cells[0].row"
    written = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "t:stock/Lamp/Quantity", "number": 2, "row": 15},
        {"address": "t:stock/Lamp/Category", "text": "Hardware", "row": 15},
        {"address": "t:stock/Fan/Quantity", "number": 1, "row": 16}]})
    assert written.ok and written.data["new_rows"] == ["Stock!B15", "Stock!B16"]


def test_expectations_and_the_revision_through_the_tool(box):
    session, _ = _session(box)
    refused = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "Stock!B10", "text": "Pen", "expect_value": "Pencil"},
        {"address": "Stock!B11", "text": "Pad", "expect": "empty"},
        {"address": "Stock!H10", "date": "2026-02-01", "expect_value": "2026-01-02"},
        {"address": "Stock!D10", "number": 1, "expect_value": "12"}]})
    assert refused.error.details["code"] == "not_expected", refused.to_json()
    assert refused.error.details["held"] == {"Stock!B11": "text", "Stock!H10": "number",
                                             "Stock!D10": "number"}
    assert "Notebook" not in json.dumps(refused.to_json())
    ok = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "revision": 0, "cells": [
        {"address": "Stock!B10", "text": "Pen", "expect": "text", "expect_value": "Pencil"},
        {"address": "Stock!D10", "blank": True, "expect_value": "120"},
        {"address": "Stock!G10", "bool": False, "expect_value": "TRUE"},
        {"address": "Stock!H10", "date": "2026-02-01", "expect_value": "2026-01-01"},
        {"address": "Stock!I10", "number": 0.1, "expect": "formula",
         "expect_value": "0.06"}], "replace_formulas": True})
    assert ok.ok and ok.version == 1, ok.to_json()
    stale = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "revision": 0, "cells": [
        {"address": "Stock!B10", "text": "Pens"}]})
    assert stale.error.details["code"] == "revision" and stale.error.field == "revision"
    bad = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "Stock!B10", "text": "Pens", "expect": "colour"}]})
    assert bad.error.code == "invalid_arguments", bad.to_json()


def test_a_session_policy_the_workbook_cannot_meet_refuses_the_open(box):
    session, opened = _session(box, {"tables": ["shelves"]})
    assert not opened.ok and opened.error.code == "refused"
    assert opened.error.details == {"code": "bad_policy", "path": "tables[0].table"}


def test_a_read_only_session_policy_and_narrowing_an_open_session(box):
    session, _ = _session(box, {"cells": []})
    refused = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "Stock!B4", "text": "Ann"}]})
    assert refused.error.details["code"] == "not_writable"
    assert box.dispatch(session, "describe", {"doc": "d1"}).data["writable"] == "nothing"
    open_session, _ = _session(box)
    assert box.dispatch(open_session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "Stock!B4", "text": "Ann"}]}).ok
    narrow(open_session, {"cells": ["Stock!D10:D14"]})
    refused = box.dispatch(open_session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "Stock!B4", "text": "Bo"}]})
    assert refused.error.details["code"] == "not_writable"
    view = open_session.view(roles=["edit"])
    narrow(view, {"cells": ["Stock!D10"]})
    assert box.dispatch(open_session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "Stock!D11", "number": 1}]}).error.details["code"] == "not_writable"


def test_a_batch_op_gets_no_further_than_a_call(box):
    session, _ = _session(box, {"cells": ["Stock!B4"]})
    result = box.dispatch(session, "batch", {"ops": [
        {"tool": "xlsx_set_cells", "arguments": {"doc": "d1", "cells": [
            {"address": "Stock!B4", "text": "Ann"}]}},
        {"tool": "xlsx_set_cells", "arguments": {"doc": "d1", "cells": [
            {"address": "Stock!D10", "number": 1}]}}]})
    assert result.ok and "ops[1]" in result.summary, result.to_json()
    assert json.dumps(result.data).count("not_writable") >= 1, result.data
    workbook = session.entry("d1").document
    assert workbook.cell("Stock!D10").v == 120


# -- atomic, and nothing of the content in the log ---------------------------------------------


def test_a_batch_rolled_back_takes_its_new_rows_away_and_a_later_call_sees_them():
    data, wb = _open()
    with wb.batch():
        wb.set_cells({"t:stock/Lamp/Quantity": 2}, new_rows=True)
        assert wb.table("t:stock").rows == 6           # seen inside the batch already
        wb.set_cells({"t:stock/Lamp/Category": "Hardware"})
    assert wb.version == 1 and wb.cell("Stock!C15").v == "Hardware"
    wb.undo()
    with pytest.raises(EditError):
        with wb.batch():
            wb.set_cells({"t:stock/Lamp/Quantity": 2}, new_rows=True)
            wb.set_cells({"Stock!B4": {"value": "Ann", "expect": {"type": "text"}}})
    assert wb.to_bytes() == data and wb.version == 0
    assert wb.table("t:stock").rows == 5 and wb.calculation_status == "cached_unverified"
    assert _region(wb)["free"] == 3


def test_no_key_or_value_reaches_the_log(caplog):
    caplog.set_level("DEBUG")
    _, wb = _open({"tables": [{"table": "stock", "new_rows": True}]})
    wb.set_cells({"t:stock/Lampshade/Quantity": 2, "t:stock/Kettle/Category": "Brassware"},
                 new_rows=True)
    for cells in ({"t:stock/Lampshade/Item": {"value": "Lampshades",
                                               "expect": {"value": "Lampshadez"}}},
                  {"Stock!B4": "Clipboard"}, {"t:stock/Lampshades/Item": "Lampshade"},
                  {"t:stock/Kettle/Quantity": {"value": 9, "expect": {"value": "Teapot"}}}):
        with pytest.raises(EditError) as raised:
            wb.set_cells(cells)
        assert not any(word in str(raised.value) for word in
                       ("Lampshade", "Clipboard", "Teapot", "Kettle"))
    text = " ".join(record.getMessage() for record in caplog.records)
    assert not any(word in text for word in ("Lampshade", "Brassware", "Kettle", "Teapot"))


def test_an_expected_value_given_as_text_is_compared_as_a_read_gives_it():
    _, wb = _open(tables=())
    assert wb.set_cells({
        "Stock!B10": {"value": "Pen", "expect": {"shown": "Pencil"}},
        "Stock!D10": {"value": 1, "expect": {"shown": "120"}},
        "Stock!E11": {"value": 1, "expect": {"type": "number", "shown": "2.5"}},
        "Stock!G10": {"value": False, "expect": {"shown": "true"}},
        "Stock!H10": {"date": "2026-02-01", "expect": {"shown": "2026-01-01"}}}).changed == 5
    error = _refused(wb, {"Stock!D11": {"value": 1, "expect": {"shown": "forty"}},
                          "Stock!B12": {"value": "x", "expect": {"shown": "kettle"}},
                          "Stock!B4": {"value": "x", "expect": {"shown": ""}}})
    assert error.code == "not_expected" and len(error.addresses) == 3
    assert _refused(wb, {"Stock!B4": {"value": "x", "expect": {"shown": 3}}}).code \
        == "bad_value"
    assert _refused(wb, {"Stock!B4": {"value": "x", "expect": {"shown": "a",
                                                               "value": "a"}}}).code \
        == "bad_value"
