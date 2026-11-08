"""Merged cells (production R56): ``wb.merge`` and ``wb.unmerge``, as Excel does them.

On the corner shop (``tests/shop.py``).  What Excel for Mac 16.106 did to the same cells
(round 44, lane S's probe; :mod:`xlsx_agent.edit.merges`): every cell of a merged area takes
the first cell's format, only the first cell's value stays, and with the first cell empty
the first value in reading order moves into it.
"""

from __future__ import annotations

import io

import pytest
from lxml import etree

from xlsx_agent import Workbook
from xlsx_agent.edit import EditError

from referents import unzipped
from shop import Shop

X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _sheet(wb: Workbook, part: str = "xl/worksheets/sheet3.xml"):
    return etree.fromstring(unzipped(wb.to_bytes())[part])


def _merges(wb: Workbook, part: str = "xl/worksheets/sheet3.xml") -> list[str]:
    return [m.get("ref") for m in _sheet(wb, part).iter(X + "mergeCell")]


def _styles(wb: Workbook, part: str = "xl/worksheets/sheet3.xml") -> dict[str, str | None]:
    return {c.get("r"): c.get("s") for c in _sheet(wb, part).iter(X + "c")}


def test_a_merge_gives_every_cell_the_first_cell_s_format():
    wb = Workbook.open(Shop().to_bytes())
    data = wb.to_bytes()
    change = wb.merge("Stock!A9:C10", drop=True)   # headers in style 4, row 10 their inputs
    styles = _styles(wb, "xl/worksheets/sheet1.xml")
    assert {styles[r] for r in ("A9", "B9", "C9", "A10", "B10", "C10")} == {"4"}
    assert "A9:C10" in _merges(wb, "xl/worksheets/sheet1.xml")
    assert wb.cell("Stock!A9").v == "Id" and wb.cell("Stock!B9").t == "blank"
    assert set(change.dropped) == {"Stock!B9", "Stock!C9", "Stock!A10", "Stock!B10",
                                   "Stock!C10"}
    assert wb.validate() == []
    wb.save(io.BytesIO())
    wb.undo()
    assert wb.to_bytes() == data


def test_a_merge_over_content_is_refused_listing_every_cell():
    wb = Workbook.open(Shop().to_bytes())
    wb.set_cells({"Totals!A10": "first", "Totals!B10": 5, "Totals!A11": "below"})
    before = wb.to_bytes()
    with pytest.raises(EditError) as raised:
        wb.merge("Totals!A10:C11")
    assert raised.value.code == "merge_content"
    assert raised.value.addresses == ("Totals!B10", "Totals!A11")
    assert "drop=True" in str(raised.value)
    assert wb.to_bytes() == before


def test_with_the_first_cell_empty_the_first_value_moves_into_it_as_excel_does():
    wb = Workbook.open(Shop().to_bytes())
    wb.set_cells({"Totals!F10": "right", "Totals!H11": 7})
    change = wb.merge("Totals!E10:H11", drop=True)
    assert change.moved == "Totals!F10" and change.dropped == ("Totals!H11",)
    assert wb.cell("Totals!E10").v == "right" and wb.cell("Totals!F10").t == "blank"
    wb.set_cells({"Totals!K11": {"formula": "=1+1"}})
    with pytest.raises(EditError) as raised:
        wb.merge("Totals!J10:K11", drop=True)
    assert raised.value.code == "merge_content"


def test_an_overlap_a_table_and_an_array_formula_are_refused_with_the_way_out():
    wb = Workbook.open(Shop(array=True).to_bytes())
    with pytest.raises(EditError) as raised:
        wb.merge("Stock!H1:J1")
    assert raised.value.code == "merge_overlap" and raised.value.addresses == ("Stock!A1:I1",)
    assert "Stock!A1:J1" in str(raised.value)
    with pytest.raises(EditError) as raised:
        wb.merge("Totals!C5:C6")
    assert raised.value.code == "merge_array"
    tabled = Workbook.open(Shop(table=True).to_bytes())
    with pytest.raises(EditError) as raised:
        tabled.merge("Lists!B5:C6")
    assert raised.value.code == "merge_table"


def test_a_merge_already_made_and_a_single_cell_are_no_change():
    wb = Workbook.open(Shop().to_bytes())
    data = wb.to_bytes()
    assert wb.merge("Stock!A1:I1").changed is False
    assert wb.merge("Stock!B4").changed is False
    assert wb.to_bytes() == data


def test_an_unmerge_keeps_formats_or_fills_them_from_the_first_cell():
    wb = Workbook.open(Shop().to_bytes())
    data = wb.to_bytes()
    change = wb.unmerge("Stock!C1")                 # a cell inside A1:I1
    assert change.areas == ("Stock!A1:I1",)
    assert _merges(wb, "xl/worksheets/sheet1.xml") == ["A2:I2"]
    assert _styles(wb, "xl/worksheets/sheet1.xml").get("C1") is None
    wb.undo()
    assert wb.to_bytes() == data
    filled = wb.unmerge("Stock!A1:I2", fill="anchor")
    assert set(filled.areas) == {"Stock!A1:I1", "Stock!A2:I2"}
    styles = _styles(wb, "xl/worksheets/sheet1.xml")
    assert {styles[f"{c}1"] for c in "ABCDEFGHI"} == {"1"}
    assert sheet_has_no_merges(wb)
    assert wb.validate() == []
    wb.save(io.BytesIO())
    assert wb.unmerge("Stock!A1").changed is False


def sheet_has_no_merges(wb: Workbook) -> bool:
    return _sheet(wb, "xl/worksheets/sheet1.xml").find(X + "mergeCells") is None


def test_a_protected_sheet_and_the_host_policy_refuse_a_merge():
    wb = Workbook.open(Shop(protected=True).to_bytes())
    with pytest.raises(EditError) as raised:
        wb.merge("Totals!A10:B10")
    assert raised.value.code == "protected"
    wb = Workbook.open(Shop().to_bytes(), policy={"cells": ["Totals!A10"]})
    with pytest.raises(EditError) as raised:
        wb.merge("Totals!A10:B10")
    assert raised.value.code == "not_writable" and raised.value.addresses == ("Totals!B10",)


def test_a_range_without_its_sheet_is_refused():
    wb = Workbook.open(Shop().to_bytes())
    with pytest.raises(EditError) as raised:
        wb.merge("A1:B2")
    assert raised.value.code == "merge_range"
