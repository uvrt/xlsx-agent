"""The Excel oracle on AutoFilters, hidden rows and columns, and pivot refresh (round 44,
R55, R61).

Opt in with ``pytest -m oracle`` (tests/oracle.py says how Excel is driven; a held lock or an
Excel in use skips).  For each filter, Excel's own (its object model's ``autofilter range``,
``autofilter mode`` false, ``show all data``: ``tools/excel_filter.applescript``) on the
made-up stock workbook (``tests/filter_book.py``) is set beside the same filter set here and
re-saved by Excel: the two files must say the same -- the ``autoFilter`` and its criteria,
the hidden rows, ``filterMode`` and the ``_FilterDatabase`` name -- and Excel must open ours
without a repair.  Criteria on numbers are integers in ``General``, whose text is the same
in every region (Excel's text for ``1.5`` in ``0.00`` is the region's, ``1,50`` in the Dutch
one measured).  Excel does not apply a saved filter on open (measured), so its re-save of
ours shows the rows as hidden here.
"""

from __future__ import annotations

import io
import sys
import zipfile

import pytest
from lxml import etree

from xlsx_agent import Workbook

import filter_book
import oracle

X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


@pytest.fixture(scope="module")
def excel():
    if sys.platform != "darwin" or not oracle.available():
        pytest.skip("the Excel oracle needs Microsoft Excel for Mac")
    try:
        with oracle.session() as current:
            yield current
    except oracle.ExcelBusy as busy:
        pytest.skip(str(busy))


def _facts(data: bytes) -> dict:
    """What a file says of its filters: each sheet's ``autoFilter`` (ref and criteria),
    hidden rows and ``filterMode``, each table's ``autoFilter``, and the filter names."""
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    book = etree.fromstring(parts["xl/workbook.xml"])
    out: dict = {"names": sorted((n.get("name"), n.get("localSheetId"), n.get("hidden"), n.text)
                                 for n in book.iter(X + "definedName"))}

    def criteria(node) -> list | None:
        if node is None:
            return None
        return [node.get("ref")] + [
            (c.get("colId"), c.find(X + "filters").get("blank"),
             [f.get("val") for f in c.iter(X + "filter")]) for c in node.iter(X + "filterColumn")]

    for name, data in sorted(parts.items()):
        if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
            root = etree.fromstring(data)
            sheet_pr = root.find(X + "sheetPr")
            out[name] = {"filter": criteria(root.find(X + "autoFilter")),
                         "hidden": [r.get("r") for r in root.iter(X + "row")
                                    if r.get("hidden") in ("1", "true")],
                         "hidden_columns": [(c.get("min"), c.get("max")) for c in
                                            root.iter(X + "col")
                                            if c.get("hidden") in ("1", "true")],
                         "filter_mode": None if sheet_pr is None
                         else sheet_pr.get("filterMode")}
        elif name.startswith("xl/tables/") and name.endswith(".xml"):
            out[name] = criteria(etree.fromstring(data).find(X + "autoFilter"))
    return out


#: Excel's steps (tools/excel_filter.applescript), and the same done here.
CASES = {
    "values": (["filter|Stock|A1:F8|1|Lamp~Desk"],
               lambda wb: wb.set_autofilter("Stock", "A1:F8", {"Item": ["Lamp", "Desk"]})),
    "blanks": (["filter|Stock|A1:F8|3|="],
               lambda wb: wb.set_autofilter("Stock", "A1:F8", {"Shelf": [None]})),
    "value_and_blanks": (["filter|Stock|A1:F8|3|A~="],
                         lambda wb: wb.set_autofilter("Stock", "A1:F8", {"Shelf": ["A", None]})),
    "integers": (["filter|Stock|A1:F8|2|3~12"],
                 lambda wb: wb.set_autofilter("Stock", "A1:F8", {"Qty": [3, 12]})),
    "buttons": (["filter|Stock|A1:F8||"], lambda wb: wb.set_autofilter("Stock!A1:F8")),
    "table": (["filter|Orders|A1:C6|2|Lamp"],
              lambda wb: wb.set_autofilter("t:Orders", criteria={"Item": "Lamp"})),
    "off": (["filter|Stock|A1:F8|1|Lamp", "off|Stock"],
            lambda wb: (wb.set_autofilter("Stock", "A1:F8", {"Item": "Lamp"}),
                        wb.clear_autofilter("Stock"))),
    "clear": (["filter|Stock|A1:F8|1|Lamp", "showall|Stock"],
              lambda wb: (wb.set_autofilter("Stock", "A1:F8", {"Item": "Lamp"}),
                          wb.clear_autofilter("Stock", keep=True))),
}


@pytest.mark.oracle
@pytest.mark.parametrize("hidden", [(), (2, 3, 10)], ids=["plain", "rows_hidden_by_hand"])
@pytest.mark.parametrize("case", sorted(CASES))
def test_a_filter_set_here_says_what_excel_s_own_says(excel, case, hidden):
    steps, ours = CASES[case]
    data = filter_book.stock_book(hidden=hidden)
    theirs = excel.operate(data, steps, name=f"filter-{case}-{len(hidden)}",
                           script=oracle.FILTER_SCRIPT)
    if theirs.outcome == "error" and "-609" in theirs.detail:
        pytest.skip(f"Excel lost its connection during the step: {theirs.detail[:120]}")
    assert theirs.ok, theirs.detail
    wb = Workbook.open(data)
    ours(wb)
    assert wb.validate() == []
    resaved = excel.resave(wb.to_bytes(), name=f"filter-ours-{case}-{len(hidden)}")
    assert resaved.ok, resaved.detail
    assert _facts(resaved.saved.read_bytes()) == _facts(theirs.saved.read_bytes())


@pytest.mark.oracle
def test_criteria_excel_did_not_apply_on_open_stay_as_written(excel):
    wb = Workbook.open(filter_book.stock_book())
    wb.set_autofilter("Stock", "A1:F8", {"Item": ["Lamp"]}, hide=False)
    resaved = excel.resave(wb.to_bytes(), name="filter-left-to-excel")
    assert resaved.ok, resaved.detail
    facts = _facts(resaved.saved.read_bytes())["xl/worksheets/sheet1.xml"]
    assert facts["filter"] == ["A1:F8", ("0", None, ["Lamp"])]
    assert facts["hidden"] == [] and facts["filter_mode"] == "1"


@pytest.mark.oracle
def test_rows_and_columns_hidden_by_hand_open_and_stay_hidden(excel):
    wb = Workbook.open(filter_book.stock_book())
    wb.hide_rows("Stock", [3, "5:6", 12])
    wb.hide_columns("Stock", ["B", "D:E"])
    resaved = excel.resave(wb.to_bytes(), name="filter-hidden-by-hand")
    assert resaved.ok, resaved.detail
    facts = _facts(resaved.saved.read_bytes())["xl/worksheets/sheet1.xml"]
    assert facts["hidden"] == ["3", "5", "6", "12"]
    assert facts["hidden_columns"] == [("2", "2"), ("4", "5")]


@pytest.mark.oracle
def test_a_pivot_cache_flagged_here_opens_and_keeps_the_flag(excel):
    wb = Workbook.open(filter_book.pivot_book())
    wb.refresh_pivots_on_load(True)
    resaved = excel.resave(wb.to_bytes(), name="filter-pivot-refresh")
    assert resaved.ok, resaved.detail
    with zipfile.ZipFile(resaved.saved) as archive:
        cache = etree.fromstring(archive.read("xl/pivotCache/pivotCacheDefinition1.xml"))
        summary = etree.fromstring(archive.read("xl/worksheets/sheet2.xml"))
    assert cache.get("refreshOnLoad") == "1"
    # measured (Excel for Mac 16.106): the flag is kept, but the pivot is not refreshed on
    # open -- it still shows the stale 20, not the source's 10
    shown = {c.get("r"): c.findtext(X + "v") for c in summary.iter(X + "c")}
    assert shown["A4"] == "20"
