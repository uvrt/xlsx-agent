"""The Excel oracle on sheets, merged cells and escaped text (round 44, lane S: R52, R56, R63b).

Opt in with ``pytest -m oracle`` (tests/oracle.py says how Excel is driven; a held lock or
an Excel in use skips).  Each made-up workbook (``tests/referents.py``: the corner shop with
every kind of referent naming ``Lists``, a chart and a pivot table) is changed as a program
would -- a sheet added, renamed, hidden, cleared, deleted (forced), cells merged and
unmerged, text with characters XML cannot carry written -- and Excel for Mac must open it
without asking to repair it, and its own re-save must keep what was written: every formula,
defined name, chart reference and pivot source after a rename, the states, the merged
areas, the text as given.
"""

from __future__ import annotations

import io
import sys
import zipfile

import pytest
from lxml import etree

from xlsx_agent import Workbook
from xlsx_agent.formula.sheets import rename

import oracle
import referents
from shop import Shop

X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


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
    """Every formula by sheet name and cell, the defined names, the chart references and
    the pivot sources of a file, whatever its part names (Excel renumbers them)."""
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    book = etree.fromstring(parts["xl/workbook.xml"])
    rels = etree.fromstring(parts["xl/_rels/workbook.xml.rels"])
    targets = {r.get("Id"): "xl/" + r.get("Target").lstrip("/").removeprefix("xl/")
               for r in rels}
    out: dict = {"names": sorted((n.get("name"), n.get("localSheetId"), n.text)
                                 for n in book.iter(X + "definedName")),
                 "sheets": [(s.get("name"), s.get("state") or "visible")
                            for s in book.iter(X + "sheet")]}
    names = [s.get("name") for s in book.iter(X + "sheet")]

    def spelled(text: str) -> str:
        # Excel writes a sheet's name bare where it needs no quotes ('Lists'!B3 -> Lists!B3).
        for name in names:
            text = rename(text, name, name) or text
        return text

    for sheet in book.iter(X + "sheet"):
        tree = etree.fromstring(parts[targets[sheet.get(R_ID)]])
        out[sheet.get("name")] = sorted((f.getparent().get("r"), spelled(f.text))
                                        for f in tree.iter(X + "f") if f.text)
        out[sheet.get("name") + ":merges"] = sorted(m.get("ref")
                                                     for m in tree.iter(X + "mergeCell"))
    out["charts"] = sorted({f.text for name, data in parts.items()
                            if name.startswith("xl/charts/chart") and name.endswith(".xml")
                            for f in etree.fromstring(data).iter(C + "f")})
    out["pivots"] = sorted(w.get("sheet") for name, data in parts.items()
                           if "pivotCacheDefinition" in name and name.endswith(".xml")
                           for w in etree.fromstring(data).iter(X + "worksheetSource"))
    return out


def _changed() -> dict[str, bytes]:
    """Each change, as a file."""
    out = {}
    wb = Workbook.open(referents.book(chart=True, pivot=True))
    wb.rename_sheet("Lists", "Tax rates")
    out["renamed"] = wb.to_bytes()
    wb = Workbook.open(referents.book(chart=True, pivot=True))
    wb.add_sheet("Rota", 1)
    wb.add_sheet("Rota 2")
    out["added"] = wb.to_bytes()
    wb = Workbook.open(referents.book(chart=True, pivot=True))
    wb.set_sheet_state("Stock", "hidden")
    wb.set_sheet_state("Orders", "veryHidden")
    out["hidden"] = wb.to_bytes()
    wb = Workbook.open(referents.book(chart=True, pivot=True, table_formula=True))
    wb.clear_sheet("Totals")
    wb.clear_sheet("Stock")
    wb.clear_sheet("Lists", force=True)
    out["cleared"] = wb.to_bytes()
    wb = Workbook.open(referents.book(chart=True, pivot=True, table_formula=True))
    wb.delete_sheet("Lists", force=True)
    out["deleted"] = wb.to_bytes()
    wb = Workbook.open(Shop().to_bytes())
    wb.set_cells({"Totals!F10": "right", "Totals!H11": 7})
    wb.merge("Stock!A9:C10", drop=True)
    wb.merge("Totals!E10:H11", drop=True)
    wb.unmerge("Stock!A2", fill="anchor")
    out["merged"] = wb.to_bytes()
    return out


TEXTS = {"Totals!A20": "a\x01b", "Totals!A21": "_x0041_", "Totals!A22": "e\rf",
         "Totals!A23": "o\x0bp", "Totals!A24": "_x005F_", "Totals!A25": "c\x1fd"}


@pytest.mark.oracle
def test_excel_opens_every_changed_workbook_without_a_repair(excel):
    for name, data in _changed().items():
        outcome = excel.resave(data, name=f"r44s-{name}")
        assert outcome.ok, (name, outcome)
        if name in ("renamed", "deleted", "added", "hidden", "merged"):
            ours, theirs = _facts(data), _facts(outcome.saved.read_bytes())
            for key in ours:
                if key == "charts" and name == "deleted":
                    continue              # Excel drops a #REF! category it has no data for
                assert ours[key] == theirs[key], (name, key)


@pytest.mark.oracle
def test_excel_keeps_every_reference_a_rename_rewrote(excel):
    data = _changed()["renamed"]
    outcome = excel.resave(data, name="r44s-renamed")
    assert outcome.ok, outcome
    saved = _facts(outcome.saved.read_bytes())
    assert ("Categories", None, "'Tax rates'!$A$2:$A$5") in saved["names"]
    assert ("Span", None, "'Stock:Tax rates'!$B$2") in saved["names"]
    assert ("B6", "SUM('Stock:Tax rates'!B2)") in saved["Totals"]
    assert saved["charts"] == ["'Tax rates'!$A$2:$A$5", "'Tax rates'!$B$1",
                               "'Tax rates'!$B$2:$B$5"]
    assert saved["pivots"] == ["Tax rates"]
    again = Workbook.open(outcome.saved.read_bytes())
    assert again.cell("Stock!I10").v == 0.06          # Excel computed through the new name


@pytest.mark.oracle
@pytest.mark.parametrize("shared", [True, False])
def test_excel_reads_escaped_text_as_it_was_given(excel, shared):
    wb = Workbook.open(Shop(shared=shared).to_bytes())
    wb.set_cells({**TEXTS, "Totals!B20": {"formula": '="_x0041_"&CHAR(1)'}})
    outcome = excel.resave(wb.to_bytes(), name=f"r44s-text-{shared}")
    assert outcome.ok, outcome
    again = Workbook.open(outcome.saved.read_bytes())
    for address, text in TEXTS.items():
        assert again.cell(address).v == text, address
    assert again.cell("Totals!B20").v == "_x0041_\x01"
    assert again.cell("Totals!B20").formula_text == '"_x0041_"&CHAR(1)'


@pytest.mark.oracle
def test_excel_keeps_sheets_named_with_outer_spaces_and_what_was_done_to_them(excel):
    """R69: sheets ``Stock`` and ``Stock `` (and `` Lead``) side by side, each call given the
    spaced name; Excel opens the result without a repair and its re-save keeps the names,
    the edits on the spaced sheet and the formulas naming it."""
    from build import WorkbookBuilder

    book = WorkbookBuilder()
    for name, word in (("Stock", "twin"), ("Stock ", "here"), (" Lead", "lead")):
        book.sheet(name).rows("A1", [["SKU", "Qty", "Bin"], [f"{word}-1", 1, "x"],
                                     [f"{word}-2", 2, "y"], [f"{word}-3", 3, "z"]])
    book.sheet("Sums")
    wb = Workbook.open(book.to_bytes())
    wb.set_cells({"Sums!A1": {"formula": "=SUM('Stock '!B2:B4)"},
                  "Sums!A2": {"formula": "=SUM(' Lead'!B2:B4)+SUM(Stock!B2:B4)"},
                  "'Stock '!E1": "noted"})
    wb.insert_rows("Stock ", 2)
    wb.merge("'Stock '!E1:F1")
    wb.set_comment("'Stock '!A1", "count again")
    wb.set_autofilter("Stock ", "A1:C5")
    wb.freeze_panes("Stock ", rows=1)
    wb.hide_columns(" Lead", ["C"])
    wb.rename_sheet(" Lead", "Lead  ")
    data = wb.to_bytes()
    outcome = excel.resave(data, name="r46x-spaces")
    assert outcome.ok, outcome
    saved = outcome.saved.read_bytes()
    assert oracle.sheet_names(saved) == ["Stock", "Stock ", "Lead  ", "Sums"]
    ours, theirs = _facts(data), _facts(saved)
    assert theirs["Sums"] == ours["Sums"]
    assert ("A2", "SUM('Lead  '!B2:B4)+SUM(Stock!B2:B4)") in theirs["Sums"]
    again = Workbook.open(saved)
    assert again.merges("Stock ") == wb.merges("Stock ")
    assert len(again.comments("Stock ")) == 1 and again.comments("Stock") == []
    assert again.cell("Sums!A1").v == 6 and again.cell("'Stock '!A3").v == "here-1"
    assert again.cell("Stock!A2").v == "twin-1"
