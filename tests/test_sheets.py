"""Sheets (E5, production R52): add, rename, delete, clear, hide -- as Excel does them.

On made-up workbooks (``tests/referents.py``: the corner shop with every kind of referent
naming its ``Lists`` sheet).  The expected texts are what Excel for Mac 16.106 wrote for
the same operation on the same workbook (round 44, lane S's probes; the module notes of
:mod:`xlsx_agent.edit.sheets` and :mod:`xlsx_agent.formula.sheets` say what was measured).
"""

from __future__ import annotations

import io

import pytest
from lxml import etree

from ooxml_edit.tools import Toolbox

from xlsx_agent import Workbook
from xlsx_agent.edit import EditError
from xlsx_agent.edit.sheets import SheetError, name_problems
from xlsx_agent.formula.sheets import delete, drop_tables, fold_name, refers_to, rename
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS

import referents
from referents import unzipped
from shop import Shop

X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
ORDER = ["Stock", "Lists", "Totals", "Orders"]


def _open(**options) -> tuple[bytes, Workbook]:
    data = referents.book(**options)
    return data, Workbook.open(data)


def _tree(wb: Workbook, part: str):
    return etree.fromstring(unzipped(wb.to_bytes())[part])


def _formulas(wb: Workbook, part: str) -> dict[str, str]:
    return {f.getparent().get("r"): f.text for f in _tree(wb, part).iter(X + "f") if f.text}


def _names(wb: Workbook) -> dict[tuple[str, str | None], str]:
    return {(n.get("name"), n.get("localSheetId")): n.text
            for n in _tree(wb, "xl/workbook.xml").iter(X + "definedName")}


def _saved(wb: Workbook) -> None:
    report = wb.save(io.BytesIO())
    assert report.issues == ()


# -- references in formulas (xlsx_agent.formula.sheets) ----------------------------------------

#: (formula, renamed Lists -> Tax rates, Lists deleted): what Excel wrote (measured).
MEASURED = [
    ("VLOOKUP(C10,Lists!$A$2:$B$5,2,FALSE)", "VLOOKUP(C10,'Tax rates'!$A$2:$B$5,2,FALSE)",
     "VLOOKUP(C10,#REF!,2,FALSE)"),
    ("SUM(Stock:Lists!B2)", "SUM('Stock:Tax rates'!B2)", "SUM(Stock!B2)"),
    ('INDIRECT("Lists!B2")', None, None),
    ("'Lists'!B3+Local", "'Tax rates'!B3+Local", "#REF!+Local"),
    ("COUNTA(Lists!A:A)+SUM(Lists!2:2)", "COUNTA('Tax rates'!A:A)+SUM('Tax rates'!2:2)",
     "COUNTA(#REF!)+SUM(#REF!)"),
    ("OFFSET(Lists!$A$2,0,0,COUNTA(Lists!$A:$A)-1,1)",
     "OFFSET('Tax rates'!$A$2,0,0,COUNTA('Tax rates'!$A:$A)-1,1)",
     "OFFSET(#REF!,0,0,COUNTA(#REF!)-1,1)"),
    ("Lists!$A$2,Totals!$B$1", "'Tax rates'!$A$2,Totals!$B$1", "#REF!,Totals!$B$1"),
    ("B1>Lists!$B$2", "B1>'Tax rates'!$B$2", "B1>#REF!"),
    ("Stock:Totals!$B$2", None, None),
    ("[1]Lists!A1", None, None),
]


@pytest.mark.parametrize("text, renamed, deleted", MEASURED)
def test_references_are_renamed_and_deleted_as_excel_writes_them(text, renamed, deleted):
    assert rename(text, "Lists", "Tax rates") == renamed
    assert delete(text, "Lists", ORDER) == deleted


def test_a_3d_run_is_quoted_whole_and_shrinks_inward_from_either_end():
    assert rename("SUM(Stock:Lists!B2)", "Lists", "a'b") == "SUM('Stock:a''b'!B2)"
    assert delete("Stock:Totals!$B$2", "Stock", ORDER) == "Lists:Totals!$B$2"
    assert delete("Stock:Lists!$B$2", "Stock", ORDER) == "Lists!$B$2"
    assert delete("Stock:Totals!$B$2", "Lists", ORDER) is None
    assert refers_to("Stock:Totals!$B$2", "Lists", ORDER)


def test_names_are_matched_as_excel_compares_them():
    assert rename("lists!A1+LISTS!B2", "Lists", "Rates") == "Rates!A1+Rates!B2"
    assert fold_name("a\u00adb") == fold_name("AB")
    assert rename("Lists!#REF!+Lists!Local", "Lists", "Q3 sales") == \
        "'Q3 sales'!#REF!+'Q3 sales'!Local"
    assert delete("Lists!Local*2", "Lists", ORDER) == "#REF!*2"


def test_a_structured_reference_to_a_table_that_goes_is_ref():
    assert drop_tables("SUM(Rates[Tax])+rates[[#This Row],[Tax]]", {"rates"}) == \
        "SUM(#REF!)+#REF!"
    assert drop_tables("SUM(Other[Tax])", {"rates"}) is None


# -- names ----------------------------------------------------------------------------------------

#: Refused by Excel's object model (measured), and taken.
REFUSED = ["a:b", "a/b", "a\\b", "a?b", "a*b", "a[b", "a]b", "'ab", "ab'", "", "x" * 32,
           "STOCK", "stock"]
TAKEN = ["a'b", " ", " ab", "ab ", "History", "x" * 31, 'a"b', "a!b", "a#b", "a,b", "a;b",
         "a<b", "a{b", "R1C1", "A1", "TRUE", "a\x7fb"]


@pytest.mark.parametrize("name", REFUSED)
def test_a_name_excel_refuses_is_refused(name):
    wb = Workbook.open(Shop().to_bytes())
    assert name_problems(wb, name)
    before = wb.to_bytes()
    with pytest.raises(SheetError) as raised:
        wb.add_sheet(name)
    assert raised.value.code in ("bad_sheet_name", "sheet_exists")
    assert wb.to_bytes() == before


@pytest.mark.parametrize("name", TAKEN)
def test_a_name_excel_takes_is_taken(name):
    wb = Workbook.open(Shop().to_bytes())
    assert name_problems(wb, name) == []


@pytest.mark.parametrize("name", ["a\x01b", "a\tb", "a\nb", "a\rb", "_x0041_"])
def test_control_characters_and_escape_like_text_are_refused_in_a_name(name):
    wb = Workbook.open(Shop().to_bytes())
    with pytest.raises(SheetError) as raised:
        wb.rename_sheet("Totals", name)
    assert raised.value.code == "bad_sheet_name"


# -- add ------------------------------------------------------------------------------------------


def test_a_sheet_is_added_as_excel_makes_one_and_the_active_tab_stays():
    _, wb = _open()
    _set_active(wb, 2)
    data = wb.to_bytes()
    change = wb.add_sheet("Rota", 1)
    assert change.sheets == ("Stock", "Rota", "Lists", "Totals", "Orders")
    assert change.id == 5
    sheet = wb.sheet("Rota")
    root = _tree(wb, sheet.part)
    assert [etree.QName(n).localname for n in root] == \
        ["dimension", "sheetViews", "sheetFormatPr", "sheetData", "pageMargins"]
    assert root.find(X + "dimension").get("ref") == "A1"
    margins = root.find(X + "pageMargins").attrib
    assert dict(margins) == {"left": "0.7", "right": "0.7", "top": "0.75", "bottom": "0.75",
                             "header": "0.3", "footer": "0.3"}
    assert root.find(X + "sheetFormatPr").get("defaultRowHeight") == "14.5"
    names = _names(wb)
    assert ("Local", "2") in names and ("_xlnm.Print_Area", "0") in names
    view = _tree(wb, "xl/workbook.xml").find(f"{X}bookViews/{X}workbookView")
    assert view.get("activeTab") == "3"           # Totals, still active
    assert wb.validate() == []
    _saved(wb)
    wb.undo()
    assert wb.to_bytes() == data


def test_a_sheet_is_added_at_the_end_by_default_and_refused_off_the_tabs():
    _, wb = _open()
    assert wb.add_sheet("Rota").sheets[-1] == "Rota"
    with pytest.raises(SheetError) as raised:
        wb.add_sheet("More", 9)
    assert raised.value.code == "bad_index"


def _set_active(wb: Workbook, index: int) -> None:
    tree = wb.package.tree(wb.part)
    tree.find(f"{X}bookViews/{X}workbookView").set("activeTab", str(index))
    wb.package.mark_dirty(wb.part)
    wb._forget()


# -- rename ---------------------------------------------------------------------------------------


def test_a_rename_rewrites_every_reference_excel_rewrites():
    data, wb = _open(chart=True, pivot=True)
    change = wb.rename_sheet("Lists", "Tax rates")
    assert change.sheet == "Tax rates" and change.id == 2
    names = _names(wb)
    assert names[("Categories", None)] == "'Tax rates'!$A$2:$A$5"
    assert names[("Span", None)] == "'Stock:Tax rates'!$B$2"
    assert names[("Local", "1")] == "'Tax rates'!$B$2"
    totals = _formulas(wb, "xl/worksheets/sheet3.xml")
    assert totals["B6"] == "SUM('Stock:Tax rates'!B2)"
    assert totals["C6"] == 'INDIRECT("Lists!B2")'           # a string: no reference
    assert totals["D6"] == "'Tax rates'!B3+Local"
    assert _formulas(wb, "xl/worksheets/sheet1.xml")["I10"] == \
        "VLOOKUP(C10,'Tax rates'!$A$2:$B$5,2,FALSE)"
    stock = etree.tostring(_tree(wb, "xl/worksheets/sheet1.xml")).decode()
    assert "'Tax rates'!$A$2:$A$5</xm:f>" in stock          # the x14 validation
    orders = etree.tostring(_tree(wb, "xl/worksheets/sheet4.xml")).decode()
    assert "'Tax rates'!$B$5</xm:f>" in orders
    totals_xml = _tree(wb, "xl/worksheets/sheet3.xml")
    assert totals_xml.find(f"{X}conditionalFormatting/{X}cfRule/{X}formula").text == \
        "B1>'Tax rates'!$B$2"
    assert totals_xml.find(f"{X}hyperlinks/{X}hyperlink").get("location") == "Lists!A1"
    chart = _tree(wb, "xl/charts/chart1.xml")
    assert [f.text for f in chart.iter(C + "f")] == [
        "'Tax rates'!$B$1", "'Tax rates'!$B$1", "'Tax rates'!$A$2:$A$5",
        "'Tax rates'!$B$2:$B$5"]
    cache = _tree(wb, "xl/pivotCache/pivotCacheDefinition1.xml")
    assert cache.find(f"{X}cacheSource/{X}worksheetSource").get("sheet") == "Tax rates"
    kinds = {r.kind for r in change.rewritten}
    assert kinds == {"name", "formula", "validation", "conditional_format", "chart",
                     "pivot_cache"}
    assert [r.as_dict() for r in change.not_rewritten] == [{"kind": "hyperlink",
                                                            "a": "Totals!A6"}]
    assert wb.sheet("Tax rates").id == 2
    assert wb.cell("'Tax rates'!A2").v == "Stationery"
    assert wb.validate() == []
    _saved(wb)
    wb.undo()
    assert wb.to_bytes() == data


def test_a_rename_to_the_same_name_is_no_change_and_a_taken_name_is_refused():
    _, wb = _open()
    change = wb.rename_sheet("Lists", "Lists")
    assert change.changed is False and change.note
    with pytest.raises(SheetError) as raised:
        wb.rename_sheet("Lists", "TOTALS")
    assert raised.value.code == "sheet_exists" and "Totals" in raised.value.options
    assert wb.rename_sheet("Lists", "LISTS").sheet == "LISTS"    # its own name, in capitals


def test_a_rename_brings_the_titles_in_docprops_up_to_date():
    found = referents.parts()
    found["docProps/app.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Properties xmlns="http:'
        '//schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="'
        'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TitlesOfParts>'
        '<vt:vector size="3" baseType="lpstr"><vt:lpstr>Stock</vt:lpstr><vt:lpstr>Lists'
        '</vt:lpstr><vt:lpstr>Lists!Local</vt:lpstr></vt:vector></TitlesOfParts></Properties>'
    ).encode()
    wb = Workbook.open(referents.zipped(found))
    wb.rename_sheet("Lists", "Tax rates")
    app = unzipped(wb.to_bytes())["docProps/app.xml"].decode()
    assert "<vt:lpstr>Tax rates</vt:lpstr>" in app
    assert "<vt:lpstr>'Tax rates'!Local</vt:lpstr>" in app


def test_a_table_formula_is_rewritten_by_a_rename():
    found = Shop(table=True).parts()
    body = found["xl/tables/table1.xml"].decode().replace(
        '<tableColumn id="2" name="Tax"/>',
        '<tableColumn id="2" name="Tax"><calculatedColumnFormula>Totals!$B$3'
        '</calculatedColumnFormula></tableColumn>')
    found["xl/tables/table1.xml"] = body.encode()
    wb = Workbook.open(referents.zipped(found))
    change = wb.rename_sheet("Totals", "Sums")
    assert any(r.kind == "table_formula" for r in change.rewritten)
    assert b"<calculatedColumnFormula>Sums!$B$3" in unzipped(wb.to_bytes())[
        "xl/tables/table1.xml"]


# -- delete ---------------------------------------------------------------------------------------


def test_a_delete_is_refused_listing_every_referent_by_address():
    data, wb = _open(chart=True, pivot=True)
    with pytest.raises(SheetError) as raised:
        wb.delete_sheet("Lists")
    error = raised.value
    assert error.code == "referred"
    found = {(r.kind, r.address) for r in error.referents}
    assert {("name", "Categories"), ("formula", "Stock!I10"), ("formula", "Totals!B6"),
            ("validation", "Stock!C10:C200"), ("conditional_format", "Totals!B1"),
            ("hyperlink", "Totals!A6"), ("validation", "Orders!K2:K50"),
            ("chart", "Totals!xl/charts/chart1.xml"),
            ("pivot_cache", "xl/pivotCache/pivotCacheDefinition1.xml")} <= found
    assert ("formula", "Totals!C6") not in found        # INDIRECT's string
    assert ("name", "Lists!Local") not in found         # goes with the sheet
    assert "clear" in str(error)
    assert wb.to_bytes() == data
    assert {(r.kind, r.address) for r in wb.sheet_referents("Lists")} == found


def test_a_forced_delete_writes_ref_as_excel_does():
    _, wb = _open(chart=True, table_formula=True)
    _set_active(wb, 1)
    data = wb.to_bytes()
    change = wb.delete_sheet("Lists", force=True)
    assert change.sheets == ("Stock", "Totals", "Orders")
    names = _names(wb)
    assert names[("Categories", None)] == "#REF!"
    assert names[("Span", None)] == "Stock!$B$2"
    assert names[("Both", None)] == "#REF!,Totals!$B$1"
    assert ("Local", "1") not in names and "name:Local" in change.removed
    assert ("_xlnm.Print_Area", "0") in names
    totals_part = wb.sheet("Totals").part
    totals = _formulas(wb, totals_part)
    assert totals["B6"] == "SUM(Stock!B2)" and totals["D6"] == "#REF!+Local"
    assert totals["E6"] == "COUNTA(#REF!)+SUM(#REF!)" and totals["F6"] == "SUM(#REF!)"
    assert totals["C6"] == 'INDIRECT("Lists!B2")'
    assert [f.text for f in _tree(wb, "xl/charts/chart1.xml").iter(C + "f")] == ["#REF!"] * 4
    assert {r.kind for r in change.not_rewritten} == {"hyperlink"}
    view = _tree(wb, "xl/workbook.xml").find(f"{X}bookViews/{X}workbookView")
    assert view.get("activeTab") == "1"                 # Totals, now at Lists' place
    assert "xl/tables/table1.xml" not in wb.package.part_names
    assert wb.calculation_status == "invalidated"
    assert wb.validate() == []
    _saved(wb)
    wb.undo()
    assert wb.to_bytes() == data


def test_a_deleted_sheet_s_pivot_table_takes_its_unused_cache_with_it():
    _, wb = _open(pivot=True)
    wb.rename_sheet("Lists", "Kept")
    change = wb.delete_sheet("Totals", force=True)
    assert any(item.startswith("pivot_cache:") for item in change.removed)
    root = _tree(wb, "xl/workbook.xml")
    assert root.find(X + "pivotCaches") is None
    assert not [p for p in wb.package.part_names if "pivot" in p]
    assert wb.validate() == []
    _saved(wb)


def test_the_last_sheet_and_the_last_visible_sheet_are_never_deleted():
    wb = Workbook.open(Shop().to_bytes())          # Lists is hidden
    wb.set_sheet_state("Totals", "hidden")
    with pytest.raises(SheetError) as raised:
        wb.delete_sheet("Stock", force=True)
    assert raised.value.code == "last_visible"
    only = Workbook.open(_one_sheet())
    with pytest.raises(SheetError) as raised:
        only.delete_sheet("Stock")
    assert raised.value.code == "last_sheet"


def _one_sheet() -> bytes:
    found = Shop().parts()
    body = found["xl/workbook.xml"].decode()
    body = body.replace('<sheet name="Lists" sheetId="2" state="hidden" r:id="rId2"/>', '')
    body = body.replace('<sheet name="Totals" sheetId="3" r:id="rId3"/>', '')
    found["xl/workbook.xml"] = body.encode()
    return referents.zipped(found)


# -- clear ----------------------------------------------------------------------------------------


def test_a_clear_removes_what_clear_all_removes_and_keeps_the_rest():
    data, wb = _open()
    change = wb.clear_sheet("Stock")
    root = _tree(wb, "xl/worksheets/sheet1.xml")
    assert root.find(f"{X}sheetData/{X}row/{X}c") is None
    for gone in ("mergeCells", "conditionalFormatting", "dataValidations"):
        assert root.find(X + gone) is None
    assert b"x14:dataValidation" not in etree.tostring(root)
    for kept in ("sheetPr", "sheetViews", "sheetFormatPr", "cols", "printOptions",
                 "pageMargins", "pageSetup", "headerFooter", "rowBreaks"):
        assert root.find(X + kept) is not None, kept
    rows = root.findall(f"{X}sheetData/{X}row")
    assert [(r.get("r"), r.get("ht"), r.get("customHeight")) for r in rows] == \
        [("1", "21", "1")]
    assert root.find(X + "dimension").get("ref") == "A1"
    assert b"corner-shop" in etree.tostring(root)              # an unknown extension stays
    assert change.cells == 58
    assert _names(wb)[("_xlnm.Print_Area", "0")] == "Stock!$A$1:$I$14"
    assert _formulas(wb, "xl/worksheets/sheet3.xml")["B6"] == "SUM(Stock:Lists!B2)"
    assert wb.validate() == []
    _saved(wb)
    again = wb.clear_sheet("Stock")
    assert again.changed is False
    wb.undo()
    assert wb.to_bytes() == data


def test_a_clear_takes_the_sheet_s_hyperlinks_rules_and_pivot_table_and_keeps_its_chart():
    _, wb = _open(chart=True, pivot=True)
    change = wb.clear_sheet("Totals")
    root = _tree(wb, "xl/worksheets/sheet3.xml")
    assert root.find(X + "hyperlinks") is None and root.find(X + "drawing") is not None
    assert "xl/charts/chart1.xml" in wb.package.part_names
    assert not [p for p in wb.package.part_names if "pivot" in p]
    assert "hyperlinks" in change.removed
    assert wb.validate() == []
    _saved(wb)


def test_a_clear_of_a_table_others_name_is_refused_unless_forced():
    data, wb = _open(table_formula=True)
    with pytest.raises(SheetError) as raised:
        wb.clear_sheet("Lists")
    assert [(r.kind, r.address) for r in raised.value.referents] == [("formula", "Totals!F6")]
    assert wb.to_bytes() == data
    change = wb.clear_sheet("Lists", force=True)
    assert "table:Rates" in change.removed
    assert _formulas(wb, "xl/worksheets/sheet3.xml")["F6"] == "SUM(#REF!)"
    assert "xl/tables/table1.xml" not in wb.package.part_names
    assert wb.validate() == []
    _saved(wb)


def test_a_protected_sheet_is_not_cleared():
    wb = Workbook.open(Shop(protected=True).to_bytes())
    with pytest.raises(EditError) as raised:
        wb.clear_sheet("Totals")
    assert raised.value.code == "protected"


# -- hide and show -------------------------------------------------------------------------------


def test_hiding_the_active_sheet_moves_the_active_tab_right_else_left():
    _, wb = _open()
    wb.set_sheet_state("Stock", "hidden")
    view = _tree(wb, "xl/workbook.xml").find(f"{X}bookViews/{X}workbookView")
    assert view.get("activeTab") == "1"
    assert _tree(wb, "xl/worksheets/sheet1.xml").find(
        f"{X}sheetViews/{X}sheetView").get("tabSelected") is None
    assert _tree(wb, "xl/worksheets/sheet2.xml").find(
        f"{X}sheetViews/{X}sheetView").get("tabSelected") == "1"
    _, wb = _open()
    _set_active(wb, 3)
    wb.set_sheet_state("Orders", "veryHidden")
    view = _tree(wb, "xl/workbook.xml").find(f"{X}bookViews/{X}workbookView")
    assert view.get("activeTab") == "2"
    assert wb.sheet("Orders").state == "veryHidden"
    assert wb.validate() == []


def test_the_last_visible_sheet_is_never_hidden_and_a_state_already_so_is_no_change():
    wb = Workbook.open(Shop().to_bytes())
    wb.set_sheet_state("Stock", "hidden")
    with pytest.raises(SheetError) as raised:
        wb.set_sheet_state("Totals", "hidden")
    assert raised.value.code == "last_visible"
    assert wb.set_sheet_state("Lists", "hidden").changed is False
    assert wb.set_sheet_state("Lists", "visible").changed is True
    assert 'state=' not in _tree(wb, "xl/workbook.xml").find(
        f"{X}sheets/{X}sheet[@name='Lists']").attrib
    with pytest.raises(SheetError):
        wb.set_sheet_state("Lists", "shown")


# -- guards ---------------------------------------------------------------------------------------


def test_a_locked_structure_refuses_every_change_of_sheets_but_a_clear():
    wb = Workbook.open(Shop(locked=True).to_bytes())
    for call in (lambda: wb.add_sheet("Rota"), lambda: wb.rename_sheet("Totals", "Sums"),
                 lambda: wb.delete_sheet("Totals", force=True),
                 lambda: wb.set_sheet_state("Totals", "hidden")):
        with pytest.raises(SheetError) as raised:
            call()
        assert raised.value.code == "structure_locked"
    assert wb.clear_sheet("Totals").changed


def test_the_host_policy_must_cover_a_sheet_whole():
    wb = Workbook.open(Shop().to_bytes(), policy={"cells": ["Totals!A1:B3", "Lists!"]})
    with pytest.raises(SheetError) as raised:
        wb.clear_sheet("Totals")
    assert raised.value.code == "not_writable"
    assert wb.set_sheet_state("Lists", "visible").changed


def test_sheets_are_named_by_id_quoted_and_withheld_ones_reach_the_host():
    _, wb = _open()
    wb.withhold_sheets(["Orders"])
    assert wb.rename_sheet("#4", "Book of orders").sheet == "Book of orders"
    assert wb.rename_sheet("'Book of orders'", "Orders").sheet == "Orders"
    with pytest.raises(SheetError) as raised:
        wb.set_sheet_state("Nope", "hidden")
    assert raised.value.code == "no_sheet" and "Orders" in raised.value.options


def test_steps_in_one_batch_are_undone_together():
    data, wb = _open()
    with wb.batch():
        wb.add_sheet("Rota")
        wb.rename_sheet("Rota", "Rota 2")
        wb.set_cells({"'Rota 2'!A1": "shift"})
        wb.set_sheet_state("Lists", "hidden")
    assert wb.cell("'Rota 2'!A1").v == "shift"
    _saved(wb)
    wb.undo()
    assert wb.to_bytes() == data


# -- the tool -------------------------------------------------------------------------------------


@pytest.fixture
def session():
    toolbox = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS)
    current = toolbox.session()
    blob = current.add_blob(referents.book(chart=True), name="shop.xlsx")
    assert toolbox.dispatch(current, "open_document", {"blob": blob}).ok
    return toolbox, current


def _call(session, **args):
    toolbox, current = session
    return toolbox.dispatch(current, "xlsx_manage_sheets", {"doc": "d1", **args})


def test_the_tool_refuses_a_delete_with_the_referents_and_the_way_out(session):
    result = _call(session, action="delete", sheet="Lists")
    assert not result.ok
    details = result.error.details
    assert details["code"] == "referred"
    assert {"kind": "formula", "a": "Totals!B6"} in details["referents"]
    assert "clear" in result.error.message and "force" in result.error.message
    forced = _call(session, action="delete", sheet="Lists", force=True)
    assert forced.ok and "Lists" not in forced.data["sheets"]


def test_the_tool_takes_intent_and_a_no_op_is_no_error(session):
    added = _call(session, action="add", sheet="Rota", position=1)
    assert added.ok and added.data["sheets"][0] == "Rota" and added.data["notes"]
    again = _call(session, action="hide", sheet="Lists")
    assert again.ok
    same = _call(session, action="hide", sheet="Lists")
    assert same.ok and same.data["changed"] is False and same.changed == []
    missing = _call(session, action="clear", sheet="Nope")
    assert not missing.ok and "Stock" in missing.error.valid_options
    bad = _call(session, action="rename", sheet="Rota", name="a/b")
    assert not bad.ok and bad.error.field == "name"


def test_the_tool_never_reaches_or_names_a_withheld_sheet(session):
    toolbox, current = session
    current.documents["d1"].document.withhold_sheets(["Orders"])
    missing = _call(session, action="hide", sheet="Orders")
    assert not missing.ok and "Orders" not in missing.error.valid_options
    refused = _call(session, action="delete", sheet="Lists")
    assert refused.error.details.get("referents_withheld") == 1
    assert all(not r["a"].startswith("Orders!") for r in refused.error.details["referents"])
    taken = _call(session, action="rename", sheet="Lists", name="orders")
    assert not taken.ok and "Orders" not in taken.error.message


def test_the_tool_is_deferred_in_the_edit_group():
    toolbox = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS)
    core = {d["name"] for d in toolbox.definitions("anthropic", groups="core")}
    assert "xlsx_manage_sheets" not in core
    every = {d["name"] for d in toolbox.definitions("anthropic")}
    assert "xlsx_manage_sheets" in every


def test_a_long_list_of_referents_is_cut_and_counted_never_silently():
    from ooxml_edit.tools.limits import Limits

    _, wb = _open()
    change = wb.rename_sheet("Lists", "Tax rates")
    shown = change.as_dict(limit=3)
    assert len(shown["rewritten"]) == 3
    assert shown["rewritten_cut"] == len(change.rewritten) - 3
    toolbox = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, limits=Limits(max_list_items=2))
    current = toolbox.session()
    blob = current.add_blob(referents.book(), name="shop.xlsx")
    assert toolbox.dispatch(current, "open_document", {"blob": blob}).ok
    refused = toolbox.dispatch(current, "xlsx_manage_sheets",
                               {"doc": "d1", "action": "delete", "sheet": "Lists"})
    details = refused.error.details
    assert len(details["referents"]) == 2
    assert details["referents_cut"] == len(Workbook.open(referents.book()).sheet_referents("Lists")) - 2
