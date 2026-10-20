"""The edit foundation (0.2, E1): history, atomic batches, undo and redo to the bytes, the
reports, the structure check, the checked save, and the shared tools a workbook gains.

Every fixture is the corner shop's stock list (``tests/shop.py``), made in code; the cells
are written with E3's ``set_cells`` (its own checks: ``tests/test_set_cells.py``).
"""

from __future__ import annotations

import datetime as dt
import io
import json
import logging
import zipfile

import pytest

from ooxml_edit.sheets import Formula, outside_footprint
from ooxml_edit.tools import Result, Toolbox, integer, number, string, tool

from xlsx_agent import Workbook
from xlsx_agent.address import parse_address
from xlsx_agent.edit import EditError, compress
from xlsx_agent.edit.report import MAX_RANGES
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS

from shop import Shop

VARIANTS = {"plain": {}, "rless": {"rless": True}, "1904": {"date1904": True},
            "inline": {"shared": False}, "table": {"table": True}}

#: What one step writes: each kind of value, a cell made, a row made, a blank.
WRITES = {
    "Stock!B10": "Spare fan",
    "Stock!D10": 12,
    "Stock!E11": 1.25,
    "Stock!G12": True,
    "Stock!H10": dt.date(2026, 10, 20),
    "Stock!C11": "=not a formula",
    "Stock!J10": Formula("D10*2"),
    "Stock!B15": "Torch, spare",
    "Stock!D14": None,
}


def _parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _pack(parts: dict[str, bytes]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(zipfile.ZipInfo(name, date_time=(2026, 1, 2, 3, 4, 6)), data)
    return buffer.getvalue()


def _edit(wb: Workbook, writes=WRITES):
    with wb.batch() as edit:
        for address, value in writes.items():
            wb.set_cells({address: value})
    return edit


# -- undo, redo, no-op -------------------------------------------------------------------------


@pytest.mark.parametrize("options", VARIANTS.values(), ids=VARIANTS.keys())
def test_undo_gives_the_original_bytes_and_redo_the_edited(options):
    original = Shop(**options).to_bytes()
    wb = Workbook.open(original)
    assert wb.version == 0 and wb.calculation_status == "cached_unverified"
    edit = _edit(wb)
    edited = wb.to_bytes()
    assert edited != original and wb.version == edit.report.version == 1
    assert wb.undo() and wb.version == 0 and wb.to_bytes() == original
    assert wb.calculation_status == "cached_unverified"
    assert not wb.undo()
    assert wb.redo() and wb.version == 1 and wb.to_bytes() == edited
    assert wb.calculation_status == "invalidated"
    assert not wb.redo()


def test_two_steps_undo_one_at_a_time():
    original = Shop().to_bytes()
    wb = Workbook.open(original)
    wb.set_cells({"Stock!D10": 7})
    first = wb.to_bytes()
    wb.set_cells({"Lists!B2": 0.07})
    second = wb.to_bytes()
    assert wb.version == 2
    assert wb.undo() and wb.to_bytes() == first and wb.version == 1
    assert wb.undo() and wb.to_bytes() == original
    assert wb.redo() and wb.redo() and wb.to_bytes() == second and wb.version == 2


def test_a_no_op_is_the_original_bytes_and_no_step():
    original = Shop().to_bytes()
    wb = Workbook.open(original)
    with wb.batch() as edit:
        assert not wb.set_cells({"Stock!D10": 120}).changed     # what it holds already
        assert not wb.set_cells({"Stock!B4": None}).changed     # a styled blank stays one
        wb.read("Stock!A1:C3")                        # a read inside a batch
    assert not edit.report.changed and edit.report.parts == () and edit.report.cells == {}
    assert wb.version == 0 and not wb.history.can_undo() and wb.to_bytes() == original
    with wb.batch():
        pass
    assert wb.version == 0 and wb.to_bytes() == original


def test_reads_after_a_write_see_it():
    wb = Workbook.open(Shop().to_bytes())
    assert wb.cell("Stock!D10").v == 120
    with wb.batch():
        wb.set_cells({"Stock!D10": 99})
    assert wb.cell("Stock!D10").v == 99
    wb.undo()
    assert wb.cell("Stock!D10").v == 120


# -- atomic ------------------------------------------------------------------------------------


@pytest.mark.parametrize("bad", [
    ("Stock!D11", float("nan")),                    # not finite
    ("Stock!B12", "x" * 32_768),                    # longer than a cell holds
    ("Stock!F10", 1),                               # a shared formula group's master
    ("Stock!D13", object()),                        # not a cell value
], ids=["nan", "too-long", "shared-master", "type"])
def test_one_invalid_write_rolls_back_every_write(bad):
    original = Shop().to_bytes()
    wb = Workbook.open(original)
    before = wb.validate()
    address, value = bad
    with pytest.raises((ValueError, TypeError)):
        with wb.batch():
            wb.set_cells({"Stock!B10": "Spare fan"})        # a new shared string
            wb.set_cells({"Stock!J10": Formula("D10*2")})   # drops the calculation chain
            wb.set_cells({address: value})
            wb.set_cells({"Stock!D12": 4})
    assert wb.to_bytes() == original and wb.version == 0
    assert not wb.history.can_undo() and not wb.package.changed_parts()
    # the calculation metadata is back: the chain, its relationship and content type, calcPr
    assert wb.package.has_part("xl/calcChain.xml")
    assert any(rel.type.endswith("/calcChain")
               for rel in wb.package.relationships("xl/workbook.xml").values())
    assert wb.package.content_type("xl/calcChain.xml") is not None
    assert wb.calculation_status == "cached_unverified" and wb.calculation == []
    assert wb.validate() == before and wb.cell("Stock!B10").v == "Pencil"


def test_a_rollback_after_a_step_keeps_that_step():
    wb = Workbook.open(Shop().to_bytes())
    wb.set_cells({"Stock!D10": 7})
    kept = wb.to_bytes()
    with pytest.raises(ValueError):
        with wb.batch():
            wb.set_cells({"Stock!D11": 8})
            wb.set_cells({"Stock!D12": float("inf")})
    assert wb.to_bytes() == kept and wb.version == 1 and wb.cell("Stock!D11").v == 40


def test_an_error_caught_inside_a_batch_still_rolls_it_back():
    original = Shop().to_bytes()
    wb = Workbook.open(original)
    with pytest.raises(EditError) as raised:
        with wb.batch():
            wb.set_cells({"Stock!D10": 7})
            try:
                wb.set_cells({"Stock!D11": float("nan")})
            except ValueError:
                pass
            wb.set_cells({"Stock!D12": 9})
    assert raised.value.code == "rolled_back"
    assert wb.to_bytes() == original and wb.version == 0


def test_a_savepoint_rolls_back_alone():
    wb = Workbook.open(Shop().to_bytes())
    with wb.batch() as edit:
        wb.set_cells({"Stock!D10": 7})
        with pytest.raises(ValueError):
            with wb.savepoint():
                wb.set_cells({"Stock!D11": 8})
                wb.set_cells({"Stock!D12": float("nan")})
        wb.set_cells({"Stock!D13": 9})
    assert edit.report.cells == {"Stock": ["D10", "D13"]}
    assert [wb.cell(a).v for a in ("Stock!D10", "Stock!D11", "Stock!D13")] == [7, 40, 9]
    assert wb.version == 1


def test_undo_waits_for_the_batch():
    wb = Workbook.open(Shop().to_bytes())
    with pytest.raises(ValueError):     # the EditError rolls the batch back too
        with wb.batch():
            wb.set_cells({"Stock!D10": 7})
            wb.undo()
    assert wb.version == 0


def test_a_strict_workbook_is_not_written():
    data = Shop().to_bytes()
    parts = _parts(data)
    for name in ("xl/workbook.xml",):
        parts[name] = parts[name].replace(
            b"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            b"http://purl.oclc.org/ooxml/spreadsheetml/main")
    wb = Workbook.open(_pack(parts))
    with pytest.raises(EditError) as raised:
        wb.set_cells({"Stock!D10": 7})
    assert raised.value.code == "strict" and wb.version == 0


# -- the footprint -----------------------------------------------------------------------------


@pytest.mark.parametrize("options", VARIANTS.values(), ids=VARIANTS.keys())
def test_edit_undo_redo_stay_inside_the_footprint(options):
    original = Shop(**options).to_bytes()
    wb = Workbook.open(original)
    _edit(wb)
    wb.undo()
    wb.redo()
    edited = wb.to_bytes()
    cells: dict[str, list[str]] = {}
    for address in WRITES:
        sheet, _, ref = address.partition("!")
        cells.setdefault(sheet, []).append(ref)
    assert outside_footprint(original, edited, cells=cells) == []
    before, after = _parts(original), _parts(edited)
    changed = {name for name in before if before[name] != after.get(name)}
    assert changed <= {"xl/worksheets/sheet1.xml", "xl/sharedStrings.xml", "xl/workbook.xml",
                       "xl/_rels/workbook.xml.rels", "[Content_Types].xml", "xl/calcChain.xml"}
    sheet = after["xl/worksheets/sheet1.xml"]
    for kept in (b"x14:dataValidations", b"Lists!$A$2:$A$5", b"keep me as I am",
                 b"mc:AlternateContent", b"x14ac:dyDescent", b'shop:template="stock-2"',
                 b"conditionalFormatting", b"mergeCell", b"rowBreaks", b"oddFooter"):
        assert kept in sheet, kept
    assert after["xl/printerSettings/printerSettings1.bin"] == \
        before["xl/printerSettings/printerSettings1.bin"]


# -- the report --------------------------------------------------------------------------------


def test_the_report_says_what_changed_by_range_and_part():
    wb = Workbook.open(Shop().to_bytes())
    report = _edit(wb).report
    assert report.before == 0 and report.version == 1 and report.changed
    assert report.cells == {"Stock": ["B10", "D10", "H10", "J10", "C11", "E11", "G12",
                                      "D14", "B15"]}
    assert report.cell_count == 9 and report.cells_cut == 0
    assert report.parts == ("[Content_Types].xml", "xl/_rels/workbook.xml.rels",
                            "xl/calcChain.xml", "xl/sharedStrings.xml", "xl/workbook.xml",
                            "xl/worksheets/sheet1.xml")
    assert report.calculation == "invalidated" and report.issues == ()
    assert wb.last_edit == report
    assert json.loads(json.dumps(report.as_dict()))["calculation"] == "invalidated"


def test_a_number_written_over_a_number_invalidates_too():
    wb = Workbook.open(Shop(calc_chain=False).to_bytes())
    report = _edit(wb, {"Lists!B2": 0.07}).report
    assert report.calculation == "invalidated"
    assert report.parts == ("xl/workbook.xml", "xl/worksheets/sheet2.xml")
    assert wb.calculation == ["full_calc_on_load"]


def test_ranges_are_compressed_and_capped():
    assert compress([(1, 1), (1, 2), (2, 1), (2, 2), (4, 4), (2, 5)]) == ["A1:B2", "E2", "D4"]
    assert compress([(3, 2), (4, 2), (5, 2), (5, 3)]) == ["B3:B4", "B5:C5"]
    wb = Workbook.open(Shop().to_bytes())
    writes = {f"Stock!{'KMOQ'[n % 4]}{20 + 2 * n}": n for n in range(60)}
    report = _edit(wb, writes).report
    assert report.cell_count == 60 and len(report.cells["Stock"]) == MAX_RANGES
    assert report.cells_cut == 10


# -- no content in diagnostics -----------------------------------------------------------------


def test_reports_errors_and_logs_carry_no_content(caplog):
    marker = "MARKER-7q"
    caplog.set_level(logging.DEBUG)
    wb = Workbook.open(Shop().to_bytes())
    texts = []
    report = _edit(wb, {"Stock!B10": f"{marker} text", "Stock!J10":
                        Formula(f'"{marker}"&B10'), "Stock!B11": f"={marker}"}).report
    texts.append(json.dumps(report.as_dict()))
    for bad in (f"{marker}" * 5000, float("nan")):
        with pytest.raises(Exception) as raised:
            with wb.batch():
                wb.set_cells({"Stock!B12": f"{marker} again"})
                wb.set_cells({"Stock!B13": bad})
        texts.append(str(raised.value))
    wb.undo()
    wb.redo()
    saved = wb.save(io.BytesIO())
    texts.append(json.dumps(saved.as_dict()))
    with pytest.raises(EditError) as raised:     # a save that does not read back
        from xlsx_agent.edit import verify
        verify(wb, Shop().to_bytes(), "xlsx")
    texts.append(str(raised.value))
    texts.append(str(wb.validate()))
    texts.extend(record.getMessage() for record in caplog.records)
    assert any("edit 0 -> 1" in record.getMessage() for record in caplog.records)
    assert any("rolled back" in record.getMessage() for record in caplog.records)
    for text in texts:
        assert marker not in text and "Spare" not in text


# -- validate ----------------------------------------------------------------------------------


@pytest.mark.parametrize("options", VARIANTS.values(), ids=VARIANTS.keys())
def test_every_fixture_validates_clean_and_stays_so_after_an_edit(options):
    wb = Workbook.open(Shop(**options).to_bytes())
    assert wb.validate() == []
    _edit(wb)
    assert wb.validate() == [] and wb.last_edit.issues == ()
    assert Workbook.open(wb.to_bytes()).validate() == []


def _broken(change, **options):
    parts = _parts(Shop(**options).to_bytes())
    change(parts)
    return Workbook.open(_pack(parts)).validate()


def _replace(part, old, new):
    def change(parts):
        assert old in parts[part], old
        parts[part] = parts[part].replace(old, new, 1)
    return change


SHEET1 = "xl/worksheets/sheet1.xml"
SHEET2 = "xl/worksheets/sheet2.xml"
WB = "xl/workbook.xml"


@pytest.mark.parametrize("change, code, address", [
    (_replace(SHEET2, b'<c r="A2" t="s"><v>', b'<c r="A2" t="s"><v>9'), "shared_string", "Lists!A2"),
    (_replace(SHEET2, b'<row r="3">', b'<row r="1">'), "row_order", "Lists!1:1"),
    (_replace(SHEET2, b'<c r="B2"', b'<c r="A2"'), "cell_order", "Lists!A2"),
    (_replace(SHEET2, b'<c r="B2"', b'<c r="B7"'), "cell_row", "Lists!B7"),
    (_replace(SHEET2, b'<c r="B2"', b'<c r="B0"'), "bad_reference", "Lists!B2"),
    (_replace(SHEET2, b'<c r="B2" s="7"', b'<c r="B2" s="70"'), "style", "Lists!B2"),
    (_replace(SHEET2, b'<dimension ref="A1:B5"/>', b'<dimension ref="A1:B3"/>'), "dimension",
     "Lists!"),
    (_replace("xl/sharedStrings.xml", b'uniqueCount="', b'uniqueCount="9'),
     "shared_string_count", "xl/sharedStrings.xml"),
    (_replace(WB, b'<definedName name="TaxRates">', b'<definedName name="categories">'),
     "name_duplicate", "categories"),
    (_replace(WB, b'localSheetId="0">Stock!$9', b'localSheetId="7">Stock!$9'), "name_scope",
     "_xlnm.Print_Titles"),
    (_replace("xl/_rels/workbook.xml.rels", b'Target="calcChain.xml"', b'Target="gone.xml"'),
     "missing_part", "xl/_rels/workbook.xml.rels#rId6"),
    (_replace("[Content_Types].xml", b'<Default Extension="bin"', b'<Default Extension="nib"'),
     "no_content_type", "xl/printerSettings/printerSettings1.bin"),
    (_replace("xl/calcChain.xml", b'<c r="B1" i="3"/>', b'<c r="B1" i="3"/><c r="A1"/>'),
     "calc_chain", "Totals!A1"),
])
def test_validate_finds_each_issue_by_code_and_address(change, code, address):
    found = _broken(change)
    assert [(p.code, p.address) for p in found if p.code == code] == [(code, address)], found


def test_validate_finds_stray_overrides_and_table_issues():
    def stray(parts):
        parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
            b"</Types>", b'<Override PartName="/xl/none.xml" ContentType="application/xml"/>'
                         b"</Types>")
    assert ("stray_override", "xl/none.xml") in [(p.code, p.address) for p in _broken(stray)]
    table = "xl/tables/table1.xml"
    header = _broken(_replace(table, b'name="Tax"/>', b'name="Rate"/>'), table=True)
    assert [(p.code, p.address) for p in header] == [("table_header", "Lists!B1")]
    columns = _broken(_replace(table, b'ref="A1:B5" totalsRowShown', b'ref="A1:C5" totalsRowShown'),
                      table=True)
    assert "table_columns" in [p.code for p in columns]
    bad = _broken(_replace(table, b'ref="A1:B5" totalsRowShown', b'ref="nowhere" totalsRowShown'),
                  table=True)
    assert ("table_ref", table) in [(p.code, p.address) for p in bad]

    def twice(parts):
        parts["xl/tables/table2.xml"] = parts[table]
        parts["xl/worksheets/_rels/sheet2.xml.rels"] = parts[
            "xl/worksheets/_rels/sheet2.xml.rels"].replace(
            b"</Relationships>", b'<Relationship Id="rId2" Type="http://schemas.openxmlformats.'
            b'org/officeDocument/2006/relationships/table" Target="../tables/table2.xml"/>'
            b"</Relationships>")
        parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
            b"</Types>", b'<Override PartName="/xl/tables/table2.xml" ContentType="application/'
            b'vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>')
    codes = {p.code for p in _broken(twice, table=True)}
    assert {"table_name", "table_overlap"} <= codes


def test_validate_by_sheet():
    def bad(parts):
        parts[SHEET2] = parts[SHEET2].replace(b'<c r="B2" s="7"', b'<c r="B2" s="70"', 1)
    parts = _parts(Shop().to_bytes())
    bad(parts)
    wb = Workbook.open(_pack(parts))
    assert wb.validate(sheets=["Stock"]) == []
    assert [p.code for p in wb.validate(sheets=["Lists"])] == ["style"]


# -- save --------------------------------------------------------------------------------------


def test_save_reopens_and_reads_every_written_cell_back(tmp_path):
    wb = Workbook.open(Shop().to_bytes())
    _edit(wb)
    report = wb.save(tmp_path / "stock.xlsx")
    assert report.kind == "xlsx" and report.verified == len(WRITES) and report.issues == ()
    saved = (tmp_path / "stock.xlsx").read_bytes()
    assert saved == wb.to_bytes() and report.size == len(saved)
    again = Workbook.open(saved)
    assert again.cell("Stock!B10").v == "Spare fan" and again.cell("Stock!C11").t == "text"
    assert again.cell("Stock!H10").v == "2026-10-20" and again.cell("Stock!J10").formula
    assert again.calculation == ["full_calc_on_load"]


def test_a_save_that_does_not_read_back_is_refused():
    from xlsx_agent.edit import verify

    wb = Workbook.open(Shop().to_bytes())
    _edit(wb)
    with pytest.raises(EditError) as raised:
        verify(wb, Shop().to_bytes(), "xlsx")
    assert raised.value.code == "readback" and "Stock!B10" in raised.value.addresses


def test_save_as_a_template_and_back():
    wb = Workbook.open(Shop().to_bytes())
    buffer = io.BytesIO()
    assert wb.save(buffer, kind="xltx").kind == "xltx"
    template = Workbook.open(buffer.getvalue())
    assert template.kind == "xltx"
    assert wb.kind == "xlsx" and wb.to_bytes() == Shop().to_bytes()


def _with_macros() -> bytes:
    parts = _parts(Shop().to_bytes())
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
        b"spreadsheetml.sheet.main+xml", b"PLACEHOLDER", 1).replace(
        b"application/vnd.openxmlformats-officedocument.PLACEHOLDER",
        b"application/vnd.ms-excel.sheet.macroEnabled.main+xml").replace(
        b"</Types>", b'<Default Extension="vba" ContentType="application/vnd.ms-office.'
                     b'vbaProject"/></Types>')
    parts["xl/_rels/workbook.xml.rels"] = parts["xl/_rels/workbook.xml.rels"].replace(
        b"</Relationships>", b'<Relationship Id="rId9" Type="http://schemas.microsoft.com/'
        b'office/2006/relationships/vbaProject" Target="project.vba"/></Relationships>')
    parts["xl/project.vba"] = bytes(64)          # made-up bytes, never read
    return _pack(parts)


def test_a_workbook_with_a_vba_project_is_not_saved_as_xlsx(tmp_path):
    wb = Workbook.open(_with_macros())
    assert wb.kind == "xlsm"
    for target, kind in ((tmp_path / "book.xlsx", None), (io.BytesIO(), "xlsx"),
                         (io.BytesIO(), "xltx")):
        with pytest.raises(EditError) as raised:
            wb.save(target, kind=kind)
        assert raised.value.code == "vba_project"
    assert not (tmp_path / "book.xlsx").exists()
    assert wb.save(tmp_path / "book.xlsm").kind == "xlsm"
    assert (tmp_path / "book.xlsm").read_bytes() == _with_macros()


# -- the tool layer ----------------------------------------------------------------------------


@tool("test_write", "Write one cell (tests only).", {
    "doc": string("Doc."), "cell": string("Cell."), "value": number("Value."),
    "style": integer("A cell format index, unchecked.", optional=True)},
    kind="xlsx", mutates=True)
def _test_write(call, doc, cell, value, style=None):
    workbook = call.document
    if style is None:
        workbook.set_cells({cell: value})
    else:       # a format the workbook may not have: past set_cells, to the editor
        parsed = parse_address(cell)
        workbook.editor.write(workbook.sheet_of(parsed), parsed.area.first, value, style=style)
    return Result(summary=f"wrote {cell}", changed=[doc])


@pytest.fixture()
def box():
    with Toolbox([*TOOLS, _test_write], formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _session(box, data=None):
    session = box.session()
    blob = session.add_blob(data or Shop().to_bytes(), name="stock.xlsx")
    assert box.dispatch(session, "open_document", {"blob": blob}).ok
    return session


def test_save_document_hands_over_a_checked_file(box):
    session = _session(box)
    assert box.dispatch(session, "test_write", {"doc": "d1", "cell": "Stock!D10",
                                                "value": 7}).ok
    saved = box.dispatch(session, "save_document", {"doc": "d1", "name": "stock.xlsx",
                                                    "format": "xlsx"})
    assert saved.ok, saved.error
    assert saved.data["saved"] == {"verified": 1, "calculation": "invalidated"}
    assert saved.data["validate"]["new"] == [] and "data" not in saved.data
    output = session.outputs[-1]
    assert Workbook.open(output.data).cell("Stock!D10").v == 7


def test_save_document_refuses_a_new_structure_issue_only(box):
    parts = _parts(Shop().to_bytes())
    parts[SHEET2] = parts[SHEET2].replace(b'<c r="B2" s="7"', b'<c r="B2" s="70"', 1)
    session = _session(box, _pack(parts))
    assert box.dispatch(session, "open_document", {"blob": session.add_blob(
        Shop().to_bytes(), name="clean.xlsx")}).ok
    # d1: an issue it opened with, on a sheet then written: not new
    assert box.dispatch(session, "test_write", {"doc": "d1", "cell": "Lists!B3",
                                                "value": 0.2}).ok
    saved = box.dispatch(session, "save_document", {"doc": "d1", "name": "a.xlsx",
                                                    "format": "xlsx"})
    assert saved.ok and saved.data["validate"]["new"] == []
    # d2: a cell given a format the workbook does not have: new, refused
    assert box.dispatch(session, "test_write", {"doc": "d2", "cell": "Stock!D10",
                                                "value": 5, "style": 99}).ok
    refused = box.dispatch(session, "save_document", {"doc": "d2", "name": "b.xlsx",
                                                      "format": "xlsx"})
    assert refused.error.code == "refused"
    assert refused.error.valid_options == ["Problem(code='style', address='Stock!D10', count=1)"]


@pytest.mark.parametrize("arguments, field", [
    ({"name": "stock.xlsx", "format": "pptx"}, "format"),
    ({"name": "stock.xltx", "format": "xlsx"}, "name"),
])
def test_save_document_refuses_a_wrong_format(box, arguments, field):
    refused = box.dispatch(_session(box), "save_document", {"doc": "d1", **arguments})
    assert refused.error.code == "invalid_arguments" and refused.error.field == field


def test_save_document_refuses_a_vba_project_as_xlsx(box):
    session = box.session()
    blob = session.add_blob(_with_macros(), name="book.xlsm")
    box.dispatch(session, "open_document", {"blob": blob})
    refused = box.dispatch(session, "save_document", {"doc": "d1", "name": "book.xlsx",
                                                      "format": "xlsx"})
    assert refused.error.code == "refused"
    assert refused.error.details["code"] == "vba_project"
    assert box.dispatch(session, "save_document", {"doc": "d1", "name": "book.xlsm",
                                                   "format": "xlsm"}).ok


def _derive(box, session):
    derived = box.dispatch(session, "xlsx_derive", {
        "doc": "d1", "from": "t:Lists!A1", "measures": [{"fn": "sum", "column": "Tax"}],
        "id": "tax"})
    assert derived.ok, derived.error


def test_a_failed_batch_leaves_the_datasets_and_version_alone(box):
    session = _session(box)
    _derive(box, session)
    entry = session.entry("d1")
    original = entry.document.to_bytes()
    failed = box.dispatch(session, "batch", {"atomic": True, "ops": [
        {"tool": "test_write", "arguments": {"doc": "d1", "cell": "Stock!D10", "value": 5}},
        {"tool": "test_write", "arguments": {"doc": "d1", "cell": "Stock!F10", "value": 5}}]})
    assert not failed.ok
    assert entry.version == 0 and entry.document.to_bytes() == original
    stale = box.dispatch(session, "xlsx_read", {"target": "d:tax"})
    assert stale.ok and stale.data["dataset"]["stale"] is False
    # a non-atomic batch keeps the op that applied, as one step
    kept = box.dispatch(session, "batch", {"ops": [
        {"tool": "test_write", "arguments": {"doc": "d1", "cell": "Stock!D10", "value": 5}},
        {"tool": "test_write", "arguments": {"doc": "d1", "cell": "Stock!F10", "value": 5}},
        {"tool": "test_write", "arguments": {"doc": "d1", "cell": "Stock!D11", "value": 6}}]})
    assert kept.ok and entry.version == 1 and kept.data["failed"] == [1]
    wb = entry.document
    assert (wb.cell("Stock!D10").v, wb.cell("Stock!F10").formula, wb.cell("Stock!D11").v) == \
        (5, True, 6)
    assert box.dispatch(session, "xlsx_read", {"target": "d:tax"}).data["dataset"]["stale"] is True
    assert box.dispatch(session, "undo", {"doc": "d1"}).ok
    assert entry.version == 0 and wb.to_bytes() == original
    assert box.dispatch(session, "xlsx_read", {"target": "d:tax"}).data["dataset"]["stale"] is False
    assert box.dispatch(session, "undo", {"doc": "d1", "redo": True}).ok
    assert box.dispatch(session, "xlsx_read", {"target": "d:tax"}).data["dataset"]["stale"] is True


def test_undo_by_sheet_leaves_later_changes_on_other_sheets(box):
    session = _session(box)
    entry = session.entry("d1")

    def write(cell, value):
        assert box.dispatch(session, "test_write", {"doc": "d1", "cell": cell,
                                                    "value": value}).ok
    write("Stock!D10", 1)       # the first write: calcPr, the chain -- every sheet's parts
    write("Lists!B2", 0.5)      # a number over a number: the sheet's own part only
    write("Stock!D11", 2)
    undone = box.dispatch(session, "undo", {"doc": "d1", "scope": "Lists"})
    assert undone.ok, undone.error
    wb = entry.document
    assert (wb.cell("Lists!B2").v, wb.cell("Stock!D10").v, wb.cell("Stock!D11").v) == \
        (0.06, 1, 2)
    assert box.dispatch(session, "undo", {"doc": "d1", "scope": "Stock"}).ok
    assert wb.cell("Stock!D11").v == 40
    refused = box.dispatch(session, "undo", {"doc": "d1", "scope": "Stock"})
    assert refused.error.code == "entangled"
    assert wb.cell("Stock!D10").v == 1
    unknown = box.dispatch(session, "undo", {"doc": "d1", "scope": "Nowhere"})
    assert unknown.error.code == "not_found"


def test_a_callers_own_undo_leaves_another_callers_later_change(box):
    session = _session(box)
    entry = session.entry("d1")
    a = session.view(context={"agent": "a"})
    b = session.view(context={"agent": "b"})

    def write(view, cell, value):
        assert box.dispatch(view, "test_write", {"doc": "d1", "cell": cell,
                                                 "value": value}).ok
    write(a, "Stock!D10", 1)        # the first write changes what every sheet shares
    write(b, "Lists!B2", 0.5)
    refused = box.dispatch(a, "undo", {"doc": "d1"})
    assert refused.error.code == "entangled"
    write(a, "Stock!D11", 2)        # the sheet's own part only
    write(b, "Lists!B3", 0.25)
    assert box.dispatch(a, "undo", {"doc": "d1"}).ok
    wb = entry.document
    assert (wb.cell("Stock!D10").v, wb.cell("Stock!D11").v, wb.cell("Lists!B2").v,
            wb.cell("Lists!B3").v) == (1, 40, 0.5, 0.25)
    assert box.dispatch(a, "undo", {"doc": "d1", "redo": True}).ok
    assert wb.cell("Stock!D11").v == 2 and wb.cell("Lists!B3").v == 0.25
