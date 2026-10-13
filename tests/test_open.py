"""R1, "Open": every fixture round-trips byte for byte; the refusals carry their reason.

Every workbook here is built with tests/build.py.  The files that are refused -- compound
files, BIFF, ``.xlsb``, other packages -- are built from their minimal bytes below, from
[MS-CFB], [MS-XLS] and [MS-XLSB]; no Office file is used.
"""

import io
import struct
import warnings
import zipfile
from datetime import date

import pytest
from ooxml_common.kinds import main_content_type
from ooxml_edit.opc import OpcPackage
from ooxml_edit.tools.limits import Limits

from build import (CT_WORKSHEET, NS_MAIN, NS_R, Error, Formula, Inline, Rich, Run, WorkbookBuilder)
from xlsx_agent import REFUSALS, KindMismatchWarning, OpenError, Sheet, Workbook
from xlsx_agent.package import XLSB_CONTENT_TYPE, compound_file_names

# -- fixtures ------------------------------------------------------------------------------------


def _inventory() -> WorkbookBuilder:
    book = WorkbookBuilder()
    stock = book.sheet("Stock")
    stock.rows("A1", [["SKU", "Price", "Added", "Note"],
                      ["SKU-0001", 12.5, date(2026, 1, 5), Rich([Run("new", bold=True), Run(" item")])],
                      ["SKU-0002", Error("#N/A"), None, Inline("inline text")],
                      ["SKU-0003", 7, date(2026, 2, 1), True]])
    stock.cell("F1", Formula("SUM(B2:B4)", cached=19.5))
    stock.cell("F2", Formula("NOW()"))
    stock.table("Products", "A1:D4", ["SKU", "Price", "Added", "Note"], write_header=False)
    stock.merge("H1:I1")
    stock.comment("A2", "checked")
    stock.validation_list("D2:D4", ["new", "old"])
    stock.hide_row(3)
    stock.hide_column("C")
    book.define_name("Prices", "Stock!$B$2:$B$4")
    return book


def _states() -> WorkbookBuilder:
    book = WorkbookBuilder(protected=True)
    book.sheet("Front").cell("A1", "front")
    book.sheet("Back", state="hidden", sheet_id=7).cell("A1", 1)
    locked = book.sheet("Locked", state="veryHidden", sheet_id=3)
    locked.cell("A1", 2)
    locked.protected = True
    return book


def _messy() -> WorkbookBuilder:
    book = WorkbookBuilder(date1904=True, calc={"fullCalcOnLoad": "1"})
    sheet = book.sheet("Sparse")
    sheet.omit_refs = True
    sheet.dimension = "A1:B2"
    sheet.rows("A1", [[1, 2, 3], [4, None, 6]])
    sheet.cell("C1000", "far")
    book.sheet("Empty").dimension = False
    return book


def _kinds():
    for kind in ("xlsm", "xltx", "xltm"):
        book = WorkbookBuilder(kind=kind)
        book.sheet("Sheet1").cell("A1", kind)
        yield kind, book


FIXTURES = {"inventory": _inventory(), "states": _states(), "messy": _messy(), "empty": WorkbookBuilder(),
            **{kind: book for kind, book in _kinds()}}


def _zip(parts: dict[str, bytes]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(zipfile.ZipInfo(name, (2026, 1, 1, 0, 0, 0)), data, zipfile.ZIP_DEFLATED)
    return buffer.getvalue()


def _content_types(main_part: str, main_type: str) -> bytes:
    return (b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            b'<Default Extension="xml" ContentType="application/xml"/>'
            + f'<Override PartName="/{main_part}" ContentType="{main_type}"/></Types>'.encode())


def _root_rels(main_part: str, rel_type: str = NS_R + "/officeDocument") -> bytes:
    return (b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + f'<Relationship Id="rId1" Type="{rel_type}" Target="{main_part}"/></Relationships>'.encode())


def compound_file(names, *, types=None, truncate: bool = False) -> bytes:
    """A minimal compound file ([MS-CFB] 2.2): 512-byte sectors, one FAT sector (sector 0)
    and a directory (sectors 1...) holding a root entry and an empty stream per name."""
    entries = [("Root Entry", 5)] + [(name, (types or {}).get(name, 2)) for name in names]
    directory_sectors = -(-len(entries) // 4)
    header = bytearray(512)
    header[0:8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    struct.pack_into("<HHHHH", header, 0x18, 0x3E, 3, 0xFFFE, 9, 6)
    struct.pack_into("<IIIIIIII", header, 0x2C, 1, 1, 0, 4096, 0xFFFFFFFE, 0, 0xFFFFFFFE, 0)
    struct.pack_into("<109I", header, 0x4C, 0, *([0xFFFFFFFF] * 108))
    fat = [0xFFFFFFFD] + [n + 1 for n in range(1, directory_sectors)] + [0xFFFFFFFE]
    fat += [0xFFFFFFFF] * (128 - len(fat))
    directory = bytearray(512 * directory_sectors)
    for index, (name, kind) in enumerate(entries):
        encoded = name.encode("utf-16-le") + b"\x00\x00"
        directory[index * 128:index * 128 + len(encoded)] = encoded
        struct.pack_into("<HB", directory, index * 128 + 0x40, len(encoded), kind)
    data = bytes(header) + struct.pack("<128I", *fat) + bytes(directory)
    return data[:600] if truncate else data


# -- round trip ----------------------------------------------------------------------------------


@pytest.mark.parametrize("name", sorted(FIXTURES))
def test_every_fixture_round_trips_byte_for_byte(name):
    data = FIXTURES[name].to_bytes()
    wb = Workbook.open(data)
    assert wb.sheets is not None  # reading the sheets changes nothing
    assert wb.to_bytes() == data
    # and the package itself writes every part back as it was read
    assert wb.package.changed_parts() == frozenset()
    with zipfile.ZipFile(io.BytesIO(data)) as before, \
            zipfile.ZipFile(io.BytesIO(wb.package.to_bytes())) as after:
        assert before.namelist() == after.namelist()
        for entry in before.namelist():
            assert before.read(entry) == after.read(entry), entry


def test_round_trip_from_a_path_a_file_and_a_package(tmp_path):
    data = FIXTURES["inventory"].to_bytes()
    path = tmp_path / "inventory.xlsx"
    path.write_bytes(data)
    assert Workbook.open(path).to_bytes() == data
    assert Workbook.open(str(path)).to_bytes() == data
    assert Workbook.open(io.BytesIO(data)).to_bytes() == data
    assert Workbook.open(bytearray(data)).to_bytes() == data
    package = OpcPackage.open(data)
    wb = Workbook.open(package)
    assert wb.package is package
    assert wb.to_bytes() == package.to_bytes()
    assert [sheet.name for sheet in wb.sheets] == ["Stock"]


def test_kind_is_what_the_main_part_declares():
    assert Workbook.open(FIXTURES["inventory"].to_bytes()).kind == "xlsx"
    for kind in ("xlsm", "xltx", "xltm"):
        assert Workbook.open(FIXTURES[kind].to_bytes()).kind == kind


def test_an_extension_of_another_kind_opens_with_a_warning(tmp_path):
    path = tmp_path / "book.xltx"
    path.write_bytes(FIXTURES["inventory"].to_bytes())
    with pytest.warns(KindMismatchWarning, match=r"\.xltx"):
        wb = Workbook.open(path)
    assert wb.kind == "xlsx"
    same = tmp_path / "book.xlsx"
    same.write_bytes(FIXTURES["inventory"].to_bytes())
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        Workbook.open(same)


def test_an_edited_package_writes_its_changes():
    wb = Workbook.open(FIXTURES["inventory"].to_bytes())
    wb.package.replace_part("xl/worksheets/sheet1.xml", wb.package.read("xl/worksheets/sheet1.xml") + b"\n")
    assert wb.to_bytes() == wb.package.to_bytes()
    assert wb.to_bytes() != FIXTURES["inventory"].to_bytes()


# -- sheets --------------------------------------------------------------------------------------


def test_sheets_list_id_name_state_and_protection():
    wb = Workbook.open(FIXTURES["states"].to_bytes())
    assert wb.sheets == [
        Sheet("Front", 1, "visible", False, "worksheet", "xl/worksheets/sheet1.xml"),
        Sheet("Back", 7, "hidden", False, "worksheet", "xl/worksheets/sheet2.xml"),
        Sheet("Locked", 3, "veryHidden", True, "worksheet", "xl/worksheets/sheet3.xml"),
    ]
    assert [sheet.visible for sheet in wb.sheets] == [True, False, False]
    assert wb.protected
    assert wb.sheet("locked").id == 3
    with pytest.raises(KeyError, match="Front"):
        wb.sheet("Nowhere")
    assert repr(wb.sheets[0]) == "Sheet(name='Front', id=1, state='visible', protected=False, type='worksheet')"


def test_an_unprotected_workbook_and_empty_workbook():
    assert not Workbook.open(FIXTURES["inventory"].to_bytes()).protected
    empty = Workbook.open(FIXTURES["empty"].to_bytes())
    assert empty.sheets == [] and not empty.protected


def _with_part(book: WorkbookBuilder, changes: dict[str, bytes]) -> bytes:
    parts = book.parts()
    parts.update(changes)
    return _zip(parts)


def test_protection_after_many_rows_is_found_and_protection_off_is_not():
    book = WorkbookBuilder()
    big = book.sheet("Big")
    big.rows("A1", [[n, f"row {n}"] for n in range(1, 3001)])
    big.protected = True
    book.sheet("Off").cell("A1", 1)
    parts = book.parts()
    off = parts["xl/worksheets/sheet2.xml"]
    parts["xl/worksheets/sheet2.xml"] = off.replace(b"</sheetData>", b'</sheetData><sheetProtection sheet="0"/>')
    wb = Workbook.open(_zip(parts))
    assert [sheet.protected for sheet in wb.sheets] == [True, False]


def test_a_cell_that_names_the_element_is_not_protection():
    book = WorkbookBuilder()
    book.sheet("Text").cell("A1", Inline('<sheetProtection sheet="1"/>'))
    assert not Workbook.open(book.to_bytes()).sheets[0].protected


def test_chart_sheets_and_a_missing_part():
    book = WorkbookBuilder()
    book.sheet("Data").cell("A1", 1)
    book.sheet("Chart")
    parts = book.parts()
    chart = (b'<?xml version="1.0" encoding="UTF-8"?><chartsheet xmlns="' + NS_MAIN.encode()
             + b'"><sheetPr/><sheetProtection content="1" objects="1"/></chartsheet>')
    parts["xl/worksheets/sheet2.xml"] = chart
    parts["xl/_rels/workbook.xml.rels"] = parts["xl/_rels/workbook.xml.rels"].replace(
        b'relationships/worksheet" Target="worksheets/sheet2.xml', b'relationships/chartsheet" Target="worksheets/sheet2.xml')
    del parts["xl/worksheets/sheet1.xml"]
    wb = Workbook.open(_zip(parts))
    assert [(s.name, s.type, s.protected, s.part) for s in wb.sheets] == [
        ("Data", "worksheet", False, None), ("Chart", "chartsheet", True, "xl/worksheets/sheet2.xml")]


def test_a_strict_workbook_is_read_like_a_transitional_one():
    book = _states()
    strict = {}
    for name, data in book.parts().items():
        data = data.replace(NS_MAIN.encode(), b"http://purl.oclc.org/ooxml/spreadsheetml/main")
        data = data.replace(NS_R.encode(), b"http://purl.oclc.org/ooxml/officeDocument/relationships")
        strict[name] = data
    data = _zip(strict)
    wb = Workbook.open(data)
    assert [(s.name, s.id, s.state, s.protected) for s in wb.sheets] == [
        ("Front", 1, "visible", False), ("Back", 7, "hidden", False), ("Locked", 3, "veryHidden", True)]
    assert wb.protected
    assert wb.to_bytes() == data


def test_a_sheet_id_that_is_no_number():
    parts = _states().parts()
    parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace(b'sheetId="7"', b'sheetId="seven"')
    assert Workbook.open(_zip(parts)).sheets[1].id is None


def _laughs(root: str) -> bytes:
    """A billion-laughs DOCTYPE: ten to the ninth copies of "lol" if it were expanded."""
    entities = b'<!ENTITY a0 "lol">' + b"".join(
        b'<!ENTITY a%d "%s">' % (n, (b"&a%d;" % (n - 1)) * 10) for n in range(1, 10))
    return (b'<?xml version="1.0"?><!DOCTYPE ' + root.encode() + b" [" + entities + b"]><" + root.encode()
            + b' xmlns="' + NS_MAIN.encode() + b'"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>&a9;</t>'
            b'</is></c></row></sheetData><sheetProtection sheet="1"/></' + root.encode() + b">")


def test_an_entity_bomb_in_a_sheet_is_refused_at_open_without_expanding():
    book = WorkbookBuilder()
    book.sheet("Laughs")
    data = _with_part(book, {"xl/worksheets/sheet1.xml": _laughs("worksheet")})
    error = _refused(data, "damaged")
    assert "xl/worksheets/sheet1.xml" in str(error) and "lol" not in str(error)


def test_an_entity_bomb_in_the_workbook_part_is_refused():
    book = WorkbookBuilder()
    book.sheet("Stock")
    data = _with_part(book, {"xl/workbook.xml": _laughs("workbook")})
    assert "xl/workbook.xml" in str(_refused(data, "damaged"))


def test_a_sheet_that_is_not_well_formed_is_refused():
    book = WorkbookBuilder()
    book.sheet("Broken").protected = True
    parts = book.parts()
    parts["xl/worksheets/sheet1.xml"] = parts["xl/worksheets/sheet1.xml"][:-20]
    assert "xl/worksheets/sheet1.xml" in str(_refused(_zip(parts), "damaged"))


# -- refusals ------------------------------------------------------------------------------------


def _refused(source, code, **kwargs) -> OpenError:
    with pytest.raises(OpenError) as caught:
        Workbook.open(source, **kwargs)
    assert caught.value.code == code
    assert str(caught.value).startswith(REFUSALS[code])
    return caught.value


def test_an_encrypted_workbook_is_refused():
    data = compound_file(["\x06DataSpaces", "EncryptionInfo", "EncryptedPackage"], types={"\x06DataSpaces": 1})
    assert "EncryptedPackage" in compound_file_names(data)
    _refused(data, "encrypted")


def test_an_xls_is_refused():
    _refused(compound_file(["Workbook", "\x05SummaryInformation"]), "xls")
    _refused(compound_file(["Book"]), "xls")
    # BIFF4 without a compound file: the BOF record first
    _refused(b"\x09\x04\x06\x00\x00\x00\x10\x00" + bytes(64), "xls")


def test_another_compound_file_is_not_a_workbook():
    _refused(compound_file(["WordDocument", "1Table"]), "not_a_workbook")


def test_a_damaged_compound_file():
    _refused(compound_file(["Workbook", "a", "b", "c", "d"], truncate=True), "damaged")
    _refused(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + bytes(20), "damaged")


def test_a_directory_over_several_sectors_is_read():
    names = [f"Stream{n}" for n in range(9)] + ["EncryptedPackage"]
    assert compound_file_names(compound_file(names)) >= set(names)
    _refused(compound_file(names), "encrypted")


def test_an_xlsb_is_refused():
    parts = {"[Content_Types].xml": _content_types("xl/workbook.bin", XLSB_CONTENT_TYPE),
             "_rels/.rels": _root_rels("xl/workbook.bin"),
             "xl/workbook.bin": b"\x83\x01\x00"}
    _refused(_zip(parts), "xlsb")


def test_other_packages_and_files_are_not_workbooks():
    docx = {"[Content_Types].xml": _content_types("word/document.xml", main_content_type("docx")),
            "_rels/.rels": _root_rels("word/document.xml"),
            "word/document.xml": b'<w:document xmlns:w="urn:w"/>'}
    assert "a .docx package" in str(_refused(_zip(docx), "not_a_workbook"))
    _refused(_zip({"[Content_Types].xml": _content_types("a.xml", "application/xml")}), "not_a_workbook")
    _refused(b"plain text, not a workbook", "not_a_workbook")
    _refused(b"", "not_a_workbook")
    _refused(b"PK\x03\x04" + bytes(40), "damaged")
    package = OpcPackage.open(_zip(docx))
    _refused(package, "not_a_workbook")


def test_the_zip_limits_are_ooxml_edits():
    data = FIXTURES["inventory"].to_bytes()
    _refused(data, "limit", limits=Limits(max_package_bytes=100))
    _refused(data, "limit", limits=Limits(max_parts=3))
    _refused(data, "limit", limits=Limits(max_uncompressed_bytes=1000))
    bomb = _with_part(WorkbookBuilder(), {"xl/padding.xml": b"<a>" + b" " * (4 * 1024 * 1024) + b"</a>"})
    assert len(bomb) < 100_000
    assert "ratio" in str(_refused(bomb, "limit"))


def test_a_path_over_the_size_limit_is_refused_before_it_is_read(tmp_path, monkeypatch):
    path = tmp_path / "big.xlsx"
    path.write_bytes(FIXTURES["inventory"].to_bytes())
    import builtins

    real_open = builtins.open

    def no_open(file, *args, **kwargs):
        if str(file) == str(path):
            raise AssertionError("the file was read")
        return real_open(file, *args, **kwargs)

    monkeypatch.setattr(builtins, "open", no_open)
    _refused(path, "limit", limits=Limits(max_package_bytes=100))


def test_refusals_name_no_content():
    marker = "MARKER-7f3a"
    book = WorkbookBuilder()
    sheet = book.sheet("Stock")
    sheet.cell("A1", marker)
    sheet.cell("A2", Formula(f'"{marker}"', cached=marker))
    sheet.comment("A1", marker)
    data = book.to_bytes()
    for limits in (Limits(max_package_bytes=100), Limits(max_parts=2)):
        with pytest.raises(OpenError) as caught:
            Workbook.open(data, limits=limits)
        assert marker not in str(caught.value)


def test_not_a_refusal_code():
    with pytest.raises(ValueError):
        OpenError("nonsense")


def test_an_unknown_source_type():
    with pytest.raises(TypeError):
        Workbook.open(12)
    with pytest.raises(TypeError):
        Workbook.open(io.StringIO("text"))


def test_worksheet_content_type_is_untouched():
    """The workbook reads its sheets from relationships, not content types: a sheet whose
    override is missing is still listed."""
    parts = FIXTURES["inventory"].parts()
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(CT_WORKSHEET.encode(), b"application/xml")
    assert [s.name for s in Workbook.open(_zip(parts)).sheets] == ["Stock"]
