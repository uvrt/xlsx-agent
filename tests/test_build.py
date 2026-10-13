"""tests/build.py, the test-only SpreadsheetML writer: R0's "done when".

A built workbook opens as a ZIP with the right content types and relationships, every part
parses, and each feature of the builder is written where ECMA-376 puts it.
"""

import io
import zipfile
from datetime import date, datetime, time

import pytest
from lxml import etree
from ooxml_common.kinds import kind_of, main_content_type
from ooxml_edit.opc import OpcPackage

from build import (CT_COMMENTS, CT_SHARED_STRINGS, CT_STYLES, CT_TABLE, CT_WORKSHEET, NS_CONTENT_TYPES,
                   NS_MAIN, NS_PKG_RELS, NS_R, REL_COMMENTS, REL_OFFICE_DOCUMENT, REL_SHARED_STRINGS,
                   REL_STYLES, REL_TABLE, REL_WORKSHEET, Error, Formula, Inline, IsoDate, Rich, Run,
                   WorkbookBuilder, date_serial)

NS = {"x": NS_MAIN, "r": NS_R, "ct": NS_CONTENT_TYPES, "rel": NS_PKG_RELS}
_PARSER = etree.XMLParser(resolve_entities=False, no_network=True)

#: CT_Worksheet's child sequence (ECMA-376 Part 1), the elements the builder writes among them.
WORKSHEET_ORDER = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData",
                   "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter",
                   "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr",
                   "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions",
                   "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
                   "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
                   "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls",
                   "webPublishItems", "tableParts", "extLst"]


class Built:
    """A built workbook's bytes, opened: parts parsed, relationships resolved."""

    def __init__(self, data: bytes) -> None:
        self.data = data
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            assert archive.testzip() is None
            self.names = archive.namelist()
            self.raw = {name: archive.read(name) for name in self.names}
        self.trees = {name: etree.fromstring(raw, _PARSER) for name, raw in self.raw.items()}

    def tree(self, name: str) -> etree._Element:
        return self.trees[name]

    def content_type(self, name: str) -> str | None:
        types = self.tree("[Content_Types].xml")
        for node in types.iterfind("ct:Override", NS):
            if node.get("PartName") == "/" + name:
                return node.get("ContentType")
        extension = name.rpartition(".")[2]
        for node in types.iterfind("ct:Default", NS):
            if node.get("Extension").lower() == extension.lower():
                return node.get("ContentType")
        return None

    def rels(self, source: str) -> dict[str, tuple[str, str]]:
        """``rId`` -> (type, target part name) for a part's relationships ("" for the package)."""
        folder, _, base = source.rpartition("/")
        rels_name = f"{folder + '/' if folder else ''}_rels/{base}.rels"
        if rels_name not in self.trees:
            return {}
        found = {}
        for node in self.tree(rels_name).iterfind("rel:Relationship", NS):
            found[node.get("Id")] = (node.get("Type"), _resolve(folder, node.get("Target")))
        return found

    def related(self, source: str, rel_type: str) -> list[str]:
        return [target for kind, target in self.rels(source).values() if kind == rel_type]

    def sheet_part(self, name: str) -> str:
        workbook = self.tree("xl/workbook.xml")
        node = workbook.find(f"x:sheets/x:sheet[@name='{name}']", NS)
        return self.rels("xl/workbook.xml")[node.get(f"{{{NS_R}}}id")][1]

    def sheet(self, name: str) -> etree._Element:
        return self.tree(self.sheet_part(name))

    def cell(self, sheet: str, ref: str) -> etree._Element:
        node = self.sheet(sheet).find(f"x:sheetData/x:row/x:c[@r='{ref}']", NS)
        assert node is not None, ref
        return node

    def shared_string(self, index: int) -> etree._Element:
        return self.tree("xl/sharedStrings.xml").findall("x:si", NS)[index]

    def number_format(self, style: int) -> tuple[int, str | None]:
        styles = self.tree("xl/styles.xml")
        xf = styles.findall("x:cellXfs/x:xf", NS)[style]
        number = int(xf.get("numFmtId"))
        code = styles.find(f"x:numFmts/x:numFmt[@numFmtId='{number}']", NS)
        return number, None if code is None else code.get("formatCode")


def _resolve(folder: str, target: str) -> str:
    parts = (folder.split("/") if folder else []) + target.split("/")
    out: list[str] = []
    for part in parts:
        if part == "..":
            out.pop()
        elif part and part != ".":
            out.append(part)
    return "/".join(out)


def _text(node: etree._Element) -> str:
    return "".join(t.text or "" for t in node.iter(f"{{{NS_MAIN}}}t")
                   if t.getparent().tag != f"{{{NS_MAIN}}}rPh")


def inventory() -> WorkbookBuilder:
    """A workbook using every feature of the builder at once."""
    book = WorkbookBuilder()
    stock = book.sheet("Stock")
    stock.rows("A1", [["SKU", "Item", "Price", "Added", "Active"],
                      ["SKU-0001", "Hinge", 12.5, date(2026, 1, 5), True],
                      ["SKU-0002", "Bracket", 4, date(2026, 2, 9), False],
                      ["SKU-0003", "Hinge", Error("#N/A"), None, True]])
    stock.table("Products", "A1:E5", ["SKU", "Item", "Price", "Added", "Active"], write_header=False,
                totals_labels={"SKU": "Total"}, totals_functions={"Price": "sum"})
    stock.cell("C5", Formula("SUBTOTAL(109,[Price])", cached=16.5))
    stock.cell("G1", Rich(["Plain ", Run("bold", bold=True, size=11, color="FF112233")]))
    stock.cell("G2", Inline("  padded  "))
    stock.cell("G3", Inline(Rich([Run("in", italic=True), "line"])))
    stock.cell("G4", Formula("C2*2"))
    stock.cell("G5", Formula("C2&\"x\"", cached="12.5x"))
    stock.cell("G6", Formula("C2>1", cached=True))
    stock.cell("G7", Formula("1/0", cached=Error("#DIV/0!")))
    stock.cell("G8", IsoDate(date(2026, 3, 4)))
    stock.cell("G9", None, style=book.style(num_fmt="0.00%"))
    stock.cell("G10", 0.25, style=book.style(num_fmt="0.00%"))
    stock.cell("G11", 45000, style=book.style(num_fmt_id=14))
    stock.merge("I1:J2")
    stock.comment("C2", "Checked against the delivery note")
    stock.comment("C3", Rich([Run("Reviewer:", bold=True), " recount"]), author="Second reviewer")
    stock.comment("D2", "Second note", author="Reviewer")
    stock.validation_list("E2:E4", ["Yes", "No", 'Say "maybe"'], allow_blank=True)
    stock.validation_list("B2:B4", source="Lists!$A$1:$A$2")
    stock.validation("C2:C4", "decimal", "0", "1000", operator="between")
    lists = book.sheet("Lists", state="hidden")
    lists.rows("A1", [["Hinge"], ["Bracket"]])
    book.sheet("Archive", state="veryHidden").cell("A1", "Old stock")
    sales = book.sheet("Q3 sales")
    sales.rows("A1", [[1, 2], [3, 4]])
    sales.table("Sales", "A1:B2", ["North", "South"], header_row=False)
    book.define_name("Prices", "Stock!$C$2:$C$4")
    book.define_name("Rate", "0.21", sheet="Q3 sales", hidden=True)
    return book


@pytest.fixture(scope="module")
def built() -> Built:
    return Built(inventory().to_bytes())


# -- the package ---------------------------------------------------------------------------------

def test_package_is_a_zip_with_content_types_first(built):
    assert built.names[0] == "[Content_Types].xml"
    assert "_rels/.rels" in built.names
    assert len(set(built.names)) == len(built.names)


def test_every_part_parses_and_has_its_content_type(built):
    expected = {"xl/workbook.xml": main_content_type("xlsx"), "xl/styles.xml": CT_STYLES,
                "xl/sharedStrings.xml": CT_SHARED_STRINGS, "xl/comments1.xml": CT_COMMENTS,
                "xl/tables/table1.xml": CT_TABLE, "xl/tables/table2.xml": CT_TABLE}
    for number in range(1, 5):
        expected[f"xl/worksheets/sheet{number}.xml"] = CT_WORKSHEET
    for name in built.names:
        if name == "[Content_Types].xml":
            continue
        if name.endswith(".rels"):
            assert built.content_type(name) == "application/vnd.openxmlformats-package.relationships+xml"
        else:
            assert built.content_type(name) == expected.pop(name), name
    assert expected == {}
    # Every override names a part that is there.
    for node in built.tree("[Content_Types].xml").iterfind("ct:Override", NS):
        assert node.get("PartName")[1:] in built.names


def test_relationships_resolve_and_reach_every_part(built):
    assert built.related("", REL_OFFICE_DOCUMENT) == ["xl/workbook.xml"]
    reached, todo = set(), [""]
    while todo:
        source = todo.pop()
        for _, target in built.rels(source).values():
            assert target in built.names, (source, target)
            if target not in reached:
                reached.add(target)
                todo.append(target)
    parts = {name for name in built.names if name != "[Content_Types].xml" and not name.endswith(".rels")}
    assert reached == parts
    assert built.related("xl/workbook.xml", REL_STYLES) == ["xl/styles.xml"]
    assert built.related("xl/workbook.xml", REL_SHARED_STRINGS) == ["xl/sharedStrings.xml"]
    assert len(built.related("xl/workbook.xml", REL_WORKSHEET)) == 4


def test_ooxml_edit_opens_it_as_an_xlsx(built):
    package = OpcPackage.open(built.data)
    main = package.main_document_part()
    assert main == "xl/workbook.xml"
    assert kind_of(package.content_type(main)) == "xlsx"
    assert package.unreachable_parts() == set()
    assert sorted(package.part_names) == sorted(built.names)
    assert package.to_bytes() == built.data


@pytest.mark.parametrize("kind", ["xlsx", "xlsm", "xltx", "xltm"])
def test_kind_sets_the_main_content_type(kind):
    built = Built(_one_cell(kind))
    assert kind_of(built.content_type("xl/workbook.xml")) == kind


def _one_cell(kind: str = "xlsx", **options) -> bytes:
    book = WorkbookBuilder(kind=kind, **options)
    book.sheet("Sheet").cell("A1", 1)
    return book.to_bytes()


def test_output_is_deterministic():
    assert inventory().to_bytes() == inventory().to_bytes()
    with zipfile.ZipFile(io.BytesIO(_one_cell())) as archive:
        assert {info.date_time for info in archive.infolist()} == {(1980, 1, 1, 0, 0, 0)}


def test_without_strings_there_is_no_shared_string_part():
    built = Built(_one_cell())
    assert "xl/sharedStrings.xml" not in built.names
    assert built.related("xl/workbook.xml", REL_SHARED_STRINGS) == []


# -- sheets --------------------------------------------------------------------------------------

def test_sheets_states_and_ids(built):
    sheets = built.tree("xl/workbook.xml").findall("x:sheets/x:sheet", NS)
    assert [(s.get("name"), s.get("sheetId"), s.get("state")) for s in sheets] == [
        ("Stock", "1", None), ("Lists", "2", "hidden"), ("Archive", "3", "veryHidden"),
        ("Q3 sales", "4", None)]
    assert built.sheet_part("Q3 sales") == "xl/worksheets/sheet4.xml"


def test_sheet_ids_may_be_given_and_names_are_unique():
    book = WorkbookBuilder()
    book.sheet("One", sheet_id=7)
    assert book.sheet("Two").sheet_id == 8
    with pytest.raises(ValueError):
        book.sheet("One")
    with pytest.raises(ValueError):
        book.sheet("Three", sheet_id=7)
    with pytest.raises(ValueError):
        book.sheet("Four", state="gone")


def test_worksheet_children_in_schema_order(built):
    for number in range(1, 5):
        tags = [etree.QName(child).localname for child in built.tree(f"xl/worksheets/sheet{number}.xml")]
        assert tags == sorted(tags, key=WORKSHEET_ORDER.index)
    assert [etree.QName(c).localname for c in built.sheet("Stock")] == [
        "dimension", "sheetData", "mergeCells", "dataValidations", "tableParts"]


def test_dimension_is_the_used_range(built):
    # From the cells; the merge over I1:J2 holds none.
    assert built.sheet("Stock").find("x:dimension", NS).get("ref") == "A1:G11"
    assert built.sheet("Archive").find("x:dimension", NS).get("ref") == "A1"


def test_cells_in_row_and_column_order(built):
    for row in built.sheet("Stock").iterfind("x:sheetData/x:row", NS):
        refs = [c.get("r") for c in row]
        assert all(ref.endswith(row.get("r")) for ref in refs)
        assert refs == sorted(refs, key=lambda ref: (len(ref), ref))
    rows = [int(r.get("r")) for r in built.sheet("Stock").iterfind("x:sheetData/x:row", NS)]
    assert rows == sorted(rows)


# -- cells of every type -------------------------------------------------------------------------

def test_numbers(built):
    price = built.cell("Stock", "C2")
    assert price.get("t") is None and price.findtext("x:v", namespaces=NS) == "12.5"
    assert built.cell("Stock", "C3").findtext("x:v", namespaces=NS) == "4"


def test_shared_strings_counted_and_deduplicated(built):
    sst = built.tree("xl/sharedStrings.xml")
    items = sst.findall("x:si", NS)
    assert int(sst.get("uniqueCount")) == len(items)
    shared = [c for sheet in built.names if sheet.startswith("xl/worksheets/sheet")
              for c in built.tree(sheet).iter(f"{{{NS_MAIN}}}c") if c.get("t") == "s"]
    assert int(sst.get("count")) == len(shared)
    first, second = built.cell("Stock", "B2"), built.cell("Stock", "B4")
    assert first.get("t") == "s" and first.findtext("x:v", namespaces=NS) == second.findtext("x:v", namespaces=NS)
    assert _text(built.shared_string(int(first.findtext("x:v", namespaces=NS)))) == "Hinge"


def test_rich_shared_string(built):
    node = built.cell("Stock", "G1")
    si = built.shared_string(int(node.findtext("x:v", namespaces=NS)))
    runs = si.findall("x:r", NS)
    assert [r.findtext("x:t", namespaces=NS) for r in runs] == ["Plain ", "bold"]
    assert runs[0].find("x:rPr", NS) is None
    properties = runs[1].find("x:rPr", NS)
    assert [etree.QName(p).localname for p in properties] == ["b", "color", "sz"]
    assert properties.find("x:color", NS).get("rgb") == "FF112233"
    assert properties.find("x:sz", NS).get("val") == "11"
    space = runs[0].find("x:t", NS).get("{http://www.w3.org/XML/1998/namespace}space")
    assert space == "preserve"


def test_inline_strings(built):
    plain = built.cell("Stock", "G2")
    assert plain.get("t") == "inlineStr" and plain.find("x:v", NS) is None
    t = plain.find("x:is/x:t", NS)
    assert t.text == "  padded  " and t.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"
    rich = built.cell("Stock", "G3")
    assert rich.get("t") == "inlineStr"
    assert _text(rich.find("x:is", NS)) == "inline"
    assert rich.find("x:is/x:r/x:rPr/x:i", NS) is not None


def test_phonetic_runs():
    book = WorkbookBuilder()
    book.sheet("Names").cell("A1", Rich(["東京"], phonetic=[(0, 2, "トウキョウ")]))
    si = Built(book.to_bytes()).shared_string(0)
    ph = si.find("x:rPh", NS)
    assert (ph.get("sb"), ph.get("eb"), ph.findtext("x:t", namespaces=NS)) == ("0", "2", "トウキョウ")
    assert [etree.QName(c).localname for c in si] == ["r", "rPh", "phoneticPr"]
    assert _text(si) == "東京"


def test_booleans_and_errors(built):
    yes, no = built.cell("Stock", "E2"), built.cell("Stock", "E3")
    assert (yes.get("t"), yes.findtext("x:v", namespaces=NS)) == ("b", "1")
    assert (no.get("t"), no.findtext("x:v", namespaces=NS)) == ("b", "0")
    error = built.cell("Stock", "C4")
    assert (error.get("t"), error.findtext("x:v", namespaces=NS)) == ("e", "#N/A")


def test_dates_via_a_style(built):
    added = built.cell("Stock", "D2")
    assert added.get("t") is None
    assert float(added.findtext("x:v", namespaces=NS)) == date_serial(date(2026, 1, 5))
    assert built.number_format(int(added.get("s"))) == (164, "yyyy-mm-dd")
    # The same format is the same style.
    assert built.cell("Stock", "D3").get("s") == added.get("s")


def test_iso_date_cell(built):
    node = built.cell("Stock", "G8")
    assert (node.get("t"), node.findtext("x:v", namespaces=NS)) == ("d", "2026-03-04")


def test_style_only_cell_is_empty(built):
    node = built.cell("Stock", "G9")
    assert node.get("s") and len(node) == 0 and node.get("t") is None


def test_formulas_with_and_without_cached_values(built):
    total = built.cell("Stock", "C5")
    assert total.findtext("x:f", namespaces=NS) == "SUBTOTAL(109,[Price])"
    assert total.findtext("x:v", namespaces=NS) == "16.5" and total.get("t") is None
    missing = built.cell("Stock", "G4")
    assert missing.findtext("x:f", namespaces=NS) == "C2*2" and missing.find("x:v", NS) is None
    text = built.cell("Stock", "G5")
    assert (text.get("t"), text.findtext("x:v", namespaces=NS)) == ("str", "12.5x")
    flag = built.cell("Stock", "G6")
    assert (flag.get("t"), flag.findtext("x:v", namespaces=NS)) == ("b", "1")
    error = built.cell("Stock", "G7")
    assert (error.get("t"), error.findtext("x:v", namespaces=NS)) == ("e", "#DIV/0!")
    for node in (total, missing, text, flag, error):
        assert [etree.QName(c).localname for c in node] in (["f", "v"], ["f"])


def test_shared_and_array_formulas():
    book = WorkbookBuilder()
    sheet = book.sheet("Calc")
    sheet.rows("A1", [[1], [2], [3]])
    sheet.cell("B1", Formula("A1*2", cached=2, kind="shared", ref="B1:B3", si=0))
    sheet.cell("B2", Formula("", cached=4, kind="shared", si=0))
    sheet.cell("C1", Formula("A1:A3*3", cached=3, kind="array", ref="C1:C3"))
    built = Built(book.to_bytes())
    master = built.cell("Calc", "B1").find("x:f", NS)
    assert (master.get("t"), master.get("ref"), master.get("si"), master.text) == ("shared", "B1:B3", "0", "A1*2")
    member = built.cell("Calc", "B2").find("x:f", NS)
    assert (member.get("t"), member.get("si"), member.get("ref"), member.text) == ("shared", "0", None, None)
    array = built.cell("Calc", "C1").find("x:f", NS)
    assert (array.get("t"), array.get("ref")) == ("array", "C1:C3")
    with pytest.raises(ValueError):
        book.sheet("Bad").cell("A1", Formula("1", kind="dynamic"))
        book.to_bytes()


def test_values_the_builder_refuses():
    book = WorkbookBuilder()
    sheet = book.sheet("Bad")
    sheet.cell("A1", float("nan"))
    with pytest.raises(ValueError):
        book.to_bytes()
    sheet.cell("A1", object())
    with pytest.raises(TypeError):
        book.to_bytes()
    with pytest.raises(ValueError):
        sheet.cell("1A", 1)


# -- styles and number formats -------------------------------------------------------------------

def test_number_formats_custom_and_built_in(built):
    percent = built.cell("Stock", "G10")
    number, code = built.number_format(int(percent.get("s")))
    assert number >= 164 and code == "0.00%"
    assert built.cell("Stock", "G9").get("s") == percent.get("s")
    assert built.number_format(int(built.cell("Stock", "G11").get("s"))) == (14, None)
    styles = built.tree("xl/styles.xml")
    formats = styles.find("x:numFmts", NS)
    assert int(formats.get("count")) == len(formats)
    xfs = styles.find("x:cellXfs", NS)
    assert int(xfs.get("count")) == len(xfs)
    assert xfs[0].get("numFmtId") == "0"
    assert [etree.QName(c).localname for c in styles] == [
        "numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles"]


def test_style_index_is_reused():
    book = WorkbookBuilder()
    assert book.style(num_fmt="0.0") == book.style(num_fmt="0.0") == 1
    assert book.style(num_fmt_id=14) == 2
    assert book.style(num_fmt="0.000") == 3
    with pytest.raises(ValueError):
        book.style()


def test_date_serials_in_both_systems():
    # ECMA-376 Part 1 §18.17.4.1: the 1900 system counts a 29 February 1900 (serial 60).
    assert date_serial(date(1900, 1, 1)) == 1
    assert date_serial(date(1900, 2, 28)) == 59
    assert date_serial(date(1900, 3, 1)) == 61
    assert date_serial(date(1904, 1, 1), date1904=True) == 0
    assert date_serial(date(1904, 1, 2), date1904=True) == 1
    assert date_serial(date(2026, 1, 5)) - date_serial(date(2026, 1, 5), date1904=True) == 1462
    assert date_serial(datetime(1900, 3, 1, 12)) == 61.5
    assert date_serial(time(6)) == 0.25
    with pytest.raises(ValueError):
        date_serial(date(1899, 12, 31))
    with pytest.raises(ValueError):
        date_serial(date(1903, 12, 31), date1904=True)


def test_date1904_workbook():
    book = WorkbookBuilder(date1904=True)
    book.sheet("Dates").cell("A1", date(2026, 1, 5))
    built = Built(book.to_bytes())
    assert built.tree("xl/workbook.xml").find("x:workbookPr", NS).get("date1904") == "1"
    value = float(built.cell("Dates", "A1").findtext("x:v", namespaces=NS))
    assert value == date_serial(date(2026, 1, 5), date1904=True)
    assert Built(_one_cell()).tree("xl/workbook.xml").find("x:workbookPr", NS) is None


def test_datetime_and_time_get_their_own_formats():
    book = WorkbookBuilder()
    sheet = book.sheet("Times")
    sheet.cell("A1", datetime(2026, 1, 5, 6, 30))
    sheet.cell("A2", time(18))
    built = Built(book.to_bytes())
    assert built.number_format(int(built.cell("Times", "A1").get("s")))[1] == "yyyy-mm-dd hh:mm:ss"
    assert built.number_format(int(built.cell("Times", "A2").get("s")))[1] == "hh:mm:ss"
    assert built.cell("Times", "A2").findtext("x:v", namespaces=NS) == "0.75"


# -- tables --------------------------------------------------------------------------------------

def test_table_parts_and_relationships(built):
    tables = built.related("xl/worksheets/sheet1.xml", REL_TABLE)
    assert tables == ["xl/tables/table1.xml"]
    part = built.sheet("Stock").find("x:tableParts", NS)
    assert part.get("count") == "1"
    rel_id = part.find("x:tablePart", NS).get(f"{{{NS_R}}}id")
    assert built.rels("xl/worksheets/sheet1.xml")[rel_id] == (REL_TABLE, "xl/tables/table1.xml")
    assert built.related("xl/worksheets/sheet4.xml", REL_TABLE) == ["xl/tables/table2.xml"]


def test_table_with_header_and_totals_rows(built):
    table = built.tree("xl/tables/table1.xml")
    assert (table.get("id"), table.get("name"), table.get("displayName"), table.get("ref")) == (
        "1", "Products", "Products", "A1:E5")
    assert table.get("headerRowCount") is None and table.get("totalsRowCount") == "1"
    assert table.find("x:autoFilter", NS).get("ref") == "A1:E4"
    columns = table.find("x:tableColumns", NS)
    assert columns.get("count") == "5"
    assert [(c.get("id"), c.get("name")) for c in columns] == [
        ("1", "SKU"), ("2", "Item"), ("3", "Price"), ("4", "Added"), ("5", "Active")]
    assert columns[0].get("totalsRowLabel") == "Total"
    assert columns[2].get("totalsRowFunction") == "sum"


def test_table_without_header_row(built):
    table = built.tree("xl/tables/table2.xml")
    assert table.get("id") == "2" and table.get("headerRowCount") == "0"
    assert table.find("x:autoFilter", NS) is None
    assert built.cell("Q3 sales", "A1").findtext("x:v", namespaces=NS) == "1"


def test_table_writes_its_header_cells():
    book = WorkbookBuilder()
    sheet = book.sheet("Stock")
    sheet.rows("A2", [["SKU-0001", 3]])
    sheet.table("Items", "A1:B2", ["SKU", "Qty"], style_name="Banded")
    built = Built(book.to_bytes())
    assert _text(built.shared_string(int(built.cell("Stock", "B1").findtext("x:v", namespaces=NS)))) == "Qty"
    assert built.tree("xl/tables/table1.xml").find("x:tableStyleInfo", NS).get("name") == "Banded"


def test_table_refusals():
    sheet = WorkbookBuilder().sheet("S")
    with pytest.raises(ValueError):
        sheet.table("T", "A1:C3", ["One", "Two"])
    with pytest.raises(ValueError):
        sheet.table("T", "A1:B3", ["One", "Two"], totals_functions={"Two": "median"})
    with pytest.raises(ValueError):
        sheet.table("T", "A1:B3", ["One", "Two"], totals_labels={"Three": "x"})
    with pytest.raises(ValueError):
        sheet.table("T", "A1:B2", ["One", "Two"], totals_row=True)


# -- names, merges, comments, validation ---------------------------------------------------------

def test_defined_names_workbook_and_sheet_scope(built):
    names = built.tree("xl/workbook.xml").findall("x:definedNames/x:definedName", NS)
    assert [(n.get("name"), n.get("localSheetId"), n.get("hidden"), n.text) for n in names] == [
        ("Prices", None, None, "Stock!$C$2:$C$4"), ("Rate", "3", "1", "0.21")]
    with pytest.raises(ValueError):
        WorkbookBuilder().define_name("X", "1", sheet="Nowhere")


def test_workbook_children_in_schema_order():
    book = WorkbookBuilder(date1904=True, protected=True, calc={"fullCalcOnLoad": "1"})
    book.sheet("S").cell("A1", 1)
    book.define_name("One", "S!$A$1")
    workbook = Built(book.to_bytes()).tree("xl/workbook.xml")
    assert [etree.QName(c).localname for c in workbook] == [
        "workbookPr", "workbookProtection", "sheets", "definedNames", "calcPr"]
    assert workbook.find("x:workbookProtection", NS).get("lockStructure") == "1"
    assert workbook.find("x:calcPr", NS).get("fullCalcOnLoad") == "1"


def test_merges(built):
    merges = built.sheet("Stock").find("x:mergeCells", NS)
    assert merges.get("count") == "1" and merges.find("x:mergeCell", NS).get("ref") == "I1:J2"


def test_comments(built):
    assert built.related("xl/worksheets/sheet1.xml", REL_COMMENTS) == ["xl/comments1.xml"]
    comments = built.tree("xl/comments1.xml")
    assert [a.text for a in comments.findall("x:authors/x:author", NS)] == ["Reviewer", "Second reviewer"]
    entries = comments.findall("x:commentList/x:comment", NS)
    assert [(c.get("ref"), c.get("authorId")) for c in entries] == [("C2", "0"), ("C3", "1"), ("D2", "0")]
    assert _text(entries[0].find("x:text", NS)) == "Checked against the delivery note"
    assert entries[1].find("x:text/x:r/x:rPr/x:b", NS) is not None
    assert built.related("xl/worksheets/sheet2.xml", REL_COMMENTS) == []


def test_data_validation_lists(built):
    rules = built.sheet("Stock").find("x:dataValidations", NS)
    assert rules.get("count") == "3"
    inline, source, decimal = rules
    assert (inline.get("type"), inline.get("sqref"), inline.get("allowBlank")) == ("list", "E2:E4", "1")
    assert inline.findtext("x:formula1", namespaces=NS) == '"Yes,No,Say ""maybe"""'
    assert source.findtext("x:formula1", namespaces=NS) == "Lists!$A$1:$A$2"
    assert source.get("allowBlank") is None
    assert (decimal.get("type"), decimal.get("operator")) == ("decimal", "between")
    assert [decimal.findtext(f"x:formula{n}", namespaces=NS) for n in (1, 2)] == ["0", "1000"]
    sheet = WorkbookBuilder().sheet("S")
    with pytest.raises(ValueError):
        sheet.validation_list("A1", ["a,b"])
    with pytest.raises(ValueError):
        sheet.validation_list("A1")


# -- the messy cases -----------------------------------------------------------------------------

def test_rows_and_cells_without_references():
    book = WorkbookBuilder()
    sheet = book.sheet("Loose")
    sheet.omit_refs = True
    sheet.rows("A1", [[1, 2, None, 4], [5]])
    sheet.cell("B4", 6)
    built = Built(book.to_bytes())
    rows = built.sheet("Loose").findall("x:sheetData/x:row", NS)
    assert [r.get("r") for r in rows] == [None, None, "4"]
    assert [c.get("r") for c in rows[0]] == [None, None, "D1"]
    assert [c.get("r") for c in rows[1]] == [None]
    assert [c.get("r") for c in rows[2]] == ["B4"]


def test_dimension_given_or_left_out():
    book = WorkbookBuilder()
    wrong = book.sheet("Wrong")
    wrong.rows("A1", [[1, 2], [3, 4]])
    wrong.dimension = "A1"
    absent = book.sheet("Absent")
    absent.cell("C3", 1)
    absent.dimension = False
    book.sheet("Empty")
    built = Built(book.to_bytes())
    assert built.sheet("Wrong").find("x:dimension", NS).get("ref") == "A1"
    assert built.sheet("Absent").find("x:dimension", NS) is None
    empty = built.sheet("Empty")
    assert [etree.QName(c).localname for c in empty] == ["sheetData"] and len(empty[0]) == 0


def test_hidden_rows_columns_and_protection():
    book = WorkbookBuilder()
    sheet = book.sheet("Locked")
    sheet.protected = True
    sheet.rows("A1", [[1, 2, 3], [4, 5, 6]])
    sheet.hide_row(2).hide_row(9).hide_column("B")
    root = Built(book.to_bytes()).sheet("Locked")
    rows = {r.get("r"): r for r in root.iterfind("x:sheetData/x:row", NS)}
    assert rows["2"].get("hidden") == "1" and rows["1"].get("hidden") is None
    assert rows["9"].get("hidden") == "1" and len(rows["9"]) == 0
    col = root.find("x:cols/x:col", NS)
    assert (col.get("min"), col.get("max"), col.get("hidden")) == ("2", "2", "1")
    assert root.find("x:sheetProtection", NS).get("sheet") == "1"
    tags = [etree.QName(c).localname for c in root]
    assert tags == sorted(tags, key=WORKSHEET_ORDER.index)
