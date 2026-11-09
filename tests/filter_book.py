"""Made-up workbooks for AutoFilters, hidden rows and pivot refresh (round 44, R55, R61).

``stock_book()``: ``Stock`` holds ``A1:F8`` -- Item, Qty, Shelf (two blank), Price (``0.00``),
Value (a formula with its cached value), Arrived (dates) -- and ``A10`` below it;
``Orders`` a table ``Orders`` (``A1:C6``: Order, Item, Qty).  ``hidden=`` rows of
``Stock`` hidden by hand.  ``pivot_book()``: ``Data`` (``A1:B5``, the quantities summing to
10) and ``Summary``, a pivot table ``Totals`` whose cache (records 5, 5, 5, 5) and shown
value (20) are stale; ``refresh`` writes its ``refreshOnLoad`` (``None``: no attribute).
Every part is written here (ECMA-376 Part 1); nothing comes from an application.
"""

from __future__ import annotations

from datetime import date

from build import Formula, WorkbookBuilder, zip_parts

X = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = R + "/"
CT = "application/vnd.openxmlformats-officedocument."
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

STOCK = [["Item", "Qty", "Shelf", "Price", "Value", "Arrived"],
         ["Lamp", 3, "A", 1.5, Formula("B2*D2", 4.5), date(2026, 1, 5)],
         ["Desk", 12, "B", 24, Formula("B3*D3", 288), date(2026, 2, 7)],
         ["Chair", 7, None, 0.35, Formula("B4*D4", 2.45), date(2026, 2, 9)],
         ["Lamp", 5, "A", 2, Formula("B5*D5", 10), date(2025, 12, 1)],
         ["Shelf", 120, "C", 1.5, Formula("B6*D6", 180), date(2026, 3, 1)],
         ["lamp", 4, "B", 3, Formula("B7*D7", 12), date(2026, 1, 20)],
         ["Desk", 9, None, 24, Formula("B8*D8", 216), date(2026, 1, 5)]]


def stock_builder(*, hidden: tuple[int, ...] = ()) -> WorkbookBuilder:
    book = WorkbookBuilder()
    money = book.style(num_fmt="0.00")
    day = book.style(num_fmt_id=14)
    sheet = book.sheet("Stock")
    sheet.rows("A1", STOCK)
    for row in range(2, 9):
        sheet.cell(f"D{row}", STOCK[row - 1][3], style=money)
        sheet.cell(f"F{row}", STOCK[row - 1][5], style=day)
    sheet.cell("A10", "below")
    for row in hidden:
        sheet.hide_row(row)
    orders = book.sheet("Orders")
    orders.rows("A1", [["Order", "Item", "Qty"], [1, "Lamp", 2], [2, "Desk", 1], [3, "Lamp", 4],
                       [4, "Chair", 1], [5, "Shelf", 3]])
    orders.table("Orders", "A1:C6", ["Order", "Item", "Qty"], write_header=False)
    return book


def stock_book(*, hidden: tuple[int, ...] = ()) -> bytes:
    return stock_builder(hidden=hidden).to_bytes()


def _rels(*rels: tuple[str, str, str]) -> bytes:
    body = "".join(f'<Relationship Id="{i}" Type="{REL}{t}" Target="{target}"/>'
                   for i, t, target in rels)
    return (DECL + f'<Relationships xmlns="{PKG}">{body}</Relationships>').encode()


def pivot_book(refresh: "bool | None" = None) -> bytes:
    book = WorkbookBuilder()
    data = book.sheet("Data")
    data.rows("A1", [["Item", "Qty"], ["Lamp", 1], ["Desk", 2], ["Chair", 3], ["Shelf", 4]])
    summary = book.sheet("Summary")
    summary.cell("A3", "Sum of Qty")
    summary.cell("A4", 20)
    parts = book.parts()
    flag = {None: "", True: ' refreshOnLoad="1"', False: ' refreshOnLoad="0"'}[refresh]
    items = "".join(f'<s v="{n}"/>' for n in ["Lamp", "Desk", "Chair", "Shelf"])
    parts["xl/pivotCache/pivotCacheDefinition1.xml"] = (
        DECL + f'<pivotCacheDefinition xmlns="{X}" xmlns:r="{R}" r:id="rId1"{flag} '
        'refreshedBy="Clerk" refreshedDate="46000" createdVersion="6" refreshedVersion="6" '
        'minRefreshableVersion="3" recordCount="4"><cacheSource type="worksheet">'
        '<worksheetSource ref="A1:B5" sheet="Data"/></cacheSource><cacheFields count="2">'
        f'<cacheField name="Item" numFmtId="0"><sharedItems count="4">{items}</sharedItems>'
        '</cacheField><cacheField name="Qty" numFmtId="0"><sharedItems '
        'containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" '
        'minValue="5" maxValue="5"/></cacheField></cacheFields></pivotCacheDefinition>'
    ).encode()
    parts["xl/pivotCache/pivotCacheRecords1.xml"] = (
        DECL + f'<pivotCacheRecords xmlns="{X}" count="4">'
        + "".join(f'<r><x v="{i}"/><n v="5"/></r>' for i in range(4))
        + '</pivotCacheRecords>').encode()
    parts["xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"] = _rels(
        ("rId1", "pivotCacheRecords", "pivotCacheRecords1.xml"))
    parts["xl/pivotTables/pivotTable1.xml"] = (
        DECL + f'<pivotTableDefinition xmlns="{X}" name="Totals" cacheId="1" '
        'applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" '
        'applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" '
        'dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" '
        'useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" '
        'outlineData="1" multipleFieldFilters="0"><location ref="A3:A4" firstHeaderRow="1" '
        'firstDataRow="1" firstDataCol="0"/><pivotFields count="2"><pivotField showAll="0"/>'
        '<pivotField dataField="1" showAll="0"/></pivotFields><rowItems count="1"><i/>'
        '</rowItems><colItems count="1"><i/></colItems><dataFields count="1"><dataField '
        'name="Sum of Qty" fld="1" baseField="0" baseItem="0"/></dataFields>'
        '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" '
        'showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>'
    ).encode()
    parts["xl/pivotTables/_rels/pivotTable1.xml.rels"] = _rels(
        ("rId1", "pivotCacheDefinition", "../pivotCache/pivotCacheDefinition1.xml"))
    parts["xl/worksheets/_rels/sheet2.xml.rels"] = _rels(
        ("rId1", "pivotTable", "../pivotTables/pivotTable1.xml"))
    parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace(
        b"</sheets>", b'</sheets><pivotCaches><pivotCache cacheId="1" r:id="rId20"/>'
        b"</pivotCaches>")
    parts["xl/_rels/workbook.xml.rels"] = parts["xl/_rels/workbook.xml.rels"].replace(
        b"</Relationships>", f'<Relationship Id="rId20" Type="{REL}pivotCacheDefinition" '
        'Target="pivotCache/pivotCacheDefinition1.xml"/></Relationships>'.encode())
    types = "".join(f'<Override PartName="/xl/{p}" ContentType="{CT}spreadsheetml.{t}+xml"/>'
                    for p, t in [("pivotCache/pivotCacheDefinition1.xml",
                                  "pivotCacheDefinition"),
                                 ("pivotCache/pivotCacheRecords1.xml", "pivotCacheRecords"),
                                 ("pivotTables/pivotTable1.xml", "pivotTable")])
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
        b"</Types>", types.encode() + b"</Types>")
    return zip_parts(parts)
