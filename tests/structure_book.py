"""A made-up workbook for inserting, deleting and moving rows and columns (round 44): a
stock sheet with everything that refers to its cells, built in code.

``Stock`` (the sheet the operations work on) holds a table ``Goods`` (``A1:F7``: Item, Qty,
Price, Value, Shelf, Link) whose ``Value`` column is a shared formula group, an array formula
(``G2:G3``, in a hidden column), formulas of every reference form (relative, anchored,
ranges, whole columns and rows, a structured reference, another sheet's cell), merged cells
(``A10:C10``, ``E10:F11``), data validations and conditional formats with formulas,
hyperlinks (a place in the workbook, a web address), comments with their VML drawing, a
sheet filter (``J1:K7``), frozen panes, column widths, a hidden column, a tall row, a hidden
row and a styled row.  ``Orders`` refers to ``Stock`` from outside: every reference form
again, a 3-D reference (``Stock:Notes!B2``), a structured reference, another workbook's
price, and a chart whose series read ``Stock``'s cells.  ``Notes`` holds the 3-D reference's
last sheet; ``Pivot`` a pivot table whose cache reads ``Stock!A1:C7``.  Defined names: a
cell, a block, a whole column, a relative name, a sheet-scoped name, a 3-D name, an
``OFFSET``, the print area and titles and the sheet filter's database.

``StructureBook(pivot=False, chart=False, external=False, comments=False)`` leaves those
parts out.  Every part is written here (ECMA-376 Part 1, §18), nothing comes from an
application; what Excel makes of each operation on it is the oracle's answer
(``tests/test_structure_oracle.py``).
"""

from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass
from xml.sax.saxutils import escape

X = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
_SML = "application/vnd.openxmlformats-officedocument."

#: The table's rows: item, quantity, price, shelf, a link's text.
GOODS = [
    ("Pencil", 120, 0.35, "A", "see orders"),
    ("Notebook", 40, 2.5, "B", "supplier"),
    ("Kettle", 3, 24.0, "A", ""),
    ("Mug", 18, 4.75, "C", "note"),
    ("Torch", 7, 9.99, "B", ""),
    ("Basket", 11, 6.0, "C", ""),
]
#: The sheet filter's block (J1:K7): a code and a count.
CODES = [("K-1", 4), ("K-2", 9), ("K-3", 1), ("K-4", 6), ("K-5", 2), ("K-6", 8)]

#: Stock's own formulas, column H: cell -> formula.
STOCK_FORMULAS = {
    "H1": "SUM(B2:B7)",
    "H2": "B4",
    "H3": "$B$4",
    "H4": "SUM(B4:B5)",
    "H5": "Orders!A1",
    "H6": "SUM(B:B)",
    "H7": "SUM(3:3)",
    "H8": "SUM(Goods[Qty])",
    "H9": "SUM(B2:B1048576)",
    "H10": "Notes!B2",
    "H11": "Rel*1",
    "H12": "SUM($B$2:B4)",
}
#: Orders' formulas into Stock, column B: cell -> formula.
ORDERS_FORMULAS = {
    "B1": "Stock!B3",
    "B2": "Stock!$B$3",
    "B3": "SUM(Stock!B2:B7)",
    "B4": "SUM(Stock!B4:B5)",
    "B5": "Stock!B5",
    "B6": "SUM(Stock!B:B)",
    "B7": "SUM(Stock!2:7)",
    "B8": "SUM(Stock:Notes!B2)",
    "B9": "SUM(Stock!B2:B4)",
    "B10": "SUM(Stock!B5:B9)",
    "B11": "Rate*2",
    "B12": "SUM(Goods[Price])",
    "B13": 'INDIRECT("Stock!B3")',
    "B14": "[1]Prices!$B$2",
    "B15": "SUM(Stock!$B$2:$B$7)",
    "B16": "SUM(Stock!C2,Stock!D3)",
    "B17": "Stock!XFD1",
    "B18": "COUNTA(Goods[[#All],[Item]])",
    "B19": "SUM(Block)",
    "B20": "Stock!E2&Stock!J3",
    "B21": "SUM(Stock!A1:K12)",
    "B22": "Local",
    "B23": "ROWS(Dyn)",
    "B24": "SUM(Stock!C:D)",
    "B25": "SUM(Stock!4:5)",
    "B26": "SUM(Stock!K2:K7)",
    "B27": "SUM(Stock!A2:F7)",
    "B28": "SUM(Stock!A2:G7)",
    "B29": "SUM(Stock!B2:B6)",
    "B30": "SUM(Stock!B3:B7)",
    "B31": "SUM(Stock!B1:B7)",
    "B32": "SUM(Stock!B2:B8)",
    "B33": "COUNTA(Stock!B10:C11)",
    "B34": "SUM(Stock!B1:B1048576)",
    "B35": "SUM(Stock!A3:XFD3)",
    "B36": "COUNT(Stock!B:D)",
}

#: The names the workbook defines: (name, local sheet index or None, hidden, formula).
NAMES = [
    ("_xlnm._FilterDatabase", 0, True, "Stock!$J$1:$K$7"),
    ("_xlnm.Print_Area", 0, False, "Stock!$A$1:$H$12"),
    ("_xlnm.Print_Titles", 0, False, "Stock!$A:$A,Stock!$1:$1"),
    ("Block", None, False, "Stock!$B$2:$C$7"),
    ("Dyn", None, False, "OFFSET(Stock!$A$1,0,0,COUNTA(Stock!$A:$A),1)"),
    ("Local", 1, False, "Stock!$B$5"),
    ("Rate", None, False, "Stock!$C$2"),
    ("Rel", None, False, "Stock!B3"),
    ("Span3D", None, False, "Stock:Notes!$B$2"),
    ("WholeQty", None, False, "Stock!$B:$B"),
]


def _cell(ref: str, value, style: int | None = None) -> str:
    s = f' s="{style}"' if style else ""
    if value is None:
        return f'<c r="{ref}"{s}/>'
    if isinstance(value, str):
        return f'<c r="{ref}"{s} t="inlineStr"><is><t>{escape(value)}</t></is></c>'
    return f'<c r="{ref}"{s}><v>{_num(value)}</v></c>'


def _formula(ref: str, text: str, value=0, extra: str = "") -> str:
    if isinstance(value, str):
        return f'<c r="{ref}" t="str"><f{extra}>{escape(text)}</f><v>{escape(value)}</v></c>'
    return f'<c r="{ref}"><f{extra}>{escape(text)}</f><v>{_num(value)}</v></c>'


def _num(value: float) -> str:
    value = round(value, 10)
    return str(int(value)) if float(value).is_integer() else repr(value)


@dataclass
class StructureBook:
    pivot: bool = True
    chart: bool = True
    external: bool = True
    comments: bool = True
    #: The array formula over G2:G3 (Stock).
    array: bool = True
    #: Cells at the grid's last row and column on Notes (A1048576, XFD1).
    edge: bool = False

    # -- sheets --------------------------------------------------------------------------------

    def stock(self) -> str:
        rows: dict[int, list[str]] = {n: [] for n in range(1, 13)}
        for column, text in zip("ABCDEF", ("Item", "Qty", "Price", "Value", "Shelf", "Link")):
            rows[1].append(_cell(f"{column}1", text))
        for i, (item, qty, price, shelf, link) in enumerate(GOODS):
            n = i + 2
            rows[n].append(_cell(f"A{n}", item))
            rows[n].append(_cell(f"B{n}", qty))
            rows[n].append(_cell(f"C{n}", price, 1))
            extra = ' t="shared" ref="D2:D7" si="0"' if n == 2 else ' t="shared" si="0"'
            text = "B2*C2" if n == 2 else ""
            rows[n].append(f'<c r="D{n}"><f{extra}>{text}</f><v>{_num(qty * price)}</v></c>')
            rows[n].append(_cell(f"E{n}", shelf))
            if link:
                rows[n].append(_cell(f"F{n}", link))
        if self.array:
            rows[2].append('<c r="G2"><f t="array" ref="G2:G3">B2:B3*2</f><v>240</v></c>')
            rows[3].append('<c r="G3"><v>80</v></c>')
        for ref, text in STOCK_FORMULAS.items():
            rows[int(ref[1:])].append(_formula(ref, text))
        rows[1] += [_cell("J1", "Code"), _cell("K1", "Count")]
        for i, (code, count) in enumerate(CODES):
            rows[i + 2] += [_cell(f"J{i + 2}", code), _cell(f"K{i + 2}", count)]
        rows[8].insert(0, _cell("A8", "hidden row"))
        rows[10] = [_cell("A10", "merged"), _cell("E10", "block")] + rows[10]
        rows[12].insert(0, _cell("A12", "end"))

        def order(cell: str) -> tuple[int, str]:
            ref = cell.split('r="', 1)[1].split('"', 1)[0]
            letters = ref.rstrip("0123456789")
            return (len(letters), letters)

        body = []
        for n, cells in rows.items():
            attrs = ""
            if n == 4:
                attrs = ' ht="30" customHeight="1"'
            elif n == 8:
                attrs = ' hidden="1"'
            elif n == 9:
                attrs = ' s="2" customFormat="1"'
            body.append(f'<row r="{n}"{attrs}>{"".join(sorted(cells, key=order))}</row>')
        legacy = '<legacyDrawing r:id="rId3"/>' if self.comments else ""
        return (
            _DECL + f'<worksheet xmlns="{X}" xmlns:r="{R}">'
            '<dimension ref="A1:K12"/>'
            '<sheetViews><sheetView tabSelected="1" workbookViewId="0">'
            '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" '
            'state="frozen"/><selection pane="topRight" activeCell="B1" sqref="B1"/>'
            '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
            '<selection pane="bottomRight" activeCell="C5" sqref="C5"/></sheetView>'
            '</sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/>'
            '<cols><col min="1" max="1" width="20.7109375" customWidth="1"/>'
            '<col min="3" max="3" width="12.7109375" style="1" customWidth="1"/>'
            '<col min="7" max="7" width="9.140625" hidden="1"/>'
            '<col min="9" max="10" width="15.7109375" customWidth="1"/></cols>'
            f'<sheetData>{"".join(body)}</sheetData>'
            '<autoFilter ref="J1:K7"/>'
            '<mergeCells count="2"><mergeCell ref="A10:C10"/><mergeCell ref="E10:F11"/>'
            '</mergeCells>'
            '<conditionalFormatting sqref="C2:C7"><cfRule type="cellIs" dxfId="0" '
            'priority="1" operator="greaterThan"><formula>$C$2</formula></cfRule>'
            '</conditionalFormatting>'
            '<conditionalFormatting sqref="D2:D7 B10"><cfRule type="expression" dxfId="1" '
            'priority="2"><formula>D2&gt;B2*10</formula></cfRule></conditionalFormatting>'
            '<conditionalFormatting sqref="K2:K7"><cfRule type="expression" dxfId="0" '
            'priority="3"><formula>K2&gt;Orders!$A$2</formula></cfRule>'
            '</conditionalFormatting>'
            '<dataValidations count="3">'
            '<dataValidation type="list" allowBlank="1" showInputMessage="1" '
            'showErrorMessage="1" sqref="E2:E7"><formula1>"A,B,C"</formula1></dataValidation>'
            '<dataValidation type="custom" allowBlank="1" showInputMessage="1" '
            'showErrorMessage="1" sqref="B2:B7"><formula1>B2&gt;=0</formula1>'
            '</dataValidation>'
            '<dataValidation type="decimal" operator="lessThan" allowBlank="1" '
            'showInputMessage="1" showErrorMessage="1" sqref="C2:C7 C9">'
            '<formula1>$H$1</formula1></dataValidation></dataValidations>'
            '<hyperlinks><hyperlink ref="F2" location="Orders!A1" display="see orders"/>'
            '<hyperlink ref="F3" r:id="rId2"/></hyperlinks>'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" '
            'footer="0.3"/>'
            f'{legacy}<tableParts count="1"><tablePart r:id="rId1"/></tableParts>'
            '</worksheet>')

    def goods_table(self) -> str:
        names = ("Item", "Qty", "Price", "Value", "Shelf", "Link")
        columns = "".join(f'<tableColumn id="{i + 1}" name="{n}"/>'
                          for i, n in enumerate(names))
        return (_DECL + f'<table xmlns="{X}" id="1" name="Goods" displayName="Goods" '
                'ref="A1:F7" totalsRowShown="0"><autoFilter ref="A1:F7"/>'
                f'<tableColumns count="6">{columns}</tableColumns>'
                '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" '
                'showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>')

    def comments_xml(self) -> str:
        items = [("B3", "check the count"), ("F5", "ask the supplier")]
        body = "".join(f'<comment ref="{ref}" authorId="0"><text><r><t>{escape(text)}</t></r>'
                       '</text></comment>' for ref, text in items)
        return (_DECL + f'<comments xmlns="{X}"><authors><author>Clerk</author></authors>'
                f'<commentList>{body}</commentList></comments>')

    def vml(self) -> str:
        shapes = []
        for n, (row, column) in enumerate(((2, 1), (4, 5))):
            shapes.append(
                f'<v:shape id="_x0000_s{1025 + n}" type="#_x0000_t202" style="position:'
                'absolute;margin-left:80pt;margin-top:2pt;width:108pt;height:59.25pt;'
                'z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">'
                '<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/>'
                '<v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto">'
                '<div style="text-align:left"></div></v:textbox>'
                '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>'
                f'<x:Anchor>{column + 1}, 15, {row - 1}, 10, {column + 3}, 15, {row + 3}, 4'
                f'</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>{row}</x:Row>'
                f'<x:Column>{column}</x:Column></x:ClientData></v:shape>')
        return ('<xml xmlns:v="urn:schemas-microsoft-com:vml" '
                'xmlns:o="urn:schemas-microsoft-com:office:office" '
                'xmlns:x="urn:schemas-microsoft-com:office:excel">'
                '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>'
                '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" '
                'path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/>'
                '<v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
                + "".join(shapes) + "</xml>")

    def orders(self) -> str:
        body = [_cell("A1", "start")]
        lines: dict[int, list[str]] = {1: body, 2: [_cell("A2", 1)]}
        for ref, text in ORDERS_FORMULAS.items():
            if text.startswith("[1]") and not self.external:
                continue
            value = "" if "&" in text else 0
            lines.setdefault(int(ref[1:]), []).append(_formula(ref, text, value))
        rows = [f'<row r="{n}">{"".join(cells)}</row>' for n, cells in sorted(lines.items())]
        drawing = '<drawing r:id="rId1"/>' if self.chart else ""
        return (_DECL + f'<worksheet xmlns="{X}" xmlns:r="{R}"><dimension ref="A1:B36"/>'
                '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
                '<sheetFormatPr defaultRowHeight="15"/>'
                f'<sheetData>{"".join(rows)}</sheetData>'
                '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" '
                f'footer="0.3"/>{drawing}</worksheet>')

    def notes(self) -> str:
        edge = ('<row r="1"><c r="XFD1"><v>1</v></c></row>' if self.edge else "")
        last = ('<row r="1048576"><c r="A1048576"><v>1</v></c></row>' if self.edge else "")
        return (_DECL + f'<worksheet xmlns="{X}"><dimension ref="B2"/>'
                '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
                '<sheetFormatPr defaultRowHeight="15"/>'
                f'<sheetData>{edge}<row r="2"><c r="B2"><v>5</v></c></row>{last}</sheetData>'
                '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" '
                'footer="0.3"/></worksheet>')

    def pivot_sheet(self) -> str:
        return (_DECL + f'<worksheet xmlns="{X}"><dimension ref="A3:A4"/>'
                '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
                '<sheetFormatPr defaultRowHeight="15"/>'
                '<sheetData><row r="3"><c r="A3" t="inlineStr"><is><t>Sum of Qty</t></is></c>'
                '</row><row r="4"><c r="A4"><v>199</v></c></row></sheetData>'
                '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" '
                'footer="0.3"/></worksheet>')

    def pivot_cache(self) -> str:
        return (_DECL + f'<pivotCacheDefinition xmlns="{X}" xmlns:r="{R}" '
                'refreshOnLoad="1" refreshedBy="Clerk" refreshedDate="46000" '
                'createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" '
                'recordCount="0">'
                '<cacheSource type="worksheet"><worksheetSource ref="A1:C7" sheet="Stock"/>'
                '</cacheSource><cacheFields count="3">'
                '<cacheField name="Item" numFmtId="0"><sharedItems/></cacheField>'
                '<cacheField name="Qty" numFmtId="0"><sharedItems containsSemiMixedTypes="0" '
                'containsString="0" containsNumber="1" containsInteger="1" minValue="3" '
                'maxValue="120"/></cacheField>'
                '<cacheField name="Price" numFmtId="0"><sharedItems containsSemiMixedTypes='
                '"0" containsString="0" containsNumber="1" minValue="0.35" maxValue="24"/>'
                '</cacheField></cacheFields></pivotCacheDefinition>')

    def pivot_table(self) -> str:
        return (_DECL + f'<pivotTableDefinition xmlns="{X}" name="Totals" cacheId="1" '
                'applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" '
                'applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" '
                'dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" '
                'useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" '
                'outline="1" outlineData="1" multipleFieldFilters="0">'
                '<location ref="A3:A4" firstHeaderRow="1" firstDataRow="1" firstDataCol="0"/>'
                '<pivotFields count="3"><pivotField showAll="0"/>'
                '<pivotField dataField="1" showAll="0"/><pivotField showAll="0"/></pivotFields>'
                '<rowItems count="1"><i/></rowItems><colItems count="1"><i/></colItems>'
                '<dataFields count="1"><dataField name="Sum of Qty" fld="1" baseField="0" '
                'baseItem="0"/></dataFields>'
                '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" '
                'showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>'
                '</pivotTableDefinition>')

    def drawing(self) -> str:
        return (_DECL + '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/'
                '2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/'
                'drawingml/2006/main"><xdr:twoCellAnchor>'
                '<xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row>'
                '<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>9</xdr:col>'
                '<xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff>'
                '</xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>'
                '<xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/>'
                '</xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
                '</xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/'
                'drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/'
                f'drawingml/2006/chart" xmlns:r="{R}" r:id="rId1"/></a:graphicData>'
                '</a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>'
                '</xdr:wsDr>')

    def chart_xml(self) -> str:
        cats = "".join(f'<c:pt idx="{i}"><c:v>{g[0]}</c:v></c:pt>' for i, g in enumerate(GOODS))
        vals = "".join(f'<c:pt idx="{i}"><c:v>{g[1]}</c:v></c:pt>' for i, g in enumerate(GOODS))
        n = len(GOODS)
        return (_DECL + '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/'
                '2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                f'xmlns:r="{R}"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>'
                '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>'
                '<c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/>'
                '<c:tx><c:strRef><c:f>Stock!$B$1</c:f><c:strCache><c:ptCount val="1"/>'
                '<c:pt idx="0"><c:v>Qty</c:v></c:pt></c:strCache></c:strRef></c:tx>'
                f'<c:cat><c:strRef><c:f>Stock!$A$2:$A$7</c:f><c:strCache><c:ptCount val="{n}"/>'
                f'{cats}</c:strCache></c:strRef></c:cat>'
                f'<c:val><c:numRef><c:f>Stock!$B$2:$B$7</c:f><c:numCache><c:formatCode>General'
                f'</c:formatCode><c:ptCount val="{n}"/>{vals}</c:numCache></c:numRef></c:val>'
                '</c:ser><c:gapWidth val="150"/><c:axId val="10"/><c:axId val="20"/></c:barChart>'
                '<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
                '<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>'
                '<c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
                '<c:delete val="0"/><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>'
                '</c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>')

    def external_link(self) -> str:
        return (_DECL + f'<externalLink xmlns="{X}"><externalBook xmlns:r="{R}" r:id="rId1">'
                '<sheetNames><sheetName val="Prices"/></sheetNames><sheetDataSet>'
                '<sheetData sheetId="0"><row r="2"><cell r="B2"><v>3.5</v></cell></row>'
                '</sheetData></sheetDataSet></externalBook></externalLink>')

    def workbook(self) -> str:
        names = "".join(
            f'<definedName name="{n}"'
            + (f' localSheetId="{local}"' if local is not None else "")
            + (' hidden="1"' if hidden else "") + f'>{escape(text)}</definedName>'
            for n, local, hidden, text in NAMES)
        sheets = ('<sheet name="Stock" sheetId="1" r:id="rId1"/>'
                  '<sheet name="Orders" sheetId="2" r:id="rId2"/>'
                  '<sheet name="Notes" sheetId="3" r:id="rId3"/>'
                  + ('<sheet name="Pivot" sheetId="4" r:id="rId4"/>' if self.pivot else ""))
        links = ' updateLinks="never"' if self.external else ""
        return (_DECL + f'<workbook xmlns="{X}" xmlns:r="{R}">'
                f'<workbookPr{links} defaultThemeVersion="166925"/>'
                '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" '
                f'windowHeight="12300"/></bookViews><sheets>{sheets}</sheets>'
                + ('<externalReferences><externalReference r:id="rId8"/></externalReferences>'
                   if self.external else "")
                + f'<definedNames>{names}</definedNames><calcPr calcId="191029"/>'
                + ('<pivotCaches><pivotCache cacheId="1" r:id="rId9"/></pivotCaches>'
                   if self.pivot else "")
                + '</workbook>')

    def styles(self) -> str:
        return (_DECL + f'<styleSheet xmlns="{X}">'
                '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.00"/></numFmts>'
                '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>'
                '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
                '<fills count="3"><fill><patternFill patternType="none"/></fill>'
                '<fill><patternFill patternType="gray125"/></fill>'
                '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/>'
                '<bgColor indexed="64"/></patternFill></fill></fills>'
                '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border>'
                '</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" '
                'borderId="0"/></cellStyleXfs><cellXfs count="3">'
                '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
                '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" '
                'applyNumberFormat="1"/>'
                '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" '
                'applyFill="1"/></cellXfs>'
                '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/>'
                '</cellStyles><dxfs count="2"><dxf><font><color rgb="FF9C0006"/></font></dxf>'
                '<dxf><fill><patternFill><bgColor rgb="FFFFEB9C"/></patternFill></fill></dxf>'
                '</dxfs><tableStyles count="0" defaultTableStyle="TableStyleMedium2"/>'
                '</styleSheet>')

    # -- the package -----------------------------------------------------------------------------

    def parts(self) -> dict[str, bytes]:
        parts: dict[str, str | bytes] = {}
        overrides = [("/xl/workbook.xml", "spreadsheetml.sheet.main+xml"),
                     ("/xl/styles.xml", "spreadsheetml.styles+xml"),
                     ("/xl/worksheets/sheet1.xml", "spreadsheetml.worksheet+xml"),
                     ("/xl/worksheets/sheet2.xml", "spreadsheetml.worksheet+xml"),
                     ("/xl/worksheets/sheet3.xml", "spreadsheetml.worksheet+xml"),
                     ("/xl/tables/table1.xml", "spreadsheetml.table+xml")]
        rels = [("rId1", "worksheet", "worksheets/sheet1.xml"),
                ("rId2", "worksheet", "worksheets/sheet2.xml"),
                ("rId3", "worksheet", "worksheets/sheet3.xml"),
                ("rId5", "styles", "styles.xml")]
        parts["xl/worksheets/sheet1.xml"] = self.stock()
        parts["xl/worksheets/sheet2.xml"] = self.orders()
        parts["xl/worksheets/sheet3.xml"] = self.notes()
        parts["xl/tables/table1.xml"] = self.goods_table()
        stock_rels = [("rId1", "table", "../tables/table1.xml")]
        external = [("rId2", "hyperlink", "https://example.com/stock")]
        if self.comments:
            stock_rels += [("rId3", "vmlDrawing", "../drawings/vmlDrawing1.vml"),
                           ("rId4", "comments", "../comments1.xml")]
            parts["xl/comments1.xml"] = self.comments_xml()
            parts["xl/drawings/vmlDrawing1.vml"] = self.vml()
            overrides.append(("/xl/comments1.xml", "spreadsheetml.comments+xml"))
        parts["xl/worksheets/_rels/sheet1.xml.rels"] = _rels(stock_rels, external)
        if self.chart:
            parts["xl/worksheets/_rels/sheet2.xml.rels"] = _rels(
                [("rId1", "drawing", "../drawings/drawing1.xml")])
            parts["xl/drawings/drawing1.xml"] = self.drawing()
            parts["xl/drawings/_rels/drawing1.xml.rels"] = _rels(
                [("rId1", "chart", "../charts/chart1.xml")])
            parts["xl/charts/chart1.xml"] = self.chart_xml()
            overrides += [("/xl/drawings/drawing1.xml", "drawing+xml"),
                          ("/xl/charts/chart1.xml", "drawingml.chart+xml")]
        if self.pivot:
            rels += [("rId4", "worksheet", "worksheets/sheet4.xml"),
                     ("rId9", "pivotCacheDefinition", "pivotCache/pivotCacheDefinition1.xml")]
            parts["xl/worksheets/sheet4.xml"] = self.pivot_sheet()
            parts["xl/worksheets/_rels/sheet4.xml.rels"] = _rels(
                [("rId1", "pivotTable", "../pivotTables/pivotTable1.xml")])
            parts["xl/pivotTables/pivotTable1.xml"] = self.pivot_table()
            parts["xl/pivotTables/_rels/pivotTable1.xml.rels"] = _rels(
                [("rId1", "pivotCacheDefinition",
                  "../pivotCache/pivotCacheDefinition1.xml")])
            parts["xl/pivotCache/pivotCacheDefinition1.xml"] = self.pivot_cache()
            overrides += [("/xl/worksheets/sheet4.xml", "spreadsheetml.worksheet+xml"),
                          ("/xl/pivotTables/pivotTable1.xml", "spreadsheetml.pivotTable+xml"),
                          ("/xl/pivotCache/pivotCacheDefinition1.xml",
                           "spreadsheetml.pivotCacheDefinition+xml")]
        if self.external:
            rels.append(("rId8", "externalLink", "externalLinks/externalLink1.xml"))
            parts["xl/externalLinks/externalLink1.xml"] = self.external_link()
            parts["xl/externalLinks/_rels/externalLink1.xml.rels"] = _rels(
                [], [("rId1", "externalLinkPath", "supplier.xlsx")])
            overrides.append(("/xl/externalLinks/externalLink1.xml",
                              "spreadsheetml.externalLink+xml"))
        parts["[Content_Types].xml"] = (
            _DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
            'content-types"><Default Extension="rels" ContentType="application/'
            'vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Default Extension="vml" ContentType="application/vnd.openxmlformats-'
            'officedocument.vmlDrawing"/>'
            + "".join(f'<Override PartName="{name}" ContentType="{_SML}{kind}"/>'
                      for name, kind in overrides) + "</Types>")
        parts["_rels/.rels"] = _rels([("rId1", "officeDocument", "xl/workbook.xml")])
        parts["xl/workbook.xml"] = self.workbook()
        parts["xl/_rels/workbook.xml.rels"] = _rels(rels)
        parts["xl/styles.xml"] = self.styles()
        return {name: data.encode("utf-8") if isinstance(data, str) else data
                for name, data in parts.items()}

    def to_bytes(self) -> bytes:
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, data in self.parts().items():
                archive.writestr(zipfile.ZipInfo(name, date_time=(2026, 1, 2, 3, 4, 6)), data,
                                 compress_type=zipfile.ZIP_DEFLATED)
        return buffer.getvalue()


def _rels(items: list[tuple[str, str, str]],
          external: "list[tuple[str, str, str]] | None" = None) -> str:
    body = "".join(f'<Relationship Id="{i}" Type="{_REL}{t}" Target="{escape(target)}"/>'
                   for i, t, target in items)
    body += "".join(f'<Relationship Id="{i}" Type="{_REL}{t}" Target="{escape(target)}" '
                    'TargetMode="External"/>' for i, t, target in external or [])
    return (_DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
            f'relationships">{body}</Relationships>')


#: GridBook's formulas on Refs, into Grid: cell -> formula.  Grid holds numbers in A1:J10,
#: nothing else: every range form against rows and columns 3 to 6.
GRID_FORMULAS = {
    **{f"A{n}": f"Grid!B{n}" for n in range(1, 10)},
    **{f"B{n}": f"Grid!{c}2" for n, c in enumerate("ABCDEFGHI", 1)},
    "C1": "SUM(Grid!B3:B6)", "C2": "SUM(Grid!B3:B3)", "C3": "SUM(Grid!B4:B5)",
    "C4": "SUM(Grid!B1:B9)", "C5": "SUM(Grid!$B$3:$B$6)", "C6": "SUM(Grid!3:6)",
    "C7": "SUM(Grid!B3:B4)", "C8": "SUM(Grid!B5:B6)", "C9": "SUM(Grid!B2:B3)",
    "C10": "SUM(Grid!B6:B7)", "C11": "SUM(Grid!B3:C6)",
    "D1": "SUM(Grid!C2:F2)", "D2": "SUM(Grid!C2:C2)", "D3": "SUM(Grid!D2:E2)",
    "D4": "SUM(Grid!A2:I2)", "D5": "SUM(Grid!$C$2:$F$2)", "D6": "SUM(Grid!C:F)",
    "D7": "SUM(Grid!C2:D2)", "D8": "SUM(Grid!E2:F2)", "D9": "SUM(Grid!B2:C2)",
    "D10": "SUM(Grid!F2:G2)",
}
#: GridBook's names: name -> formula.
GRID_NAMES = {"RowSpan": "Grid!$B$3:$B$6", "ColSpan": "Grid!$C$2:$F$2",
              "Rows36": "Grid!$3:$6", "Cols": "Grid!$C:$F"}


@dataclass
class GridBook:
    """A plain grid (``Grid``: numbers in A1:J10) and formulas into it (``Refs``): what
    Excel does to references when rows and columns move, with no table, merge or format
    in the way.  ``formats`` adds a data validation and a conditional format over B3:B6
    and C2:F2."""

    formats: bool = True

    def grid(self) -> str:
        rows = []
        for n in range(1, 11):
            cells = "".join(f'<c r="{c}{n}"><v>{n * 100 + i}</v></c>'
                            for i, c in enumerate("ABCDEFGHIJ", 1))
            rows.append(f'<row r="{n}">{cells}</row>')
        extra = ""
        if self.formats:
            extra = ('<conditionalFormatting sqref="B3:B6"><cfRule type="expression" '
                     'dxfId="0" priority="1"><formula>B3&gt;A3</formula></cfRule>'
                     '</conditionalFormatting><conditionalFormatting sqref="C2:F2">'
                     '<cfRule type="expression" dxfId="0" priority="2"><formula>C2&gt;C1'
                     '</formula></cfRule></conditionalFormatting>'
                     '<dataValidations count="2"><dataValidation type="whole" '
                     'operator="greaterThan" allowBlank="1" sqref="B3:B6"><formula1>0'
                     '</formula1></dataValidation><dataValidation type="whole" '
                     'operator="greaterThan" allowBlank="1" sqref="C2:F2"><formula1>0'
                     '</formula1></dataValidation></dataValidations>')
        return (_DECL + f'<worksheet xmlns="{X}"><dimension ref="A1:J10"/>'
                '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>'
                '<sheetFormatPr defaultRowHeight="15"/>'
                f'<sheetData>{"".join(rows)}</sheetData>{extra}'
                '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" '
                'footer="0.3"/></worksheet>')

    def refs(self) -> str:
        lines: dict[int, list[str]] = {}
        for ref, text in GRID_FORMULAS.items():
            lines.setdefault(int(ref[1:]), []).append((ref, text))
        rows = []
        for n in sorted(lines):
            cells = sorted(lines[n], key=lambda item: item[0][0])
            rows.append(f'<row r="{n}">' + "".join(_formula(r, t) for r, t in cells) + "</row>")
        return (_DECL + f'<worksheet xmlns="{X}"><dimension ref="A1:D11"/>'
                '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
                '<sheetFormatPr defaultRowHeight="15"/>'
                f'<sheetData>{"".join(rows)}</sheetData>'
                '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" '
                'footer="0.3"/></worksheet>')

    def parts(self) -> dict[str, bytes]:
        names = "".join(f'<definedName name="{n}">{escape(t)}</definedName>'
                        for n, t in GRID_NAMES.items())
        workbook = (_DECL + f'<workbook xmlns="{X}" xmlns:r="{R}">'
                    '<workbookPr defaultThemeVersion="166925"/><bookViews><workbookView/>'
                    '</bookViews><sheets><sheet name="Grid" sheetId="1" r:id="rId1"/>'
                    '<sheet name="Refs" sheetId="2" r:id="rId2"/></sheets>'
                    f'<definedNames>{names}</definedNames><calcPr calcId="191029"/>'
                    '</workbook>')
        overrides = [("/xl/workbook.xml", "spreadsheetml.sheet.main+xml"),
                     ("/xl/styles.xml", "spreadsheetml.styles+xml"),
                     ("/xl/worksheets/sheet1.xml", "spreadsheetml.worksheet+xml"),
                     ("/xl/worksheets/sheet2.xml", "spreadsheetml.worksheet+xml")]
        parts: dict[str, str] = {
            "[Content_Types].xml": (
                _DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
                'content-types"><Default Extension="rels" ContentType="application/'
                'vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                + "".join(f'<Override PartName="{name}" ContentType="{_SML}{kind}"/>'
                          for name, kind in overrides) + "</Types>"),
            "_rels/.rels": _rels([("rId1", "officeDocument", "xl/workbook.xml")]),
            "xl/workbook.xml": workbook,
            "xl/_rels/workbook.xml.rels": _rels([
                ("rId1", "worksheet", "worksheets/sheet1.xml"),
                ("rId2", "worksheet", "worksheets/sheet2.xml"),
                ("rId3", "styles", "styles.xml")]),
            "xl/worksheets/sheet1.xml": self.grid(),
            "xl/worksheets/sheet2.xml": self.refs(),
            "xl/styles.xml": StructureBook().styles(),
        }
        return {name: data.encode("utf-8") for name, data in parts.items()}

    def to_bytes(self) -> bytes:
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, data in self.parts().items():
                archive.writestr(zipfile.ZipInfo(name, date_time=(2026, 1, 2, 3, 4, 6)), data,
                                 compress_type=zipfile.ZIP_DEFLATED)
        return buffer.getvalue()


@dataclass
class FormatBook:
    """What an inserted row or column takes from its neighbours, and tables that compute:
    ``Fmt`` holds styled cells (B3 bold and filled, C3 two decimals) under a styled header
    (A1:D1), a table with a calculated column (``Calc``, F1:H5: Qty, Price, Total), a table
    with a totals row (``Sums``, J1:K6), two outlined rows (8, 9), a hyperlink to a place on
    the sheet (A2 -> Fmt!B5) and a chart anchored over A12:D20 reading B2:B5.
    ``sums=False`` leaves the totals table out, ``calculated=False`` the calculated
    column's formula (its cells keep theirs)."""

    sums: bool = True
    calculated: bool = True
    #: The table with the calculated column (``calc=False``: its cells stay, no table).
    calc: bool = True
    #: The totals row of ``Sums`` (``totals=False``: a plain table J1:K5).
    totals: bool = True

    def sheet(self) -> str:
        rows: dict[int, list[str]] = {n: [] for n in range(1, 10)}
        for column, text in zip("ABCD", ("Name", "Count", "Rate", "Note")):
            rows[1].append(_cell(f"{column}1", text, 2))
        for n in range(2, 6):
            rows[n] += [_cell(f"A{n}", f"N{n}"), _cell(f"B{n}", n * 10, 2 if n == 3 else None),
                        _cell(f"C{n}", n / 4, 1 if n == 3 else None), _cell(f"D{n}", "x")]
        for column, text in zip("FGH", ("Qty", "Price", "Total")):
            rows[1].append(_cell(f"{column}1", text))
        calc = "Calc[[#This Row],[Qty]]*Calc[[#This Row],[Price]]"
        for n in range(2, 6):
            rows[n] += [_cell(f"F{n}", n), _cell(f"G{n}", n + 0.5),
                        _formula(f"H{n}", calc, n * (n + 0.5)) if self.calc
                        else _cell(f"H{n}", n * (n + 0.5))]
        if self.sums:
            rows[1] += [_cell("J1", "Name"), _cell("K1", "Amount")]
            for n in range(2, 6):
                rows[n] += [_cell(f"J{n}", f"S{n}"), _cell(f"K{n}", n * 3)]
            if self.totals:
                rows[6] = [_cell("J6", "Total"),
                           _formula("K6", "SUBTOTAL(109,Sums[Amount])", 42)]
        rows[8] = [_cell("A8", "grouped")]
        rows[9] = [_cell("A9", "grouped too")]
        body = []
        for n, cells in rows.items():
            attrs = ' outlineLevel="1"' if n in (8, 9) else ""
            body.append(f'<row r="{n}"{attrs}>{"".join(cells)}</row>')
        return (_DECL + f'<worksheet xmlns="{X}" xmlns:r="{R}"><sheetPr>'
                '<outlinePr summaryBelow="1"/></sheetPr><dimension ref="A1:K9"/>'
                '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>'
                '<sheetFormatPr defaultRowHeight="15" outlineLevelRow="1"/>'
                f'<sheetData>{"".join(body)}</sheetData>'
                '<hyperlinks><hyperlink ref="A2" location="Fmt!B5" display="N2"/></hyperlinks>'
                '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" '
                'footer="0.3"/><drawing r:id="rId3"/>'
                + _table_parts(["rId1"] * self.calc + ["rId2"] * self.sums)
                + '</worksheet>')

    def calc_table(self) -> str:
        calc = "Calc[[#This Row],[Qty]]*Calc[[#This Row],[Price]]"
        return (_DECL + f'<table xmlns="{X}" id="1" name="Calc" displayName="Calc" '
                'ref="F1:H5" totalsRowShown="0"><autoFilter ref="F1:H5"/>'
                '<tableColumns count="3"><tableColumn id="1" name="Qty"/>'
                '<tableColumn id="2" name="Price"/><tableColumn id="3" name="Total">'
                + (f'<calculatedColumnFormula>{escape(calc)}</calculatedColumnFormula>'
                   if self.calculated else "")
                + '</tableColumn></tableColumns><tableStyleInfo name="TableStyleMedium2" '
                'showFirstColumn="0" showLastColumn="0" showRowStripes="1" '
                'showColumnStripes="0"/></table>')

    def sums_table(self) -> str:
        if not self.totals:
            return (_DECL + f'<table xmlns="{X}" id="2" name="Sums" displayName="Sums" '
                    'ref="J1:K5" totalsRowShown="0"><autoFilter ref="J1:K5"/>'
                    '<tableColumns count="2"><tableColumn id="1" name="Name"/>'
                    '<tableColumn id="2" name="Amount"/></tableColumns><tableStyleInfo '
                    'name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" '
                    'showRowStripes="1" showColumnStripes="0"/></table>')
        return (_DECL + f'<table xmlns="{X}" id="2" name="Sums" displayName="Sums" '
                'ref="J1:K6" totalsRowCount="1"><autoFilter ref="J1:K5"/>'
                '<tableColumns count="2"><tableColumn id="1" name="Name" '
                'totalsRowLabel="Total"/><tableColumn id="2" name="Amount" '
                'totalsRowFunction="sum"/></tableColumns><tableStyleInfo '
                'name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" '
                'showRowStripes="1" showColumnStripes="0"/></table>')

    def parts(self) -> dict[str, bytes]:
        base = StructureBook()
        workbook = (_DECL + f'<workbook xmlns="{X}" xmlns:r="{R}">'
                    '<workbookPr defaultThemeVersion="166925"/><bookViews><workbookView/>'
                    '</bookViews><sheets><sheet name="Fmt" sheetId="1" r:id="rId1"/></sheets>'
                    '<calcPr calcId="191029"/></workbook>')
        chart = base.chart_xml().replace("Stock!$B$1", "Fmt!$B$1").replace(
            "Stock!$A$2:$A$7", "Fmt!$A$2:$A$5").replace("Stock!$B$2:$B$7", "Fmt!$B$2:$B$5")
        drawing = base.drawing().replace("<xdr:col>3</xdr:col>", "<xdr:col>0</xdr:col>", 1) \
            .replace("<xdr:row>1</xdr:row>", "<xdr:row>11</xdr:row>", 1) \
            .replace("<xdr:col>9</xdr:col>", "<xdr:col>4</xdr:col>", 1) \
            .replace("<xdr:row>15</xdr:row>", "<xdr:row>20</xdr:row>", 1)
        overrides = [("/xl/workbook.xml", "spreadsheetml.sheet.main+xml"),
                     ("/xl/styles.xml", "spreadsheetml.styles+xml"),
                     ("/xl/worksheets/sheet1.xml", "spreadsheetml.worksheet+xml"),
                     *([("/xl/tables/table1.xml", "spreadsheetml.table+xml")]
                       if self.calc else []),
                     *([("/xl/tables/table2.xml", "spreadsheetml.table+xml")]
                       if self.sums else []),
                     ("/xl/drawings/drawing1.xml", "drawing+xml"),
                     ("/xl/charts/chart1.xml", "drawingml.chart+xml")]
        parts: dict[str, str] = {
            "[Content_Types].xml": (
                _DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
                'content-types"><Default Extension="rels" ContentType="application/'
                'vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                + "".join(f'<Override PartName="{name}" ContentType="{_SML}{kind}"/>'
                          for name, kind in overrides) + "</Types>"),
            "_rels/.rels": _rels([("rId1", "officeDocument", "xl/workbook.xml")]),
            "xl/workbook.xml": workbook,
            "xl/_rels/workbook.xml.rels": _rels([
                ("rId1", "worksheet", "worksheets/sheet1.xml"),
                ("rId3", "styles", "styles.xml")]),
            "xl/worksheets/sheet1.xml": self.sheet(),
            "xl/worksheets/_rels/sheet1.xml.rels": _rels([
                *([("rId1", "table", "../tables/table1.xml")] if self.calc else []),
                *([("rId2", "table", "../tables/table2.xml")] if self.sums else []),
                ("rId3", "drawing", "../drawings/drawing1.xml")]),
            **({"xl/tables/table1.xml": self.calc_table()} if self.calc else {}),
            **({"xl/tables/table2.xml": self.sums_table()} if self.sums else {}),
            "xl/drawings/drawing1.xml": drawing,
            "xl/drawings/_rels/drawing1.xml.rels": _rels(
                [("rId1", "chart", "../charts/chart1.xml")]),
            "xl/charts/chart1.xml": chart,
            "xl/styles.xml": base.styles(),
        }
        return {name: data.encode("utf-8") for name, data in parts.items()}

    def to_bytes(self) -> bytes:
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, data in self.parts().items():
                archive.writestr(zipfile.ZipInfo(name, date_time=(2026, 1, 2, 3, 4, 6)), data,
                                 compress_type=zipfile.ZIP_DEFLATED)
        return buffer.getvalue()


def _table_parts(ids: list[str]) -> str:
    if not ids:
        return ""
    return (f'<tableParts count="{len(ids)}">'
            + "".join(f'<tablePart r:id="{i}"/>' for i in ids) + "</tableParts>")
