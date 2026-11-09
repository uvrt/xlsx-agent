"""A made-up workbook for reordering a table's columns (round 46): a shop sheet whose table
has everything that follows a table column when Excel moves one, built in code.

``Shop`` holds a title (``B1``, merged over ``B1:C1``) and the table ``Lines`` (``B3:G10``:
Code, Qty, Price, Total, Note, Bin; header row 3, data rows 4-9, a totals row 10) whose
``Total`` is a calculated column, whose totals row has a label, two functions and a custom
formula, whose filter keeps some bins (criteria on ``Bin``) and whose sort state sorts by
``Qty``.  ``Price`` has a conditional format, ``Bin`` a data validation, a ``Note`` cell a
hyperlink; every table column its own width.  Next to the table (``I``) and on ``Sums``,
formulas read the table's cells by A1 references of every shape (a cell, a column's data,
several columns, a whole column, the whole table, a column with its header and totals,
beyond the table) and by structured references (a column, a range of columns, the totals
row, ``#Data``); defined names read a cell and a column's data, and a chart's series reads
``Code`` and ``Price``.  A cell under the table (``B12``) and one beside its header
(``H3``) sit in the sheet's other rows.

Every part is written here (ECMA-376 Part 1, §18), nothing from an application; what Excel
makes of moving a table column is the oracle's answer (``tests/test_table_columns.py``).
"""

from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass
from xml.sax.saxutils import escape

from structure_book import _DECL, _SML, R, X, StructureBook, _cell, _formula, _rels

#: The table's columns and its rows: code, quantity, price, note, bin.
COLUMNS = ("Code", "Qty", "Price", "Total", "Note", "Bin")
LINES = [
    ("P-1", 12, 1.5, "fragile", "A"),
    ("P-2", 4, 7.25, "", "B"),
    ("P-3", 30, 0.8, "boxed", "A"),
    ("P-4", 9, 12.0, "", "B"),
    ("P-5", 16, 3.2, "spare", "A"),
    ("P-6", 2, 25.0, "", "B"),
]
#: The calculated column's formula (the cells' and the table's).
TOTAL = "Lines[[#This Row],[Qty]]*Lines[[#This Row],[Price]]"
#: Formulas next to the table, column I: cell -> formula.
SHOP_FORMULAS = {
    "I4": "D4*2",
    "I5": "SUM(C4:C9)",
    "I6": "SUM(C4:E9)",
    "I7": "SUM(D:D)",
    "I8": "SUM(Lines[[Qty]:[Price]])",
    "I9": "B12",
}
#: Formulas on Sums: cell -> formula.
SUMS_FORMULAS = {
    "A1": "Shop!D5",
    "A2": "SUM(Shop!D4:D9)",
    "A3": "SUM(Shop!C4:E9)",
    "A4": "SUM(Shop!D:D)",
    "A5": "SUM(Lines[Price])",
    "A6": "SUM(Lines[[Qty]:[Price]])",
    "A7": "Lines[[#Totals],[Total]]",
    "A8": "COUNTA(Shop!B3:G10)",
    "A9": "SUM(Shop!D3:D10)",
    "A10": "SUM(Shop!D2:D12)",
    "A11": "INDEX(Lines[#Data],1,2)",
    "A12": "Shop!$F$4&Shop!G$5",
    "A13": "SUM(Shop!C5:D6)",
    "A14": "SUM(Shop!B4:G4)",
    "A15": "PriceBody",
}
#: More formulas on Sums (``edges=True``): references that reach past the table's rows.
EDGE_FORMULAS = {
    "A16": "SUM(Shop!D5:D12)",
    "A17": "SUM(Shop!D1:D4)",
    "A18": "SUM(Shop!C2:E9)",
    "A19": "Shop!D11",
    "A20": "COUNTA(Shop!C3:D3)",
    "A21": "SUM(Shop!C10:D11)",
}
#: Ranges reaching past the table's rows in each of its columns (``straddle=True``).
STRADDLE_FORMULAS = {
    **{f"C{n}": f"SUM(Shop!{c}5:{c}12)" for n, c in enumerate("BCDEFG", 1)},
    **{f"D{n}": f"COUNTA(Shop!{c}1:{c}4)" for n, c in enumerate("BCDEFG", 1)},
}
#: Widths of the table's columns B..G and of I.
WIDTHS = {2: 14.7109375, 3: 8.7109375, 4: 10.7109375, 5: 12.7109375, 6: 20.7109375,
          7: 6.7109375, 9: 11.7109375}


@dataclass
class TableColumnsBook:
    #: The totals row (``totals=False``: the table is ``B3:G9``).
    totals: bool = True
    #: The filter's criteria and the sort state.
    criteria: bool = True
    #: The chart reading Code and Price.
    chart: bool = True
    #: References and rules reaching past the table's rows (:data:`EDGE_FORMULAS`; a
    #: conditional format over ``F2:F11`` with a relative formula, one over ``C1:C12``
    #: without, a validation over ``E2:E9 E11:E12``).
    edges: bool = False
    #: Comments on ``D5`` (Price) and ``F12`` (under the table), with their VML drawing.
    comments: bool = False
    #: On Sums, for every table column, a range from inside the table's rows down past
    #: them and one from above them into them (:data:`STRADDLE_FORMULAS`).
    straddle: bool = False

    @property
    def last(self) -> int:
        return 10 if self.totals else 9

    def shop(self) -> str:
        rows: dict[int, list[str]] = {n: [] for n in range(1, 13)}
        rows[1].append(_cell("B1", "Shop lines", 1))
        for column, name in zip("BCDEFG", COLUMNS):
            rows[3].append(_cell(f"{column}3", name))
        rows[3].append(_cell("H3", "side"))
        for i, (code, qty, price, note, bin_) in enumerate(LINES):
            n = i + 4
            rows[n] += [_cell(f"B{n}", code), _cell(f"C{n}", qty), _cell(f"D{n}", price, 1),
                        _formula(f"E{n}", TOTAL, qty * price)]
            if note:
                rows[n].append(_cell(f"F{n}", note))
            rows[n].append(_cell(f"G{n}", bin_))
        if self.totals:
            qty = sum(line[1] for line in LINES)
            rows[10] += [_cell("B10", "Total"),
                         _formula("C10", "SUBTOTAL(109,Lines[Qty])", qty),
                         _formula("D10", "MAX(Lines[Price])", 25),
                         _formula("E10", "SUBTOTAL(109,Lines[Total])", 0)]
        for ref, text in SHOP_FORMULAS.items():
            rows[int(ref[1:])].append(_formula(ref, text))
        rows[12] += [_cell("B12", "under"), _cell("F12", 5)]
        body = "".join(f'<row r="{n}">{"".join(cells)}</row>' for n, cells in rows.items()
                       if cells)
        cols = "".join(f'<col min="{c}" max="{c}" width="{w}" customWidth="1"/>'
                       for c, w in WIDTHS.items())
        return (
            _DECL + f'<worksheet xmlns="{X}" xmlns:r="{R}">'
            '<dimension ref="B1:I12"/>'
            '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>'
            f'<sheetFormatPr defaultRowHeight="15"/><cols>{cols}</cols>'
            f'<sheetData>{body}</sheetData>'
            '<mergeCells count="1"><mergeCell ref="B1:C1"/></mergeCells>'
            '<conditionalFormatting sqref="D4:D9"><cfRule type="cellIs" dxfId="0" '
            'priority="1" operator="greaterThan"><formula>5</formula></cfRule>'
            '</conditionalFormatting>'
            + ('<conditionalFormatting sqref="F2:F11"><cfRule type="expression" dxfId="0" '
               'priority="2"><formula>LEN(F2)&gt;C2</formula></cfRule></conditionalFormatting>'
               '<conditionalFormatting sqref="C1:C12"><cfRule type="duplicateValues" '
               'dxfId="1" priority="3"/></conditionalFormatting>' if self.edges else "")
            + '<dataValidations count="{0}"><dataValidation type="list" allowBlank="1" '
            'showInputMessage="1" showErrorMessage="1" sqref="G4:G9"><formula1>"A,B"'
            '</formula1></dataValidation>'.format(2 if self.edges else 1)
            + ('<dataValidation type="whole" operator="greaterThan" allowBlank="1" '
               'showErrorMessage="1" sqref="E2:E9 E11:E12"><formula1>0</formula1></dataValidation>'
               if self.edges else "")
            + '</dataValidations>'
            '<hyperlinks><hyperlink ref="F4" location="Sums!A1" display="fragile"/>'
            '</hyperlinks>'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" '
            'footer="0.3"/>'
            + ('<drawing r:id="rId2"/>' if self.chart else "")
            + ('<legacyDrawing r:id="rId3"/>' if self.comments else "")
            + '<tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>')

    def sums(self) -> str:
        formulas = {**SUMS_FORMULAS, **(EDGE_FORMULAS if self.edges else {}),
                    **(STRADDLE_FORMULAS if self.straddle else {})}
        lines: dict[int, list[str]] = {}
        for ref, text in formulas.items():
            lines.setdefault(int(ref[1:]), []).append(_formula(ref, text))
        rows = "".join(f'<row r="{n}">{"".join(cells)}</row>'
                       for n, cells in sorted(lines.items()))
        last = "D" if self.straddle else "A"
        return (_DECL + f'<worksheet xmlns="{X}" xmlns:r="{R}">'
                f'<dimension ref="A1:{last}{max(lines)}"/>'
                '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
                f'<sheetFormatPr defaultRowHeight="15"/><sheetData>{rows}</sheetData>'
                '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" '
                'footer="0.3"/></worksheet>')

    def table(self) -> str:
        columns = []
        for i, name in enumerate(COLUMNS):
            attrs = ""
            inner = ""
            if self.totals:
                attrs = {"Code": ' totalsRowLabel="Total"', "Qty": ' totalsRowFunction="sum"',
                         "Price": ' totalsRowFunction="custom"',
                         "Total": ' totalsRowFunction="sum"'}.get(name, "")
            if name == "Total":
                inner = f'<calculatedColumnFormula>{escape(TOTAL)}</calculatedColumnFormula>'
            if self.totals and name == "Price":
                inner += '<totalsRowFormula>MAX(Lines[Price])</totalsRowFormula>'
            columns.append(f'<tableColumn id="{i + 1}" name="{name}"{attrs}>{inner}'
                           '</tableColumn>' if inner else
                           f'<tableColumn id="{i + 1}" name="{name}"{attrs}/>')
        last = self.last
        totals = ' totalsRowCount="1"' if self.totals else ' totalsRowShown="0"'
        criteria = sort = ""
        if self.criteria:
            criteria = ('<filterColumn colId="5"><filters><filter val="A"/><filter val="B"/>'
                        '</filters></filterColumn>')
            sort = ('<sortState ref="B4:G9"><sortCondition descending="1" ref="C4:C9"/>'
                    '</sortState>')
        return (_DECL + f'<table xmlns="{X}" id="1" name="Lines" displayName="Lines" '
                f'ref="B3:G{last}"{totals}><autoFilter ref="B3:G9">{criteria}</autoFilter>{sort}'
                f'<tableColumns count="6">{"".join(columns)}</tableColumns>'
                '<tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" '
                'showLastColumn="0" showRowStripes="1" showColumnStripes="1"/></table>')

    def chart_xml(self) -> str:
        return (StructureBook().chart_xml()
                .replace("Stock!$B$1", "Shop!$D$3")
                .replace("Stock!$A$2:$A$7", "Shop!$B$4:$B$9")
                .replace("Stock!$B$2:$B$7", "Shop!$D$4:$D$9"))

    def workbook(self) -> str:
        names = ('<definedName name="PriceBody">Shop!$D$4:$D$9</definedName>'
                 '<definedName name="QtyCell">Shop!$C$5</definedName>')
        return (_DECL + f'<workbook xmlns="{X}" xmlns:r="{R}">'
                '<workbookPr defaultThemeVersion="166925"/><bookViews><workbookView/>'
                '</bookViews><sheets><sheet name="Shop" sheetId="1" r:id="rId1"/>'
                '<sheet name="Sums" sheetId="2" r:id="rId2"/></sheets>'
                f'<definedNames>{names}</definedNames><calcPr calcId="191029"/></workbook>')

    def comments_xml(self) -> str:
        items = [("D5", "check the price"), ("F12", "spare stock")]
        body = "".join(f'<comment ref="{ref}" authorId="0"><text><r><t>{escape(text)}</t></r>'
                       '</text></comment>' for ref, text in items)
        return (_DECL + f'<comments xmlns="{X}"><authors><author>Clerk</author></authors>'
                f'<commentList>{body}</commentList></comments>')

    def vml(self) -> str:
        # StructureBook's shapes are on (row 2, column 1) and (row 4, column 5), zero-based:
        # B3 and F5 there; here D5 and F12.
        return StructureBook().vml().replace("<x:Row>2</x:Row>", "<x:Row>4</x:Row>") \
            .replace("<x:Column>1</x:Column>", "<x:Column>3</x:Column>") \
            .replace("2, 15, 1, 10, 4, 15, 5, 4", "4, 15, 3, 10, 6, 15, 7, 4") \
            .replace("<x:Row>4</x:Row>\n", "") \
            .replace("6, 15, 3, 10, 8, 15, 7, 4", "6, 15, 10, 10, 8, 15, 14, 4") \
            .replace("<x:Row>4</x:Row><x:Column>5</x:Column>",
                     "<x:Row>11</x:Row><x:Column>5</x:Column>")

    def parts(self) -> dict[str, bytes]:
        base = StructureBook()
        overrides = [("/xl/workbook.xml", "spreadsheetml.sheet.main+xml"),
                     ("/xl/styles.xml", "spreadsheetml.styles+xml"),
                     ("/xl/worksheets/sheet1.xml", "spreadsheetml.worksheet+xml"),
                     ("/xl/worksheets/sheet2.xml", "spreadsheetml.worksheet+xml"),
                     ("/xl/tables/table1.xml", "spreadsheetml.table+xml")]
        shop_rels = [("rId1", "table", "../tables/table1.xml")]
        parts: dict[str, str] = {
            "xl/workbook.xml": self.workbook(),
            "xl/_rels/workbook.xml.rels": _rels([
                ("rId1", "worksheet", "worksheets/sheet1.xml"),
                ("rId2", "worksheet", "worksheets/sheet2.xml"),
                ("rId3", "styles", "styles.xml")]),
            "xl/worksheets/sheet1.xml": self.shop(),
            "xl/worksheets/sheet2.xml": self.sums(),
            "xl/tables/table1.xml": self.table(),
            "xl/styles.xml": base.styles(),
            "_rels/.rels": _rels([("rId1", "officeDocument", "xl/workbook.xml")]),
        }
        if self.chart:
            shop_rels.append(("rId2", "drawing", "../drawings/drawing1.xml"))
            drawing = base.drawing().replace("<xdr:col>3</xdr:col>", "<xdr:col>10</xdr:col>", 1) \
                .replace("<xdr:col>9</xdr:col>", "<xdr:col>16</xdr:col>", 1)
            parts["xl/drawings/drawing1.xml"] = drawing
            parts["xl/drawings/_rels/drawing1.xml.rels"] = _rels(
                [("rId1", "chart", "../charts/chart1.xml")])
            parts["xl/charts/chart1.xml"] = self.chart_xml()
            overrides += [("/xl/drawings/drawing1.xml", "drawing+xml"),
                          ("/xl/charts/chart1.xml", "drawingml.chart+xml")]
        if self.comments:
            shop_rels += [("rId3", "vmlDrawing", "../drawings/vmlDrawing1.vml"),
                          ("rId4", "comments", "../comments1.xml")]
            parts["xl/comments1.xml"] = self.comments_xml()
            parts["xl/drawings/vmlDrawing1.vml"] = self.vml()
            overrides.append(("/xl/comments1.xml", "spreadsheetml.comments+xml"))
        parts["xl/worksheets/_rels/sheet1.xml.rels"] = _rels(shop_rels)
        parts["[Content_Types].xml"] = (
            _DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
            'content-types"><Default Extension="rels" ContentType="application/'
            'vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Default Extension="vml" ContentType="application/vnd.openxmlformats-'
            'officedocument.vmlDrawing"/>'
            + "".join(f'<Override PartName="{name}" ContentType="{_SML}{kind}"/>'
                      for name, kind in overrides) + "</Types>")
        return {name: data.encode("utf-8") for name, data in parts.items()}

    def to_bytes(self) -> bytes:
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, data in self.parts().items():
                archive.writestr(zipfile.ZipInfo(name, date_time=(2026, 1, 2, 3, 4, 6)), data,
                                 compress_type=zipfile.ZIP_DEFLATED)
        return buffer.getvalue()


def table_facts(data: bytes) -> dict:
    """What a workbook says about its tables beyond :func:`structure_facts.facts`: each
    table's range and style, each table column's id, name, totals function, label and
    formula, and its data format (the dxf's number format code); the filter's criteria by
    column (their XML, namespaces aside); the sort state's range and keys."""
    import re
    from xml.etree import ElementTree

    z = zipfile.ZipFile(io.BytesIO(data))
    styles = ElementTree.fromstring(z.read("xl/styles.xml"))
    dxfs = styles.find(f"{{{X}}}dxfs")

    def dxf(index: str | None) -> str | None:
        if index is None or dxfs is None:
            return None
        found = dxfs[int(index)].find(f"{{{X}}}numFmt")
        return None if found is None else found.get("formatCode")

    out = {}
    for name in sorted(n for n in z.namelist() if n.startswith("xl/tables/table")):
        root = ElementTree.fromstring(z.read(name))
        columns = []
        for column in root.iter(f"{{{X}}}tableColumn"):
            item = {k: column.get(k) for k in ("id", "name", "totalsRowFunction",
                                                "totalsRowLabel") if column.get(k)}
            item["data_format"] = dxf(column.get("dataDxfId"))
            item["totals_formula"] = column.findtext(f"{{{X}}}totalsRowFormula")
            columns.append(item)
        filt = root.find(f"{{{X}}}autoFilter")
        criteria = {}
        for column in (filt.findall(f"{{{X}}}filterColumn") if filt is not None else []):
            text = ElementTree.tostring(column[0], encoding="unicode") if len(column) else ""
            text = re.sub(r"ns\d+:|\s*xmlns:ns\d+=\"[^\"]*\"", "", text)
            criteria[column.get("colId")] = re.sub(r"\s+(/?)>", r"\1>", text)
        sort = root.find(f"{{{X}}}sortState")
        out[root.get("name")] = {
            "ref": root.get("ref"), "style": root.find(f"{{{X}}}tableStyleInfo").get("name"),
            "columns": columns, "filter": filt.get("ref") if filt is not None else None,
            "criteria": criteria,
            "sort": None if sort is None else [sort.get("ref")] + [
                c.get("ref") for c in sort.findall(f"{{{X}}}sortCondition")]}
    return out


__all__ = ["COLUMNS", "LINES", "TableColumnsBook", "table_facts"]
