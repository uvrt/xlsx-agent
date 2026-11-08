"""A made-up workbook whose every kind of referent names one sheet: the corner shop's
stock list (``tests/shop.py``) with its ``Orders`` sheet and ``Rates`` table, the ``Lists``
sheet visible, and more that refers to ``Lists``:

* defined names: ``Categories``, ``TaxRates``, ``Shelves`` (the shop's), ``Local`` (local
  to ``Lists``), ``Span`` (a 3-D run ``Stock:Lists``), ``Both`` (a union over two sheets);
* on ``Totals``: a 3-D ``SUM(Stock:Lists!B2)`` (``B6``), ``INDIRECT("Lists!B2")`` (a string,
  no reference; ``C6``), ``'Lists'!B3+Local`` (``D6``), whole columns and rows of ``Lists``
  (``E6``), a conditional format reading ``Lists`` (``B1``) and a hyperlink to it (``A6``);
* the shop's own: ``VLOOKUP`` into ``Lists`` on ``Stock``, an ``x14`` list validation on
  ``Stock`` and an ``x14`` bound on ``Orders``;
* ``chart=True``: a column chart on ``Totals`` whose title, series name, categories and
  values are ``Lists`` cells; ``pivot=True``: a pivot table on ``Totals`` whose cache's
  source is ``Lists!A1:B5``; ``table_formula=True``: ``Totals!F6`` holding
  ``SUM(Rates[Tax])``.

Every part is written here; Excel for Mac opens it without a repair (the oracle, round 44
lane S), once the shop's list validation through ``Rates[Tax]`` (which Excel repairs) is
left out.
"""

from __future__ import annotations

import io
import re
import zipfile

from shop import Shop

C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
REL = R + "/"
CT = "application/vnd.openxmlformats-officedocument."


def _patch(parts: dict, part: str, old: str, new: str) -> None:
    body = parts[part].decode()
    assert old in body, (part, old)
    parts[part] = body.replace(old, new, 1).encode()


def zipped(parts: dict) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(zipfile.ZipInfo(name, date_time=(2026, 1, 2, 3, 4, 6)), data,
                             compress_type=zipfile.ZIP_DEFLATED)
    return out.getvalue()


def unzipped(data: bytes) -> dict:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _rels(*items: tuple[str, str, str]) -> bytes:
    body = "".join(f'<Relationship Id="{i}" Type="{REL}{t}" Target="{target}"/>'
                   for i, t, target in items)
    return (DECL + f'<Relationships xmlns="{PR}">{body}</Relationships>').encode()


def chart_xml(sheet: str = "Lists") -> str:
    s = sheet
    cats = "".join(f'<c:pt idx="{i}"><c:v>{n}</c:v></c:pt>'
                   for i, n in enumerate(["Stationery", "Kitchen", "Hardware", "Garden"]))
    vals = "".join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>'
                   for i, v in enumerate([0.06, 0.21, 0.21, 0.21]))
    name = (f'<c:strRef><c:f>{s}!$B$1</c:f><c:strCache><c:ptCount val="1"/>'
            '<c:pt idx="0"><c:v>Tax</c:v></c:pt></c:strCache></c:strRef>')
    return (DECL + f'<c:chartSpace xmlns:c="{C}" xmlns:a="{A}" xmlns:r="{R}"><c:chart>'
            f'<c:title><c:tx>{name}</c:tx><c:overlay val="0"/></c:title>'
            '<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:barChart>'
            '<c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
            f'<c:ser><c:idx val="0"/><c:order val="0"/><c:tx>{name}</c:tx>'
            f'<c:cat><c:strRef><c:f>{s}!$A$2:$A$5</c:f><c:strCache><c:ptCount val="4"/>'
            f'{cats}</c:strCache></c:strRef></c:cat>'
            f'<c:val><c:numRef><c:f>{s}!$B$2:$B$5</c:f><c:numCache>'
            f'<c:formatCode>General</c:formatCode><c:ptCount val="4"/>{vals}</c:numCache>'
            '</c:numRef></c:val></c:ser><c:gapWidth val="150"/><c:axId val="111"/>'
            '<c:axId val="222"/></c:barChart><c:catAx><c:axId val="111"/><c:scaling>'
            '<c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>'
            '<c:crossAx val="222"/></c:catAx><c:valAx><c:axId val="222"/><c:scaling>'
            '<c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>'
            '<c:crossAx val="111"/></c:valAx></c:plotArea><c:plotVisOnly val="1"/>'
            '</c:chart></c:chartSpace>')


def drawing_xml() -> str:
    return (DECL + f'<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{A}"><xdr:twoCellAnchor>'
            '<xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row>'
            '<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>10</xdr:col>'
            '<xdr:colOff>0</xdr:colOff><xdr:row>16</xdr:row><xdr:rowOff>0</xdr:rowOff>'
            '</xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>'
            '<xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
            '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
            f'<a:graphic><a:graphicData uri="{C}"><c:chart xmlns:c="{C}" xmlns:r="{R}" '
            'r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/>'
            '</xdr:twoCellAnchor></xdr:wsDr>')


def pivot_xml(sheet: str = "Lists") -> tuple[str, str, str]:
    x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    items = "".join(f'<s v="{n}"/>' for n in ["Stationery", "Kitchen", "Hardware", "Garden"])
    cache = (DECL + f'<pivotCacheDefinition xmlns="{x}" xmlns:r="{R}" r:id="rId1" '
             'refreshOnLoad="1" recordCount="4"><cacheSource type="worksheet">'
             f'<worksheetSource ref="A1:B5" sheet="{sheet}"/></cacheSource>'
             '<cacheFields count="2"><cacheField name="Category" numFmtId="0">'
             f'<sharedItems count="4">{items}</sharedItems></cacheField>'
             '<cacheField name="Tax" numFmtId="0"><sharedItems containsSemiMixedTypes="0" '
             'containsString="0" containsNumber="1" minValue="0.06" maxValue="0.21"/>'
             '</cacheField></cacheFields></pivotCacheDefinition>')
    records = (DECL + f'<pivotCacheRecords xmlns="{x}" count="4">'
               '<r><x v="0"/><n v="0.06"/></r><r><x v="1"/><n v="0.21"/></r>'
               '<r><x v="2"/><n v="0.21"/></r><r><x v="3"/><n v="0.21"/></r>'
               '</pivotCacheRecords>')
    table = (DECL + f'<pivotTableDefinition xmlns="{x}" name="Summary" cacheId="5" '
             'applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" '
             'applyPatternFormats="0" applyAlignmentFormats="0" '
             'applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" '
             'minRefreshableVersion="3" createdVersion="6" indent="0" outline="1" '
             'outlineData="1"><location ref="A12:B17" firstHeaderRow="1" firstDataRow="1" '
             'firstDataCol="1"/><pivotFields count="2"><pivotField axis="axisRow" '
             'showAll="0"><items count="5"><item x="0"/><item x="1"/><item x="2"/>'
             '<item x="3"/><item t="default"/></items></pivotField><pivotField '
             'dataField="1" showAll="0"/></pivotFields><rowFields count="1"><field x="0"/>'
             '</rowFields><rowItems count="5"><i><x/></i><i><x v="1"/></i><i><x v="2"/></i>'
             '<i><x v="3"/></i><i t="grand"><x/></i></rowItems><colItems count="1"><i/>'
             '</colItems><dataFields count="1"><dataField name="Sum of Tax" fld="1" '
             'baseField="0" baseItem="0"/></dataFields></pivotTableDefinition>')
    return cache, records, table


def parts(*, chart: bool = False, pivot: bool = False, table_formula: bool = False) -> dict:
    found = Shop(orders=True, table=True).parts()
    orders = found["xl/worksheets/sheet4.xml"].decode()
    orders = re.sub(r'<dataValidation type="list" allowBlank="1" sqref="H2:H50">'
                    r'<formula1>Rates\[Tax\]</formula1></dataValidation>', '', orders)
    found["xl/worksheets/sheet4.xml"] = orders.replace('<dataValidations count="9">',
                                                       '<dataValidations count="8">').encode()
    _patch(found, "xl/workbook.xml",
           '<sheet name="Lists" sheetId="2" state="hidden" r:id="rId2"/>',
           '<sheet name="Lists" sheetId="2" r:id="rId2"/>')
    _patch(found, "xl/workbook.xml", '<definedName name="Markup">1.4</definedName>',
           '<definedName name="Markup">1.4</definedName>'
           '<definedName name="Local" localSheetId="1">Lists!$B$2</definedName>'
           '<definedName name="Span">Stock:Lists!$B$2</definedName>'
           '<definedName name="Both">Lists!$A$2,Totals!$B$1</definedName>')
    extra = '<c r="F6"><f>SUM(Rates[Tax])</f><v>0.69</v></c>' if table_formula else ""
    _patch(found, "xl/worksheets/sheet3.xml", '</row></sheetData>',
           '</row><row r="6"><c r="A6" t="inlineStr"><is><t>go</t></is></c>'
           '<c r="B6"><f>SUM(Stock:Lists!B2)</f><v>0</v></c>'
           '<c r="C6"><f>INDIRECT("Lists!B2")</f><v>0.06</v></c>'
           "<c r=\"D6\"><f>'Lists'!B3+Local</f><v>0.27</v></c>"
           '<c r="E6"><f>COUNTA(Lists!A:A)+SUM(Lists!2:2)</f><v>2</v></c>'
           f'{extra}</row></sheetData>')
    _patch(found, "xl/worksheets/sheet3.xml", '<pageMargins',
           '<conditionalFormatting sqref="B1"><cfRule type="expression" priority="1" '
           'dxfId="0"><formula>B1&gt;Lists!$B$2</formula></cfRule></conditionalFormatting>'
           '<hyperlinks><hyperlink ref="A6" location="Lists!A1" display="go"/></hyperlinks>'
           '<pageMargins')
    _patch(found, "xl/worksheets/sheet3.xml", '<dimension ref="A1:B3"/>',
           '<dimension ref="A1:F6"/>')
    found.pop("xl/calcChain.xml")
    _patch(found, "xl/_rels/workbook.xml.rels",
           f'<Relationship Id="rId6" Type="{REL}calcChain" Target="calcChain.xml"/>', '')
    _patch(found, "[Content_Types].xml",
           f'<Override PartName="/xl/calcChain.xml" ContentType="{CT}spreadsheetml.calcChain'
           '+xml"/>', '')
    sheet_rels = []
    types = []
    if chart:
        found["xl/charts/chart1.xml"] = chart_xml().encode()
        found["xl/drawings/drawing1.xml"] = drawing_xml().encode()
        found["xl/drawings/_rels/drawing1.xml.rels"] = _rels(
            ("rId1", "chart", "../charts/chart1.xml"))
        sheet_rels.append(("rId1", "drawing", "../drawings/drawing1.xml"))
        _patch(found, "xl/worksheets/sheet3.xml", '</worksheet>',
               '<drawing r:id="rId1"/></worksheet>')
        types += [("/xl/charts/chart1.xml", "drawingml.chart+xml"),
                  ("/xl/drawings/drawing1.xml", "drawing+xml")]
    if pivot:
        cache, records, table = pivot_xml()
        found["xl/pivotCache/pivotCacheDefinition1.xml"] = cache.encode()
        found["xl/pivotCache/pivotCacheRecords1.xml"] = records.encode()
        found["xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"] = _rels(
            ("rId1", "pivotCacheRecords", "pivotCacheRecords1.xml"))
        found["xl/pivotTables/pivotTable1.xml"] = table.encode()
        found["xl/pivotTables/_rels/pivotTable1.xml.rels"] = _rels(
            ("rId1", "pivotCacheDefinition", "../pivotCache/pivotCacheDefinition1.xml"))
        sheet_rels.append(("rId2", "pivotTable", "../pivotTables/pivotTable1.xml"))
        _patch(found, "xl/_rels/workbook.xml.rels", '</Relationships>',
               f'<Relationship Id="rId20" Type="{REL}pivotCacheDefinition" '
               'Target="pivotCache/pivotCacheDefinition1.xml"/></Relationships>')
        _patch(found, "xl/workbook.xml", '<calcPr calcId="191029"/>',
               '<calcPr calcId="191029"/><pivotCaches><pivotCache cacheId="5" r:id="rId20"/>'
               '</pivotCaches>')
        types += [("/xl/pivotCache/pivotCacheDefinition1.xml",
                   "spreadsheetml.pivotCacheDefinition+xml"),
                  ("/xl/pivotCache/pivotCacheRecords1.xml", "spreadsheetml.pivotCacheRecords+xml"),
                  ("/xl/pivotTables/pivotTable1.xml", "spreadsheetml.pivotTable+xml")]
    if sheet_rels:
        found["xl/worksheets/_rels/sheet3.xml.rels"] = _rels(*sheet_rels)
    if types:
        _patch(found, "[Content_Types].xml", '</Types>',
               "".join(f'<Override PartName="{name}" ContentType="{CT}{kind}"/>'
                       for name, kind in types) + '</Types>')
    return found


def book(**options) -> bytes:
    """The workbook (see the module) as a file."""
    return zipped(parts(**options))
