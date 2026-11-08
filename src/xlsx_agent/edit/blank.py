"""A new workbook from nothing, written the way Excel writes a new blank workbook (R51).

Every part is authored here; nothing is copied from Excel's template or a user's files.
**Measured** (Excel for Mac 16.106, Dutch interface, 2026-11-08, round 44 lane T): Excel
made a new workbook through its AppleScript (``make new workbook``) and saved it as a
``.xlsx``; its parts were read back.  It writes nine parts, and so does :func:`build`:

* ``[Content_Types].xml``, ``_rels/.rels`` (``app``, ``core``, ``officeDocument``, in
  Excel's order and ids);
* ``xl/workbook.xml``: ``fileVersion`` (``appName="xl"``, ``lastEdited``/``lowestEdited``
  7, ``rupBuild`` 10208), ``workbookPr defaultThemeVersion="202300"``, ``bookViews`` with
  one ``workbookView``, the sheets, ``calcPr calcId="181029"`` and the two extensions
  Excel writes (``x15:workbookPr chartTrackingRefBase``, the calculation features);
* ``xl/_rels/workbook.xml.rels`` (styles ``rId3``, theme ``rId2``, sheets from ``rId1``:
  a sheet beyond the first takes the next free id);
* ``xl/worksheets/sheetN.xml``: ``dimension ref="A1"``, a ``sheetView`` (the first sheet
  ``tabSelected``), ``sheetFormatPr baseColWidth="10" defaultRowHeight="16"`` with its
  ``x14ac:dyDescent``, an empty ``sheetData``, the page margins (0.7, 0.7, 0.75, 0.75,
  0.3, 0.3);
* ``xl/styles.xml``: the one font (the theme's minor font, Aptos Narrow, 12 pt, colour
  ``theme="1"``, family 2, scheme minor), the two fills (``none``, ``gray125``), one empty
  border, one cell style format and one cell format, the Normal cell style
  (``builtinId="0"``), no ``dxfs`` and ``tableStyles`` with ``TableStyleMedium2`` and
  ``PivotStyleLight16`` as the defaults, and the slicer and timeline style extensions;
* ``xl/theme/theme1.xml``: the Office theme -- PowerPoint's (:mod:`ooxml_common.drawingml.
  theme`, its colours, faces, gradients and line widths) but for the body face, **Aptos
  Narrow**, and the Thai script's faces, **Tahoma** for both (:data:`SCRIPT_FONTS`), with
  the line defaults (``a:objectDefaults/a:lnDef``) and the theme family extension;
  compared element by element with Excel's;
* ``docProps/core.xml`` (the creator and last modifier: the ``author`` given, or none;
  created and modified now, UTC) and ``docProps/app.xml`` (the sheet names, in Excel's
  element order).

**Left out on purpose**, as being Excel's session and machine rather than the workbook's:
the ``x15ac:absPath`` (the folder the file was saved in), ``xr:revisionPtr`` (a co-authoring
document id), the window's position and size on ``workbookView``, the ``xr:uid`` of the
sheet and the view; ``Application`` names this library, as docx-agent and pptx-agent do.
**Localised names** are English Excel's: the sheet ``Sheet1`` (Dutch Excel writes
``Blad1``), the cell style ``Normal`` (``Standaard``), the theme ``Office Theme``
(``Kantoorthema``) and ``Worksheets`` in ``app.xml`` (``Werkbladen``).

**What depends on the platform.**  Excel for Mac's Normal font is 12 pt, its row 16 points
and its base column width 10; Excel for Windows makes a new workbook in 11 pt (not
measured here: there is no Windows Excel on the test machine).  :data:`FONT_SIZE`,
:data:`ROW_HEIGHT` and :data:`BASE_COLUMN_WIDTH` hold the measured Mac values in one place.

No shared-strings part: Excel writes none until a cell holds text, and a text written into
the new workbook adds one (:func:`xlsx_agent.edit.tables.ensure_shared_strings`).
"""

from __future__ import annotations

import io
import zipfile
from datetime import datetime, timezone
from typing import Sequence
from xml.sax.saxutils import escape, quoteattr

from ooxml_common.drawingml import theme

_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_R = NS_R + "/"
REL_OFFICE_DOCUMENT = _R + "officeDocument"
REL_WORKSHEET = _R + "worksheet"
REL_STYLES = _R + "styles"
REL_THEME = _R + "theme"
REL_SHARED_STRINGS = _R + "sharedStrings"
REL_CORE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
REL_APP = _R + "extended-properties"

_SML = "application/vnd.openxmlformats-officedocument.spreadsheetml."
CT_WORKBOOK = _SML + "sheet.main+xml"
CT_WORKSHEET = _SML + "worksheet+xml"
CT_STYLES = _SML + "styles+xml"
CT_SHARED_STRINGS = _SML + "sharedStrings+xml"
CT_THEME = "application/vnd.openxmlformats-officedocument.theme+xml"
CT_CORE = "application/vnd.openxmlformats-package.core-properties+xml"
CT_APP = "application/vnd.openxmlformats-officedocument.extended-properties+xml"

#: What ``app.xml`` names as the application (the siblings name themselves too).
APPLICATION = "xlsx-agent"

#: The Normal font's size in points, the default row height in points and the base column
#: width in characters: Excel for Mac's, measured (see the module).
FONT_SIZE = 12
ROW_HEIGHT = 16
BASE_COLUMN_WIDTH = 10
#: ``x14ac:dyDescent`` of the sheet format, as measured.
DY_DESCENT = "0.2"

#: The theme's body (minor) face in Excel; its headings face is PowerPoint's.
MINOR_FONT = "Aptos Narrow"
MAJOR_FONT = theme.OFFICE_MAJOR_FONT
#: Excel's per-script faces: PowerPoint's (:data:`ooxml_common.drawingml.theme.
#: POWERPOINT_SCRIPT_FONTS`) but for Thai, which Excel writes as Tahoma for both.
SCRIPT_FONTS: tuple[tuple[str, str, str], ...] = tuple(
    (script, "Tahoma" if script == "Thai" else major, "Tahoma" if script == "Thai" else minor)
    for script, major, minor in theme.POWERPOINT_SCRIPT_FONTS)
THEME_NAME = "Office Theme"
THEME_FAMILY = ("Office Theme", "{2E142A2C-CD16-42D6-873A-C26D2A0506FA}",
                "{1BDDFF52-6CD6-40A5-AB3C-68EB2F1E4D0A}")

#: The sheet a new workbook has, as English Excel names it.
DEFAULT_SHEET = "Sheet1"
#: The table style a new workbook's tables take (``tableStyles defaultTableStyle``).
DEFAULT_TABLE_STYLE = "TableStyleMedium2"
DEFAULT_PIVOT_STYLE = "PivotStyleLight16"

_X14AC = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
_XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
_X15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def workbook_xml(sheets: Sequence[str]) -> str:
    entries = "".join(f'<sheet name={quoteattr(name)} sheetId="{n}" r:id="rId{_sheet_rid(n)}"/>'
                      for n, name in enumerate(sheets, 1))
    return (_DECL + f'<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_R}" xmlns:mc="{_MC}" '
            f'mc:Ignorable="x15" xmlns:x15="{_X15}">'
            '<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="10208"/>'
            '<workbookPr defaultThemeVersion="202300"/>'
            '<bookViews><workbookView/></bookViews>'
            f"<sheets>{entries}</sheets>"
            '<calcPr calcId="181029"/>'
            '<extLst><ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}" '
            f'xmlns:x15="{_X15}"><x15:workbookPr chartTrackingRefBase="1"/></ext>'
            '<ext uri="{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}" '
            'xmlns:xcalcf="http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures">'
            "<xcalcf:calcFeatures>"
            + "".join(f'<xcalcf:feature name="microsoft.com:{name}"/>'
                      for name in ("RD", "Single", "FV", "CNMTM", "LET_WF", "LAMBDA_WF",
                                   "ARRAYTEXT_WF"))
            + "</xcalcf:calcFeatures></ext></extLst></workbook>")


def _sheet_rid(number: int) -> int:
    """The workbook relationship id of sheet ``number``: Excel's ``rId1`` for the first,
    then past the styles' and theme's ``rId3``."""
    return 1 if number == 1 else number + 2


def worksheet_xml(selected: bool) -> str:
    view = '<sheetView tabSelected="1" workbookViewId="0"/>' if selected \
        else '<sheetView workbookViewId="0"/>'
    return (_DECL + f'<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_R}" xmlns:mc="{_MC}" '
            f'mc:Ignorable="x14ac" xmlns:x14ac="{_X14AC}">'
            f'<dimension ref="A1"/><sheetViews>{view}</sheetViews>'
            f'<sheetFormatPr baseColWidth="{BASE_COLUMN_WIDTH}" defaultRowHeight="{ROW_HEIGHT}" '
            f'x14ac:dyDescent="{DY_DESCENT}"/><sheetData/>'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" '
            'footer="0.3"/></worksheet>')


def styles_xml() -> str:
    return (_DECL + f'<styleSheet xmlns="{NS_MAIN}" xmlns:mc="{_MC}" mc:Ignorable="x14ac" '
            f'xmlns:x14ac="{_X14AC}">'
            f'<fonts count="1" x14ac:knownFonts="1"><font><sz val="{FONT_SIZE}"/>'
            f'<color theme="1"/><name val="{MINOR_FONT}"/><family val="2"/>'
            '<scheme val="minor"/></font></fonts>'
            '<fills count="2"><fill><patternFill patternType="none"/></fill>'
            '<fill><patternFill patternType="gray125"/></fill></fills>'
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border>'
            '</borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
            '</cellStyleXfs>'
            '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
            '</cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/>'
            '</cellStyles><dxfs count="0"/>'
            f'<tableStyles count="0" defaultTableStyle="{DEFAULT_TABLE_STYLE}" '
            f'defaultPivotStyle="{DEFAULT_PIVOT_STYLE}"/>'
            '<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" '
            'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
            '<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext>'
            '<ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" '
            f'xmlns:x15="{_X15}"><x15:timelineStyles '
            'defaultTimelineStyle="TimeSlicerStyleLight1"/></ext></extLst></styleSheet>')


def _gradient(stops) -> str:
    out = ['<a:gradFill rotWithShape="1"><a:gsLst>']
    for position, modifiers in stops:
        mods = "".join(f'<a:{name} val="{value}"/>' for name, value in modifiers)
        out.append(f'<a:gs pos="{position}"><a:schemeClr val="phClr">{mods}</a:schemeClr>'
                   "</a:gs>")
    out.append('</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>')
    return "".join(out)


def theme_xml() -> str:
    """Excel's Office theme (see the module), element by element as Excel writes it."""
    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    out = [_DECL, f'<a:theme xmlns:a="{a}" name="{THEME_NAME}"><a:themeElements>'
           '<a:clrScheme name="Office">']
    for name, value in theme.OFFICE_THEME_COLORS:
        if value.startswith("sys:"):
            _, system, last = value.split(":")
            out.append(f'<a:{name}><a:sysClr val="{system}" lastClr="{last}"/></a:{name}>')
        else:
            out.append(f'<a:{name}><a:srgbClr val="{value}"/></a:{name}>')
    out.append('</a:clrScheme><a:fontScheme name="Office">')
    for tag, face, column in (("majorFont", MAJOR_FONT, 1), ("minorFont", MINOR_FONT, 2)):
        out.append(f'<a:{tag}><a:latin typeface="{face}" panose="{theme.APTOS_PANOSE}"/>'
                   '<a:ea typeface=""/><a:cs typeface=""/>')
        out += [f'<a:font script="{entry[0]}" typeface="{entry[column]}"/>'
                for entry in SCRIPT_FONTS]
        out.append(f"</a:{tag}>")
    out.append('</a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst>'
               '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>')
    out += [_gradient(stops) for stops in theme.OFFICE_FILL_GRADIENTS]
    out.append("</a:fillStyleLst><a:lnStyleLst>")
    for width in theme.POWERPOINT_LINE_WIDTHS:
        out.append(f'<a:ln w="{width}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>'
                   '<a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/>'
                   '<a:miter lim="800000"/></a:ln>')
    out.append('</a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>'
               '<a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst>'
               '<a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" '
               'rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr>'
               "</a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>"
               '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
               '<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/>'
               '<a:satMod val="170000"/></a:schemeClr></a:solidFill>')
    out.append(_gradient(theme.OFFICE_BACKGROUND_GRADIENT))
    family, family_id, vid = THEME_FAMILY
    out.append("</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults>"
               "<a:lnDef><a:spPr/><a:bodyPr/><a:lstStyle/><a:style>"
               '<a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>'
               '<a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>'
               '<a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef>'
               '<a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef>'
               "</a:style></a:lnDef></a:objectDefaults>"
               '<a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}">'
               '<thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" '
               f'name="{family}" id="{family_id}" vid="{vid}"/></a:ext></a:extLst></a:theme>')
    return "".join(out)


def core_xml(author: str | None, title: str | None, created: datetime) -> str:
    stamp = created.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    fields = []
    if title is not None:
        fields.append(f"<dc:title>{escape(title)}</dc:title>")
    if author is not None:
        fields.append(f"<dc:creator>{escape(author)}</dc:creator>"
                      f"<cp:lastModifiedBy>{escape(author)}</cp:lastModifiedBy>")
    fields.append(f'<dcterms:created xsi:type="dcterms:W3CDTF">{stamp}</dcterms:created>'
                  f'<dcterms:modified xsi:type="dcterms:W3CDTF">{stamp}</dcterms:modified>')
    return (_DECL + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/'
            'metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" '
            'xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            + "".join(fields) + "</cp:coreProperties>")


def app_xml(sheets: Sequence[str]) -> str:
    titles = "".join(f"<vt:lpstr>{escape(name)}</vt:lpstr>" for name in sheets)
    return (_DECL + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/'
            'extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/'
            f'2006/docPropsVTypes"><Application>{APPLICATION}</Application>'
            "<DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>"
            '<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant>'
            "<vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant>"
            f"<vt:i4>{len(sheets)}</vt:i4></vt:variant></vt:vector></HeadingPairs>"
            f'<TitlesOfParts><vt:vector size="{len(sheets)}" baseType="lpstr">{titles}'
            "</vt:vector></TitlesOfParts><Company></Company><LinksUpToDate>false</LinksUpToDate>"
            "<SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged>"
            "<AppVersion>16.0300</AppVersion></Properties>")


#: What a sheet's name may not hold: the characters Excel's sheet-name rule refuses.
_NOT_IN_SHEET_NAMES = "[]:*?/\\"
#: The longest sheet name Excel takes.
MAX_SHEET_NAME = 31


def sheet_name_problems(names: Sequence[str]) -> list[str]:
    """Why each of ``names`` cannot name a new workbook's sheet: not text, empty, longer
    than :data:`MAX_SHEET_NAME`, holding one of ``[]:*?/\\``, starting or ending with an
    apostrophe, or the same as another (case aside).  Conservative: Excel's rule as its
    rename refuses, the characters and the length not measured in this round."""
    out: list[str] = []
    seen: set[str] = set()
    for n, name in enumerate(names):
        if not isinstance(name, str) or not name.strip():
            out.append(f"sheets[{n}]: a sheet's name is text that is not empty")
            continue
        if len(name) > MAX_SHEET_NAME:
            out.append(f"sheets[{n}]: longer than {MAX_SHEET_NAME} characters")
        if any(char in _NOT_IN_SHEET_NAMES for char in name):
            out.append(f"sheets[{n}]: holds one of [ ] : * ? / \\")
        if name.startswith("'") or name.endswith("'"):
            out.append(f"sheets[{n}]: starts or ends with an apostrophe")
        if name.casefold() in seen:
            out.append(f"sheets[{n}]: the same name as an earlier sheet (case aside)")
        seen.add(name.casefold())
    return out


def _rels(relationships: Sequence[tuple[str, str, str]]) -> str:
    items = "".join(f'<Relationship Id="{rid}" Type="{kind}" Target="{target}"/>'
                    for rid, kind, target in relationships)
    return (_DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
            f'relationships">{items}</Relationships>')


def _content_types(overrides: Sequence[tuple[str, str]]) -> str:
    items = "".join(f'<Override PartName="/{path}" ContentType="{kind}"/>'
                    for path, kind in overrides)
    return (_DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.'
            'relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
            f"{items}</Types>")


def build(sheets: Sequence[str] = (DEFAULT_SHEET,), *, author: str | None = None,
          title: str | None = None, created: datetime | None = None) -> bytes:
    """The whole package of a new workbook with ``sheets`` (names, checked by the caller),
    as bytes; the first sheet's tab selected, as Excel's.  Deterministic for a given
    ``created`` (now, UTC, when ``None``)."""
    created = (created or datetime.now(timezone.utc)).replace(microsecond=0)
    if created.tzinfo is None:
        created = created.replace(tzinfo=timezone.utc)
    names = list(sheets)
    workbook_rels = [("rId3", REL_STYLES, "styles.xml"), ("rId2", REL_THEME, "theme/theme1.xml")]
    workbook_rels += [(f"rId{_sheet_rid(n)}", REL_WORKSHEET, f"worksheets/sheet{n}.xml")
                      for n in range(len(names), 0, -1)]
    parts: list[tuple[str, str]] = [
        ("_rels/.rels", _rels([("rId3", REL_APP, "docProps/app.xml"),
                               ("rId2", REL_CORE, "docProps/core.xml"),
                               ("rId1", REL_OFFICE_DOCUMENT, "xl/workbook.xml")])),
        ("docProps/app.xml", app_xml(names)),
        ("docProps/core.xml", core_xml(author, title, created)),
        ("xl/_rels/workbook.xml.rels", _rels(workbook_rels)),
        ("xl/styles.xml", styles_xml()),
        ("xl/theme/theme1.xml", theme_xml()),
        ("xl/workbook.xml", workbook_xml(names)),
    ]
    parts += [(f"xl/worksheets/sheet{n}.xml", worksheet_xml(n == 1))
              for n in range(1, len(names) + 1)]
    overrides = [("xl/workbook.xml", CT_WORKBOOK)]
    overrides += [(f"xl/worksheets/sheet{n}.xml", CT_WORKSHEET) for n in range(1, len(names) + 1)]
    overrides += [("xl/theme/theme1.xml", CT_THEME), ("xl/styles.xml", CT_STYLES),
                  ("docProps/core.xml", CT_CORE), ("docProps/app.xml", CT_APP)]
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for path, data in [("[Content_Types].xml", _content_types(overrides)), *parts]:
            entry = zipfile.ZipInfo(path, date_time=(1980, 1, 1, 0, 0, 0))
            entry.compress_type = zipfile.ZIP_DEFLATED
            entry.create_system = 3     # ZipInfo stamps the platform: the same bytes everywhere
            archive.writestr(entry, data.encode("utf-8"))
    return buffer.getvalue()


__all__ = ["APPLICATION", "BASE_COLUMN_WIDTH", "DEFAULT_SHEET", "DEFAULT_TABLE_STYLE",
           "FONT_SIZE", "MAX_SHEET_NAME", "ROW_HEIGHT", "SCRIPT_FONTS", "app_xml", "build",
           "core_xml", "sheet_name_problems",
           "styles_xml", "theme_xml", "workbook_xml", "worksheet_xml"]
