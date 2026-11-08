"""What a workbook says about its cells' places, read from its parts: for comparing what
xlsx-agent's insert, delete and move make of a workbook with what Excel makes of it (round
44, ``tests/test_structure_oracle.py``).

:func:`facts` reads a package (bytes) into plain data -- each cell's formula text (a shared
formula's member computed from its master), type and value; rows' heights, hidden and styled
rows; column widths; merged cells; data validations and conditional formats with their
formulas; hyperlinks; comments and their VML anchors; filters; panes; tables; defined names;
chart series references; pivot cache sources -- in the same form whoever wrote the file, so
two files are compared fact by fact (:func:`differences`).  Read with the standard library
and xlsx-agent's formula lexer only.
"""

from __future__ import annotations

import io
import posixpath
import re
import zipfile
import zlib
from xml.etree import ElementTree

from xlsx_agent.formula import shift

M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PR = "{http://schemas.openxmlformats.org/package/2006/relationships}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
XM = "{http://schemas.microsoft.com/office/excel/2006/main}"
VX = "{urn:schemas-microsoft-com:office:excel}"
_REF = re.compile(r"([A-Z]+)([0-9]+)")


def _col(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def _split(ref: str) -> tuple[int, int]:
    m = _REF.fullmatch(ref)
    assert m, ref
    return int(m.group(2)), _col(m.group(1))


def _rels(z: zipfile.ZipFile, part: str) -> dict[str, tuple[str, str]]:
    folder, base = posixpath.split(part)
    name = posixpath.join(folder, "_rels", base + ".rels")
    if name not in z.namelist():
        return {}
    out = {}
    for rel in ElementTree.fromstring(z.read(name)).iter(PR + "Relationship"):
        target = rel.get("Target", "")
        if rel.get("TargetMode") != "External":
            target = (target.lstrip("/") if target.startswith("/")
                      else posixpath.normpath(posixpath.join(folder, target)))
        out[rel.get("Id", "")] = (rel.get("Type", "").rpartition("/")[2], target)
    return out


def _shared_strings(z: zipfile.ZipFile, wb_rels) -> list[str]:
    for kind, target in wb_rels.values():
        if kind == "sharedStrings":
            root = ElementTree.fromstring(z.read(target))
            return ["".join(t.text or "" for t in si.iter(M + "t"))
                    for si in root.findall(M + "si")]
    return []


def facts(data: bytes) -> dict:
    """Every fact of a workbook this module compares (see the module)."""
    z = zipfile.ZipFile(io.BytesIO(data))
    wb = ElementTree.fromstring(z.read("xl/workbook.xml"))
    wb_rels = _rels(z, "xl/workbook.xml")
    strings = _shared_strings(z, wb_rels)
    styles = _styles(z, wb_rels)
    out: dict = {"sheets": {}, "names": {}, "charts": [], "pivots": []}
    sheet_names = []
    for sheet in wb.iter(M + "sheet"):
        name = sheet.get("name")
        sheet_names.append(name)
        part = wb_rels[sheet.get(RNS + "id")][1]
        out["sheets"][name] = _sheet(z, part, strings, styles)
    for node in wb.iter(M + "definedName"):
        local = node.get("localSheetId")
        key = (sheet_names[int(local)] + "!" if local is not None else "") + node.get("name")
        out["names"][key] = node.text
    for name in z.namelist():
        if name.startswith("xl/charts/chart") and name.endswith(".xml"):
            root = ElementTree.fromstring(z.read(name))
            out["charts"].append([f.text for f in root.iter(C + "f")])
        if name.startswith("xl/pivotCache/pivotCacheDefinition"):
            root = ElementTree.fromstring(z.read(name))
            for source in root.iter(M + "worksheetSource"):
                out["pivots"].append(dict(source.attrib))
    return out


def _sheet(z: zipfile.ZipFile, part: str, strings: list[str], styles: list) -> dict:
    root = ElementTree.fromstring(z.read(part))
    rels = _rels(z, part)
    cells: dict[str, tuple] = {}
    masters: dict[str, tuple[int, int, str]] = {}
    column_styles = {}
    for col in root.iter(M + "col"):
        if col.get("style") is not None:
            for n in range(int(col.get("min")), int(col.get("max")) + 1):
                column_styles[n] = styles[int(col.get("style"))]
    rows: dict[int, dict] = {}
    for row in root.iter(M + "row"):
        props = {k: row.get(k) for k in ("ht", "customHeight", "hidden", "customFormat")
                 if row.get(k) not in (None, "0", "false")}
        if row.get("customFormat") in ("1", "true"):
            props["s"] = styles[int(row.get("s", "0"))]
        if row.get("customHeight") in (None, "0", "false"):
            props.pop("ht", None)
        if props:
            rows[int(row.get("r"))] = props
        for c in row.findall(M + "c"):
            ref = c.get("r")
            f = c.find(M + "f")
            v = c.find(M + "v")
            t = c.get("t", "n")
            value = v.text if v is not None else None
            if t == "s" and value is not None:
                value, t = strings[int(value)], "str"
            elif t == "inlineStr":
                value, t = "".join(x.text or "" for x in c.iter(M + "t")), "str"
            elif t == "str":
                t = "str"
            if value is not None and t == "n":
                value = float(value)
            formula = None
            if f is not None:
                kind = f.get("t", "normal")
                if kind == "shared":
                    si = f.get("si")
                    if f.text:
                        r, k = _split(ref)
                        masters[si] = (r, k, f.text)
                        formula = f.text
                    else:
                        r0, k0, text = masters[si]
                        r, k = _split(ref)
                        formula = shift(text, r - r0, k - k0)
                elif kind == "array":
                    formula = "{" + (f.text or "") + "}@" + f.get("ref", "")
                else:
                    formula = f.text
            style = styles[int(c.get("s", "0"))]
            if formula is None and value is None and (
                    style is None or style == _fallback(row, styles, column_styles,
                                                        _split(ref)[1])):
                continue
            cells[ref] = (formula, t if value is not None else "blank", value, style)
    cols = [{k: col.get(k) for k in ("min", "max", "width", "hidden", "style")
             if col.get(k) is not None} for col in root.iter(M + "col")]
    for col in cols:
        if "style" in col:
            col["style"] = styles[int(col["style"])]
        if "width" in col:
            col["width"] = round(float(col["width"]), 2)
    pane = root.find(f"{M}sheetViews/{M}sheetView/{M}pane")
    validations = []
    for dv in root.iter(M + "dataValidation"):
        for cell, texts in _per_cell(dv.get("sqref", ""), (_text(dv.find(M + "formula1")),
                                                           _text(dv.find(M + "formula2")))):
            validations.append((cell, dv.get("type")) + texts)
    for dv in root.iter("{http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}"
                        "dataValidation"):
        validations.append((" ".join(sorted(_text(dv.find(XM + "sqref")).split())),
                            dv.get("type"), None, None))
    formats = []
    for cf in root.iter(M + "conditionalFormatting"):
        for rule in cf.findall(M + "cfRule"):
            texts = tuple(f.text for f in rule.findall(M + "formula"))
            for cell, here in _per_cell(cf.get("sqref", ""), texts):
                formats.append((cell, rule.get("type"), here))
    links = []
    for link in root.iter(M + "hyperlink"):
        rid = link.get(RNS + "id")
        links.append((link.get("ref"), link.get("location") or (rels.get(rid, ("", ""))[1]
                                                                 if rid else None)))
    merges = sorted(m.get("ref") for m in root.iter(M + "mergeCell"))
    tables = []
    comments: list = []
    vml: list = []
    for rid, (kind, target) in rels.items():
        if kind == "table":
            t = ElementTree.fromstring(z.read(target))
            af = t.find(M + "autoFilter")
            tables.append({"name": t.get("name"), "ref": t.get("ref"),
                           "autoFilter": af.get("ref") if af is not None else None,
                           "columns": [c.get("name") for c in t.iter(M + "tableColumn")],
                           "calculated": [_text(c.find(M + "calculatedColumnFormula"))
                                          for c in t.iter(M + "tableColumn")]})
        elif kind == "comments":
            croot = ElementTree.fromstring(z.read(target))
            comments = sorted((c.get("ref"), "".join(x.text or "" for x in c.iter(M + "t")))
                              for c in croot.iter(M + "comment"))
        elif kind == "vmlDrawing":
            text = z.read(target).decode("utf-8", "replace")
            for row, column in re.findall(r"<x:Row>(\d+)</x:Row>\s*<x:Column>(\d+)</x:Column>",
                                          text):
                vml.append((int(row), int(column)))
            vml.sort()
    af = root.find(M + "autoFilter")
    return {"cells": cells, "rows": rows, "cols": cols,
            "pane": None if pane is None else {k: pane.get(k) for k in
                                                ("xSplit", "ySplit", "topLeftCell", "state")
                                                if pane.get(k) is not None},
            "validations": sorted(validations, key=repr), "formats": sorted(formats, key=repr),
            "links": sorted(links, key=repr), "merges": merges, "tables": tables,
            "comments": comments, "vml": vml,
            "autoFilter": af.get("ref") if af is not None else None}


def _styles(z: zipfile.ZipFile, wb_rels) -> list:
    """Each cell format as what it shows (number format, font, fill, border, alignment),
    ``None`` for one that shows as the default: Excel writes equal formats under new
    indexes."""
    for kind, target in wb_rels.values():
        if kind == "styles":
            root = ElementTree.fromstring(z.read(target))
            fonts = [_font(f) for f in root.iter(M + "font")]
            fills = [ElementTree.tostring(f) for f in root.iter(M + "fill")]
            formats = {n.get("numFmtId"): n.get("formatCode") for n in root.iter(M + "numFmt")}
            xfs = root.find(M + "cellXfs")
            out = []
            for xf in (xfs if xfs is not None else []):
                number = xf.get("numFmtId", "0")
                key = (formats.get(number, _BUILTIN.get(number, number)),
                       fonts[int(xf.get("fontId", "0"))] if xf.get("fontId", "0") != "0"
                       else None,
                       fills[int(xf.get("fillId", "0"))] if xf.get("fillId", "0") != "0"
                       else None, xf.get("borderId", "0"),
                       tuple(sorted((xf.find(M + "alignment").attrib
                                     if xf.find(M + "alignment") is not None else {}).items())))
                out.append(None if key == ("0", None, None, "0", ()) else
                           f"nf={key[0]}" + (f" font={zlib.crc32(key[1]):x}" if key[1] else "")
                           + (f" fill={zlib.crc32(key[2]):x}" if key[2] else "")
                           + (f" border={key[3]}" if key[3] != "0" else "")
                           + (f" align={key[4]}" if key[4] else ""))
            return out
    return [None]


#: Built-in number formats by id, as far as the fixtures use them (ECMA-376 Part 1,
#: §18.8.30): Excel may write a custom ``0.00`` as built-in 2.
_BUILTIN = {"1": "0", "2": "0.00", "3": "#,##0", "4": "#,##0.00", "9": "0%", "10": "0.00%"}


def _fallback(row, styles, column_styles, column):
    if row.get("customFormat") in ("1", "true"):
        return styles[int(row.get("s", "0"))]
    return column_styles.get(column)


def _font(node) -> bytes:
    """A font as what it shows: bold, italic, size, colour, name (Excel adds family and
    scheme on a re-save)."""
    parts = []
    for tag in ("b", "i", "u", "strike", "sz", "color", "name"):
        child = node.find(M + tag)
        if child is not None:
            parts.append(f"{tag}={sorted(child.attrib.items())}")
    return ";".join(parts).encode()


def _per_cell(sqref: str, formulas: tuple) -> list[tuple[str, tuple]]:
    """A rule's cells, each with the formulas as they read at that cell (a rule's
    formulas are written for the top-left cell of its range's bounding box)."""
    areas = []
    for piece in sqref.split():
        first, _, last = piece.partition(":")
        last = last or first
        r1, c1 = _split(first.replace("$", ""))
        r2, c2 = _split(last.replace("$", ""))
        areas.append((r1, c1, r2, c2))
    if not areas:
        return []
    top = min(a[0] for a in areas), min(a[1] for a in areas)
    out = []
    for r1, c1, r2, c2 in areas:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                try:
                    here = tuple(shift(f, r - top[0], c - top[1]) if f else f
                                 for f in formulas)
                except ValueError:
                    here = ("(off the grid)",)
                out.append((_name(r, c), here))
    return out


def _name(row: int, column: int) -> str:
    letters = ""
    while column:
        column, rest = divmod(column - 1, 26)
        letters = chr(65 + rest) + letters
    return f"{letters}{row}"


def _text(node) -> str | None:
    return None if node is None else node.text


def differences(ours: dict, theirs: dict, path: str = "") -> list[str]:
    """Where two :func:`facts` differ, as ``path: ours != theirs`` lines."""
    out: list[str] = []
    if isinstance(ours, dict) and isinstance(theirs, dict):
        for key in sorted(set(ours) | set(theirs), key=str):
            if key not in ours:
                out.append(f"{path}/{key}: missing in ours, theirs {theirs[key]!r}")
            elif key not in theirs:
                out.append(f"{path}/{key}: ours {ours[key]!r}, missing in theirs")
            else:
                out += differences(ours[key], theirs[key], f"{path}/{key}")
        return out
    if ours != theirs:
        out.append(f"{path}: ours {ours!r} != theirs {theirs!r}")
    return out
