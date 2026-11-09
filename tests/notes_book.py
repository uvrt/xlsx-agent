"""Made-up workbooks for the notes tests (production R57, :mod:`xlsx_agent.edit.notes`).

``plain()`` is a new workbook (``Workbook.new``) with a ``Stock`` and an ``Orders`` sheet,
Orders given a wide, a narrow and a wide column and a tall row: the sheets Excel for Mac
16.106 put its own notes on (round 44, lane C's probe), whose shapes and anchors
:data:`MEASURED` holds.  ``hidden()`` hides Orders' narrow column and row 3, as the second
probe did.  ``threaded()`` gives Stock!B2 a threaded comment ([MS-XLSX] 2.1.x: the
``threadedComments`` part, its ``persons`` list and the plain comment Excel writes beside
it), written here by hand.  ``unshaped()`` is a workbook another program wrote: a comment
on Stock!A2 with no VML drawing.
"""

from __future__ import annotations

import io
import zipfile

from xlsx_agent import Workbook

from build import WorkbookBuilder

#: Each note Excel made (cell -> its shape's margin-left and margin-top in points, and its
#: x:Anchor), measured: plain() for every cell but those of HIDDEN, which hidden() gave.
MEASURED = {
    "Stock!B3": (145, 22, (2, 15, 1, 6, 3, 54, 5, 11)),
    "Stock!B2": (145, 6, (2, 15, 0, 6, 3, 54, 4, 11)),
    "Stock!A1": (80, 2, (1, 15, 0, 2, 2, 54, 4, 7)),
    "Stock!A1048576": (80, 16777139, (1, 15, 1048571, 3, 2, 54, 1048575, 8)),
    "Stock!XFD1": (1064776, 2, (16381, 11, 0, 2, 16382, 50, 4, 7)),
    "Orders!B4": (205, 38, (2, 15, 2, 6, 3, 84, 5, 13)),
    "Orders!C2": (240, 6, (3, 15, 0, 6, 3, 119, 3, 27)),
    "Orders!A1": (80, 2, (1, 15, 0, 2, 1, 119, 3, 23)),
}
HIDDEN = {
    "Orders!B2": (205, 6, (3, 15, 0, 6, 3, 119, 4, 13)),
    "Orders!A5": (80, 52, (1, 15, 3, 20, 1, 119, 7, 11)),
    "Orders!B3": (205, 22, (3, 15, 1, 6, 3, 119, 5, 13)),
}

TC = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
REL_TC = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"
REL_PERSON = "http://schemas.microsoft.com/office/2017/10/relationships/person"
PERSON = "{6C9A1F2E-0B3D-4E5A-9C7B-1D2E3F4A5B6C}"
THREAD = "{1A2B3C4D-5E6F-4A1B-8C2D-3E4F5A6B7C8D}"


def mac_new(sheets) -> bytes:
    """A new workbook with Excel for Mac's own defaults (Aptos Narrow 12, base width 10,
    rows of 16 points), as the workbook Excel put its notes on had: ``Workbook.new`` writes
    Excel for Windows' since round 46."""
    def change(parts):
        parts["xl/styles.xml"] = parts["xl/styles.xml"].replace(b'<sz val="11"/>',
                                                                b'<sz val="12"/>', 1)
        for name in list(parts):
            if name.startswith("xl/worksheets/sheet"):
                parts[name] = parts[name].replace(
                    b'<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>',
                    b'<sheetFormatPr baseColWidth="10" defaultRowHeight="16" '
                    b'x14ac:dyDescent="0.2"/>')
    return _edit(Workbook.new(sheets).to_bytes(), change)


def plain() -> bytes:
    wb = Workbook.open(mac_new(["Stock", "Orders"]))
    wb.set_cells({"Stock!A1": "Item", "Stock!B1": "Count", "Stock!A2": "Bolts",
                  "Stock!B2": 40, "Stock!A3": "Nuts", "Stock!B3": 12, "Orders!A1": "Order",
                  "Orders!B4": 7})
    wb.set_column_widths("Orders!B:B", 20)
    wb.set_column_widths("Orders!C:C", 5)
    wb.set_column_widths("Orders!D:D", 30)
    wb.set_row_heights("Orders!4:4", 30)
    return wb.to_bytes()


def _edit(data: bytes, change) -> bytes:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    change(parts)
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, body in parts.items():
            archive.writestr(name, body)
    return out.getvalue()


def hidden() -> bytes:
    def change(parts):
        sheet = parts["xl/worksheets/sheet2.xml"].decode()
        sheet = sheet.replace('<col min="3" max="3" width="5.83203125" customWidth="1"/>',
                              '<col min="3" max="3" width="5.83203125" hidden="1" '
                              'customWidth="1"/>')
        sheet = sheet.replace('<row r="4"', '<row r="3" hidden="1"/><row r="4"')
        parts["xl/worksheets/sheet2.xml"] = sheet.encode()
    return _edit(plain(), change)


def threaded() -> bytes:
    """Stock!B2 holds a threaded comment (with its plain comment and note shape)."""
    wb = Workbook.open(plain())
    wb.set_comment("Stock!B2", "[Threaded comment]\n\nComment:\n    Is this count right?",
                   author=f"tc={THREAD}")
    data = wb.to_bytes()

    def change(parts):
        parts["xl/threadedComments/threadedComment1.xml"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            f'<ThreadedComments xmlns="{TC}" xmlns:x="http://schemas.openxmlformats.org/'
            'spreadsheetml/2006/main"><threadedComment ref="B2" dT="2026-11-01T10:00:00.00" '
            f'personId="{PERSON}" id="{THREAD}"><text>Is this count right?</text>'
            '</threadedComment></ThreadedComments>').encode()
        parts["xl/persons/person.xml"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            f'<personList xmlns="{TC}" xmlns:x="http://schemas.openxmlformats.org/'
            f'spreadsheetml/2006/main"><person displayName="Ann Example" id="{PERSON}" '
            'userId="Ann Example" providerId="None"/></personList>').encode()
        types = parts["[Content_Types].xml"].decode().replace(
            "</Types>",
            '<Override PartName="/xl/threadedComments/threadedComment1.xml" '
            'ContentType="application/vnd.ms-excel.threadedcomments+xml"/>'
            '<Override PartName="/xl/persons/person.xml" '
            'ContentType="application/vnd.ms-excel.person+xml"/></Types>')
        parts["[Content_Types].xml"] = types.encode()
        rels = parts["xl/worksheets/_rels/sheet1.xml.rels"].decode().replace(
            "</Relationships>",
            f'<Relationship Id="rId9" Type="{REL_TC}" '
            'Target="../threadedComments/threadedComment1.xml"/></Relationships>')
        parts["xl/worksheets/_rels/sheet1.xml.rels"] = rels.encode()
        book = parts["xl/_rels/workbook.xml.rels"].decode().replace(
            "</Relationships>",
            f'<Relationship Id="rId99" Type="{REL_PERSON}" Target="persons/person.xml"/>'
            "</Relationships>")
        parts["xl/_rels/workbook.xml.rels"] = book.encode()
    return _edit(data, change)


def unshaped() -> bytes:
    builder = WorkbookBuilder()
    stock = builder.sheet("Stock")
    stock.rows("A1", [["Item", "Count"], ["Bolts", 4], ["Nuts", 5]])
    stock.comment("A2", "old note", author="Clerk")
    builder.sheet("Orders").cell("A1", "Order")
    return builder.to_bytes()


__all__ = ["HIDDEN", "MEASURED", "hidden", "plain", "threaded", "unshaped"]
