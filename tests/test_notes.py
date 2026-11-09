"""Notes on cells (production R57): ``set_comment``, ``set_comments``, ``remove_comment``,
written as Excel for Mac 16.106 wrote its own (round 44, lane C's probe:
:mod:`xlsx_agent.edit.notes`; the measured shapes are tests/notes_book.py's).

Tool design rules: a refusal names every problem and the way out (2), a note already so is
``changed`` false (3), many notes are one call and one step (5), and what is written is
what Excel writes, measured (6).
"""

from __future__ import annotations

import io
import re
import zipfile

import pytest
from lxml import etree

from xlsx_agent import Workbook
from xlsx_agent.edit import EditError
from xlsx_agent.edit.policy import WritePolicy

import notes_book
from shop import Shop

X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
V = "{urn:schemas-microsoft-com:vml}"
XL = "{urn:schemas-microsoft-com:office:excel}"
O = "{urn:schemas-microsoft-com:office:office}"


def parts(wb: Workbook) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(wb.to_bytes())) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def shapes(data: dict[str, bytes], vml: str) -> dict[str, tuple]:
    """Each note shape of a VML part by its cell: (id, margin-left, margin-top, anchor)."""
    root = etree.fromstring(data[vml])
    out = {}
    for shape in root.iter(V + "shape"):
        client = shape.find(XL + "ClientData")
        row, column = int(client.find(XL + "Row").text), int(client.find(XL + "Column").text)
        style = shape.get("style")
        left = int(re.search(r"margin-left:(\d+)pt", style).group(1))
        top = int(re.search(r"margin-top:(\d+)pt", style).group(1))
        anchor = tuple(int(n) for n in client.find(XL + "Anchor").text.split(","))
        out[(row, column)] = (shape.get("id"), left, top, anchor)
    return out


def _cell(address: str) -> tuple[int, int]:
    from ooxml_edit.sheets import Cell

    cell = Cell.parse(address.split("!")[1])
    return cell.row - 1, cell.column - 1


def test_a_first_note_writes_the_parts_excel_writes():
    wb = Workbook.open(notes_book.plain())
    change = wb.set_comment("Stock!B3", "Check count")
    assert change.added == ("Stock!B3",) and change.changed
    data = parts(wb)
    types = data["[Content_Types].xml"].decode()
    assert '<Default Extension="vml" ContentType="application/vnd.openxmlformats-' \
           'officedocument.vmlDrawing"/>' in types
    assert '<Override PartName="/xl/comments1.xml" ContentType="application/vnd.' \
           'openxmlformats-officedocument.spreadsheetml.comments+xml"/>' in types
    rels = data["xl/worksheets/_rels/sheet1.xml.rels"].decode()
    assert "relationships/comments" in rels and "relationships/vmlDrawing" in rels
    sheet = etree.fromstring(data["xl/worksheets/sheet1.xml"])
    names = [etree.QName(child).localname for child in sheet]
    assert names[-1] == "legacyDrawing" and names.index("pageMargins") < names.index(
        "legacyDrawing")
    comments = etree.fromstring(data["xl/comments1.xml"])
    comment = comments.find(f"{X}commentList/{X}comment")
    assert (comment.get("ref"), comment.get("authorId"), comment.get("shapeId")) == \
        ("B3", "0", "0")
    run = comment.find(f"{X}text/{X}r")
    font = [(etree.QName(n).localname, dict(n.attrib)) for n in run.find(X + "rPr")]
    assert font == [("b", {}), ("sz", {"val": "10"}), ("color", {"rgb": "FF000000"}),
                    ("rFont", {"val": "Tahoma"}), ("family", {"val": "2"})]
    vml = etree.fromstring(data["xl/drawings/vmlDrawing1.vml"])
    assert vml.find(f"{O}shapelayout/{O}idmap").get("data") == "1"
    assert vml.find(V + "shapetype").get("id") == "_x0000_t202"
    shape = vml.find(V + "shape")
    assert shape.get("id") == "_x0000_s1025" and "visibility:hidden" in shape.get("style")
    assert "width:104pt;height:69pt" in shape.get("style")
    assert shape.get("fillcolor") == "#ffffe1 [80]"
    [note] = wb.comments("Stock")
    assert (str(note.cell.row), note.text, note.author, note.threaded) == \
        ("3", "Check count", "", False)
    assert wb.validate() == []
    wb.save(io.BytesIO())


@pytest.mark.parametrize("book, measured", [("plain", notes_book.MEASURED),
                                            ("hidden", notes_book.HIDDEN)])
def test_every_shape_sits_where_excel_put_it(book, measured):
    wb = Workbook.open(getattr(notes_book, book)())
    wb.set_comments({address: "note" for address in measured})
    data = parts(wb)
    for address, (left, top, anchor) in measured.items():
        sheet = address.split("!")[0]
        number = {"Stock": 1, "Orders": 2}[sheet] if book == "plain" else 1
        found = shapes(data, f"xl/drawings/vmlDrawing{number}.vml")[_cell(address)]
        assert found[1:] == (left, top, anchor), address


def test_text_is_stored_as_excel_stores_it():
    wb = Workbook.open(notes_book.plain())
    text = "Line one\nTab\there & <x> \x01end"
    wb.set_comment("Stock!B2", text)
    stored = parts(wb)["xl/comments1.xml"].decode()
    assert "Line one\nTab\there &amp; &lt;x&gt; _x0001_end" in stored
    assert wb.comments("Stock")[0].text == text
    wb.set_comment("Stock!C2", " padded ")
    assert 'xml:space="preserve"> padded </t>' in parts(wb)["xl/comments1.xml"].decode()
    assert {c.text for c in wb.comments("Stock")} == {text, " padded "}


def test_notes_in_reading_order_and_shape_ids_rising():
    wb = Workbook.open(notes_book.plain())
    wb.set_comment("Stock!B3", "three")
    wb.set_comment("Stock!B2", "two")
    wb.set_comment("Stock!A1", "one")
    data = parts(wb)
    refs = [c.get("ref") for c in etree.fromstring(data["xl/comments1.xml"]).iter(
        X + "comment")]
    assert refs == ["A1", "B2", "B3"]
    ids = {cell: found[0] for cell, found in shapes(data, "xl/drawings/vmlDrawing1.vml").items()}
    assert ids == {(2, 1): "_x0000_s1025", (1, 1): "_x0000_s1026", (0, 0): "_x0000_s1027"}
    wb.remove_comment("Stock!B3")
    wb.set_comment("Stock!C3", "after delete")          # the highest id plus one, as Excel
    ids = {cell: found[0] for cell, found in
           shapes(parts(wb), "xl/drawings/vmlDrawing1.vml").items()}
    assert ids[(2, 2)] == "_x0000_s1028" and (2, 1) not in ids


def test_a_note_replaced_keeps_its_shape_and_its_author():
    wb = Workbook.open(notes_book.plain())
    wb.set_comment("Stock!B3", "first", author="Ann")
    before = shapes(parts(wb), "xl/drawings/vmlDrawing1.vml")
    change = wb.set_comment("Stock!B3", "second")
    assert change.replaced == ("Stock!B3",) and not change.added
    assert shapes(parts(wb), "xl/drawings/vmlDrawing1.vml") == before
    [note] = wb.comments("Stock")
    assert (note.text, note.author) == ("second", "Ann")
    wb.set_comment("Stock!B3", "second", author="Bo")
    assert wb.comments("Stock")[0].author == "Bo"
    authors = etree.fromstring(parts(wb)["xl/comments1.xml"]).find(X + "authors")
    assert [a.text for a in authors] == ["Bo"]              # Ann is no one's any more


def test_a_note_already_so_changes_nothing():
    wb = Workbook.open(notes_book.plain())
    wb.set_comments({"Stock!B3": "same", "Stock!B2": "two"}, author="Ann")
    version, data = wb.version, wb.to_bytes()
    change = wb.set_comments({"Stock!B3": "same", "Stock!B2": "two"})
    assert not change.changed and change.unchanged == ("Stock!B3", "Stock!B2")
    assert change.note == "every note given is already so"
    assert wb.version == version and wb.to_bytes() == data
    assert not wb.remove_comment("Stock!D4").changed      # no note there: no error
    assert wb.remove_comment("Stock!D4").note == "no note there"


def test_many_notes_one_step_and_undo_exact():
    wb = Workbook.open(notes_book.plain())
    original = wb.to_bytes()
    wb.set_comments({f"Stock!C{n}": f"row {n}" for n in range(1, 41)})
    assert len(wb.comments("Stock")) == 40
    edited = wb.to_bytes()
    assert wb.undo() and wb.to_bytes() == original
    assert wb.redo() and wb.to_bytes() == edited
    with wb.batch():
        wb.set_cells({"Stock!C50": 5})
        wb.set_comment("Stock!C50", "with its value")
    assert wb.undo() and wb.to_bytes() == edited
    wb.set_comment("Stock!D1", "")                     # an empty note, as Excel allows
    empty = wb.to_bytes()
    assert wb.undo() and wb.redo() and wb.to_bytes() == empty
    assert [c.text for c in wb.comments("Stock") if c.cell.column == 4] == [""]


def test_removing_the_last_note_takes_the_parts_away():
    wb = Workbook.open(notes_book.plain())
    original = parts(wb)
    wb.set_comments({"Stock!B3": "a", "Stock!B2": "b", "Orders!A1": "c"})
    change = wb.remove_comment("Stock!A1:B3")
    assert change.removed == ("Stock!B2", "Stock!B3")
    after = parts(wb)
    assert "xl/comments1.xml" not in after and "xl/drawings/vmlDrawing1.vml" not in after
    assert "xl/worksheets/_rels/sheet1.xml.rels" not in after
    assert after["xl/worksheets/sheet1.xml"] == original["xl/worksheets/sheet1.xml"]
    assert "legacyDrawing" in after["xl/worksheets/sheet2.xml"].decode()
    wb.remove_comment(["Orders!"])
    assert set(parts(wb)) == set(original)
    assert wb.comments("Orders") == [] and wb.validate() == []


def test_every_problem_is_refused_before_anything_changes():
    wb = Workbook.open(Shop(protected=True).to_bytes())
    data = wb.to_bytes()
    with pytest.raises(EditError) as raised:
        wb.set_comments({"B3": "no sheet", "Stock!A1:B2": "a range", "Stock!C3": 5,
                         "Stock!C4": "x" * 32_768, "Totals!B2": "protected",
                         "Nowhere!A1": "no such sheet"})
    codes = [code for code, _ in raised.value.problems]
    assert codes == ["bad_address", "bad_address", "bad_value", "text_too_long", "no_sheet",
                     "protected"]
    assert "Stock!B3" in str(raised.value)            # the way out: an address that works
    assert wb.to_bytes() == data
    with pytest.raises(EditError) as raised:
        wb.set_comments([("Stock!C3", "a"), ("Stock!C3", "b")])
    assert raised.value.code == "duplicate_cell"
    with pytest.raises(EditError) as raised:
        wb.set_comment("Stock!C3", "a", author="two\nlines")
    assert raised.value.code == "bad_value"


def test_the_write_policy_covers_notes():
    wb = Workbook.open(notes_book.plain(), policy=WritePolicy(cells=("Stock!A1:A10",)))
    wb.set_comment("Stock!A2", "allowed")
    with pytest.raises(EditError) as raised:
        wb.set_comment("Stock!B2", "not allowed")
    assert raised.value.code == "not_writable" and raised.value.addresses == ("Stock!B2",)


def test_a_threaded_comment_is_never_written_over():
    data = notes_book.threaded()
    wb = Workbook.open(data)
    [note] = wb.comments("Stock")
    assert note.threaded and note.text.startswith("[Threaded comment]")
    with pytest.raises(EditError) as raised:
        wb.set_comment("Stock!B2", "over it")
    assert raised.value.code == "threaded_comment" and "Stock!B2" in raised.value.addresses
    with pytest.raises(EditError) as raised:
        wb.remove_comment("Stock!A1:C3")
    assert raised.value.code == "threaded_comment"
    assert wb.to_bytes() == data
    wb.set_comment("Stock!C3", "beside it")
    after = parts(wb)
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        assert after["xl/threadedComments/threadedComment1.xml"] == archive.read(
            "xl/threadedComments/threadedComment1.xml")
    assert {(c.text, c.threaded) for c in wb.comments("Stock")} >= {("beside it", False)}


def test_a_comment_written_without_a_shape_is_given_one():
    wb = Workbook.open(notes_book.unshaped())
    change = wb.set_comment("Stock!B5", "beside old")
    assert change.added == ("Stock!B5",) and change.shaped == ("Stock!A2",)
    data = parts(wb)
    vml = [name for name in data if name.endswith(".vml")]
    assert len(vml) == 1 and set(shapes(data, vml[0])) == {(1, 0), (4, 1)}
    assert {c.text for c in wb.comments("Stock")} == {"old note", "beside old"}
    assert {c.author for c in wb.comments("Stock")} == {"Clerk", ""}


def test_rows_inserted_and_deleted_move_new_notes_with_their_cells():
    wb = Workbook.open(notes_book.plain())
    wb.set_comments({"Stock!B3": "three", "Stock!B5": "five"})
    wb.insert_rows("Stock", 2, 2)
    cells = sorted((c.cell.row, c.cell.column) for c in wb.comments("Stock"))
    assert cells == [(5, 2), (7, 2)]
    found = shapes(parts(wb), "xl/drawings/vmlDrawing1.vml")
    assert set(found) == {(4, 1), (6, 1)}
    assert found[(4, 1)][3][2] == 3 and found[(4, 1)][3][6] == 7      # rows moved by 2
    wb.delete_rows("Stock", 5)
    assert [(c.cell.row, c.text) for c in wb.comments("Stock")] == [(6, "five")]
    assert set(shapes(parts(wb), "xl/drawings/vmlDrawing1.vml")) == {(5, 1)}
    wb.set_comment("Stock!B2", "new after")             # a note after the move: its own id
    assert len(shapes(parts(wb), "xl/drawings/vmlDrawing1.vml")) == 2
    assert wb.validate() == []


def test_a_sheet_s_notes_survive_a_reopen_and_a_save():
    wb = Workbook.open(notes_book.plain())
    wb.set_comments({"Stock!B3": "three", "Orders!B4": "four"}, author="Ann")
    out = io.BytesIO()
    wb.save(out)
    again = Workbook.open(out.getvalue())
    assert [(c.text, c.author) for c in again.comments("Orders")] == [("four", "Ann")]
    again.set_comment("Stock!B2", "two")
    assert len(again.comments("Stock")) == 2
    assert again.validate() == []
