"""Markdown rich text in cells (X3b): ``{"markdown": ...}`` written as runs, read back.

On the corner shop's stock list (``tests/shop.py``): Markdown written as a shared string of
runs, each with the cell's whole font and its marks; Markdown without marks and literal
text as plain text; lines, bullets and numbers shown in the cell, the cell given a
wrapping format, a fixed-height row named; every guard and the atomicity of set_cells; the
footprint (shared strings and cell formats appended); and the opt-in Markdown read, which
writes back as the same cell.
"""

from __future__ import annotations

import io
import zipfile

import pytest
from lxml import etree

from ooxml_edit.markdown import Run
from ooxml_edit.sheets import Cell, outside_footprint
from ooxml_edit.tools import Toolbox

from xlsx_agent import Workbook
from xlsx_agent.edit import EditError
from xlsx_agent.markdown import BULLET, cell_markdown, cell_runs
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS

from shop import FIRST_INPUT, Shop

X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
B10 = f"Stock!B{FIRST_INPUT}"


def _open(**options) -> tuple[bytes, Workbook]:
    data = Shop(**options).to_bytes()
    return data, Workbook.open(data)


def _part(data: bytes, name: str) -> etree._Element:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return etree.fromstring(archive.read(name))


def _runs(wb: Workbook, address: str):
    sheet, _, ref = address.partition("!")
    return wb.editor.cells.sheet(sheet).runs(Cell.parse(ref))


# -- writing ------------------------------------------------------------------------------------


def test_markdown_is_written_as_runs_with_the_cells_font():
    before, wb = _open()
    written = wb.set_cells({B10: {"markdown": "**Total** due *now*, ~~not~~ ***both***"}})
    assert written.changed == 1
    assert wb.cell(B10).v == "Total due now, not both"
    assert _runs(wb, B10) == (Run("Total", bold=True), Run(" due "), Run("now", italic=True),
                              Run(", "), Run("not", strike=True), Run(" "),
                              Run("both", bold=True, italic=True))
    after = wb.to_bytes()
    item = _part(after, "xl/sharedStrings.xml").findall(X + "si")[-1]
    for run in item.findall(X + "r"):
        props = run.find(X + "rPr")
        assert props.find(X + "rFont").get("val") == "Aptos"   # the cell's font, kept
        assert props.find(X + "sz") is not None
    assert outside_footprint(before, after, cells={"Stock": [B10.partition("!")[2]]}) == []


def test_text_stays_literal_and_markdown_without_marks_is_plain():
    _, wb = _open()
    wb.set_cells({B10: "**not bold**", f"Stock!B{FIRST_INPUT + 1}": {"markdown": "a \\*b\\*"},
                  f"Stock!B{FIRST_INPUT + 2}": {"markdown": "`**code**` stays"}})
    assert wb.cell(B10).v == "**not bold**" and _runs(wb, B10) is None
    assert wb.cell(f"Stock!B{FIRST_INPUT + 1}").v == "a *b*"
    assert _runs(wb, f"Stock!B{FIRST_INPUT + 1}") is None
    assert wb.cell(f"Stock!B{FIRST_INPUT + 2}").v == "**code** stays"


def test_lines_and_lists_show_in_the_cell_and_wrap():
    _, wb = _open()
    markdown = ("Order **list**\n\n\n- one\n* two\n  + nested *it*\n    3. deep\n    9. deeper\n"
                "  more of two\n1) first\n1) second")
    wb.set_cells({B10: {"markdown": markdown}})
    shown = wb.cell(B10).v
    assert shown == (f"Order list\n\n{BULLET}one\n{BULLET}two\n    {BULLET}nested it\n"
                     "        3. deep\n        4. deeper\n  more of two\n"
                     "1) first\n2) second")
    styles = wb.editor.cells.styles
    style = wb.editor.cells.sheet("Stock").style(Cell.parse(B10.partition("!")[2]))
    assert styles.wraps_text(style) and not styles.wraps_text(5)
    assert styles.number_format(style) == styles.number_format(5)
    count = styles.count
    wb.set_cells({f"Stock!B{FIRST_INPUT + 1}": {"markdown": "- a\n- b"}})
    assert styles.count == count                       # the wrapping format is reused
    assert wb.cell(f"Stock!B{FIRST_INPUT + 1}").v == f"{BULLET}a\n{BULLET}b"


def test_one_line_keeps_the_format_and_a_fixed_height_row_is_named():
    _, wb = _open()
    written = wb.set_cells({B10: {"markdown": "**one** line"},
                            "Stock!A1": {"markdown": "Title\n*second line*"}})
    style = wb.editor.cells.sheet("Stock").style(Cell.parse(B10.partition("!")[2]))
    assert style == 5
    assert written.fixed_height == ("Stock!A1",)
    assert written.as_dict()["fixed_height"] == ["Stock!A1"]
    sheet = _part(wb.to_bytes(), "xl/worksheets/sheet1.xml")
    first = sheet.find(f"{X}sheetData/{X}row")
    assert first.get("ht") == "21" and first.get("customHeight") == "1"   # never changed


def test_markdown_meets_every_guard_and_refuses_before_anything_changes():
    _, wb = _open()
    before, version = wb.to_bytes(), wb.version
    with pytest.raises(EditError) as raised:
        wb.set_cells({B10: {"markdown": "**fine**"},
                      f"Stock!F{FIRST_INPUT + 1}": {"markdown": "**over a formula**"},
                      f"Stock!B{FIRST_INPUT + 2}": {"markdown": 12},
                      f"Stock!B{FIRST_INPUT + 3}": {"markdown": "\n\n  \n"},
                      f"Stock!B{FIRST_INPUT + 4}": {"markdown": "*x* " + "x" * 33_000},
                      f"Stock!B{FIRST_INPUT + 5}": {"markdown": "a", "text": "b"}})
    codes = {code for code, _ in raised.value.problems}
    assert codes == {"formula_cell", "bad_value", "text_too_long"}
    assert len(raised.value.problems) == 5
    assert wb.to_bytes() == before and wb.version == version
    written = wb.set_cells({f"Stock!F{FIRST_INPUT + 1}": {"markdown": "**over**"}},
                           replace_formulas=True)
    assert written.changed == 1 and wb.cell(f"Stock!F{FIRST_INPUT + 1}").formula is False


def test_a_refused_markdown_write_inside_a_batch_rolls_it_back():
    data, wb = _open()
    with pytest.raises(EditError):
        with wb.batch():
            wb.set_cells({B10: {"markdown": "**a**\n- b"}})
            wb.set_cells({f"Stock!F{FIRST_INPUT}": {"markdown": "**c**"}})
    assert wb.to_bytes() == data


def test_undo_gives_the_original_bytes_and_the_same_markdown_twice_is_no_change():
    data, wb = _open()
    wb.set_cells({B10: {"markdown": "**Spare** fan\n- one"}})
    edited = wb.to_bytes()
    assert wb.set_cells({B10: {"markdown": "**Spare** fan\n- one"}}).changed == 0
    wb.undo()
    assert wb.to_bytes() == data
    wb.redo()
    assert wb.to_bytes() == edited


def test_inline_strings_and_a_saved_file_read_back():
    _, wb = _open(shared=False)
    wb.set_cells({B10: {"markdown": "**Spare** fan"}})
    report = wb.save(io.BytesIO())
    assert report.verified >= 1
    assert wb.cell(B10, markdown=True).v == "**Spare** fan"


# -- reading ------------------------------------------------------------------------------------


ROUND_TRIPS = [
    "**Total** due *now*",
    "~~gone~~ and ***both***",
    "Title\n\n- one\n- **two**\n  - nested\n    1. deep\n    2. deeper\n- three",
    "3) x\n4) y\n   more of y\n\nafter",
    "a \\*literal\\* star, C:\\\\path, \\_under\\_ and \\`tick\\`",
    "snake_case_name and 2*3*4",
]


@pytest.mark.parametrize("markdown", ROUND_TRIPS)
def test_written_markdown_reads_back_and_writes_back_as_the_same_cell(markdown):
    _, wb = _open()
    wb.set_cells({B10: {"markdown": markdown}})
    reopened = Workbook.open(wb.to_bytes())
    read = reopened.cell(B10, markdown=True).v
    assert cell_runs(read) == _runs(wb, B10) or (
        _runs(wb, B10) is None and cell_runs(read) == (Run(wb.cell(B10).v),))
    _, other = _open()
    other.set_cells({B10: {"markdown": read}})
    assert other.cell(B10).v == wb.cell(B10).v
    assert _runs(other, B10) == _runs(wb, B10)


def test_a_read_gives_plain_text_unless_asked_and_escapes_plain_cells():
    _, wb = _open()
    wb.set_cells({B10: {"markdown": "**Spare** fan"}, f"Stock!B{FIRST_INPUT + 1}": "a*b*c",
                  f"Stock!B{FIRST_INPUT + 2}": f"{BULLET}x\n{BULLET}y"})
    rows = list(wb.read(f"Stock!B{FIRST_INPUT}:B{FIRST_INPUT + 2}"))
    assert [row[0].v for row in rows] == ["Spare fan", "a*b*c", f"{BULLET}x\n{BULLET}y"]
    rows = list(wb.read(f"Stock!B{FIRST_INPUT}:B{FIRST_INPUT + 2}", markdown=True))
    assert [row[0].v for row in rows] == ["**Spare** fan", "a\\*b\\*c", "- x\n- y"]
    assert rows[0][0].t == "text"
    numbers = [v for v in wb.cells("Stock", markdown=True) if v.t == "number"]
    assert numbers and all(not isinstance(v.v, str) for v in numbers)


def test_what_markdown_cannot_hold_reads_as_text_lines():
    assert cell_markdown((Run("3. x\n9. y"),)) == "3\\. x\n9\\. y"
    assert cell_runs(cell_markdown((Run("3. x\n9. y"),))) == (Run("3. x\n9. y"),)
    assert cell_markdown((Run("- not a list"),)) == "\\- not a list"
    assert cell_markdown((Run("a", bold=True), Run("(b)"))) == "**a**(b)"


def test_a_redacted_term_across_runs_drops_the_marks_not_the_redaction():
    _, wb = _open()
    wb.set_cells({B10: {"markdown": "Init**ech** fan"}, f"Stock!B{FIRST_INPUT + 1}":
                  {"markdown": "**Initech** fan"}})
    reopened = Workbook.open(wb.to_bytes())
    reopened.redact([{"redact": ["Initech"], "with": "***"}])
    first = reopened.cell(B10, markdown=True).v
    second = reopened.cell(f"Stock!B{FIRST_INPUT + 1}", markdown=True).v
    assert "Init" not in first and first == "\\*\\*\\* fan"
    assert second == "**\\*\\*\\*** fan"


# -- the tools ----------------------------------------------------------------------------------


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _session(box):
    session = box.session()
    blob = session.add_blob(Shop().to_bytes(), name="stock.xlsx")
    assert box.dispatch(session, "open_document", {"blob": blob}).ok
    return session


def test_xlsx_set_cells_markdown_and_xlsx_read_markdown(box):
    session = _session(box)
    result = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": B10, "markdown": "**Spare** fan\n- one"},
        {"address": "Stock!A1", "markdown": "Title\n*sub*"},
        {"address": f"Stock!B{FIRST_INPUT + 1}", "text": "**literal**"}]})
    assert result.ok, result.to_json()
    assert result.data["fixed_height"] == ["Stock!A1"]
    assert any("fixed height" in w for w in result.warnings)
    plain = box.dispatch(session, "xlsx_read", {"doc": "d1", "target": B10})
    assert plain.data["rows"][0][0]["v"] == f"Spare fan\n{BULLET}one"
    read = box.dispatch(session, "xlsx_read", {"doc": "d1", "markdown": True,
                                               "target": f"Stock!B{FIRST_INPUT}:B"
                                                         f"{FIRST_INPUT + 1}"})
    assert [row[0]["v"] for row in read.data["rows"]] == ["**Spare** fan\n- one",
                                                          "\\*\\*literal\\*\\*"]
    both = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": B10, "markdown": "**a**", "text": "a"}]})
    assert not both.ok and "markdown" in str(both.to_json())
