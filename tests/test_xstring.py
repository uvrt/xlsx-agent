"""Characters XML cannot carry (production R63b): written ``_xHHHH_`` as Excel writes them.

What Excel for Mac 16.106 wrote for text typed into cells through its object model (round 44,
lane S's probe): ``a<U+0001>b`` as ``a_x0001_b``, a carriage return as ``_x000D_``, the
literal text ``_x0041_`` as ``_x005F_x0041_`` (and ``_xabcd_`` too, ``_X0041_`` and ``_x00``
as they are), a formula's string result ``q_x0001_r``, a formula's text ``"_x005F_x0041_"``.
Before this, ``set_cells`` raised lxml's own ``ValueError``.
"""

from __future__ import annotations

import io
import re

import pytest

from ooxml_edit.tools import Toolbox

from xlsx_agent import Workbook
from xlsx_agent.edit import EditError
from xlsx_agent.read.strings import decode_xstring, encode_xstring

from referents import unzipped
from shop import Shop

#: (text, as Excel stored it): measured.
MEASURED = [
    ("a\x01b", "a_x0001_b"), ("c\x1fd", "c_x001F_d"), ("e\rf", "e_x000D_f"),
    ("o\x0bp", "o_x000B_p"), ("_x0041_", "_x005F_x0041_"), ("k_x0041_l", "k_x005F_x0041_l"),
    ("_x005F_", "_x005F_x005F_"), ("_xabcd_", "_x005F_xabcd_"), ("_x00", "_x00"),
    ("x_y", "x_y"), ("_X0041_", "_X0041_"), ("g\th", "g\th"), ("i\nj", "i\nj"),
]


@pytest.mark.parametrize("text, stored", MEASURED)
def test_text_is_stored_as_excel_stores_it_and_read_back(text, stored):
    assert encode_xstring(text) == stored
    assert decode_xstring(stored) == text


@pytest.mark.parametrize("shared", [True, False])
def test_set_cells_writes_what_xml_cannot_carry_and_reads_it_back(shared):
    wb = Workbook.open(Shop(shared=shared).to_bytes())
    cells = {"Totals!A20": "a\x01b", "Totals!A21": "_x0041_", "Totals!A22": "e\rf",
             "Totals!A23": "x\uffff\x00y", "Totals!A24": "plain",
             "Totals!A25": {"markdown": "**a\x02** b"},
             "Totals!B20": {"formula": '="_x0041_"&CHAR(1)'}}
    written = wb.set_cells(cells)
    assert set(written.escaped) == set(cells) - {"Totals!A24"}
    assert written.as_dict()["escaped"]
    report = wb.save(io.BytesIO())
    assert report.verified == len(cells)
    again = Workbook.open(wb.to_bytes())
    for address, value in cells.items():
        if isinstance(value, str):
            assert again.cell(address).v == value
    assert again.cell("Totals!A25").v == "a\x02 b"
    assert again.cell("Totals!B20").formula_text == '"_x0041_"&CHAR(1)'
    parts = unzipped(wb.to_bytes())
    stored = parts["xl/sharedStrings.xml" if shared else "xl/worksheets/sheet3.xml"]
    assert b"a_x0001_b" in stored and b"_x005F_x0041_" in stored and b"e_x000D_f" in stored
    assert b'"_x005F_x0041_"&amp;CHAR(1)' in parts["xl/worksheets/sheet3.xml"]


def test_an_expectation_compares_the_text_as_read():
    wb = Workbook.open(Shop().to_bytes())
    wb.set_cells({"Totals!A20": "a\x01b"})
    wb.set_cells({"Totals!A20": {"value": "new", "expect": {"value": "a\x01b"}}})
    wb.set_cells({"Totals!A20": {"value": "a\x01b"}})
    wb.set_cells({"Totals!A20": {"value": "z", "expect": {"shown": "a\x01b"}}})
    assert wb.cell("Totals!A20").v == "z"


def test_text_too_long_once_escaped_is_refused_naming_why():
    wb = Workbook.open(Shop().to_bytes())
    with pytest.raises(EditError) as raised:
        wb.set_cells({"Totals!A20": "\x01" * 5000})
    assert raised.value.code == "text_too_long" and "_xHHHH_" in str(raised.value)


def test_a_formula_s_cached_string_and_text_are_read_decoded():
    found = Shop().parts()
    body = found["xl/worksheets/sheet3.xml"].decode()
    body = re.sub(r"</sheetData>", '<row r="9"><c r="A9" t="str"><f>"q"&amp;CHAR(1)&amp;"r"</f>'
                  '<v>q_x0001_r</v></c><c r="B9" t="str"><f>"_x005F_x0041_"</f>'
                  '<v>_x005F_x0041_</v></c></row></sheetData>', body)
    found["xl/worksheets/sheet3.xml"] = body.encode()
    from referents import zipped

    wb = Workbook.open(zipped(found))
    assert wb.cell("Totals!A9").v == "q\x01r"
    assert wb.cell("Totals!B9").v == "_x0041_"
    assert wb.cell("Totals!B9").formula_text == '"_x0041_"'


def test_the_tool_writes_control_characters_and_says_so():
    from xlsx_agent.tools import FORMAT, GROUPS, TOOLS

    toolbox = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS)
    session = toolbox.session()
    blob = session.add_blob(Shop().to_bytes(), name="shop.xlsx")
    assert toolbox.dispatch(session, "open_document", {"blob": blob}).ok
    result = toolbox.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "Totals!A20", "text": "a\x01b"}]})
    assert result.ok and result.data["escaped"] == ["Totals!A20"]


def test_a_table_header_holding_a_control_character_is_escaped_too():
    wb = Workbook.new()
    wb.write_table("Sheet1", "A1", ["a\x01b", "c"], [[1, 2]], as_table=True)
    assert wb.validate() == []
    again = Workbook.open(wb.to_bytes())
    assert again.tables[0].columns[0] == "a\x01b"
    assert again.cell("Sheet1!A1").v == "a\x01b"
    assert b'name="a_x0001_b"' in unzipped(wb.to_bytes())["xl/tables/table1.xml"]
