"""Follow-up C of round 18: sheet-name quoting, measured in Excel.

Excel for Mac 16.106 on macOS, 2026-10-27, through tests/oracle.py: 253 made-up sheet names,
each the target of a formula ``='Name'!A1`` written into a workbook Excel opened and saved,
and also set through Excel's object model and read back.  These lists are what Excel's
saved ``<f>`` text wrote.  Excel read the quoted form of every one of them, giving the
referred cell's value -- which is why a name may be quoted where Excel leaves it bare
(``IRREGULAR``), but never bare where Excel quotes it.
"""

import pytest

from ooxml_edit.sheets import Area

from xlsx_agent import (RangeAddress, SheetRef, canonical_address, needs_quotes, parse_address,
                        quote_sheet)
from xlsx_agent.formula import tokens

#: Excel wrote these bare, and so does an address.
BARE = [
    'Stock', 'Sheet1', 'Q3_sales', 'Sales2026', 'data.2026', 'a.b', '_data', '_1', '_A1',
    'ABCD1', 'XFE1', 'A1048577', 'AB0', 'RR', 'CC', 'true1', 'Überblick', 'café', '数据',
    'Δέλτα', 'Ωmega', '日本1', 'Rx', 'RCx', 'A1x', 'A1_', 'AB12c', 'XFD9x', 'a1b', 'R0', 'C0',
    'R1048577', 'ABC', 'XFD', 'XFE', 'A', 'Z', 'AA', 'a_.b', 'x_.y', '__', '_', 'a.', 'Ä',
    'aµb', 'aªb', 'a٣b', 'ﬁx', 'e\u0301', 'été1', 'Ä1', 'R0x', 'C0x', 'R1048577x', 'C16385x',
    'C99999999', 'R12345678', 'RCC', 'a.A1', 'a._', '_a', '_a_', 'x._y', 'a.1', 'TRUEx',
    'FALSE_', 'xTRUE', 'Привет', 'שלום', 'مرحبا', 'नमस\u094dत\u0947', 'สว\u0e31สด\u0e35',
    '한국', 'ひらがな', 'カタカナ', 'Ａｂｃ', 'a１b', 'ǅx', 'ʰx', 'Բար', 'ქარ', 'ሰላም', 'தம\u0bbfழ\u0bcd',
    '𠀀x', 'ß', 'ıx', 'øæ',
]

#: Excel wrote these quoted.
QUOTED = [
    'TOT2026', '3rd', '2024', '1.5', '1a', 'A1', 'XFD9', 'ab12', 'R', 'C', 'R1C1', 'RC', 'R2',
    'C3', 'RC1', 'R1C', 'R1C1x', 'TRUE', 'False', 'Q3 sales', "Bob's", 'a-b', 'a+b', 'a&b',
    '(x)', '#3', 'a,b', 'a;b', 'a=b', 'a%b', 'a$b', 'a@b', 'a~b', 'a^b', '.hidden', '.5',
    '_.x', '\xa0x', 'a\xa0b', 'xfd9', 'r', 'c', 'rc', 'true', 'Abc1', 'R1x', 'C1x', 'A1.b',
    'R1C1_', 'R1C1.x', 'R2C3D', 'R1048576', 'C16384', 'C16385', 'R1C16385', '_.', '_a.b',
    '_._', 'a..b', '_x.', '€x', 'a©b', 'a\u3000b', 'a、b', 'a¬b', '٣a', 'Ab\xa0', 'R1048576x',
    'C99999', 'C65536', 'C65537', 'C1000000', 'R1Cx', 'RC1x', 'R1R', 'C1C', 'CR1', 'R1a',
    'C1a', 'R1_', 'R1.x', 'Rx1', 'R01', 'C01', 'R1C01', 'R1C1C1', 'A01', 'A1.', 'A1.2',
    'AB12.c', 'XFE1.b', 'Z9.z', 'AB0.c', 'a.b.c', '.', 'a...b', '__.', '_1.2', 'Q3.sales',
    '1_', 'TRUE.x', '\u0301a', 'a®b', 'a¢b', 'a£b', 'a¥b', '£x', '¥x', 'a«b', 'a»b', 'a‘b',
    'a’b', 'a„b', 'a•b', 'a¦b', 'a，b', 'a；b', 'a（b', '©x', 'a¬', '¬a', 'a©', 'é©', '。x', '、x',
]

#: Excel wrote these bare, but quotes others of their kind (``a©b``, ``a£b``, ``a«b``, ``€x``,
#: ``a，b``): non-ASCII symbols and punctuation, and names starting with something not a
#: letter.  Quoted here, which Excel reads.
IRREGULAR = [
    'x²', 'a·b', '²x', '·x', 'a€b', 'a§b', 'a…b', 'a—b', 'a“b', 'a×b', 'a÷b', 'a¿b', 'a°b',
    'a±b', 'a😀b', 'a。b', 'Ⅻ', 'a\u200bb', '１ab', 'ⓐx', 'a☃b', 'a™b', 'a‰b', 'a€', 'a¤b',
    'a†b', 'a¶b', 'a¨b', 'a¯b', 'a´b', 'a¸b', 'a¼b', 'a\u2002b', 'a\xadb', 'a−b', '§x', '…x',
    '×x', '°x', 'a¡b', 'a·',
]


def test_the_lists_are_the_measured_names():
    assert len(BARE) + len(QUOTED) + len(IRREGULAR) == 253
    assert len(set(BARE) | set(QUOTED) | set(IRREGULAR)) == 253


@pytest.mark.parametrize("name", BARE)
def test_a_name_excel_writes_bare_is_bare(name):
    assert not needs_quotes(name)
    assert quote_sheet(name) == name
    assert canonical_address(f"'{name}'!b4") == f"{name}!B4"


@pytest.mark.parametrize("name", QUOTED + IRREGULAR)
def test_a_name_excel_quotes_is_quoted(name):
    assert needs_quotes(name)
    written = "'" + name.replace("'", "''") + "'"
    assert quote_sheet(name) == written
    assert canonical_address(f"{written}!b4") == f"{written}!B4"


@pytest.mark.parametrize("name", BARE + QUOTED + IRREGULAR)
def test_every_measured_name_round_trips(name):
    address = RangeAddress(SheetRef(name=name), Area.parse("A1:B2"))
    assert parse_address(address.text) == address


@pytest.mark.parametrize("name", BARE + IRREGULAR)
def test_the_formula_lexer_reads_a_bare_name_as_excel_writes_it(name):
    found = list(tokens(f"{name}!A1+1"))
    assert found[0].text == f"{name}!A1"


def test_compatibility_forms_are_read_as_excel_would():
    # Not measured: a full-width cell reference, boolean or R1C1 reference is quoted, as
    # its compatibility form (NFKC) is one.  Excel compares sheet names in some such form:
    # a workbook with sheets a¹b and a１b did not open (nor one with a\u00adb and a\ufeffb).
    for name in ("\uff21\uff11", "\uff34\uff32\uff35\uff25", "\uff32\uff11\uff23\uff11"):
        assert needs_quotes(name)
