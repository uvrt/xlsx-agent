"""R2, "Addresses": a table of forms parses and formats back to itself; quoting round-trips;
invalid forms are refused with a reason.

The workbooks are built with tests/build.py.
"""

import pytest
from ooxml_edit.sheets import MAX_COLUMN, MAX_ROW, Area, Cell

from build import WorkbookBuilder
from xlsx_agent import (FORMS, AddressError, DatasetAddress, NameNotFound, RangeAddress,
                        RefAddress, SheetRef, TableAddress, Workbook, format_address,
                        needs_quotes, parse_address, quote_sheet)

# -- the table of forms ------------------------------------------------------------------------

#: Every form in ROADMAP.md's "Addresses", canonical: (text, class, kind).
CANONICAL = [
    ("Stock!F2", RangeAddress, "cell"),
    ("Stock!A1:H40", RangeAddress, "range"),
    ("Stock!A:A", RangeAddress, "range"),
    ("Stock!A:C", RangeAddress, "range"),
    ("Stock!2:5", RangeAddress, "range"),
    ("Stock!7:7", RangeAddress, "range"),
    ("Stock!XFD1048576", RangeAddress, "cell"),
    ("Stock!A1:XFD1048576", RangeAddress, "range"),
    ("'Q3 sales'!B2:B9", RangeAddress, "range"),
    ("'Bob''s list'!A1", RangeAddress, "cell"),
    ("'Q3 sales'!C:D", RangeAddress, "range"),
    ("#3!F2", RangeAddress, "cell"),
    ("#3!A1:B9", RangeAddress, "range"),
    ("#12!4:4", RangeAddress, "range"),
    ("#0!A1", RangeAddress, "cell"),
    ("Stock!", RangeAddress, "sheet"),
    ("'Q3 sales'!", RangeAddress, "sheet"),
    ("'Bob''s list'!", RangeAddress, "sheet"),
    ("'A1'!", RangeAddress, "sheet"),
    ("#3!", RangeAddress, "sheet"),
    ("t:Products", TableAddress, "table"),
    ("t:Stock!A1", TableAddress, "region"),
    ("t:Stock!Prices", TableAddress, "table"),
    ("t:'Q3 sales'!Rate[Price]", TableAddress, "column"),
    ("t:#3!Rate/#2", TableAddress, "row"),
    ("t:'Q3 sales'!B2", TableAddress, "region"),
    ("t:#3!A1", TableAddress, "region"),
    ("t:Products[Price]", TableAddress, "column"),
    ("t:Products[Unit price]", TableAddress, "column"),
    ("t:Stock!A1[Price]", TableAddress, "column"),
    ("t:Products/SKU-0042/Price", TableAddress, "value"),
    ("t:Products/'A/7'/Price", TableAddress, "value"),
    ("t:Products/'it''s'/Price", TableAddress, "value"),
    ("t:Products/'#5'/Price", TableAddress, "value"),
    ("t:Products/Blue widget/Unit price", TableAddress, "value"),
    ("t:Products/SKU-0042", TableAddress, "row"),
    ("t:Products/#5", TableAddress, "row"),
    ("t:Products/#5/Price", TableAddress, "value"),
    ("t:Products/#9:#16", TableAddress, "rows"),
    ("t:Products/#9:", TableAddress, "rows"),
    ("t:Products/#9:#16/Price", TableAddress, "values"),
    ("t:Products/#9:/*", TableAddress, "rows"),
    ("t:Stock!A1/#2:#3", TableAddress, "rows"),
    ("t:Stock!A1/SKU-0042/Price", TableAddress, "value"),
    ("t:Products/42/Price", TableAddress, "value"),
    ("d:by_category", DatasetAddress, "dataset"),
    ("d:by_category[average]", DatasetAddress, "column"),
    ("d:by_category/Hardware/average", DatasetAddress, "value"),
    ("d:by_category/Hardware", DatasetAddress, "row"),
    ("d:by_category/#2", DatasetAddress, "row"),
    ("d:by_category/#2:#5", DatasetAddress, "rows"),
    ("d:by_category/#2:/average", DatasetAddress, "values"),
    ("$step1", RefAddress, "ref"),
    ("$step1/SKU-0042/Price", RefAddress, "ref"),
]


@pytest.mark.parametrize("text,cls,kind", CANONICAL)
def test_every_form_parses_and_formats_back_to_itself(text, cls, kind):
    address = parse_address(text)
    assert isinstance(address, cls)
    assert address.kind == kind
    assert address.text == text == str(address) == format_address(address)
    assert parse_address(address.text) == address
    assert cls.parse(text) == address


def test_a_whole_sheet_is_the_whole_grid_clipped_to_the_used_range():
    whole = parse_address("'Q3 sales'!")
    assert whole.span == "sheet" and whole.sheet == SheetRef(name="Q3 sales")
    assert whole.area == Area(Cell(1, 1), Cell(MAX_ROW, MAX_COLUMN))
    assert whole.clipped(Area.parse("B2:D9")) == Area.parse("B2:D9")
    assert whole.clipped(None) is None
    assert whole.on(SheetRef(id=3)).text == "#3!"
    assert RangeAddress.parse("#3!").sheet == SheetRef(id=3)
    assert parse_address(" Stock! ").text == "Stock!"
    assert parse_address("'Stock'!").text == "Stock!"
    assert parse_address("Überblick!").text == "'Überblick'!"
    # the same cells as a rectangle are a range, not the sheet
    assert parse_address("Stock!A1:XFD1048576") != parse_address("Stock!")
    with pytest.raises(ValueError, match="whole grid"):
        RangeAddress(SheetRef(name="Stock"), Area.parse("A1:B2"), "sheet")


@pytest.mark.parametrize("name", ["Stock", "Q3 sales", "Bob's list", "A1", "TRUE", "#3",
                                  "R1C1", "it''s", "Ünïcödé 2026"])
def test_a_whole_sheet_round_trips_any_sheet_name(name):
    text = SheetRef(name=name).text + "!"
    address = parse_address(text)
    assert address.kind == "sheet" and address.sheet.name == name and address.text == text


def test_every_example_in_forms_is_canonical():
    for text in FORMS.values():
        assert parse_address(text).text == text


def test_what_the_forms_hold():
    cell = parse_address("Stock!F2")
    assert cell.sheet == SheetRef(name="Stock")
    assert cell.area == Area(Cell(2, 6), Cell(2, 6)) and cell.cell == Cell(2, 6)
    assert parse_address("Stock!A1:H40").area == Area.parse("A1:H40")

    columns = parse_address("Stock!B:C")
    assert columns.span == "columns"
    assert columns.area == Area(Cell(1, 2), Cell(MAX_ROW, 3))
    rows = parse_address("Stock!2:5")
    assert rows.span == "rows"
    assert rows.area == Area(Cell(2, 1), Cell(5, MAX_COLUMN))

    assert parse_address("'Bob''s list'!A1").sheet.name == "Bob's list"
    assert parse_address("#3!F2").sheet == SheetRef(id=3)

    value = parse_address("t:Products/'A/7'/Price")
    assert (value.name, value.key, value.column, value.position) == ("Products", "A/7", "Price", None)
    assert value.region is None and not value.is_region
    row = parse_address("t:Products/#5")
    assert (row.key, row.position, row.column) == (None, 5, None)
    assert parse_address("t:Products/'#5'/Price").key == "#5"
    assert parse_address("t:Products/42/Price").key == "42"  # a key is text until R4 types it

    region = parse_address("t:'Q3 sales'!B2[Price]")
    assert region.is_region and region.name is None
    assert region.region == RangeAddress(SheetRef(name="Q3 sales"), Area.parse("B2"))
    assert region.column == "Price"

    dataset = parse_address("d:by_category/Hardware/average")
    assert (dataset.id, dataset.key, dataset.column) == ("by_category", "Hardware", "average")
    ref = parse_address("$step1/SKU-0042/Price")
    assert (ref.name, ref.rest) == ("step1", "SKU-0042/Price")


# -- normalised forms ----------------------------------------------------------------------------

@pytest.mark.parametrize("text,canonical", [
    ("Stock!$F$2", "Stock!F2"),
    ("Stock!$A$1:$H$40", "Stock!A1:H40"),
    ("Stock!f2", "Stock!F2"),
    ("Stock!H40:A1", "Stock!A1:H40"),
    ("Stock!B1:A2", "Stock!A1:B2"),
    ("Stock!A1:A1", "Stock!A1"),
    ("Stock!$C:$A", "Stock!A:C"),
    ("Stock!$5:$2", "Stock!2:5"),
    ("'Stock'!A1", "Stock!A1"),
    ("Überblick!A1", "'Überblick'!A1"),
    ("A1!B2", "'A1'!B2"),
    ("Q3-sales!A1", "'Q3-sales'!A1"),
    ("  Stock!F2\t", "Stock!F2"),
    ("t:Stock!$A$1", "t:Stock!A1"),
    ("t:'Stock'!A1", "t:Stock!A1"),
    ("t:Products/'SKU-0042'/'Price'", "t:Products/SKU-0042/Price"),
    ("t:Products[#x]", "t:Products['#x]"),
])
def test_other_spellings_normalise(text, canonical):
    address = parse_address(text)
    assert address.text == canonical
    assert parse_address(canonical) == address


# -- quoting ----------------------------------------------------------------------------------

SHEET_NAMES = [
    "Stock", "Sheet1", "_data", "data.2026", "Q3 sales", "Bob's", "'", "''", "'lead", "trail'",
    "it''s", "Hello!", "a!b!c", " padded ", "Überblick", "Zahlen 2026", "数据", "A1", "XFD1048576",
    "ab12", "R1C1", "R", "C", "rc", "R2", "C3", "TRUE", "false", "#3", "3", "3rd", "$x", "t:x",
    "d:x", "a-b", "a+b", "(x)", "a,b", "x[1]", "a/b", "1.5", "-", "tab\there",
]


@pytest.mark.parametrize("name", SHEET_NAMES)
def test_sheet_name_quoting_round_trips(name):
    for where in (Area.parse("B2"), Area.parse("A1:C9")):
        address = RangeAddress(SheetRef(name=name), where)
        again = parse_address(address.text)
        assert again == address
        assert again.sheet.name == name and again.sheet.id is None
        assert again.text == address.text
    region = TableAddress(region=RangeAddress(SheetRef(name=name), Area.parse("A1")),
                          key="K/1", column="C")
    assert parse_address(region.text) == region


def test_quotes_are_left_off_only_where_safe():
    for name in ("Stock", "Sheet1", "_data", "data.2026", "ABCD1", "Sales_Q3"):
        assert not needs_quotes(name) and quote_sheet(name) == name
    for name in ("Q3 sales", "A1", "xfd9", "R1C1", "RC", "c", "TRUE", "True", "3rd", "Überblick",
                 "#3", "a-b", "Bob's"):
        assert needs_quotes(name)
    assert quote_sheet("Bob's") == "'Bob''s'"
    assert quote_sheet("'") == "''''"
    with pytest.raises(ValueError):
        quote_sheet("")


def test_a_quoted_hash_is_a_name_not_an_id():
    assert parse_address("'#3'!A1").sheet == SheetRef(name="#3")
    assert parse_address("#3!A1").sheet == SheetRef(id=3)


KEYS = ["SKU-0042", "A/7", "it's", "''", "#5", "#", " lead", "trail ", "a[1]", "a]b", "x!y",
        "Blue widget", "数据", "42", "a''/'b", "/", "'", "t:Products"]


@pytest.mark.parametrize("key", KEYS)
def test_key_quoting_round_trips(key):
    for address in (TableAddress(name="Products", key=key, column="Price"),
                    TableAddress(name="Products", key=key),
                    TableAddress(name="Products", position=3, column=key),
                    DatasetAddress("by_category", key=key, column=key)):
        again = parse_address(address.text)
        assert again == address and again.text == address.text


COLUMNS = ["Price", "Unit price", "a[1]", "]", "[", "#Totals", "'", "it's", "A/B", "x'#y"]


@pytest.mark.parametrize("column", COLUMNS)
def test_bracketed_column_quoting_round_trips(column):
    for address in (TableAddress(name="Products", column=column),
                    DatasetAddress("by_category", column=column)):
        again = parse_address(address.text)
        assert again == address and again.text == address.text


def test_bracketed_column_escapes_as_excel_structured_references():
    assert TableAddress(name="Products", column="Size [cm]").text == "t:Products[Size '[cm']]"
    assert TableAddress(name="Products", column="#Items").text == "t:Products['#Items]"
    assert TableAddress(name="Products", column="Bob's").text == "t:Products[Bob''s]"


# -- refusals ---------------------------------------------------------------------------------

INVALID = [
    # (text, what the reason says)
    ("", "empty"),
    ("   ", "empty"),
    ("A1", "names its sheet"),
    ("A1:B2", "names its sheet"),
    ("A:A", "names its sheet"),
    ("Products", "t:Name"),
    ("t:Stock!", "a whole sheet has no t: form"),
    ("t:#3!", "a whole sheet has no t: form"),
    ("Stock!!", "not a cell or range"),
    ("!A1", "follows a sheet name"),
    ("Stock!A1:", "not a cell or range"),
    ("Stock!A1:B", "not a cell or range"),
    ("Stock!A", "not a cell or range"),
    ("Stock!R1C1", "not a cell or range"),
    ("Stock!A0", "row 0"),
    ("Stock!A1048577", "outside the sheet's rows"),
    ("Stock!0:3", "row 0"),
    ("Stock!XFE1", "past the last column"),
    ("Stock!A:XFE", "past the last column"),
    ("Stock!A1 B2", "not a cell or range"),
    ("Stock!A1,B2", "not a cell or range"),
    ("Q3 sales!A1", "is quoted"),
    ("Bob's!A1", "is quoted"),
    ("a/b!A1", "is quoted"),
    ("'Q3 sales!A1", "no closing quote"),
    ("''!A1", "never empty"),
    ("'Q3 sales'A1", "followed by !"),
    ("'Q3 sales'", "followed by !"),
    ("#x!A1", "sheet id"),
    ("#!A1", "sheet id"),
    ("#03!A1", "sheet id"),
    ("#-1!A1", "sheet id"),
    ("#4294967296!A1", "at most 4294967295"),
    ("t:", "follows the prefix"),
    ("t:#3", "names its sheet and its header"),
    ("t:Stock!A1:F20", "one cell"),
    ("t:Stock!A:A", "one cell"),
    ("t:Pro ducts", "no spaces"),
    ("t:Products]", "no spaces and none of"),
    ("t:Products[", "no closing ]"),
    ("t:Products[]", "never empty"),
    ("t:Products[Price]x", "ends the address"),
    ("t:Products[Price]/K", "ends the address"),
    ("t:Products[a'b]", "escapes one of"),
    ("t:Products[a[b]", "is written '["),
    ("t:Products/", "never empty"),
    ("t:Products//Price", "never empty"),
    ("t:Products/K/", "never empty"),
    ("t:Products/K/Price/x", "at most a row and a column"),
    ("t:Products/a'b/Price", "is quoted"),
    ("t:Products/'A/7/Price", "no closing quote"),
    ("t:Products/'A/7'x/Price", "followed by / or the end"),
    ("t:Products/''/Price", "never empty"),
    ("t:Products/ K/Price", "space at either end"),
    ("t:Products/#0", "row position"),
    ("t:Products/#05", "row position"),
    ("t:Products/#5x", "row position"),
    ("t:Products/#-1", "row position"),
    ("t:Products/#1048577", "at most 1048576"),
    ("t:Products/#", "row position"),
    ("t:Products/#0:#3", "row range"),
    ("t:Products/#3::", "row range"),
    ("t:Products/#3-#9", "row range"),
    ("t:Products/#3:#1048577", "at most 1048576"),
    ("t:Products/{#3:#5}", "keys, not positions"),
    ("t:Products x", "no spaces"),
    ("t:Products?", None),  # accepted: see below
    ("d:", "follows the prefix"),
    ("d:#1", "does not start with #"),
    ("d:Stock!A1", "none of"),
    ("d:by category", "no spaces"),
    ("$", "a ref is $"),
    ("$Step1", "a ref is $"),
    ("$1step", "a ref is $"),
    ("$step1/", "followed by more"),
    ("Stock", "t:Name"),
]


@pytest.mark.parametrize("text,reason", [case for case in INVALID if case[1] is not None])
def test_invalid_forms_are_refused_with_a_reason(text, reason):
    with pytest.raises(AddressError) as caught:
        parse_address(text)
    error = caught.value
    assert reason in str(error)
    assert error.text == text
    assert error.valid_options, "an error says what would work"


def test_a_table_name_may_hold_what_excel_names_hold():
    assert parse_address("t:Products?").name == "Products?"
    assert parse_address(r"t:\Prices.2026").name == r"\Prices.2026"


def test_a_refusal_never_quotes_a_key():
    for text in ("t:Products/SECRET'MARK/Price", "t:Products/ SECRETMARK/Price",
                 "t:Products/SECRETMARK/Price/x", "t:Stock!A1:B2/SECRETMARK",
                 "t:Products/#SECRETMARK", "t:Products['SECRETMARK]"):
        with pytest.raises(AddressError) as caught:
            parse_address(text)
        assert "SECRETMARK" not in str(caught.value)
        assert not any("SECRETMARK" in option for option in caught.value.valid_options)


def test_a_refusal_points_at_the_character():
    with pytest.raises(AddressError) as caught:
        parse_address("  Stock!A0")
    assert caught.value.position == 9
    assert "at character 9" in str(caught.value)


def test_a_refusal_offers_the_corrected_address_when_there_is_one():
    def options(text):
        with pytest.raises(AddressError) as caught:
            parse_address(text)
        return caught.value.valid_options

    assert options("Q3 sales!B2:B9") == ["'Q3 sales'!B2:B9"]
    assert options("Bob's!A1") == ["'Bob''s'!A1"]
    assert options("t:Stock!A1:F20") == ["t:Stock!A1"]
    assert options("t:'Q3 sales'!B2:B9") == ["t:'Q3 sales'!B2"]
    assert options("Products") == ["t:Products", "d:Products"]
    assert options("A1") == [FORMS["cell"], FORMS["quoted sheet"], FORMS["sheet by id"]]
    # else the forms of the place the error is in
    assert options("Stock!A0") == [FORMS[k] for k in ("cell", "range", "columns", "rows", "sheet",
                                                       "quoted sheet", "sheet by id")]
    assert "t:Products[Price]" in options("t:Products[")
    assert "d:by_category[average]" in options("d:")


def test_a_non_string_is_refused():
    with pytest.raises(AddressError, match="a string"):
        parse_address(3)


def test_the_typed_parsers_refuse_other_forms():
    with pytest.raises(AddressError, match="table address where cell was expected") as caught:
        RangeAddress.parse("t:Products")
    assert "Stock!F2" in caught.value.valid_options
    with pytest.raises(AddressError, match="cell address where table was expected"):
        TableAddress.parse("Stock!A1")
    with pytest.raises(AddressError):
        DatasetAddress.parse("t:Products")
    with pytest.raises(AddressError):
        RefAddress.parse("d:x")


# -- building addresses ------------------------------------------------------------------------

def test_constructors_refuse_what_no_address_can_say():
    with pytest.raises(ValueError):
        SheetRef()
    with pytest.raises(ValueError):
        SheetRef(name="Stock", id=3)
    with pytest.raises(ValueError):
        SheetRef(name="")
    with pytest.raises(ValueError):
        SheetRef(id=-1)
    with pytest.raises(ValueError):
        SheetRef(id=True)
    with pytest.raises(ValueError):
        RangeAddress(SheetRef(name="S"), Area.parse("A1:B2"), "rows")
    with pytest.raises(ValueError):
        RangeAddress(SheetRef(name="S"), Area(Cell(0, 1), Cell(1, 1)))
    with pytest.raises(ValueError):
        TableAddress()
    with pytest.raises(ValueError):
        TableAddress(name="P", region=RangeAddress(SheetRef(name="S"), Area.parse("A1")))
    with pytest.raises(ValueError):
        TableAddress(region=RangeAddress(SheetRef(name="S"), Area.parse("A1:B2")))
    with pytest.raises(ValueError):
        TableAddress(name="Pro ducts")
    with pytest.raises(ValueError):
        TableAddress(name="P", key="K", position=2)
    with pytest.raises(ValueError):
        TableAddress(name="P", key="")
    with pytest.raises(ValueError):
        TableAddress(name="P", position=0)
    with pytest.raises(ValueError):
        DatasetAddress("by/category")
    with pytest.raises(ValueError):
        RefAddress("Step1")


@pytest.mark.parametrize("text, canonical, found", [
    ("t:Products/#16:#9", "t:Products/#9:#16", (9, 16)),
    ("t:Products/#9:16", "t:Products/#9:#16", (9, 16)),
    ("t:Products/#9:#9", "t:Products/#9", None),
    ("t:Products/#9:9/Price", "t:Products/#9/Price", None),
    ("d:by_category/#5:#2", "d:by_category/#2:#5", (2, 5)),
])
def test_a_row_range_normalises(text, canonical, found):
    address = parse_address(text)
    assert address.text == canonical and address.row_range == found


def test_a_row_range_is_a_run_of_positions():
    address = parse_address("t:Products/#9:/Price")
    assert (address.row_range, address.column, address.position) == ((9, None), "Price", None)
    assert TableAddress(name="Products", row_range=(2, 4)).text == "t:Products/#2:#4"
    for bad in [(4, 4), (4, 2), (0, 3), (3,), (True, 4), (2, 1048577)]:
        with pytest.raises(ValueError):
            TableAddress(name="P", row_range=bad)
    with pytest.raises(ValueError):
        TableAddress(name="P", row_range=(2, 4), position=3)


def test_a_range_on_the_other_name_of_its_sheet():
    address = parse_address("'Q3 sales'!B2:B9")
    assert address.on(SheetRef(id=4)).text == "#4!B2:B9"
    assert parse_address("#4!A:B").on(SheetRef(name="Q3 sales")).text == "'Q3 sales'!A:B"


def test_whole_columns_and_rows_are_clipped_to_the_used_range():
    used = Area.parse("B2:F40")
    assert parse_address("Stock!A:C").clipped(used) == Area.parse("B2:C40")
    assert parse_address("Stock!3:5").clipped(used) == Area.parse("B3:F5")
    assert parse_address("Stock!H:H").clipped(used) is None
    assert parse_address("Stock!50:60").clipped(used) is None
    assert parse_address("Stock!A:A").clipped(None) is None
    assert parse_address("Stock!H1:H9").clipped(used) == Area.parse("H1:H9")  # cells as given
    assert parse_address("Stock!H1:H9").clipped(None) == Area.parse("H1:H9")


def test_a_ref_resolves_to_its_address():
    refs = {"step1": "t:Products", "totals": "Stock!F2"}
    assert parse_address("$step1/SKU-0042/Price").resolve(refs) == \
        parse_address("t:Products/SKU-0042/Price")
    assert parse_address("$totals").resolve(refs) == parse_address("Stock!F2")
    with pytest.raises(NameNotFound) as caught:
        parse_address("$step2").resolve(refs)
    assert caught.value.valid_options == ["$step1"]
    assert caught.value.candidates == ["$step1", "$totals"]
    assert caught.value.what == "ref"


# -- sheets in a workbook ------------------------------------------------------------------------

def _workbook() -> Workbook:
    book = WorkbookBuilder()
    book.sheet("Stock", sheet_id=3).cell("A1", "SKU")
    book.sheet("Q3 sales", sheet_id=7).cell("A1", 1)
    book.sheet("Bob's list", sheet_id=1).cell("A1", 1)
    return Workbook.open(book.to_bytes())


def test_an_address_finds_its_sheet_by_name_or_id():
    wb = _workbook()
    assert wb.sheet_of("Stock!F2").id == 3
    assert wb.sheet_of("stock!F2").name == "Stock"  # names match case aside
    assert wb.sheet_of("#3!F2").name == "Stock"
    assert wb.sheet_of("'Q3 SALES'!A:A").id == 7
    assert wb.sheet_of("'Bob''s list'!A1").id == 1
    assert wb.sheet_of("t:#7!A1").name == "Q3 sales"
    assert wb.sheet_of(SheetRef(id=1)).name == "Bob's list"
    assert wb.sheet_of(parse_address("Stock!2:5")).id == 3
    with pytest.raises(ValueError, match="names no sheet"):
        wb.sheet_of("t:Products")


def test_both_names_of_a_sheet():
    wb = _workbook()
    by_name, by_id = wb.sheet_refs(wb.sheet("Q3 sales"))
    assert (by_name.text, by_id.text) == ("'Q3 sales'", "#7")
    assert wb.sheet_of(by_name) == wb.sheet_of(by_id)


def test_an_unknown_sheet_is_not_found_with_the_nearest_names():
    wb = _workbook()
    with pytest.raises(NameNotFound) as caught:
        wb.sheet_of("'Q3 sale'!A1")
    error = caught.value
    assert isinstance(error, KeyError)
    assert error.what == "sheet" and error.name == "Q3 sale"
    assert error.valid_options == ["'Q3 sales'"]
    assert error.candidates == ["Stock", "'Q3 sales'", "'Bob''s list'"]
    assert "Q3 sale" in str(error) and not str(error).startswith('"')

    with pytest.raises(NameNotFound) as caught:
        wb.sheet_of("Nowhere!A1")
    assert caught.value.valid_options == ["Stock", "'Q3 sales'", "'Bob''s list'"]  # none near


def test_an_unknown_sheet_id_is_not_found_with_the_ids():
    wb = _workbook()
    with pytest.raises(NameNotFound) as caught:
        wb.sheet_of("#9!A1")
    assert caught.value.what == "sheet id"
    assert caught.value.valid_options == ["#3", "#7", "#1"]
