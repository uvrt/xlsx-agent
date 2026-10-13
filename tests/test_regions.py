"""R4, "Tables, names, regions": ``t:``, ``t:...[col]`` and ``t:.../key/col`` resolve on
every fixture; a merged header is flagged, never detected; names are matched case aside.

Every workbook is built with tests/build.py; the table parts the builder will not write
(a bad ``ref``, a column count other than the width) are written by hand and put into a
built package with ``zip_parts``.
"""

from datetime import date

import pytest

from build import NS_MAIN, Formula, WorkbookBuilder, zip_parts
from xlsx_agent import NameNotFound, Problem, Workbook
from xlsx_agent.address import AddressError, TableAddress, is_address_name, is_name, parse_address
from xlsx_agent.read.names import range_of
from xlsx_agent.read.regions import Region, RegionError, UnclearRegion

MARKER = "MARKER-7Q"  # cell content that must never reach a diagnostic


def _open(book: WorkbookBuilder) -> Workbook:
    return Workbook.open(book.to_bytes())


# -- fixtures ------------------------------------------------------------------------------------


def inventory() -> WorkbookBuilder:
    """A table with a header and a totals row, a headerless table, and defined names of
    every kind."""
    book = WorkbookBuilder()
    stock = book.sheet("Stock")
    stock.rows("A2", [["SKU-0001", "Bolt", 0.25, date(2026, 1, 5), True],
                      ["SKU-0002", "Nut", 0.1, date(2026, 1, 6), True],
                      ["SKU-0003", "Bolt", 12.5, date(2026, 2, 1), False],
                      ["SKU-0004", MARKER, 3.75, date(2026, 2, 2), True]])
    stock.cell("A6", "Total")
    stock.cell("C6", Formula("SUBTOTAL(109,[Price])", cached=16.6))
    stock.table("Products", "A1:E6", ["SKU", "Item", "Price", "Added", "Active"],
                totals_labels={"SKU": "Total"}, totals_functions={"Price": "sum"})
    sales = book.sheet("Q3 sales", sheet_id=3)
    sales.rows("A1", [[120, 80], [95, 101]])
    sales.table("Sales", "A1:B2", ["North", "South"], header_row=False)
    book.define_name("Prices", "Stock!$C$2:$C$5")
    book.define_name("Rate", f'"{MARKER}"&0.21', sheet="Q3 sales", hidden=True)
    book.define_name("Picks", "Stock!$A$2:$B$3", sheet="Stock")
    book.define_name("Elsewhere", "[1]Other!$A$1")
    book.define_name("Gone", "#REF!")
    book.define_name("Firsts", "'Q3 sales'!$A:$A")
    book.define_name("Twice", "Stock!$A$1,Stock!$C$1")
    book.define_name("_xlnm.Print_Area", "Stock!$A$1:$E$6", sheet="Stock")
    return book


def lists() -> WorkbookBuilder:
    """A plain block with a single-row header (one blank and two duplicate header cells),
    and a second block beside it."""
    book = WorkbookBuilder()
    sheet = book.sheet("Lists")
    sheet.rows("A1", [["Code", "Name", None, "Qty", "qty", None, "Store"],
                      [101, "Bolt", "x", 4, 1, None, "North"],
                      [102, "Nut", "y", 9, 2, None, "South"],
                      [103, MARKER, "z", 1, 3, None, "East"]])
    sheet.rows("A7", [["Code", "Note"], [1, "later"]])
    return book


def returns() -> WorkbookBuilder:
    """A two-row header with merged cells: flagged, then declared."""
    book = WorkbookBuilder()
    sheet = book.sheet("Returns")
    sheet.rows("A1", [["Item", "Amounts", None],
                      [None, "Qty", "Refund"],
                      ["R-1", 2, 5.5],
                      ["R-2", 1, 2.25],
                      ["R-3", 4, 9.0]])
    sheet.merge("A1:A2")
    sheet.merge("B1:C1")
    return book


def unclear() -> WorkbookBuilder:
    """A title over its header, a second header row without merges, a number in a header."""
    book = WorkbookBuilder()
    sheet = book.sheet("Report")
    sheet.rows("A1", [["Stock report"], ["SKU", "Qty"], ["A-1", 4], ["A-2", 5]])
    sheet.rows("E1", [["Item", "Q1", "Q2"], ["Name", "Units", "Units"], ["Bolt", 4, 5],
                      ["Nut", 6, 7]])
    sheet.rows("A10", [["Year", 2025, 2026], ["Bolt", 1, 2]])
    return book


def duplicates() -> WorkbookBuilder:
    """A table whose SKU column repeats a key, and two columns that could each be a key."""
    book = WorkbookBuilder()
    sheet = book.sheet("Dupes")
    sheet.rows("A2", [["K-1", 1, 10], ["K-2", 2, 20], ["K-1", 3, 30]])
    sheet.table("Repeats", "A1:C4", ["SKU", "Line", "Amount"])
    return book


# -- tables --------------------------------------------------------------------------------------


def test_tables_are_read_with_header_totals_columns_and_filter():
    wb = _open(inventory())
    products, sales = wb.tables
    assert (products.kind, products.name, products.sheet, products.area.ref) == (
        "table", "Products", "Stock", "A1:E6")
    assert (products.header_rows, products.totals_rows, products.id) == (1, 1, 1)
    assert products.columns == ("SKU", "Item", "Price", "Added", "Active")
    assert products.totals == (None, None, "sum", None, None)
    assert products.auto_filter.ref == "A1:E5"
    assert products.data.ref == "A2:E5" and products.header.ref == "A1:E1"
    assert products.totals_area.ref == "A6:E6" and products.rows == 4
    assert products.address.text == "t:Products"
    assert (sales.header_rows, sales.columns, sales.data.ref, sales.header) == (
        0, ("North", "South"), "A1:B2", None)
    assert sales.sheet_id == 3 and sales.auto_filter is None


def test_the_display_name_is_the_tables_address_name():
    book = WorkbookBuilder()
    book.sheet("S").rows("A1", [["A"], [1]])
    book.sheets[0].table("Table1", "A1:A2", ["A"], write_header=False, display_name="Orders")
    wb = _open(book)
    (table,) = wb.tables
    assert (table.name, table.object_name) == ("Orders", "Table1")
    assert wb.resolve("t:Orders").range.text == "S!A2"
    with pytest.raises(NameNotFound):
        wb.region("t:Table1")


def test_a_table_resolves_to_its_data_rows_column_row_and_value():
    wb = _open(inventory())
    assert wb.resolve("t:Products").range.text == "Stock!A2:E5"
    column = wb.resolve("t:Products[Price]")
    assert (column.kind, column.range.text, column.column) == ("column", "Stock!C2:C5", "Price")
    value = wb.resolve("t:Products/SKU-0003/Price")
    assert (value.kind, value.range.text, value.row, value.address) == (
        "value", "Stock!C4", 3, "t:Products/SKU-0003/Price")
    assert wb.cell("t:Products/SKU-0003/Price").v == 12.5
    row = wb.resolve("t:Products/#2")
    assert (row.kind, row.range.text, row.row) == ("row", "Stock!A3:E3", 2)
    assert wb.resolve("t:Products/SKU-0004").range.text == "Stock!A5:E5"
    assert [v.v for v in next(iter(wb.read("t:Products/#2")))][:2] == ["SKU-0002", "Nut"]
    assert [r[0].v for r in wb.read("t:Products[Price]")] == [0.25, 0.1, 12.5, 3.75]


def test_names_and_columns_are_matched_case_aside_and_given_as_the_workbook_spells_them():
    wb = _open(inventory())
    selected = wb.resolve("t:PRODUCTS/SKU-0001/price")
    assert selected.range.text == "Stock!C2"
    assert selected.address == "t:Products/SKU-0001/Price"
    assert wb.resolve("t:products[ITEM]").column == "Item"
    assert wb.region("t:stock!picks").name == "Picks"


def test_a_key_is_matched_exactly():
    wb = _open(inventory())
    with pytest.raises(NameNotFound) as caught:
        wb.resolve("t:Products/sku-0001/Price")
    assert caught.value.what == "key" and caught.value.valid_options == []


def test_a_headerless_table_resolves_by_its_own_column_names():
    wb = _open(inventory())
    assert wb.resolve("t:Sales[South]").range.text == "'Q3 sales'!B1:B2"
    assert wb.resolve("t:Sales/120/South", key="North").range.text == "'Q3 sales'!B1"
    assert wb.cell(wb.resolve("t:Sales/95/South", key="North").range).v == 101


def test_the_key_column_is_found_when_only_one_column_can_be():
    wb = _open(inventory())
    assert wb.key_columns("t:Products") == ["SKU"]  # Item repeats; the rest are not keys
    assert wb.regions.key_column(wb.region("t:Products")) == "SKU"


def test_a_key_given_or_declared_wins():
    wb = _open(duplicates())
    assert wb.key_columns("t:Repeats") == ["Line", "Amount"]
    with pytest.raises(RegionError) as caught:
        wb.resolve("t:Repeats/2/Amount")
    assert caught.value.code == "no_key_column"
    assert caught.value.valid_options == ["Line", "Amount"]
    assert wb.resolve("t:Repeats/2/Amount", key="line").range.text == "Dupes!C3"
    table = wb.table("t:Repeats", key="LINE")
    assert table.key == "Line"


def test_a_duplicate_key_is_refused_with_the_rows_never_the_first():
    wb = _open(duplicates())
    with pytest.raises(RegionError) as caught:
        wb.resolve("t:Repeats/K-1/Amount", key="SKU")
    error = caught.value
    assert error.code == "duplicate_key"
    assert error.valid_options == ["t:Repeats/#1", "t:Repeats/#3"]
    assert "K-1" not in str(error)
    assert wb.resolve("t:Repeats/K-2/Amount", key="SKU").range.text == "Dupes!C3"


def test_unknown_tables_columns_keys_and_positions_say_what_would_work():
    wb = _open(inventory())
    with pytest.raises(NameNotFound) as caught:
        wb.resolve("t:Prodcts")
    assert caught.value.what == "table" and caught.value.valid_options[0] == "t:Products"
    assert "t:Prices" in caught.value.candidates
    with pytest.raises(NameNotFound) as caught:
        wb.resolve("t:Products[Prise]")
    assert caught.value.what == "column" and caught.value.valid_options == ["Price"]
    with pytest.raises(NameNotFound) as caught:
        wb.resolve(f"t:Products/{MARKER}X/Price")
    assert caught.value.what == "key"
    with pytest.raises(RegionError) as caught:
        wb.resolve("t:Products/#9")
    assert caught.value.code == "position_out_of_range"
    assert caught.value.valid_options == ["t:Products/#4"]


# -- defined names -------------------------------------------------------------------------------


def test_defined_names_are_listed_with_scope_hidden_and_whether_they_are_ranges():
    wb = _open(inventory())
    by_name = {(n.name, n.scope): n for n in wb.names}
    assert by_name[("Prices", None)].range.text == "Stock!C2:C5"
    rate = by_name[("Rate", "Q3 sales")]
    assert (rate.hidden, rate.range, rate.reason) == (True, None, "formula")
    assert by_name[("Picks", "Stock")].range.text == "Stock!A2:B3"
    assert by_name[("Elsewhere", None)].reason == "external"
    assert by_name[("Gone", None)].reason == "ref_error"
    assert by_name[("Firsts", None)].range.text == "'Q3 sales'!A:A"
    assert by_name[("Twice", None)].reason == "formula"
    area = by_name[("_xlnm.Print_Area", "Stock")]
    assert area.builtin and area.range.text == "Stock!A1:E6"
    assert MARKER not in repr(rate)


def test_a_name_over_a_range_resolves_with_letter_columns_and_no_header():
    wb = _open(inventory())
    prices = wb.region("t:Prices")
    assert (prices.kind, prices.header_rows, prices.columns) == ("name", 0, ("Column C",))
    assert wb.resolve("t:Prices").range.text == "Stock!C2:C5"
    assert wb.resolve("t:Prices[Column C]").range.text == "Stock!C2:C5"
    assert wb.resolve("t:Prices/#2").range.text == "Stock!C3"
    firsts = wb.resolve("t:Firsts")
    assert firsts.range.text == "'Q3 sales'!A1:A2"  # whole columns, clipped to the used range


def test_a_sheet_scoped_name_is_written_with_its_sheet():
    wb = _open(inventory())
    picks = wb.resolve("t:Stock!Picks/SKU-0002/Column B", key="Column A")
    assert picks.range.text == "Stock!B3" and picks.address == "t:Stock!Picks/SKU-0002/Column B"
    assert wb.region("t:#1!Picks").scope == "Stock"
    with pytest.raises(RegionError) as caught:
        wb.region("t:Picks")
    assert caught.value.code == "scoped_name"
    assert caught.value.valid_options == ["t:Stock!Picks"]
    with pytest.raises(NameNotFound):
        wb.region("t:'Q3 sales'!Picks")


def test_a_name_that_is_not_a_range_is_refused_and_never_evaluated():
    wb = _open(inventory())
    for address, reason in [("t:'Q3 sales'!Rate", "formula"), ("t:Elsewhere", "external"),
                            ("t:Gone", "ref_error"), ("t:Twice", "formula")]:
        with pytest.raises(RegionError) as caught:
            wb.resolve(address)
        assert (caught.value.code, caught.value.reason) == ("not_a_range", reason)
        assert MARKER not in str(caught.value)


@pytest.mark.parametrize("formula,expected", [
    ("Stock!$B$2:$B$9", "Stock!B2:B9"),
    ("=Stock!B2", "Stock!B2"),
    ("'Q3 sales'!$A$1:$C$5", "'Q3 sales'!A1:C5"),
    ("stock!$2:$5", "Stock!2:5"),
    ("'Bob''s'!$A$1", "'Bob''s'!A1"),
])
def test_a_names_formula_read_as_a_range(formula, expected):
    sheets = {"stock": "Stock", "q3 sales": "Q3 sales", "bob's": "Bob's"}
    area, reason = range_of(formula, sheets)
    assert reason is None and area.text == expected


@pytest.mark.parametrize("formula,reason", [
    ("0.21", "formula"), ("SUM(Stock!B2:B9)", "formula"), ("Stock!$A$1,Stock!$C$1", "formula"),
    ("Sheet1:Sheet3!$A$1", "formula"), ("Stock!$A$1:$B$2 Stock!$B$1", "formula"),
    ("Products[Price]", "formula"), ("[1]Other!$A$1", "external"),
    ("'[1]Other'!$A$1", "external"), ("#REF!", "ref_error"), ("Stock!#REF!", "ref_error"),
    ("Gone!$A$1", "unknown_sheet"), ("Stock!A1:B", "formula"), ("", "formula"),
])
def test_a_names_formula_that_is_not_a_range(formula, reason):
    assert range_of(formula, {"stock": "Stock"}) == (None, reason)


def test_a_table_and_a_name_with_one_name_are_ambiguous_never_one_picked():
    book = inventory()
    book.define_name("products", "Stock!$A$2:$A$5")
    wb = _open(book)
    with pytest.raises(RegionError) as caught:
        wb.region("t:Products")
    assert caught.value.code == "ambiguous_name"
    assert Problem("duplicate_name", "t:products") in wb.problems


# -- detected regions ----------------------------------------------------------------------------


def test_a_single_row_header_is_detected_and_blank_or_duplicate_headers_named_by_letter():
    wb = _open(lists())
    region = wb.detect_region("t:Lists!A1")
    assert (region.kind, region.name, region.area.ref) == ("detected", None, "A1:E4")
    assert region.columns == ("Code", "Name", "Column C", "Column D", "Column E")
    assert set(region.flags) == {Problem("blank_header", "Lists!C1"),
                                 Problem("duplicate_header", "Lists!D1"),
                                 Problem("duplicate_header", "Lists!E1")}
    assert region.address.text == "t:Lists!A1"
    assert wb.resolve("t:Lists!A1").range.text == "Lists!A2:E4"
    assert wb.resolve("t:Lists!A1[Name]").range.text == "Lists!B2:B4"
    assert wb.key_columns("t:Lists!A1") == ["Code", "Name", "Column C", "Column D", "Column E"]
    assert wb.resolve("t:Lists!A1/102/Column D", key="Code").range.text == "Lists!D3"
    assert wb.resolve("t:Lists!G1/South/Store", key="Store").range.text == "Lists!G3"


def test_find_regions_lists_each_block_of_a_sheet():
    wb = _open(lists())
    found = wb.find_regions("Lists")
    assert [r.address.text for r in found] == ["t:Lists!A1", "t:Lists!G1", "t:Lists!A7"]
    assert all(isinstance(r, Region) for r in found)


def test_find_regions_leaves_tables_out():
    book = inventory()
    book.sheets[0].rows("A9", [["Bin", "Count"], ["B-1", 3]])
    wb = _open(book)
    (region,) = wb.find_regions("Stock")
    assert region.address.text == "t:Stock!A9" and region.area.ref == "A9:B10"


def test_a_merged_header_is_flagged_never_detected():
    wb = _open(returns())
    for cell in ("A1", "B1", "B2"):
        with pytest.raises(RegionError) as caught:
            wb.detect_region(f"t:Returns!{cell}")
        assert caught.value.code == "header_unclear"
        assert caught.value.reason in ("merged_header", "rows_above")
        assert "declare_region" in str(caught.value)
    with pytest.raises(RegionError) as caught:
        wb.resolve("t:Returns!A1[Item]")
    assert (caught.value.code, caught.value.reason) == ("header_unclear", "merged_header")
    assert wb.find_regions("Returns") == [UnclearRegion("t:Returns!A1", "merged_header")]


def test_a_merged_header_is_declared_with_its_names():
    wb = _open(returns())
    with pytest.raises(RegionError) as caught:
        wb.declare_region("Returned", "Returns!A1:C5", header_rows=2)
    assert caught.value.code == "headers_required"
    region = wb.declare_region("Returned", "Returns!A1:C5", header_rows=2,
                               headers=["Item", "Qty", "Refund"], key="item")
    assert (region.kind, region.key, region.data.ref) == ("declared", "Item", "A3:C5")
    assert wb.resolve("t:Returned").range.text == "Returns!A3:C5"
    assert wb.resolve("t:returned[Refund]").range.text == "Returns!C3:C5"
    assert wb.cell("t:Returned/R-2/Refund").v == 2.25
    assert wb.resolve("t:Returned/#3/Qty").range.text == "Returns!B5"
    assert "t:Returned" in wb.regions.known()


def test_a_title_a_second_header_row_and_a_number_in_a_header_are_flagged():
    wb = _open(unclear())
    for address, reason in [("t:Report!A2", "rows_above"), ("t:Report!E1", "multi_row_header"),
                            ("t:Report!A10", "header_not_text"),
                            # the title read as a header: the real header is its first row
                            ("t:Report!A1", "multi_row_header")]:
        with pytest.raises(RegionError) as caught:
            wb.detect_region(address)
        assert (caught.value.code, caught.value.reason) == ("header_unclear", reason)


def test_no_header_at_a_blank_cell():
    wb = _open(lists())
    with pytest.raises(RegionError) as caught:
        wb.detect_region("t:Lists!Z99")
    assert caught.value.code == "no_header"


# -- declared regions ----------------------------------------------------------------------------


def test_a_declared_region_reads_its_one_header_row_or_names_columns_by_letter():
    wb = _open(lists())
    one = wb.declare_region("Codes", "Lists!A1:B4")
    assert one.columns == ("Code", "Name") and one.data.ref == "A2:B4"
    none = wb.declare_region("Raw", "Lists!A2:B4", header_rows=0)
    assert none.columns == ("Column A", "Column B") and none.data.ref == "A2:B4"
    assert wb.resolve("t:Raw/Nut/Column A", key="Column B").range.text == "Lists!A3"
    again = wb.declare_region("RAW", "Lists!A2:A4", header_rows=0)  # declared again: replaced
    assert wb.region("t:Raw") == again


@pytest.mark.parametrize("name", ["TRUE", "false", "A1", "XFD1048576", "R1C1", "R", "C2",
                                  "has space", "1st", "Price?", "a/b", "", "x" * 256])
def test_a_declared_region_needs_a_name_as_ecma_376_defines_one(name):
    wb = _open(lists())
    with pytest.raises(RegionError) as caught:
        wb.declare_region(name, "Lists!A1:B4")
    assert caught.value.code == "bad_name"


def test_a_declared_region_takes_a_name_of_its_own():
    wb = _open(inventory())
    for name in ("products", "Prices"):
        with pytest.raises(RegionError) as caught:
            wb.declare_region(name, "Stock!A1:B3")
        assert caught.value.code == "name_taken"
    wb.declare_region("Picks", "Stock!A1:B3")  # Picks is local to Stock: t:Picks is free
    assert wb.region("t:Picks").kind == "declared"


def test_declared_headers_are_checked():
    wb = _open(lists())
    for headers in (["Code"], ["Code", "code"], ["Code", ""]):
        with pytest.raises(RegionError) as caught:
            wb.declare_region("Codes", "Lists!A1:B4", headers=headers)
        assert caught.value.code == "bad_headers"
    with pytest.raises(RegionError):
        wb.declare_region("Codes", "Lists!A1:B4", header_rows=5)


# -- names as ECMA-376 defines them --------------------------------------------------------------


@pytest.mark.parametrize("name", ["Products", "_private", "\\Prices.2026", "Überblick",
                                  "Stock_2026", "LOGO1000", "XFE1", "AAAA1", "a.b"])
def test_is_name_follows_the_standards_grammar(name):
    assert is_name(name)


@pytest.mark.parametrize("name", ["TRUE", "A1", "$A$1", "xfd1048576", "R1C1", "RC", "1a",
                                  ".a", "a b", "a?", "a\\b", "a-b", "", "x" * 256])
def test_is_name_refuses_what_the_grammar_does_not_allow(name):
    assert not is_name(name)


def test_an_address_takes_names_the_grammar_does_not_but_a_file_may_hold():
    assert is_address_name("Products?") and is_address_name("a-b")
    for name in ("a b", "a[b", "a/b", "a!b", "a'b", "#a", ""):
        assert not is_address_name(name)
    assert parse_address("t:Stock!Prices") == TableAddress(name="Prices", scope=parse_address(
        "Stock!A1").sheet)
    with pytest.raises(AddressError):
        parse_address("t:Stock!Pri ces")


# -- odd table parts -----------------------------------------------------------------------------


def _with_table(book: WorkbookBuilder, body: str) -> bytes:
    parts = book.parts()
    parts["xl/tables/table1.xml"] = (f'<?xml version="1.0" encoding="UTF-8"?>'
                                     f'<table xmlns="{NS_MAIN}" {body}</table>').encode()
    return zip_parts(parts)


def test_a_table_with_too_few_columns_names_the_rest_by_letter():
    book = WorkbookBuilder()
    book.sheet("S").rows("A1", [["A", "B"], [1, 2]])
    book.sheets[0].table("T", "A1:B2", ["A", "B"], write_header=False)
    data = _with_table(book, 'id="1" name="T" displayName="T" ref="A1:B2">'
                             '<tableColumns count="1"><tableColumn id="1" name="A_x0020_1"/>'
                             '</tableColumns>')
    wb = Workbook.open(data)
    (table,) = wb.tables
    assert table.columns == ("A 1", "Column B")
    assert Problem("columns_mismatch", "S!A1:B2") in wb.problems


def test_a_table_whose_ref_is_not_a_range_is_left_out_and_reported():
    book = WorkbookBuilder()
    book.sheet("S").rows("A1", [["A"], [1]])
    book.sheets[0].table("T", "A1:A2", ["A"], write_header=False)
    wb = Workbook.open(_with_table(book, 'id="1" name="T" displayName="T" ref="nowhere">'))
    assert wb.tables == []
    assert Problem("bad_ref", "S") in wb.problems


# -- every fixture -------------------------------------------------------------------------------

#: (fixture, table address, column address, value address, key, the value's cell).
EVERY_FIXTURE = [
    (inventory, "t:Products", "t:Products[Item]", "t:Products/SKU-0002/Item", None, "Stock!B3"),
    (inventory, "t:Sales", "t:Sales[North]", "t:Sales/101/North", "South", "'Q3 sales'!A2"),
    (inventory, "t:Prices", "t:Prices[Column C]", "t:Prices/#4/Column C", None, "Stock!C5"),
    (inventory, "t:Stock!Picks", "t:Stock!Picks[Column A]", "t:Stock!Picks/SKU-0001/Column B",
     "Column A", "Stock!B2"),
    (lists, "t:Lists!A1", "t:Lists!A1[Name]", "t:Lists!A1/103/Column C", "Code", "Lists!C4"),
    (lists, "t:Lists!A7", "t:Lists!A7[Note]", "t:Lists!A7/1/Note", "Code", "Lists!B8"),
    (duplicates, "t:Repeats", "t:Repeats[Amount]", "t:Repeats/3/Amount", "Line", "Dupes!C4"),
]


@pytest.mark.parametrize("fixture,table,column,value,key,cell", EVERY_FIXTURE)
def test_t_and_column_and_key_column_resolve_on_every_fixture(fixture, table, column, value,
                                                               key, cell):
    wb = _open(fixture())
    whole = wb.resolve(table)
    assert whole.range is not None and whole.region.data is not None
    picked = wb.resolve(column)
    assert picked.kind == "column" and picked.range.area.first.column == \
        picked.range.area.last.column
    found = wb.resolve(value, key=key)
    assert found.kind == "value" and found.range.text == cell
    assert wb.cell(found.range) == wb.cell(cell)
    if key is None:
        assert wb.cell(value) == wb.cell(cell)


def test_the_returns_fixture_resolves_once_declared():
    wb = _open(returns())
    wb.declare_region("Returned", "Returns!A1:C5", header_rows=2, headers=["Item", "Qty", "Refund"])
    assert wb.resolve("t:Returned/R-3/Qty", key="Item").range.text == "Returns!B5"


# -- no content in diagnostics -------------------------------------------------------------------


def test_no_cell_text_or_formula_reaches_a_diagnostic():
    messages = []
    attempts = [(inventory, f"t:Products/{MARKER}/Price"), (inventory, "t:'Q3 sales'!Rate"),
                (lists, f"t:Lists!A1/{MARKER}X/Code"), (returns, "t:Returns!A1"),
                (duplicates, f"t:Repeats/{MARKER}")]
    for fixture, address in attempts:
        wb = _open(fixture())
        try:
            wb.resolve(address, key="Name" if "Lists" in address else None)
        except (RegionError, NameNotFound) as error:
            messages.append(str(error))
            messages.extend(getattr(error, "valid_options", []))
            messages.extend(getattr(error, "candidates", []))
        else:
            raise AssertionError(f"{address} resolved")
        messages.extend(repr(p) for p in wb.problems)
    assert messages and not any(MARKER in message for message in messages)
