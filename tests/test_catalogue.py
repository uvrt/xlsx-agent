"""R6, the catalogue (``describe``'s result): what it lists, its detail levels, and its token
target on the largest fixture.

Every workbook is built with tests/build.py.  ``warehouse`` is the largest fixture of the
suite: eight wide tables over several sheets, detected and unclear blocks, names, hidden
sheets and problems; its catalogue must stay within the target and say what it cut.  Every
other fixture of the suite is described too, and must fit uncut.
"""

import importlib
import json
from datetime import date, timedelta

import pytest

from build import NS_R, Error, Formula, WorkbookBuilder, zip_parts
from xlsx_agent import Workbook
from xlsx_agent.catalogue import (DEFAULT_TOKENS, SAMPLE_ROWS, SAMPLE_TEXT, catalogue,
                                  estimate_tokens)

MARKER = "MARKER-7Q"  # cell content: allowed in a result, never in a diagnostic

#: Every other workbook the suite builds, by module and function.
FIXTURES = [("test_build", "inventory"), ("test_cells", "_values_book"),
            ("test_cells", "_formulas_book"), ("test_open", "_inventory"),
            ("test_open", "_states"), ("test_open", "_messy"), ("test_profile", "inventory"),
            ("test_regions", "inventory"), ("test_regions", "lists"),
            ("test_regions", "returns"), ("test_regions", "unclear"),
            ("test_regions", "duplicates")]


def _open(book: WorkbookBuilder) -> Workbook:
    return Workbook.open(book.to_bytes())


# -- fixtures ----------------------------------------------------------------------------------


def shop() -> WorkbookBuilder:
    """A small workbook with something of every kind the catalogue lists: a table with a
    validation list, a number stored as text and a formula without a cached value; a hidden,
    protected sheet with a detected block and a name local to it; a merged header; a name
    that is a formula; a full recalculation on load."""
    book = WorkbookBuilder(calc={"fullCalcOnLoad": "1"})
    stock = book.sheet("Stock")
    stock.rows("A1", [["SKU", "Category", "Price", "Added"],
                      ["SKU-0001", "Hardware", 12.5, date(2026, 1, 5)],
                      ["SKU-0002", "Tools", "7.25", date(2026, 1, 6)],
                      ["SKU-0003", "Hardware", 3, date(2026, 2, 1)],
                      ["SKU-0004", "Garden", Formula("C2*2"), date(2026, 3, 1)]])
    stock.table("Products", "A1:D5", ["SKU", "Category", "Price", "Added"], write_header=False)
    stock.validation_list("B2:B5", ["Hardware", "Tools", "Garden"])
    notes = book.sheet("Notes", state="hidden")
    notes.protected = True
    notes.rows("B2", [["Store", "Visits"], ["North", 4], ["South", 6]])
    returns = book.sheet("Returns")
    returns.rows("A1", [["Item", "Amounts", None], [None, "Qty", "Refund"], ["R-1", 2, 5.5]])
    returns.merge("A1:A2")
    returns.merge("B1:C1")
    book.define_name("Rate", "0.21")
    book.define_name("Visits", "Notes!$C$3:$C$4", sheet="Notes")
    return book


#: The warehouse's tables: (sheet, name, top-left); each has :data:`WIDE` columns.
TABLES = [("Stock", "Products", "A1"), ("Stock", "Parts", "A400"), ("Sales", "Orders", "A1"),
          ("Sales", "Returns", "A400"), ("Staff", "People", "A1"), ("Staff", "Shifts", "A400"),
          ("Archive", "Old", "A1"), ("Archive", "Older", "A400")]
WIDE = 24
ROWS = 300
CATEGORIES = ["Hardware", "Tools", "Garden", "Paint", "Lighting", "Plumbing"]


def _value(kind: int, row: int, table: int):
    """Row ``row``'s value in a column of kind ``kind`` (the columns cycle through eight)."""
    if kind == 0:
        return f"K{table}-{row:05d}"
    if kind == 1:
        return CATEGORIES[row % len(CATEGORIES)]
    if kind == 2:
        return round(row * 1.25 + table, 2)
    if kind == 3:
        return date(2025, 1, 1) + timedelta(days=row)
    if kind == 4:
        return f"A note about row {row} of table {table}, longer than a sample shows. {MARKER}"
    if kind == 5:
        return row % 3 == 0
    if kind == 6:
        return Formula(f"C{row}*2", cached=row * 2.5) if row % 7 else Formula(f"C{row}*2")
    return row if row % 50 else "12"  # a number stored as text now and then


def warehouse() -> WorkbookBuilder:
    """The largest fixture: eight tables of :data:`WIDE` columns and :data:`ROWS` rows over
    four sheets (one hidden, one very hidden), a sheet of detected and unclear blocks, names
    over ranges and formulas, a formula without a cached value every seventh row, numbers
    stored as text, long text and errors."""
    book = WorkbookBuilder(calc={"calcMode": "manual"})
    sheets = {"Stock": book.sheet("Stock"), "Sales": book.sheet("Sales"),
              "Staff": book.sheet("Staff", state="hidden"),
              "Archive": book.sheet("Archive", state="veryHidden")}
    sheets["Staff"].protected = True
    for number, (sheet, name, corner) in enumerate(TABLES):
        top = int(corner[1:])
        columns = [f"{name} column {k + 1}" for k in range(WIDE)]
        rows = [[_value(k % 8, r, number) for k in range(WIDE)] for r in range(1, ROWS + 1)]
        rows[5][2] = Error("#DIV/0!")
        sheets[sheet].rows(f"A{top + 1}", rows)
        last = f"{chr(64 + WIDE)}{top + ROWS}"
        sheets[sheet].table(name, f"A{top}:{last}", columns)
    report = book.sheet("Report")
    for block in range(6):
        top = block * 12 + 1
        report.rows(f"A{top}", [[f"Item {block}", "Qty", "Price"]]
                    + [[f"I-{block}-{r}", r, r * 1.5] for r in range(1, 9)])
    report.rows("F1", [["Title over its header"], ["Code", "Qty"], ["C-1", 1]])
    for k in range(10):
        book.define_name(f"Range{k}", f"Report!$B${12 * k + 2}:$B${12 * k + 9}"
                         if k < 6 else f"Stock!$C$2:$C${10 + k}")
        book.define_name(f"Formula{k}", f"SUM(Stock!$C$2:$C${10 + k})")
    return book


def _with_links(book: WorkbookBuilder) -> bytes:
    """``book`` with an external link and a VBA project related to its workbook part."""
    parts = book.parts()
    rels = parts["xl/_rels/workbook.xml.rels"].decode()
    extra = "".join(f'<Relationship Id="rIdX{k}" Type="{NS_R}/{kind}" Target="{target}"/>'
                    for k, (kind, target) in enumerate([
                        ("externalLink", "externalLinks/externalLink1.xml"),
                        ("externalLink", "externalLinks/externalLink2.xml")]))
    extra += ('<Relationship Id="rIdV" Type="http://schemas.microsoft.com/office/2006/'
              'relationships/vbaProject" Target="vbaProject.bin"/>')
    parts["xl/_rels/workbook.xml.rels"] = rels.replace("</Relationships>",
                                                       extra + "</Relationships>").encode()
    return zip_parts(parts)


@pytest.fixture(scope="module")
def large() -> Workbook:
    return _open(warehouse())


# -- what the catalogue lists ------------------------------------------------------------------


def test_the_catalogue_of_a_small_workbook_lists_everything():
    assert _open(shop()).catalogue() == {
        "kind": "xlsx", "date_system": 1900, "flags": {"full_calc_on_load": True},
        "cache": "invalidated",       # it arrived asking for a recalculation
        "sheets": [{"name": "Stock", "id": 1, "used": "A1:D5"},
                   {"name": "Notes", "id": 2, "state": "hidden", "protected": True,
                    "used": "B2:C4"},
                   {"name": "Returns", "id": 3, "used": "A1:C3"}],
        "regions": [
            {"address": "t:Products", "kind": "table", "range": "Stock!A1:D5", "rows": 4,
             "key": "SKU",
             "columns": [
                 {"name": "SKU", "type": "text", "fill": 1.0,
                  "values": [["SKU-0001", 1], ["SKU-0002", 1], ["SKU-0003", 1], ["SKU-0004", 1]],
                  "text_length": [8, 8]},
                 {"name": "Category", "type": "text", "fill": 1.0,
                  "values": [["Hardware", 2], ["Tools", 1], ["Garden", 1]],
                  "text_length": [7, 8],
                  "allowed": [{"values": ["Hardware", "Tools", "Garden"]}]},
                 {"name": "Price", "type": "number", "fill": 0.75, "min": 3, "max": 12.5,
                  "mean": 7.75, "values": [[12.5, 1], ["7.25", 1], [3, 1]],
                  "text_length": [4, 4], "formula_share": 0.25, "also": {"text": 1},
                  "flags": {"no_cached_value": 1, "number_as_text": 1}},
                 {"name": "Added", "type": "date", "fill": 1.0, "min": "2026-01-05",
                  "max": "2026-03-01", "mean": "2026-01-25",
                  "values": [["2026-01-05", 1], ["2026-01-06", 1], ["2026-02-01", 1],
                             ["2026-03-01", 1]]}],
             "sample": [["SKU-0001", "Hardware", 12.5, "2026-01-05"],
                        ["SKU-0002", "Tools", "7.25", "2026-01-06"]]},
            {"address": "t:Notes!B2", "kind": "detected", "range": "Notes!B2:C4", "rows": 2,
             "key_candidates": ["Store", "Visits"],
             "columns": [
                 {"name": "Store", "type": "text", "fill": 1.0,
                  "values": [["North", 1], ["South", 1]], "text_length": [5, 5],
                  "flags": {"hidden": 2}},
                 {"name": "Visits", "type": "number", "fill": 1.0, "min": 4, "max": 6,
                  "mean": 5.0, "values": [[4, 1], [6, 1]], "flags": {"hidden": 2}}],
             "sample": [["North", 4], ["South", 6]]},
            {"address": "t:Notes!Visits", "kind": "name", "range": "Notes!C3:C4", "rows": 2,
             "header_rows": 0, "key": "Column C",
             "columns": [{"name": "Column C", "type": "number", "fill": 1.0, "min": 4,
                          "max": 6, "mean": 5.0, "values": [[4, 1], [6, 1]],
                          "flags": {"hidden": 2}}],
             "sample": [[4], [6]]}],
        "unclear": [{"address": "t:Returns!A1", "code": "header_unclear",
                     "reason": "merged_header"}],
        "names": [{"name": "t:Rate", "not_a_range": "formula"}],
        "problems": [{"code": "no_cached_value", "address": "Stock!C5"},
                     {"code": "number_as_text", "address": "Stock!C3"}]}


def test_the_full_detail_gives_every_columns_whole_profile():
    wb = _open(shop())
    full = wb.catalogue(detail="full")
    products = full["regions"][0]
    assert products["columns"] == [s.as_dict() for s in wb.profile("t:Products")]
    assert "cut" not in full and len(products["sample"]) == SAMPLE_ROWS


def test_workbook_flags_are_listed_never_followed():
    wb = Workbook.open(_with_links(shop()))
    assert wb.catalogue()["flags"] == {"full_calc_on_load": True, "external_links": 2,
                                       "macros": True}
    manual = _open(WorkbookBuilder(calc={"calcMode": "manual"}, protected=True, date1904=True))
    head = manual.catalogue()
    assert (head["flags"], head["protected"], head["date_system"]) == (
        {"manual_calculation": True}, True, 1904)


def test_sample_text_is_cut_and_a_region_without_data_rows_has_no_sample():
    book = WorkbookBuilder()
    sheet = book.sheet("Notes")
    sheet.rows("A1", [["Id", "Text"], [1, "y" * 100]])
    sheet.rows("D1", [["Code", "Name"]])
    wb = _open(book)
    wb.declare_region("Empty", "Notes!D1:E1")
    regions = {r["address"]: r for r in wb.catalogue()["regions"]}
    assert regions["t:Notes!A1"]["sample"] == [[1, "y" * SAMPLE_TEXT + "…"]]
    assert regions["t:Notes!A1"]["columns"][1]["values"] == [["y" * SAMPLE_TEXT + "…", 1]]
    assert regions["t:Empty"]["rows"] == 0 and "sample" not in regions["t:Empty"]


def test_a_title_is_shown_as_its_tables_title_never_as_a_region():
    book = WorkbookBuilder()
    sheet = book.sheet("Report")
    sheet.cell("A1", "T" * 100)
    sheet.rows("A3", [["Ref", "Score"], ["C-1", 3]])
    sheet.cell("A6", "A caption")
    other = book.sheet("Area 1")
    other.cell("A1", "Checks of Area 1")
    other.rows("A3", [["Ref", "Score"], ["C-1", 3]])
    wb = _open(book)
    wb.declare_tables({"tables": [{"name": "checks", "sheets": "Area 1", "header_row": 3}]})
    regions = wb.catalogue()["regions"]
    assert [r["address"] for r in regions] == ["t:Report!A3", "t:checks"]
    assert regions[0]["title"] == "T" * SAMPLE_TEXT + "…"
    assert "title" not in regions[1]


def test_describe_is_the_catalogue_as_text():
    wb = _open(shop())
    assert json.loads(wb.describe()) == wb.catalogue()
    assert json.loads(wb.describe(max_tokens=200)) == wb.catalogue(max_tokens=200)


def test_detail_and_target_are_checked():
    wb = _open(shop())
    with pytest.raises(ValueError):
        wb.catalogue(detail="summary")
    with pytest.raises(ValueError):
        catalogue(wb, max_tokens=0)


# -- the token target --------------------------------------------------------------------------


def test_the_largest_fixtures_catalogue_is_within_its_target(large):
    data = large.catalogue()
    assert estimate_tokens(data) <= DEFAULT_TOKENS
    # every table and region listed; the columns left out counted, the rest named
    addresses = [r["address"] for r in data["regions"]]
    assert addresses[:len(TABLES)] == [f"t:{name}" for _, name, _ in TABLES]
    assert len(addresses) == len(TABLES) + 6 + 10
    assert data["cut"]["columns"] > 0
    assert all(region["columns"] for region in data["regions"][:len(TABLES)])
    for region in data["regions"]:
        columns = region["columns"]
        more = region.get("more_columns", [])
        if isinstance(more, list):
            assert [c["name"] for c in columns] + more == \
                [c.column for c in large.profile(region["address"])]
    # the hidden and very hidden sheets and the stale cache are said
    assert [s.get("state") for s in data["sheets"]] == [None, None, "hidden", "veryHidden", None]
    assert data["flags"] == {"manual_calculation": True}


def test_the_full_detail_of_the_largest_fixture_is_not_cut(large):
    full = large.catalogue(detail="full")
    assert "cut" not in full
    assert sum(len(r["columns"]) for r in full["regions"]) == len(TABLES) * WIDE + 6 * 3 + 10


@pytest.mark.parametrize("target", [3000, 1500, 600, 300])
def test_a_tighter_target_cuts_more_and_still_fits(large, target):
    data = large.catalogue(max_tokens=target)
    assert estimate_tokens(data) <= target
    assert data["cut"]


def test_the_cut_goes_from_samples_to_columns_to_regions(large):
    roomy = large.catalogue(max_tokens=16000)
    narrow = large.catalogue(max_tokens=6000)
    tight = large.catalogue(max_tokens=300)
    assert len(roomy["regions"][0]["sample"]) == 1
    assert roomy["cut"] == {"sample_rows": True, "next": "xlsx_read gives its rows"}
    assert narrow["cut"]["next"] == ("xlsx_profile gives a region's every column; "
                                     "xlsx_read gives its rows")
    products = narrow["regions"][0]
    assert 0 < len(products["columns"]) < WIDE and len(products["sample"]) == 1
    assert [c["name"] for c in products["columns"]] + products["more_columns"] == \
        [f"Products column {k + 1}" for k in range(WIDE)]
    assert "regions" in tight["cut"] and len(tight["regions"]) < len(roomy["regions"])


@pytest.mark.parametrize("module,name", FIXTURES, ids=[f"{m}.{n}" for m, n in FIXTURES])
def test_every_fixture_is_described_uncut(module, name):
    book = getattr(importlib.import_module(module), name)()
    data = _open(book).catalogue()
    assert estimate_tokens(data) <= DEFAULT_TOKENS and "cut" not in data
