"""R74: a merged banner over a header row is the region's title; a merged area of one
column in the data rows gives its value to each of its rows in the region's reading.

The rules are regions.py's ("Banners and merged data cells"); the shapes that stay
``merged_header`` (a two-row header with merged cells, a merged area beside another in the
banner's row, one narrower than the header, one that starts at the header row) are
checked alongside.  Every workbook is made up here with tests/build.py.
"""

import pytest

from build import WorkbookBuilder

from xlsx_agent import Problem, Workbook
from xlsx_agent.read.regions import FilledSheet, Region, RegionError, UnclearRegion
from xlsx_agent.address import SheetRef


def _open(book: WorkbookBuilder) -> Workbook:
    return Workbook.open(book.to_bytes())


ITEMS = [["Bolt", "B-01", 40, "box"], ["Nut", "B-02", 120, "bag"], ["Washer", "B-03", 75, "bag"],
         ["Hinge", "C-01", 12, "box"]]


def banner(*, gap: bool = False, wide: bool = False) -> WorkbookBuilder:
    """A merged banner over a stock list's header (``gap``: a blank row between;
    ``wide``: the banner runs two columns past the header)."""
    book = WorkbookBuilder()
    sheet = book.sheet("Depot")
    sheet.cell("A1", "Spring stock count")
    sheet.merge("A1:F1" if wide else "A1:D1")
    top = 3 if gap else 2
    sheet.rows(f"A{top}", [["Item", "Bin", "Qty", "Unit"], *ITEMS])
    return book


def rota(*, with_banner: bool = False) -> WorkbookBuilder:
    """A staff rota whose first column, the team, is merged over each team's rows; one
    row of a team has nothing but its team (merged) and a name."""
    book = WorkbookBuilder()
    sheet = book.sheet("Rota")
    top = 2 if with_banner else 1
    if with_banner:
        sheet.cell("A1", "Week 12 rota")
        sheet.merge("A1:C1")
    sheet.rows(f"A{top}", [["Team", "Name", "Shift"],
                           ["North", "Ada", "early"], [None, "Ben", "late"],
                           [None, "Cleo", None],
                           ["South", "Dev", "early"], [None, "Eve", "late"],
                           ["West", "Fay", "night"]])
    sheet.merge(f"A{top + 1}:A{top + 3}")
    sheet.merge(f"A{top + 4}:A{top + 5}")
    sheet.rows("A20", [["Note"], ["see the board"]])
    return book


# -- banners ------------------------------------------------------------------------------------


@pytest.mark.parametrize("gap,wide", [(False, False), (True, False), (False, True)])
def test_a_banner_over_a_header_is_its_title(gap, wide):
    wb = _open(banner(gap=gap, wide=wide))
    header = 3 if gap else 2
    found = wb.find_regions("Depot")
    assert len(found) == 1 and isinstance(found[0], Region)
    region = found[0]
    assert region.title == "Spring stock count"
    assert region.area.first.row == header and region.columns == ("Item", "Bin", "Qty", "Unit")
    assert region.header.ref == f"A{header}:D{header}" and region.rows == 4
    assert region.flags == ()
    assert wb.detect_region(f"t:Depot!A{header}") == region
    caption = wb.detect_region("t:Depot!A1")   # the banner itself: a caption, no region
    assert caption.caption and caption.columns == ("Spring stock count",)
    assert [row[0].v for row in wb.read(f"t:Depot!A{header}[Qty]")] == [40, 120, 75, 12]


def test_a_banner_with_text_beside_it_is_a_header_still():
    wb = _open(_beside())
    with pytest.raises(RegionError) as caught:
        wb.detect_region("t:Depot!A2")
    assert caught.value.reason == "merged_header"
    assert wb.find_regions("Depot")[0] == UnclearRegion("t:Depot!A1", "merged_header")


def _beside() -> WorkbookBuilder:
    book = WorkbookBuilder()
    sheet = book.sheet("Depot")
    sheet.cell("A1", "Spring stock count")
    sheet.merge("A1:C1")
    sheet.cell("D1", "as of May")
    sheet.rows("A2", [["Item", "Bin", "Qty", "Unit"], *ITEMS])
    return book


@pytest.mark.parametrize("shape", ["narrower", "beside_merged", "number", "header_merged"])
def test_merged_areas_that_are_no_banner_stay_unclear(shape):
    book = WorkbookBuilder()
    sheet = book.sheet("Depot")
    sheet.rows("A2", [["Item", "Bin", "Qty", "Unit"], *ITEMS])
    if shape == "narrower":       # over part of the header: a group of its columns
        sheet.cell("B1", "Where")
        sheet.merge("B1:C1")
    elif shape == "beside_merged":   # two merged areas side by side: groups, not a title
        sheet.cell("A1", "Stock")
        sheet.merge("A1:B1")
        sheet.cell("C1", "Count")
        sheet.merge("C1:D1")
    elif shape == "number":
        sheet.cell("A1", 2026)
        sheet.merge("A1:D1")
    else:                         # a merged cell in the header row itself
        sheet.merge("C2:D2")
    wb = _open(book)
    with pytest.raises(RegionError) as caught:
        wb.detect_region("t:Depot!A2")
    assert (caught.value.code, caught.value.reason) == ("header_unclear", "merged_header")


def test_a_two_row_header_with_merged_cells_stays_unclear():
    book = WorkbookBuilder()
    sheet = book.sheet("Returns")
    sheet.rows("A1", [["Item", "Amounts", None], [None, "Qty", "Refund"], ["R-1", 2, 5.5],
                      ["R-2", 1, 2.25]])
    sheet.merge("A1:A2")
    sheet.merge("B1:C1")
    wb = _open(book)
    for cell in ("A1", "B1", "B2"):
        with pytest.raises(RegionError) as caught:
            wb.detect_region(f"t:Returns!{cell}")
        assert caught.value.reason in ("merged_header", "rows_above")
    assert wb.find_regions("Returns") == [UnclearRegion("t:Returns!A1", "merged_header")]


# -- merged data cells ------------------------------------------------------------------------


@pytest.mark.parametrize("with_banner", [False, True])
def test_a_merged_data_column_fills_its_rows_in_the_region_s_reading(with_banner):
    wb = _open(rota(with_banner=with_banner))
    top = 2 if with_banner else 1
    found = wb.find_regions("Rota")
    region = found[0]
    assert isinstance(region, Region) and region.area.ref == f"A{top}:C{top + 6}"
    assert region.title == ("Week 12 rota" if with_banner else None)
    assert region.flags == (Problem("merged_cells_filled", f"Rota!A{top + 1}:A{top + 3}", 2),
                            Problem("merged_cells_filled", f"Rota!A{top + 4}:A{top + 5}", 1))
    teams = [row[0] for row in wb.read(f"t:Rota!A{top}[Team]")]
    assert [v.v for v in teams] == ["North", "North", "North", "South", "South", "West"]
    assert [v.a for v in teams][:3] == [f"Rota!A{top + 1}", f"Rota!A{top + 2}", f"Rota!A{top + 3}"]
    assert all(v.t == "text" and not v.formula for v in teams)
    # by key, a row below the merged area's first: the first's value
    assert wb.cell(f"t:Rota!A{top}/Cleo/Team").v == "North"
    assert wb.cell(f"t:Rota!A{top}/Eve/Team").v == "South"
    assert [r[0].v for r in wb.read(f"t:Rota!A{top}/#3")] == ["North"]
    # the cells themselves, read as cells, are as stored
    assert wb.cell(f"Rota!A{top + 2}").t == "blank"
    team = {c.column: c for c in wb.profile(f"t:Rota!A{top}")}["Team"]
    assert (team.rows, team.filled) == (6, 6)    # profiled as the region reads it
    assert len(found) == 2   # the note below is a region of its own


def test_a_region_sheet_with_fills_is_the_sheet():
    plain = SheetRef(name="Rota")
    filled = FilledSheet(name="Rota", fills=())
    assert filled == plain and plain == filled and hash(filled) == hash(plain)
    assert {plain: 1}[filled] == 1


def test_a_merged_area_from_the_header_row_down_stays_unclear():
    book = WorkbookBuilder()
    sheet = book.sheet("Rota")
    sheet.rows("A1", [["Team", "Name"], [None, "Ada"], ["South", "Dev"]])
    sheet.merge("A1:A2")
    wb = _open(book)
    with pytest.raises(RegionError) as caught:
        wb.detect_region("t:Rota!A1")
    assert caught.value.reason == "merged_header"


def test_a_merged_area_across_columns_in_the_first_data_row_stays_unclear():
    book = WorkbookBuilder()
    sheet = book.sheet("Rota")
    sheet.rows("A1", [["Team", "Name"], ["North", None], ["South", "Dev"]])
    sheet.merge("A2:B2")
    wb = _open(book)
    with pytest.raises(RegionError) as caught:
        wb.detect_region("t:Rota!A1")
    assert caught.value.reason == "merged_header"
