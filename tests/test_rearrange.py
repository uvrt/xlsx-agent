"""Putting a sheet's columns in an order (R59, round 46: R70, R75).

* The reported bugs, with made-up names: a mapping or an order whose columns are not next to
  each other put them in the wrong places and the result said otherwise (R70).
* A property test: random orders, mappings and drops on random made-up sheets, held to a
  pure model of what the call means -- every cell, formula (in cells, on another sheet, in
  defined names, in data validations and conditional formats), merged area and column
  width where the model puts it, and the result saying so; a range across moved columns
  that now covers other cells is listed in ``reshaped``.
* A table's columns reordered as Excel's drag of a table column does (R75): see
  ``test_table_columns.py`` for what Excel did, cell by cell.

Tool design rules: 1 (an order or a mapping, letters or numbers, mean the same), 2 (a
refusal names the table and its columns, and what to do), 3 (an order already standing is
no change), 4 (``reshaped`` is capped and counted), 5 (one call, one undo step).
"""

from __future__ import annotations

import io
import random
import re
import zipfile
from xml.sax.saxutils import escape

import pytest
from ooxml_edit.sheets import column_index, column_letters

from xlsx_agent import Workbook
from xlsx_agent.edit import EditError

from structure_book import _DECL, _SML, R, X, StructureBook, _rels
from structure_facts import facts
from table_columns_book import TableColumnsBook


# -- a made-up grid ----------------------------------------------------------------------------


def _inline(ref: str, text: str) -> str:
    return f'<c r="{ref}" t="inlineStr"><is><t>{escape(text)}</t></is></c>'


def _f(ref: str, text: str) -> str:
    return f'<c r="{ref}"><f>{escape(text)}</f><v>0</v></c>'


class Grid:
    """A sheet ``Grid`` of ``width`` columns: rows 1-6 hold text naming each cell's own
    column and row (``c5r2``), row 7 a formula in each column reading another column's
    cell, row 8 one counting another column's cells; vertical merged areas (rows 9:10),
    data validations and conditional formats whose formulas read another column, a width
    for every column.  ``Refs`` reads Grid's cells by every reference form, and defined
    names a cell each, and a chart on Refs plots two of Grid's columns.  ``spans`` are
    ranges across several columns."""

    def __init__(self, rng: random.Random, width: int) -> None:
        self.width = width
        cols = range(1, width + 1)
        self.reads = {c: rng.choice(cols) for c in cols}             # row 7: =<col>2
        self.counts = {c: rng.choice(cols) for c in cols}            # row 8: COUNTA(..)
        self.merged = sorted(rng.sample(list(cols), k=min(3, width)))
        self.checked = {c: rng.choice(cols) for c in rng.sample(list(cols), k=min(3, width))}
        self.shaded = {c: rng.choice(cols) for c in rng.sample(list(cols), k=min(3, width))}
        self.refs = [rng.choice(cols) for _ in range(6)]
        self.names = [rng.choice(cols) for _ in range(3)]
        self.series = [rng.choice(cols) for _ in range(2)]
        spans = []
        for _ in range(4):
            a, b = sorted(rng.sample(list(cols), 2))
            spans.append((a, b))
        self.spans = spans

    def grid(self) -> str:
        L = column_letters
        rows = []
        for r in range(1, 7):
            rows.append(f'<row r="{r}">' + "".join(_inline(f"{L(c)}{r}", f"c{c}r{r}")
                                                   for c in range(1, self.width + 1))
                        + "</row>")
        rows.append('<row r="7">' + "".join(_f(f"{L(c)}7", f"{L(self.reads[c])}2")
                                            for c in range(1, self.width + 1)) + "</row>")
        rows.append('<row r="8">' + "".join(
            _f(f"{L(c)}8", f"COUNTA({L(self.counts[c])}1:{L(self.counts[c])}6)")
            for c in range(1, self.width + 1)) + "</row>")
        rows.append('<row r="9">' + "".join(_inline(f"{L(c)}9", f"m{c}")
                                            for c in self.merged) + "</row>")
        cols = "".join(f'<col min="{c}" max="{c}" width="{8 + c / 4}" customWidth="1"/>'
                       for c in range(1, self.width + 1))
        merges = "".join(f'<mergeCell ref="{L(c)}9:{L(c)}10"/>' for c in self.merged)
        formats = "".join(
            f'<conditionalFormatting sqref="{L(c)}1:{L(c)}6"><cfRule type="expression" '
            f'dxfId="0" priority="{n + 1}"><formula>{L(y)}1&lt;&gt;""</formula></cfRule>'
            '</conditionalFormatting>' for n, (c, y) in enumerate(sorted(self.shaded.items())))
        checks = "".join(
            f'<dataValidation type="custom" allowBlank="1" sqref="{L(c)}1:{L(c)}6">'
            f'<formula1>LEN({L(y)}1)&gt;0</formula1></dataValidation>'
            for c, y in sorted(self.checked.items()))
        return (_DECL + f'<worksheet xmlns="{X}" xmlns:r="{R}">'
                f'<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
                f'<sheetFormatPr defaultRowHeight="15"/><cols>{cols}</cols>'
                f'<sheetData>{"".join(rows)}</sheetData>'
                f'<mergeCells count="{len(self.merged)}">{merges}</mergeCells>{formats}'
                f'<dataValidations count="{len(self.checked)}">{checks}</dataValidations>'
                '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" '
                'footer="0.3"/></worksheet>')

    def ref_formulas(self) -> dict[str, str]:
        L = column_letters
        out = {}
        for n, c in enumerate(self.refs):
            out[f"A{n + 1}"] = f"Grid!{L(c)}3"
            out[f"B{n + 1}"] = f"Grid!${L(c)}$4"
            out[f"C{n + 1}"] = f"SUM(Grid!{L(c)}1:{L(c)}6)"
            out[f"D{n + 1}"] = f"COUNTA(Grid!{L(c)}:{L(c)})"
        for n, (a, b) in enumerate(self.spans):
            out[f"E{n + 1}"] = f"COUNTA(Grid!{L(a)}2:{L(b)}2)"
        return out

    def to_bytes(self) -> bytes:
        refs = self.ref_formulas()
        rows: dict[int, list[str]] = {}
        for ref, text in refs.items():
            rows.setdefault(int(ref[1:]), []).append(_f(ref, text))
        body = "".join(f'<row r="{n}">{"".join(sorted(cells, key=lambda c: c[6:8]))}</row>'
                       for n, cells in sorted(rows.items()))
        sheet2 = (_DECL + f'<worksheet xmlns="{X}" xmlns:r="{R}"><sheetData>{body}'
                  '</sheetData><drawing r:id="rId1"/></worksheet>')
        L = column_letters
        chart = StructureBook().chart_xml() \
            .replace("Stock!$B$1", f"Grid!${L(self.series[0])}$1") \
            .replace("Stock!$A$2:$A$7", f"Grid!${L(self.series[0])}$2:${L(self.series[0])}$6") \
            .replace("Stock!$B$2:$B$7", f"Grid!${L(self.series[1])}$2:${L(self.series[1])}$6")
        names = "".join(f'<definedName name="Spot{n}">Grid!${column_letters(c)}$5'
                        '</definedName>' for n, c in enumerate(self.names))
        workbook = (_DECL + f'<workbook xmlns="{X}" xmlns:r="{R}"><sheets>'
                    '<sheet name="Grid" sheetId="1" r:id="rId1"/>'
                    '<sheet name="Refs" sheetId="2" r:id="rId2"/></sheets>'
                    f'<definedNames>{names}</definedNames><calcPr calcId="191029"/>'
                    '</workbook>')
        parts = {
            "[Content_Types].xml": (
                _DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
                'content-types"><Default Extension="rels" ContentType="application/'
                'vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                + "".join(f'<Override PartName="{name}" ContentType="{_SML}{kind}"/>'
                          for name, kind in (
                              ("/xl/workbook.xml", "spreadsheetml.sheet.main+xml"),
                              ("/xl/styles.xml", "spreadsheetml.styles+xml"),
                              ("/xl/worksheets/sheet1.xml", "spreadsheetml.worksheet+xml"),
                              ("/xl/worksheets/sheet2.xml", "spreadsheetml.worksheet+xml"),
                              ("/xl/drawings/drawing1.xml", "drawing+xml"),
                              ("/xl/charts/chart1.xml", "drawingml.chart+xml")))
                + "</Types>"),
            "_rels/.rels": _rels([("rId1", "officeDocument", "xl/workbook.xml")]),
            "xl/workbook.xml": workbook,
            "xl/_rels/workbook.xml.rels": _rels([
                ("rId1", "worksheet", "worksheets/sheet1.xml"),
                ("rId2", "worksheet", "worksheets/sheet2.xml"),
                ("rId3", "styles", "styles.xml")]),
            "xl/worksheets/sheet1.xml": self.grid(),
            "xl/worksheets/sheet2.xml": sheet2,
            "xl/worksheets/_rels/sheet2.xml.rels": _rels(
                [("rId1", "drawing", "../drawings/drawing1.xml")]),
            "xl/drawings/drawing1.xml": StructureBook().drawing(),
            "xl/drawings/_rels/drawing1.xml.rels": _rels(
                [("rId1", "chart", "../charts/chart1.xml")]),
            "xl/charts/chart1.xml": chart,
            "xl/styles.xml": StructureBook().styles(),
        }
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, data in parts.items():
                archive.writestr(name, data.encode("utf-8"))
        return buffer.getvalue()


def _formula(wb: Workbook, address: str) -> str:
    sheet, _, ref = address.partition("!")
    return facts(wb.to_bytes())["sheets"][sheet]["cells"][ref][0]


def _header(wb: Workbook, sheet: str, width: int) -> str:
    return " ".join(str(wb.cell(f"{sheet}!{column_letters(c)}1").v)
                    for c in range(1, width + 1))


# -- R70: the reported cases ----------------------------------------------------------------


@pytest.fixture()
def plain() -> Workbook:
    wb = Workbook.new(["Rota"])
    wb.set_cells({f"Rota!{column_letters(c)}1": f"h{c}" for c in range(1, 12)}
                 | {f"Rota!{column_letters(c)}2": c for c in range(1, 12)}
                 | {"Rota!L1": {"formula": "=C2*10"}, "Rota!L2": {"formula": "=SUM(I2,K2)"}})
    return wb


def test_a_mapping_swaps_columns_apart(plain):
    done = plain.rearrange_columns("Rota", {"C": "A", "A": "C"})
    assert _header(plain, "Rota", 4) == "h3 h2 h1 h4"          # c, b, a -- not c, a, b
    assert done.columns == {"A": "C", "C": "A"}
    assert _formula(plain, "Rota!L1") == "A2*10"                    # the reference follows C


def test_an_order_swaps_columns_apart(plain):
    done = plain.rearrange_columns("Rota", ["K", "I"])
    assert _header(plain, "Rota", 11).split()[7:] == ["h8", "h11", "h10", "h9"]
    assert done.columns == {"I": "K", "K": "I"}
    assert _formula(plain, "Rota!L2") == "SUM(K2,I2)"


def test_a_cycle_of_three_apart(plain):
    done = plain.rearrange_columns("Rota", {"A": "E", "E": "C", "C": "A"})
    assert _header(plain, "Rota", 5) == "h3 h2 h5 h4 h1"
    assert done.columns == {"A": "E", "C": "A", "E": "C"}


def test_the_report_is_what_the_steps_did(plain):
    done = plain.rearrange_columns("Rota", ["F", "B", "J"], drop=["D"])
    # after D goes: E..K are D..J; F -> B's place, B -> F's (now E), J -> J's (now I)
    header = _header(plain, "Rota", 10).split()
    for old, new in done.columns.items():
        if column_index(old) <= 11:                         # L holds formulas, now K
            assert header[column_index(new) - 1] == f"h{column_index(old)}"
    assert done.columns["L"] == "K"
    assert done.dropped == ("D",) and done.columns["F"] == "B"


# -- the property test ------------------------------------------------------------------------


def _model(width: int, order: list[int], dropped: list[int]) -> dict[int, int | None]:
    """Where each column's cells go: dropped first, then each listed column into the place
    of the listed column as many from the left, the others staying."""
    goes = dict(zip(order, sorted(order)))
    out: dict[int, int | None] = {}
    for c in range(1, width + 1):
        if c in dropped:
            out[c] = None
            continue
        target = goes.get(c, c)
        out[c] = target - sum(1 for d in dropped if d < target)
    return out


def _letter(where: dict[int, int | None], c: int) -> str | None:
    at = where[c]
    return None if at is None else column_letters(at)


def _expect_ref(where, c: int, row: int, sheet: str = "", anchored: bool = False) -> str:
    letter = _letter(where, c)
    if letter is None:
        return f"{sheet}#REF!"
    return f"{sheet}${letter}${row}" if anchored else f"{sheet}{letter}{row}"


@pytest.mark.parametrize("seed", range(120))
def test_random_rearrangements_match_the_model(seed):
    rng = random.Random(seed)
    width = rng.randint(3, 12)
    grid = Grid(rng, width)
    columns = list(range(1, width + 1))
    dropped = sorted(rng.sample(columns, k=rng.randint(0, min(2, width - 2))))
    kept = [c for c in columns if c not in dropped]
    listed = rng.sample(kept, k=rng.randint(0, len(kept)))
    if not listed and not dropped:
        listed = kept[:2]
    order: object = listed
    if rng.random() < 0.4 and listed:
        order = dict(zip((column_letters(c) for c in listed),
                         (column_letters(c) for c in sorted(listed))))
    elif rng.random() < 0.5:
        order = [column_letters(c) for c in listed]
    original = grid.to_bytes()
    wb = Workbook.open(original)
    done = wb.rearrange_columns("Grid", order, drop=[column_letters(d) for d in dropped])
    where = _model(width, listed, dropped)
    L = column_letters
    found = facts(wb.to_bytes())
    cells = found["sheets"]["Grid"]["cells"]
    # every cell where the model puts it, and nothing else in rows 1-6
    for c in columns:
        for r in range(1, 7):
            if where[c] is not None:
                assert cells[f"{L(where[c])}{r}"][2] == f"c{c}r{r}"
    assert sum(1 for ref in cells if int(re.sub("[A-Z]", "", ref)) <= 6) == 6 * len(kept)
    # the formulas in the moved columns read the cells they read
    for c in kept:
        at = L(where[c])
        assert cells[f"{at}7"][0] == _expect_ref(where, grid.reads[c], 2)
        n = grid.counts[c]
        expected = "COUNTA(#REF!)" if where[n] is None else \
            f"COUNTA({L(where[n])}1:{L(where[n])}6)"
        assert cells[f"{at}8"][0] == expected
    # another sheet's references and the defined names
    refs = found["sheets"]["Refs"]["cells"]
    for n, c in enumerate(grid.refs):
        assert refs[f"A{n + 1}"][0] == _expect_ref(where, c, 3, "Grid!")
        assert refs[f"B{n + 1}"][0] == _expect_ref(where, c, 4, "Grid!", anchored=True)
        letter = _letter(where, c)
        assert refs[f"C{n + 1}"][0] == ("SUM(Grid!#REF!)" if letter is None
                                        else f"SUM(Grid!{letter}1:{letter}6)")
        assert refs[f"D{n + 1}"][0] == ("COUNTA(Grid!#REF!)" if letter is None
                                        else f"COUNTA(Grid!{letter}:{letter})")
    for n, c in enumerate(grid.names):
        assert found["names"][f"Spot{n}"] == _expect_ref(where, c, 5, "Grid!", anchored=True)
    first, second = (_letter(where, c) for c in grid.series)
    assert found["charts"] == [[
        "Grid!#REF!" if first is None else f"Grid!${first}$1",
        "Grid!#REF!" if first is None else f"Grid!${first}$2:${first}$6",
        "Grid!#REF!" if second is None else f"Grid!${second}$2:${second}$6"]]
    # a range across columns: listed in ``reshaped`` exactly when it now covers other
    # columns than its own that stayed (a deletion only narrows it)
    back = {at: c for c, at in where.items() if at is not None}
    for n, (a, b) in enumerate(grid.spans):
        text = refs[f"E{n + 1}"][0]
        place = f"Refs!E{n + 1}"
        own = {c for c in range(a, b + 1) if where[c] is not None}
        found_span = re.fullmatch(r"COUNTA\(Grid!([A-Z]+)2:([A-Z]+)2\)", text)
        if found_span is None:
            assert not own and "#REF!" in text
            continue
        x, y = (column_index(found_span.group(i)) for i in (1, 2))
        covered = {back[k] for k in range(x, y + 1)}
        assert (covered != own) == (place in done.reshaped), (text, own, covered)
    # merged areas, rules and widths follow their columns
    sheet = found["sheets"]["Grid"]
    expected_merges = sorted(f"{L(where[c])}9:{L(where[c])}10" for c in grid.merged
                             if where[c] is not None)
    assert sorted(sheet["merges"]) == expected_merges
    rules = {(cell, f1) for cell, _type, f1, _f2 in sheet["validations"]}
    for c, y in grid.checked.items():
        if where[c] is None:
            continue
        for r in range(1, 7):
            read = "#REF!" if where[y] is None else f"{L(where[y])}{r}"
            assert (f"{L(where[c])}{r}", f"LEN({read})>0") in rules
    shades = {(cell, fs[0]) for cell, _type, fs in sheet["formats"]}
    for c, y in grid.shaded.items():
        if where[c] is None:
            continue
        for r in range(1, 7):
            read = "#REF!" if where[y] is None else f"{L(where[y])}{r}"
            assert (f"{L(where[c])}{r}", f'{read}<>""') in shades
    widths = {int(col["min"]): float(col["width"]) for col in sheet["cols"]
              if col["min"] == col["max"]}
    for c in kept:
        assert widths[where[c]] == pytest.approx(round(8 + c / 4, 2), abs=0.01)
    # the result says where each column went
    for c in kept:
        if where[c] != c or c in listed:
            assert done.columns[L(c)] == L(where[c])
    assert set(done.columns) <= {L(c) for c in kept}
    assert done.dropped == tuple(L(d) for d in dropped)
    # one undo step back to the bytes opened (none when nothing moved)
    assert wb.undo() == done.changed and wb.to_bytes() == original


# -- the call's own rules ---------------------------------------------------------------------


def test_order_and_mapping_letters_and_numbers_mean_the_same():
    results = []
    for order in (["E", "B"], [5, 2], {"E": "B", "B": "E"}, {5: 2, 2: 5}):
        wb = Workbook.open(Grid(random.Random(3), 6).to_bytes())
        wb.rearrange_columns("Grid", order)
        results.append(facts(wb.to_bytes()))
    assert all(found == results[0] for found in results)       # rule 1


def test_an_order_already_standing_is_no_change():
    wb = Workbook.open(Grid(random.Random(4), 6).to_bytes())
    done = wb.rearrange_columns("Grid", ["B", "D", "E"])
    assert not done.changed and done.steps == () and wb.version == 0      # rule 3


def test_one_call_is_one_undo_step():
    wb = Workbook.open(Grid(random.Random(5), 8).to_bytes())
    done = wb.rearrange_columns("Grid", ["H", "A", "D", "B"], drop=["F"])
    assert len(done.steps) > 2 and wb.version == 1                         # rule 5
    assert wb.undo() and wb.version == 0


def test_reshaped_ranges_are_listed_and_counted():
    wb = Workbook.new(["Rota"])
    wb.set_cells({"Rota!A1": 1, "Rota!B1": 2, "Rota!C1": 3, "Rota!E1": {"formula": "=SUM(A1:C1)"},
                  "Rota!E2": {"formula": "=SUM(A1:B1)"}})
    done = wb.rearrange_columns("Rota", ["C", "A"])           # C, B, A
    assert _formula(wb, "Rota!E1") != "SUM(A1:C1)"               # as Excel's moves leave it
    assert "Rota!E1" in done.reshaped and done.reshaped_count >= 1
    assert done.as_dict()["reshaped_count"] == done.reshaped_count       # rule 4


# -- tables (R75) -----------------------------------------------------------------------------


def _shop(**variant) -> Workbook:
    return Workbook.open(TableColumnsBook(**variant).to_bytes())


def test_a_tables_columns_are_reordered_in_the_table():
    wb = _shop()
    done = wb.rearrange_columns("Shop", ["F", "C", "D", "E"])      # Note before Qty
    assert [wb.cell(f"Shop!{c}3").v for c in "BCDEFG"] == \
        ["Code", "Note", "Qty", "Price", "Total", "Bin"]
    assert wb.cell("Shop!B1").v == "Shop lines" and wb.cell("Shop!F12").v == 5   # stay
    assert done.columns == {"C": "D", "D": "E", "E": "F", "F": "C"}
    assert any("t:Lines" in note for note in done.notes)
    found = facts(wb.to_bytes())["sheets"]["Shop"]
    assert found["tables"][0]["columns"] == ["Code", "Note", "Qty", "Price", "Total", "Bin"]


def test_a_table_and_the_columns_beside_it_in_one_call():
    wb = _shop(chart=False)            # the chart over K:Q would be cut: not measured
    wb.set_cells({"Shop!J4": "j", "Shop!K4": "k"})
    before = wb.version
    done = wb.rearrange_columns("Shop", ["G", "B", "K", "J"])    # G and B swap, K and J
    assert [wb.cell(f"Shop!{c}3").v for c in "BCDEFG"] == \
        ["Bin", "Qty", "Price", "Total", "Note", "Code"]
    assert (wb.cell("Shop!J4").v, wb.cell("Shop!K4").v) == ("k", "j")
    assert done.columns["G"] == "B" and done.columns["J"] == "K"
    assert wb.undo() and wb.version == before


def test_crossing_a_tables_edge_is_refused_with_the_way_out():
    wb = _shop()
    with pytest.raises(EditError) as raised:
        wb.rearrange_columns("Shop", ["I", "D"])
    error = raised.value
    assert error.code == "table_move" and wb.version == 0
    assert "t:Lines holds B:G" in str(error) and "D, I would cross" in str(error)    # rule 2
    assert ("table_move", "t:Lines") in error.problems


def test_a_filtered_column_moved_while_rows_are_hidden_is_refused():
    wb = Workbook.open(TableColumnsBook().to_bytes())
    data = TableColumnsBook().parts()
    sheet = data["xl/worksheets/sheet1.xml"].replace(b'<row r="5">', b'<row r="5" hidden="1">')
    data["xl/worksheets/sheet1.xml"] = sheet
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, part in data.items():
            archive.writestr(name, part)
    wb = Workbook.open(buffer.getvalue())
    with pytest.raises(EditError) as raised:
        wb.rearrange_columns("Shop", ["G", "B"])                  # Bin carries the filter
    assert raised.value.code == "not_measured" and "clear the filter" in str(raised.value)
    wb.rearrange_columns("Shop", ["D", "C"])                      # other columns: fine


def test_whole_table_columns_by_move_columns_point_to_rearrange():
    wb = _shop()
    with pytest.raises(EditError) as raised:
        wb.move_columns("Shop", "F", "C")
    assert raised.value.code == "table_move" and "rearrange_columns" in str(raised.value)


@pytest.mark.parametrize("seed", range(40))
def test_random_table_orders_match_the_model(seed):
    """A table's columns in a random order: the table's rows' cells, the references into
    them and the rules on them go where the model puts them; everything else stays."""
    rng = random.Random(seed)
    book = TableColumnsBook(chart=False, edges=rng.random() < 0.5,
                            comments=rng.random() < 0.5)
    wb = Workbook.open(book.to_bytes())
    before = facts(wb.to_bytes())
    columns = list(range(2, 8))                                      # B..G
    listed = rng.sample(columns, k=rng.randint(2, 6))
    order: object = [column_letters(c) for c in listed]
    if rng.random() < 0.5:
        order = dict(zip(order, (column_letters(c) for c in sorted(listed))))
    done = wb.rearrange_columns("Shop", order)
    goes = dict(zip(listed, sorted(listed)))
    L = column_letters

    def f(c: int) -> int:
        return goes.get(c, c)

    after = facts(wb.to_bytes())
    old, new = before["sheets"]["Shop"]["cells"], after["sheets"]["Shop"]["cells"]
    for ref, cell in old.items():
        letters = ref.rstrip("0123456789")
        row, c = int(ref[len(letters):]), column_index(letters)
        moved = 3 <= row <= 10 and 2 <= c <= 7
        target = f"{L(f(c))}{row}" if moved else ref
        assert new[target][1:] == cell[1:], ref                  # type, value, style
    assert after["sheets"]["Shop"]["cols"] == before["sheets"]["Shop"]["cols"]
    assert after["sheets"]["Shop"]["merges"] == before["sheets"]["Shop"]["merges"]
    names = before["sheets"]["Shop"]["tables"][0]["columns"]
    moved_names = [None] * 6
    for c in columns:
        moved_names[f(c) - 2] = names[c - 2]
    assert after["sheets"]["Shop"]["tables"][0]["columns"] == moved_names
    sums = after["sheets"]["Sums"]["cells"]
    assert sums["A1"][0] == f"Shop!{L(f(4))}5"
    assert sums["A2"][0] == f"SUM(Shop!{L(f(4))}4:{L(f(4))}9)"
    assert sums["A9"][0] == f"SUM(Shop!{L(f(4))}3:{L(f(4))}10)"
    assert sums["A4"][0] == "SUM(Shop!D:D)" and sums["A10"][0] == "SUM(Shop!D2:D12)"
    assert sums["A12"][0] == f"Shop!${L(f(6))}$4&Shop!{L(f(7))}$5"
    for ref in ("A5", "A6", "A7", "A11"):                            # structured: by name
        assert sums[ref][0] == before["sheets"]["Sums"]["cells"][ref][0]
    assert after["names"] == {"PriceBody": f"Shop!${L(f(4))}$4:${L(f(4))}$9",
                              "QtyCell": f"Shop!${L(f(3))}$5"}
    assert new["I4"][0] == f"{L(f(4))}4*2" and new["I9"][0] == "B12"
    # rules on the table's cells moved with them (a moved column may take its new
    # neighbour's as well)
    shop_before, shop_after = before["sheets"]["Shop"], after["sheets"]["Shop"]
    for cell, kind, formulas in shop_before["formats"]:
        if kind != "cellIs":
            continue
        letters = cell.rstrip("0123456789")
        moved = f"{L(f(column_index(letters)))}{cell[len(letters):]}"
        assert (moved, kind, tuple(formulas)) in \
            [(c, k, tuple(fs)) for c, k, fs in shop_after["formats"]]
    for cell, kind, f1, f2 in shop_before["validations"]:
        letters = cell.rstrip("0123456789")
        row = int(cell[len(letters):])
        moved = f"{L(f(column_index(letters)))}{row}" if 3 <= row <= 10 else cell
        assert (moved, kind, f1, f2) in [tuple(v) for v in shop_after["validations"]]
    for ref, text in shop_before["comments"]:
        letters = ref.rstrip("0123456789")
        row = int(ref[len(letters):])
        moved = f"{L(f(column_index(letters)))}{row}" if 3 <= row <= 10 else ref
        assert (moved, text) in [tuple(x) for x in shop_after["comments"]]
    # spans across the table's columns: listed when they now cover other cells
    back = {f(c): c for c in range(1, 10)}
    for ref, (a, b) in {"A3": (3, 5), "A13": (3, 4), "A14": (2, 7), "A8": (2, 7)}.items():
        x, y = (column_index(p.rstrip("0123456789"))
                for p in re.search(r"Shop!([A-Z]+\d+):([A-Z]+\d+)", sums[ref][0]).groups())
        covered = {back[k] for k in range(x, y + 1)}
        assert (covered != set(range(a, b + 1))) == (f"Sums!{ref}" in done.reshaped), ref
    assert wb.undo() == done.changed and wb.version == 0


# -- the oracle: Excel making the same moves ---------------------------------------------------

#: Rearrangements of structure_book.GridBook's Grid (numbers in A1:J10, formulas into it on
#: Refs): order, drop.
GRID_CASES = {
    "swap-apart": ({"C": "A", "A": "C"}, []),
    "two-apart": (["J", "H"], []),
    "cycle": ({"A": "E", "E": "C", "C": "A"}, []),
    "with-drop": (["F", "B", "I"], ["D"]),
}


def _excel_steps(done) -> list[str]:
    """The call's steps as Excel's own: deletes, and Cut + "Insert Cut Cells"."""
    out = []
    for step in done.steps:
        first = column_letters(step.first)
        last = column_letters(step.first + step.count - 1)
        if step.operation == "delete_columns":
            out.append(f"delete_columns|Grid|{first}:{last}|")
        else:
            before = column_letters(step.before)
            out.append(f"move|Grid|{first}:{last}|{before}:{before}")
    return out


@pytest.fixture(scope="module")
def excel():
    import sys

    import oracle

    if sys.platform != "darwin" or not oracle.available():
        pytest.skip("the Excel oracle needs Microsoft Excel for Mac")
    try:
        with oracle.session() as current:
            yield current
    except oracle.ExcelBusy as busy:
        pytest.skip(str(busy))


@pytest.mark.oracle
@pytest.mark.parametrize("name", sorted(GRID_CASES))
def test_excel_makes_the_same_of_the_same_moves(excel, name):
    from structure_book import GridBook
    from structure_facts import differences, measured

    order, drop = GRID_CASES[name]
    source = GridBook().to_bytes()
    wb = Workbook.open(source)
    done = wb.rearrange_columns("Grid", order, drop=drop)
    theirs = excel.operate(source, _excel_steps(done), name=f"rearrange-{name}")
    assert theirs.ok, theirs
    expected = measured(facts(theirs.saved.read_bytes()))
    assert differences(measured(facts(wb.to_bytes())), expected) == []
    ours = excel.resave(wb.to_bytes(), name=f"rearrange-ours-{name}")
    assert ours.ok, ours                       # opens without a repair prompt
    assert differences(measured(facts(ours.saved.read_bytes())), expected) == []
