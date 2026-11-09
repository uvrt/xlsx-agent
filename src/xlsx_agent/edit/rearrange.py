"""Putting a sheet's columns in an order (R59, round 46): ``wb.rearrange_columns("Stock",
["C", "A", "B"], drop=["E"])`` -- one step made of Excel's own column deletes and moves
(:mod:`xlsx_agent.edit.grid`), so every reference, on every sheet, is remapped exactly as
if a person had deleted and cut-and-inserted the columns in Excel one by one.

**The order.**  ``order`` lists columns (``"C"``, ``3``) in the order they are to stand; they
take the places the listed columns hold now, and columns not listed stay where they are:
``["C", "A", "B"]`` puts C in A's place, A in B's and B in C's; ``["K", "I"]`` swaps K and I
and leaves J between them.  A mapping says the same from the other side, ``{"C": "A", "A":
"C"}`` (each column to its new place; the new places must be the places the columns hold,
every one once).  ``drop`` lists columns deleted first; ``order`` and a mapping name columns
as they are before anything is dropped.

**A table's columns** are reordered among themselves as Excel does when a table column is
dragged to another place in its table (measured, Excel for Mac 16.106, 2026-11-09: Cut and
"Insert Cut Cells" of the table column's cells): only the table's rows move, so the table
keeps its name, style, filter, totals row and calculated columns, each table column its
name, id and formulas, and the sheet's column widths and the cells above and below the
table stay where they are (:mod:`.table_columns`).  A column is not moved into or out of a
table this way (Excel refuses to move other cells into a table): an order whose columns
would cross a table's edge is refused, naming the table and its columns.

**Ranges.**  A cell reference follows its cell, and so does a range inside one column; a
range across several columns changes as Excel's moves change it, which can leave it
covering other cells (``SUM(A1:C1)`` with C put first is ``SUM(B1:C1)``: Excel does the
same).  Each such place is listed in the result (``reshaped``).

**Reading and giving back.**  To read a sheet a person edited -- its columns in another
order -- a dataset by header (``wb.dataset("t:...")``, the tables' columns by name) already
ignores the order; :meth:`~xlsx_agent.Workbook.rearrange_columns` is for giving a sheet back
in a template's order.

Every step is checked when it is reached: one refused refuses the whole call and nothing is
changed.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Mapping, Sequence

from ooxml_edit.sheets import column_index, column_letters

from .grid import COLUMN, X, Change, parse_area
from .restructure import MAX_LISTED, Restructured, line_span, restructure


@dataclass(frozen=True)
class Rearranged:
    """What :func:`rearrange_columns` did: ``columns`` each column placed and each other
    column whose letters changed, old letters to new (where the steps put it; for a
    table's column, its cells in the table); ``dropped`` the columns deleted (as they were); ``steps`` each delete and move
    made (:class:`~xlsx_agent.edit.restructure.Restructured`); ``changed`` whether
    anything moved or went; ``reshaped`` where a range across moved columns now covers
    other cells (see the module); ``notes`` what the caller should know."""

    sheet: str
    columns: dict[str, str] = field(default_factory=dict)
    dropped: tuple[str, ...] = ()
    steps: tuple[Restructured, ...] = ()
    notes: tuple[str, ...] = ()

    @property
    def changed(self) -> bool:
        return any(step.changed for step in self.steps)

    @property
    def reshaped(self) -> tuple[str, ...]:
        seen: list[str] = []
        for step in self.steps:
            seen += [where for where in step.reshaped if where not in seen]
        return tuple(seen[:MAX_LISTED])

    @property
    def reshaped_count(self) -> int:
        return sum(step.reshaped_count for step in self.steps)

    def as_dict(self) -> dict[str, Any]:
        out = {"sheet": self.sheet, "changed": self.changed, "columns": dict(self.columns),
               "dropped": list(self.dropped),
               "steps": [step.as_dict() for step in self.steps],
               "formulas": sum(step.formulas for step in self.steps),
               "broken_count": sum(step.broken_count for step in self.steps)}
        if self.reshaped_count:
            out["reshaped"] = list(self.reshaped)
            out["reshaped_count"] = self.reshaped_count
        if self.notes:
            out["notes"] = list(self.notes)
        return out


@dataclass
class _Table:
    name: str
    part: str
    c1: int
    c2: int


def rearrange_columns(workbook: Any, sheet: Any, order: Any = None, *,
                      drop: "Sequence[str | int]" = (), revision: int | None = None
                      ) -> Rearranged:
    """See the module.  :class:`~xlsx_agent.edit.EditError` ``bad_position`` for an order
    that names a column twice, a mapping whose places are not the columns' own, or a
    column both kept and dropped; ``table_move`` for columns that would cross a table's
    edge; any step's refusal refuses the call."""
    from . import EditError

    if revision is not None and revision != workbook.version:
        raise EditError("revision", f"the call was made for version {revision}; the "
                        f"workbook is at version {workbook.version}")
    dropped = sorted({_column(c, "drop") for c in drop})
    if isinstance(order, Mapping):
        pairs = [(_column(k, "mapping"), _column(v, "mapping")) for k, v in order.items()]
        sources = [k for k, _ in pairs]
        targets = sorted(v for _, v in pairs)
        if len(set(sources)) != len(sources) or sorted(sources) != targets:
            raise EditError("bad_position", "mapping: each column to a new place, the places "
                            "being the columns' own, every one once (as {\"C\": \"A\", "
                            "\"A\": \"C\"}); or give order=[\"C\", \"B\", \"A\"]")
        listed = [k for k, _ in sorted(pairs, key=lambda kv: kv[1])]
    elif order is None:
        listed = []
    elif isinstance(order, (str, bytes)) or not isinstance(order, Sequence):
        raise EditError("bad_position", "order: a list of columns (\"C\", 3) in the order "
                        "they are to stand")
    else:
        listed = [_column(c, "order") for c in order]
        if len(set(listed)) != len(listed):
            twice = sorted({column_letters(c) for c in listed if listed.count(c) > 1})
            raise EditError("bad_position", f"order: {', '.join(twice)} named twice")
    both = sorted(set(listed) & set(dropped))
    if both:
        raise EditError("bad_position", "a column both placed and dropped: "
                        + ", ".join(column_letters(c) for c in both))
    if not listed and not dropped:
        raise EditError("bad_position", "nothing to do: give order (or a mapping) and/or drop")

    # Each listed column goes to the place of the listed column that many from the left.
    goes = {c: t for c, t in zip(listed, sorted(listed)) if c != t}
    try:
        found = workbook._cell_sheet(sheet)
    except (KeyError, ValueError) as error:
        raise EditError("not_a_worksheet", str(error).strip("'\"")) from None
    tables = _tables(workbook, found)
    _check_tables(goes, tables)

    steps: list[Restructured] = []
    notes: list[str] = []
    last = max([*listed, *dropped, _last_column(workbook, found)])
    # pos[c]: where the original column c stands now; cells[c]: where its cells are (a
    # table's column: its cells in the table, which move without the column).
    pos: dict[int, int | None] = {c: c for c in range(1, last + 1)}
    cells: dict[int, int | None] = dict(pos)

    def follow(change: Change, only: "set[int] | None" = None) -> None:
        for track in ((cells,) if only is not None else (pos, cells)):
            for c, at in track.items():
                if at is not None and (only is None or c in only):
                    track[c] = change.index(at)

    with workbook.editor.batch():
        # Deletes first, the rightmost first: each run of adjacent columns in one step.
        for first, end in reversed(_runs(dropped)):
            steps.append(restructure(workbook, found, "delete_columns", first,
                                     end - first + 1))
            follow(Change("delete", COLUMN, first, end - first + 1))
        # A table's columns among themselves: the table's cells, as a dragged table column.
        from .table_columns import move_table_columns

        for table in tables:
            if not any(table.c1 <= c <= table.c2 for c in goes):
                continue
            inside = [c for c in range(table.c1, table.c2 + 1) if c not in dropped]
            final = [_placed(c, goes) for c in inside]
            for at, count, before in _plan([_at(pos, c) for c in inside],
                                           [_at(pos, c) for c in final], blocks=False):
                steps.append(move_table_columns(workbook, found, table.part, at, count,
                                                before))
                follow(Change("move", COLUMN, at, count, before), set(inside))
            notes.append(f"t:{table.name}: its columns were reordered within the table's "
                         "rows, as Excel does when a table column is dragged; the column "
                         "widths and the cells above and below the table stayed")
        # The other columns: whole columns, by Excel's Cut and "Insert Cut Cells".
        outside = {c for c in goes if not any(t.c1 <= c <= t.c2 for t in tables)}
        if outside:
            kept = [c for c in range(min(outside), max(outside) + 1) if c not in dropped]
            final = [_placed(c, goes) if c in outside else c for c in kept]
            for at, count, before in _plan(
                    [_at(pos, c) for c in kept], [_at(pos, c) for c in final],
                    {_at(pos, c) for c in kept if c not in outside}):
                steps.append(restructure(workbook, found, "move_columns", at, count, before))
                follow(Change("move", COLUMN, at, count, before))
        # The steps placed every listed column where it was asked to go.
        for c, target in goes.items():
            expected = target - sum(1 for d in dropped if d < target)
            if cells[c] != expected:
                raise EditError("not_measured", f"column {column_letters(c)} would stand at "
                                f"{column_letters(cells[c] or 1)}, not "
                                f"{column_letters(expected)}: nothing was changed")
    columns = {column_letters(c): column_letters(at) for c, at in sorted(cells.items())
               if at is not None and (at != c or c in listed)}
    label = steps[0].sheet if steps else str(sheet)
    return Rearranged(sheet=label, columns=columns,
                      dropped=tuple(column_letters(c) for c in dropped), steps=tuple(steps),
                      notes=tuple(notes))


def _at(pos: "dict[int, int | None]", column: int) -> int:
    at = pos[column]
    assert at is not None
    return at


def _placed(column: int, goes: dict[int, int]) -> int:
    """The column whose cells end in ``column``'s place."""
    for source, target in goes.items():
        if target == column:
            return source
    return column


def _plan(current: list[int], order: list[int], still: "set[int] | None" = None, *,
          blocks: bool = True) -> list[tuple[int, int, int]]:
    """Moves ``(first, count, before)`` that turn the columns as they stand (``current``,
    numbered ``current[0]`` on) into ``order``: the longest run of columns already in
    order stays (every column of ``still`` among them), and each other column is cut and
    inserted before the next column after it that stays -- with its neighbours in the
    order when they stand together, as one block (``blocks``; a table's columns go one at
    a time, as measured).  A person reordering columns in Excel moves as few as that."""
    lo = current[0]
    place = {column: i for i, column in enumerate(current)}
    stays = _longest(order, place, still or set())
    state = list(current)
    moves = []
    i = 0
    while i < len(order):
        if order[i] in stays:
            i += 1
            continue
        k = 1
        while (blocks and i + k < len(order) and order[i + k] not in stays
               and state.index(order[i + k]) == state.index(order[i]) + k):
            k += 1
        after = next((c for c in order[i + k:] if c in stays), None)
        j = state.index(order[i])
        at = state.index(after) if after is not None else len(state)
        if at != j + k:
            moves.append((lo + j, k, lo + at))
            block = state[j:j + k]
            del state[j:j + k]
            if at > j:
                at -= k
            state[at:at] = block
        i += k
    assert state == order
    return moves


def _longest(order: list[int], place: dict[int, int], still: "set[int]") -> set[int]:
    """The heaviest run of ``order`` whose places increase: a column of ``still`` weighs
    more than all the others together (it must stay), any other one."""
    heavy = len(order) + 1
    size = len(order)
    tree = [(0, -1)] * (size + 1)          # Fenwick tree of (weight, index) maxima by place
    best: list[tuple[int, int]] = []
    for index, column in enumerate(order):
        at = place[column]
        weight, previous = 0, -1
        k = at                              # the best run ending at a place below ``at``
        while k > 0:
            if tree[k][0] > weight:
                weight, previous = tree[k]
            k -= k & -k
        weight += heavy if column in still else 1
        best.append((weight, previous))
        k = at + 1
        while k <= size:
            if weight > tree[k][0]:
                tree[k] = (weight, index)
            k += k & -k
    end = max(range(len(order)), key=lambda i: best[i][0], default=-1)
    out: set[int] = set()
    while end >= 0:
        out.add(order[end])
        end = best[end][1]
    return out


def _tables(workbook: Any, sheet: Any) -> list[_Table]:
    package = workbook.package
    if sheet.part is None or not package.has_part(sheet.part):
        return []
    out = []
    for rel in package.relationships(sheet.part).values():
        if rel.is_external or rel.type.rpartition("/")[2] != "table" \
                or rel.target_part is None or not package.has_part(rel.target_part):
            continue
        root = package.tree(rel.target_part)
        area = parse_area(root.get("ref", ""))
        if area is None:
            continue
        out.append(_Table(root.get("displayName") or root.get("name") or "", rel.target_part,
                          area.c1, area.c2))
    return sorted(out, key=lambda t: t.c1)


def _check_tables(goes: dict[int, int], tables: list[_Table]) -> None:
    """Columns that would go into or out of a table: refused, every table named."""
    from . import EditError

    problems = []
    for table in tables:
        def inside(c: int) -> bool:
            return table.c1 <= c <= table.c2

        crossing = sorted(c for c, target in goes.items() if inside(c) != inside(target))
        if crossing:
            first, last = column_letters(table.c1), column_letters(table.c2)
            problems.append((f"t:{table.name}", f"t:{table.name} holds {first}:{last}; "
                             + ", ".join(column_letters(c) for c in crossing)
                             + " would cross its edge (a table's columns are reordered "
                             "among themselves, as Excel drags a table column, and Excel "
                             "refuses to move other cells into a table)"))
    if problems:
        raise EditError("table_move", "; ".join(detail for _, detail in problems),
                        problems=[("table_move", where) for where, _ in problems])


def _last_column(workbook: Any, sheet: Any) -> int:
    """The sheet's last column holding a cell (1 for none): what ``columns`` reports up to."""
    if sheet.part is None or not workbook.package.has_part(sheet.part):
        return 1
    data = workbook.package.tree(sheet.part).find(X + "sheetData")
    last = 1
    for row in (data if data is not None else []):
        for cell in row.iterchildren(X + "c", reversed=True):
            letters = (cell.get("r") or "").rstrip("0123456789")
            if letters.isalpha():
                last = max(last, column_index(letters))
            break
    return last


def _column(value: Any, what: str) -> int:
    first, last = line_span(value, COLUMN, what)
    if first != last:
        from . import EditError

        raise EditError("bad_position", f"{what}: one column at a time ({value!r} is "
                        "several)")
    return first


def _runs(columns: list[int]) -> list[tuple[int, int]]:
    out: list[tuple[int, int]] = []
    for column in columns:
        if out and out[-1][1] + 1 == column:
            out[-1] = (out[-1][0], column)
        else:
            out.append((column, column))
    return out


__all__ = ["Rearranged", "rearrange_columns"]
