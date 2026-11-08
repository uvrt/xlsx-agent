"""Putting a sheet's columns in an order (R59): ``wb.rearrange_columns("Stock", ["C", "A",
"B"], drop=["E"])`` -- one step made of Excel's own column deletes and moves
(:mod:`xlsx_agent.edit.grid`), so every reference, on every sheet, is remapped exactly as
if a person had deleted and cut-and-inserted the columns in Excel one by one.

**The order.**  ``order`` lists columns (``"C"``, ``3``) in the order they are to stand; they
take the places the listed columns hold now, so columns not listed stay where they are:
``["C", "A", "B"]`` puts C in A's place, A in B's and B in C's.  A mapping says the same
from the other side, ``{"C": "A", "A": "B", "B": "C"}`` (each column to its new place; the
new places must be the places the columns hold, every one once).  ``drop`` lists columns
deleted first; ``order`` and a mapping name columns as they are before anything is dropped.

**Reading and giving back.**  To read a sheet a person edited -- its columns in another
order -- a dataset by header (``wb.dataset("t:...")``, the tables' columns by name) already
ignores the order; :meth:`~xlsx_agent.Workbook.rearrange_columns` is for giving a sheet back
in a template's order.

Every step is checked when it is reached: one refused (a table's columns: Excel refuses to
move them) refuses the whole call and nothing is changed.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Mapping, Sequence

from ooxml_edit.sheets import column_letters

from .grid import COLUMN
from .restructure import Restructured, line_span, restructure


@dataclass(frozen=True)
class Rearranged:
    """What :func:`rearrange_columns` did: ``columns`` each column that stayed, old
    letters to new; ``dropped`` the columns deleted (as they were); ``steps`` each delete
    and move made (:class:`~xlsx_agent.edit.restructure.Restructured`), ``changed``
    whether anything moved or went."""

    sheet: str
    columns: dict[str, str] = field(default_factory=dict)
    dropped: tuple[str, ...] = ()
    steps: tuple[Restructured, ...] = ()

    @property
    def changed(self) -> bool:
        return any(step.changed for step in self.steps)

    def as_dict(self) -> dict[str, Any]:
        return {"sheet": self.sheet, "changed": self.changed, "columns": dict(self.columns),
                "dropped": list(self.dropped),
                "steps": [step.as_dict() for step in self.steps],
                "formulas": sum(step.formulas for step in self.steps),
                "broken_count": sum(step.broken_count for step in self.steps)}


def rearrange_columns(workbook: Any, sheet: Any, order: Any = None, *,
                      drop: "Sequence[str | int]" = (), revision: int | None = None
                      ) -> Rearranged:
    """See the module.  :class:`~xlsx_agent.edit.EditError` ``bad_position`` for an order
    that names a column twice, a mapping whose places are not the columns' own, or a
    column both kept and dropped; any step's refusal refuses the call."""
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
                            "\"A\": \"B\", \"B\": \"C\"}); or give order=[\"C\", \"A\", \"B\"]")
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

    steps: list[Restructured] = []
    with workbook.editor.batch():
        found = workbook._cell_sheet(sheet)
        # Deletes first, the rightmost first: each run of adjacent columns in one step.
        for first, last in reversed(_runs(dropped)):
            steps.append(restructure(workbook, found, "delete_columns", first,
                                     last - first + 1))
        # Where each kept column is now.
        def now(column: int) -> int:
            return column - sum(1 for d in dropped if d < column)

        current = [now(c) for c in listed]
        slots = sorted(current)
        # Place them one by one, left to right: Excel's Cut and Insert Cut Cells.
        positions = list(current)
        for index, slot in enumerate(slots):
            at = positions[index]
            if at == slot:
                continue
            steps.append(restructure(workbook, found, "move_columns", at, 1, slot))
            # columns from slot to at-1 shift right by one
            positions = [p + 1 if slot <= p < at else p for p in positions]
            positions[index] = slot
    columns = {}
    for original, slot in zip(listed, slots):
        columns[column_letters(original)] = column_letters(slot)
    return Rearranged(sheet=steps[0].sheet if steps else str(sheet), columns=columns,
                      dropped=tuple(column_letters(c) for c in dropped), steps=tuple(steps))


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
