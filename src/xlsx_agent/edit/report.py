"""What an edit did, bounded and by address: the report every step leaves.

An :class:`EditReport` is machine-readable and never holds content (ROADMAP.md, "Design
rules", 7): the cells written, compressed into ranges per sheet and capped
(:data:`MAX_RANGES`, the rest counted); the package parts the step changed; the
workbook's calculation status after it; and the structure issues
(:mod:`xlsx_agent.edit.structure`) of the sheets it wrote -- codes with addresses.  No
value, no text, no formula is in it, nor in what :mod:`xlsx_agent.edit` logs.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Iterable, Mapping

from ..read.types import Problem, _letters

#: At most this many ranges are listed in a report, over every sheet; the rest are counted.
MAX_RANGES = 50
#: At most this many structure issues are listed; the rest are counted.
MAX_ISSUES = 50


@dataclass(frozen=True)
class EditReport:
    """One step's facts.

    ``before`` and ``version`` are the workbook's versions before and after the step (equal
    when it changed nothing).  ``cells`` maps each sheet written (as an address writes it)
    to the cells written, as ranges (``["D10:D14", "F10"]``), at most :data:`MAX_RANGES` in
    all; ``cells_cut`` counts the ranges left out and ``cell_count`` every cell written.
    ``parts`` the package parts whose bytes the step changed, sorted.  ``calculation`` is
    ``"invalidated"`` once a cell has been written (every formula's cached value may be
    stale: :attr:`Workbook.calculation_status <xlsx_agent.Workbook.calculation_status>`)
    and the workbook's status otherwise.  ``issues`` the structure issues of the sheets
    written and of the workbook (:func:`~xlsx_agent.edit.structure.check_structure`), at
    most :data:`MAX_ISSUES`, ``issues_cut`` the rest."""

    before: int
    version: int
    cells: dict[str, list[str]] = field(default_factory=dict)
    cell_count: int = 0
    cells_cut: int = 0
    parts: tuple[str, ...] = ()
    calculation: str = "cached_unverified"
    issues: tuple[Problem, ...] = ()
    issues_cut: int = 0

    @property
    def changed(self) -> bool:
        return self.version != self.before

    def as_dict(self) -> dict[str, Any]:
        """The report as plain data, for a tool result or a log record."""
        cells = {sheet: list(ranges) for sheet, ranges in self.cells.items()}
        out: dict[str, Any] = {"before": self.before, "version": self.version,
                               "cells": cells, "cell_count": self.cell_count,
                               "parts": list(self.parts),
                               "calculation": self.calculation,
                               "issues": [_issue(p) for p in self.issues]}
        if self.cells_cut:
            out["cells_cut"] = self.cells_cut
        if self.issues_cut:
            out["issues_cut"] = self.issues_cut
        return out


def _issue(problem: Problem) -> dict[str, Any]:
    out: dict[str, Any] = {"code": problem.code}
    if problem.address is not None:
        out["address"] = problem.address
    if problem.count != 1:
        out["count"] = problem.count
    return out


def compress(cells: Iterable[tuple[int, int]]) -> list[str]:
    """Cells (row, column) as few ranges as a simple sweep finds: each row's runs of
    adjacent columns, then runs of equal column spans in adjacent rows merged into one
    rectangle.  In row order, then column order: ``["A1:C3", "E5"]``."""
    by_row: dict[int, list[int]] = {}
    for row, column in set(cells):
        by_row.setdefault(row, []).append(column)
    runs: dict[int, list[tuple[int, int]]] = {}
    for row, columns in by_row.items():
        columns.sort()
        out: list[tuple[int, int]] = []
        for column in columns:
            if out and column == out[-1][1] + 1:
                out[-1] = (out[-1][0], column)
            else:
                out.append((column, column))
        runs[row] = out
    # rectangles: (first row, last row, first column, last column), grown downwards
    open_: dict[tuple[int, int], list[int]] = {}
    done: list[tuple[int, int, int, int]] = []
    for row in sorted(runs):
        spans = set(runs[row])
        for span, rect in list(open_.items()):
            if span in spans and rect[1] == row - 1:
                rect[1] = row
                spans.discard(span)
            else:
                done.append((rect[0], rect[1], span[0], span[1]))
                del open_[span]
        for span in spans:
            open_[span] = [row, row]
    done.extend((rect[0], rect[1], span[0], span[1]) for span, rect in open_.items())
    done.sort()
    return [_ref(first, left) if (first, left) == (last, right)
            else f"{_ref(first, left)}:{_ref(last, right)}"
            for first, last, left, right in done]


def _ref(row: int, column: int) -> str:
    return f"{_letters(column)}{row}"


def capped(cells: Mapping[str, Iterable[tuple[int, int]]]
           ) -> tuple[dict[str, list[str]], int, int]:
    """Each sheet's cells compressed (:func:`compress`), at most :data:`MAX_RANGES` ranges
    over all sheets: the ranges kept by sheet, how many were cut, and how many cells there
    were."""
    out: dict[str, list[str]] = {}
    room = MAX_RANGES
    cut = 0
    count = 0
    for sheet, positions in cells.items():
        positions = set(positions)
        count += len(positions)
        ranges = compress(positions)
        if room <= 0:
            cut += len(ranges)
            continue
        out[sheet] = ranges[:room]
        cut += max(0, len(ranges) - room)
        room -= len(out[sheet])
    return out, cut, count
