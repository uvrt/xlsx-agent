"""``xlsx_compare``: two open workbooks compared -- sheets by id, tables and regions by name,
rows by key, cells by type and value (ROADMAP.md, "Tools").  In the loaded core: its
definition is 329 estimated (about 487 counted by the offline proxy) and the core with it
about 4,153 counted, under the 6,000 budget; X10's words on the summary make it 330, the
core 3,474 estimated, about 5,142 counted (ROADMAP.md: the check tools stay loaded while
the budget holds; ``tests/test_tools_schema.py`` keeps it).

What it does is :mod:`xlsx_agent.compare`'s (:meth:`Workbook.compare
<xlsx_agent.Workbook.compare>`): ``a`` is the older workbook, ``b`` the newer, both open
documents; ``tables`` names the tables or regions to compare (each by the same name in
both; ``table_b`` gives ``b``'s name for a single one), every table and declared region of
one name in both by default; ``key`` names the key column, else each region's declared or
only key column in both (none or several: refused, ``no_key_column``, the candidates as
``valid_options``); ``columns`` limits the cells compared.

**The result.**  ``tables``, first: each pair compared, its addresses and ranges in both,
its key column, its ``label`` column (the text that names a row for a reader:
:mod:`xlsx_agent.compare`, "Labels"; none for a table without one), its rows in each, how many were paired by key and of those unchanged, and
its ``totals``; on the first page, its **summary** (X10): ``columns``, each column with a
cell changed and how many (``changed``), for numbers, percents, dates and times how many
went ``up`` and ``down`` and the largest move each way by key (``max_up``:
``"t:Stock/SKU-0002/Price +0.01"``, a date's in days, ``+3d``; :class:`~xlsx_agent.compare.
ColumnSummary`), text, booleans and lists ``changed`` only; and ``added`` and ``removed``,
the rows by key (``t:Stock/SKU-0009``, :data:`KEYS_SHOWN` each, the rest counted in
``more_added``, ``more_removed``).  It is what a model reads to skip the pages: at most
``Limits.max_list_items`` columns a table (``more_columns``), and the tables' summaries
within half the page (the rest left out, counted in ``cut``'s ``summaries``).  A column
hidden from the model is no column of its table: never compared, counted or named.
``totals``: every difference counted by kind, always, over all of them.  ``differences``:
the differences themselves, in order -- sheets (``sheet``: added, removed, renamed, by id),
tables in one workbook only (``table``), columns (``column``: added, removed, renamed),
keys two rows of one workbook share (``duplicate_key``: never paired, each workbook's rows
by position, :data:`POSITIONS_SHOWN` at most, the rest counted in ``more``), rows without
a key (``unkeyed``), rows added and removed (``row``, with ``at``, its address by key in
its workbook, its ``label``, and its cells with their sheet addresses), and cells changed
(``cell``: ``changed``, ``type`` -- a number and the same number as text --, ``added``,
``removed``; with ``at``, its address by key in ``b``, ``t:Stock/SKU-0002/Price``, its
row's ``label``, and the cell in ``a`` and in ``b``, typed with their sheet addresses).
A row's ``label`` (T19b) is there only when the row has one: a model that names a row
by its key alone misnames it.  A key longer than
:data:`~.read.TEXT_SHOWN` characters is no address: the difference gives its ``table``,
its ``key`` cut with ``key_length``, and a cell's ``column``.

**Bounded and paged** as the other read tools are (:mod:`.read`'s cursor and sizes): at
most ``Limits.max_list_items`` differences a page within ``Limits.max_result_chars``; text
longer than :data:`~.read.TEXT_SHOWN` characters cut with its ``length``; a row's cells cut
to what a page holds (``more_cells``: read the row by its address).  ``cut`` says what was
left, ``next_cursor`` gives the next page of the same comparison.  ``total`` is the number
of differences.  The comparison is made once and kept for its next pages (the workbooks
do not change while open for reading; a cursor of another comparison is refused).

**Content.**  Keys and values are in results only; the summary names documents and counts,
errors name tables, columns and codes (ROADMAP.md, "Design rules", 7).
"""

from __future__ import annotations

import json
from typing import Any, Callable

from ooxml_edit.tools import CORE, Result, ToolError, array, boolean, string, tool

from ..address import TableAddress, parse_address
from ..compare import Diff, compare
from ..read.types import Value
from .docformat import KIND, handler
from .query import KEPT
from .read import _MARGIN, CURSOR, TEXT_SHOWN, _cursor, _read_cursor, _size

#: The most row positions a ``duplicate_key`` difference lists per workbook.
POSITIONS_SHOWN = 8
#: The most rows added, and removed, a table's summary names by key.
KEYS_SHOWN = 10


@tool("xlsx_compare",
      "Compare two open workbooks, a (older) and b (newer): sheets by id, tables by name, "
      "rows by key (added, removed), cells by type and value (old and new, by key; number "
      "vs text is a type change), columns added, removed, renamed. A key two rows share is "
      "reported, never paired. First: changes per column (numbers, dates: up, down, "
      "largest); totals; paged.",
      {"a": string("Older document id."),
       "b": string("Newer document id."),
       "tables": array(string(), "Tables or regions, same name in both (t:Products). "
                       "Default every table.", optional=True),
       "table_b": string("b's name for the one table given, if it differs.", optional=True),
       "key": string("Key column. Default the declared or only key column in both.",
                     optional=True),
       "columns": array(string(), "Compare only these columns' cells.", optional=True),
       "cells": boolean("Every cell by address instead: formulas by text, even if values "
                        "match; names, validations, merges, conditional formats; changes "
                        "outside the write policy.", optional=True),
       "formula_text": boolean("With cells: formulas' text.", optional=True),
       "cursor": CURSOR},
      group=CORE, kind=KIND, documents=("a", "b"))
@handler
def xlsx_compare(call: Any, a: str, b: str, tables: list[str] | None = None,
                 table_b: str | None = None, key: str | None = None,
                 columns: list[str] | None = None, cells: bool = False,
                 formula_text: bool = False, cursor: str | None = None) -> Result:
    """``xlsx_compare``'s workbook handler (see the module's notes)."""
    old = call.document
    entry = call.entries.get(b)
    if entry is None or entry.kind != KIND:
        raise ToolError("invalid_arguments", f"xlsx_compare compares two workbooks; {b} is a "
                        f"{getattr(entry, 'kind', 'missing')} document", field="b")
    new = entry.document
    for name in tables or ():
        _whole(name, "tables")
    if table_b is not None:
        if not tables or len(tables) != 1:
            raise ToolError("invalid_arguments", "table_b is b's name for the one table in "
                            "tables", field="table_b")
        _whole(table_b, "table_b")
    if formula_text and not cells:
        raise ToolError("invalid_arguments", "formula_text lists the formulas of the cell "
                        "compare: give cells too", field="formula_text")
    pairs = None if not tables else \
        ([(tables[0], table_b)] if table_b is not None else list(dict.fromkeys(tables)))
    # A cursor belongs to one comparison: the two documents as they are, and the arguments.
    versions = [call.entry.version, entry.version]
    ident = json.dumps(["compare", a, b, versions, pairs, key, columns, bool(cells),
                        bool(formula_text)], ensure_ascii=False)
    diff = _diff(old, new, ident, pairs, key, columns, bool(cells), bool(formula_text))

    items = list(_items(diff))
    total = len(items)
    start, _ = _read_cursor(cursor, ident, total, 1, f"{a} and {b}")
    chars = call.limits.max_result_chars
    limit = call.limits.max_list_items
    totals = diff.totals()
    data: dict[str, Any] = {"a": a, "b": b}
    shown_tables = [_table(t) for t in diff.tables[:limit]]
    data["tables"] = shown_tables
    unsummed = _summaries(shown_tables, diff.tables, chars // 2, limit) if start == 0 else 0
    data["totals"] = totals
    if diff.cells is not None and start == 0:
        data["cells"] = diff.cells.as_dict()
    data["differences"] = []
    summary_room = len(_summary(a, b, diff, totals, 1, total, total, True))
    budget = chars - _MARGIN - _size({**data, "cut": _CUT_ROOM}) - summary_room
    used = 0
    texts = 0
    index = start
    full = False
    while index < total and len(data["differences"]) < limit:
        shaped, cut = items[index](TEXT_SHOWN, None)
        size = _size(shaped) + 1
        if used + size > budget:
            if data["differences"]:
                full = True
                break
            shaped, cut = _fit(items[index], budget - used)
            size = _size(shaped) + 1
        data["differences"].append(shaped)
        used += size
        texts += cut
        index += 1

    out_cut: dict[str, Any] = {}
    next_cursor = None
    if index < total:
        out_cut["differences"] = total - index
        if full:
            out_cut["chars"] = True
        next_cursor = _cursor(ident, index, 0)
    if texts:
        out_cut["text"] = texts
    if len(diff.tables) > len(shown_tables):
        out_cut["tables"] = len(diff.tables) - len(shown_tables)
    if unsummed:
        out_cut["summaries"] = unsummed
    if out_cut:
        data["cut"] = out_cut
    summary = _summary(a, b, diff, totals, start + 1, index, total, next_cursor is not None)
    return Result(summary=summary, data=data, total=total, next_cursor=next_cursor)


def _whole(name: str, field: str) -> None:
    address = parse_address(name)
    if not isinstance(address, TableAddress) or not address.is_whole:
        raise ToolError("invalid_arguments", "xlsx_compare takes tables or regions "
                        "(t:Products); name columns in columns", field=field)


def _diff(old: Any, new: Any, ident: str, pairs: Any, key: str | None,
          columns: list[str] | None, cells: bool = False, formula_text: bool = False) -> Diff:
    """The comparison, kept with the older workbook's last queries for its next pages."""
    kept = old._queries
    found = kept.get(ident)
    if found is None:
        found = compare(old, new, tables=pairs, key=key, columns=columns, cells=cells,
                        formula_text=formula_text)
        kept[ident] = found
        while len(kept) > KEPT:
            kept.popitem(last=False)
    else:
        kept.move_to_end(ident)
    return found


def _table(table: Any) -> dict[str, Any]:
    out: dict[str, Any] = {"a": table.a}
    if table.b != table.a:
        out["b"] = table.b
    out.update({"range_a": table.range_a, "range_b": table.range_b, "key": table.key})
    if table.label is not None:
        out["label"] = table.label
    if table.formula_key:
        out["formula_key"] = True
    out.update({"rows": [table.rows_a, table.rows_b], "paired": table.paired,
                "unchanged": table.unchanged})
    totals = table.totals()
    if totals:
        out["totals"] = totals
    return out


def _summary_of(table: Any, limit: int) -> dict[str, Any]:
    """A table's summary for the first page: its columns' changes counted
    (:meth:`~xlsx_agent.compare.TableDiff.by_column`, ``limit`` columns at most, the
    rest counted in ``more_columns``) and its rows added and removed by key
    (:data:`KEYS_SHOWN` each at most, the rest counted in ``more_added`` and
    ``more_removed``); nothing for a table without such changes."""
    out: dict[str, Any] = {}
    columns = table.by_column()
    if columns:
        out["columns"] = {c.column: c.as_dict() for c in columns[:limit]}
        if len(columns) > limit:
            out["more_columns"] = len(columns) - limit
    for name in ("added", "removed"):
        rows = getattr(table, name)
        if rows:
            out[name] = [row.at for row in rows[:KEYS_SHOWN]]
            if len(rows) > KEYS_SHOWN:
                out[f"more_{name}"] = len(rows) - KEYS_SHOWN
    return out


def _summaries(shown: list[dict], tables: Any, room: int, limit: int) -> int:
    """Each shown table's summary added to it while they fit ``room`` characters together;
    how many tables with one were left without it."""
    used = 0
    left = 0
    for out, table in zip(shown, tables):
        summary = _summary_of(table, limit)
        if not summary:
            continue
        size = _size(summary)
        if left or used + size > room:
            left += 1
            continue
        out.update(summary)
        used += size
    return left


# -- the differences, each shaped when its page is made ------------------------------------------

#: An item: ``(cap, cells)`` -> the difference as data, with text cut at ``cap`` characters
#: and a row's cells cut at ``cells`` (``None``: all), and how many texts were cut.
Item = Callable[[int, "int | None"], "tuple[dict, int]"]


def _items(diff: Diff):
    for change in diff.sheets:
        yield lambda cap, cells, change=change: ({"sheet": change.change,
                                                 **_drop(change.as_dict(), "change")}, 0)
    for name in diff.tables_added:
        yield lambda cap, cells, name=name: ({"table": "added", "b": name}, 0)
    for name in diff.tables_removed:
        yield lambda cap, cells, name=name: ({"table": "removed", "a": name}, 0)
    for table in diff.tables:
        where = table.a
        for change in table.columns:
            yield lambda cap, cells, change=change, where=where: (
                {"table": where, "column": change.change, **_drop(change.as_dict(), "change")}, 0)
        for duplicate in table.duplicates:
            yield lambda cap, cells, d=duplicate, where=where: _duplicate(where, d, cap)
        for side, count in (("a", table.unkeyed_a), ("b", table.unkeyed_b)):
            if count.count:
                yield lambda cap, cells, side=side, count=count, where=where: (
                    {"table": where, "unkeyed": side, **count.as_dict()}, 0)
        for row in (*table.added, *table.removed):
            yield lambda cap, cells, row=row, where=where: _row(where, row, cap, cells)
        for change in table.changed:
            yield lambda cap, cells, change=change, where=where: _cell_change(where, change, cap)
    if diff.cells is not None and diff.cells.formula_text:
        for change in diff.cells.structure:
            if change.a or change.b or change.withheld:
                yield lambda cap, cells, change=change: _structure(change, cap)
        for edit in diff.cells.edits:
            if edit.a is not None or edit.b is not None or edit.withheld:
                yield lambda cap, cells, edit=edit: _formula(edit, cap)


def _drop(data: dict, name: str) -> dict:
    return {k: v for k, v in data.items() if k != name}


def _text(text: str, cap: int) -> tuple[str, int]:
    return (text[:cap], 1) if len(text) > cap else (text, 0)


def _value(value: Value, cap: int) -> tuple[dict, int]:
    out = value.as_dict()
    if isinstance(out["v"], str) and len(out["v"]) > cap:
        out["length"] = len(out["v"])
        out["v"] = out["v"][:cap]
        return out, 1
    return out, 0


def _duplicate(where: str, duplicate: Any, cap: int) -> tuple[dict, int]:
    key, cut = _text(duplicate.key, cap)
    out: dict[str, Any] = {"table": where, "duplicate_key": key}
    if cut:
        out["key_length"] = len(duplicate.key)
    for side in ("a", "b"):
        rows = getattr(duplicate, side)
        out[side] = list(rows[:POSITIONS_SHOWN])
        if len(rows) > POSITIONS_SHOWN:
            out[f"more_{side}"] = len(rows) - POSITIONS_SHOWN
    return out, cut


def _named(where: str, key: str, at: str, cap: int, column: str | None = None
           ) -> tuple[dict, int]:
    """A difference's place: its address by key (``at``), or, for a key longer than
    ``cap``, its table, its key cut and its length (and its column)."""
    if len(key) <= cap:
        return {"at": at}, 0
    out: dict[str, Any] = {"table": where, "key": key[:cap], "key_length": len(key)}
    if column is not None:
        out["column"] = column
    return out, 1


def _row(where: str, row: Any, cap: int, cells: int | None) -> tuple[dict, int]:
    named, cut = _named(where, row.key, row.at, cap)
    out: dict[str, Any] = {"row": row.change, **named}
    if row.label is not None:
        out["label"] = row.label
    shown = row.cells if cells is None else row.cells[:cells]
    values = []
    for value in shown:
        shaped, n = _value(value, cap)
        values.append(shaped)
        cut += n
    out["cells"] = values
    if len(shown) < len(row.cells):
        out["more_cells"] = len(row.cells) - len(shown)
    return out, cut


def _cell_change(where: str, change: Any, cap: int) -> tuple[dict, int]:
    named, cut = _named(where, change.key, change.at, cap, change.column)
    out: dict[str, Any] = {"cell": change.change, **named}
    if change.label is not None:
        out["label"] = change.label
    for side in ("a", "b"):
        shaped, n = _value(getattr(change, side), cap)
        out[side] = shaped
        cut += n
    return out, cut


def _formula(edit: Any, cap: int) -> tuple[dict, int]:
    """A cell of the cell compare with its formulas (``formula_text``): its address, kind,
    and each side's formula, cut at ``cap`` characters with its length."""
    out: dict[str, Any] = {"formula": edit.kind, "at": edit.address}
    cut = 0
    if edit.withheld:
        out["withheld"] = True
    for side in ("a", "b"):
        text = getattr(edit, side)
        if text is None:
            continue
        shown, n = _text(text, cap)
        out[side] = shown
        if n:
            out[f"{side}_length"] = len(text)
        cut += n
    return out, cut


def _structure(change: Any, cap: int) -> tuple[dict, int]:
    """A structural change with its formulas (``formula_text``): its kind, where, and each
    side's formulas, each cut at ``cap`` characters."""
    out: dict[str, Any] = {"structure": change.kind, "at": change.address}
    cut = 0
    if change.withheld:
        out["withheld"] = True
    for side in ("a", "b"):
        texts = getattr(change, side)
        if not texts:
            continue
        shown = []
        for text in texts:
            kept, n = _text(text, cap)
            shown.append(kept)
            cut += n
        out[side] = shown
    return out, cut


def _fit(item: Item, room: int) -> tuple[dict, int]:
    """A page's first difference, too big for the page as it is: its text cut, then a row's
    cells, until it fits ``room`` characters."""
    cap = TEXT_SHOWN
    while True:
        shaped, cut = item(cap, None)
        if _size(shaped) + 1 <= room or cap == 0:
            break
        cap = cap // 2 if cap > 64 else 0
    cells = len(shaped.get("cells", ()))
    while _size(shaped) + 1 > room and cells > 0:
        cells //= 2
        shaped, cut = item(cap, cells)
    return shaped, cut


#: A ``cut`` at its longest, to keep room for it in a page.
_CUT_ROOM = {"differences": 10 ** 9, "chars": True, "text": 10 ** 6, "tables": 10 ** 5,
             "summaries": 10 ** 5}


def _summary(a: str, b: str, diff: Diff, totals: dict[str, int], first: int, last: int,
             total: int, more: bool) -> str:
    text = f"{a} and {b}: {len(diff.tables)} tables compared"
    if diff.cells is not None:
        text += f", cells of {len(diff.cells.sheets)} sheets"
        if diff.cells.outside is not None:
            outside = len(diff.cells.outside) + len(diff.cells.structure) \
                + diff.cells.sheets_outside
            text += f" ({outside} outside the write policy)"
    if not totals:
        return text + "; no differences"
    counts = ", ".join(f"{n} {name.replace('_', ' ')}" for name, n in totals.items())
    text += f"; {counts}"
    if not total:
        return text
    text += f"; differences {first}-{last} of {total}"
    return text + "; more with next_cursor" if more else text


TOOLS = [xlsx_compare]
