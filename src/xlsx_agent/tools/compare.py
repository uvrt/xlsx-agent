"""``xlsx_compare``: two open workbooks compared -- sheets by id, tables and regions by name,
rows by key, cells by type and value (ROADMAP.md, "Tools").  In the loaded core: its
definition is 329 estimated (about 487 counted by the offline proxy) and the core with it
about 4,153 counted, under the 6,000 budget (ROADMAP.md: the check tools stay loaded while
the budget holds; ``tests/test_tools_schema.py`` keeps it).

What it does is :mod:`xlsx_agent.compare`'s (:meth:`Workbook.compare
<xlsx_agent.Workbook.compare>`): ``a`` is the older workbook, ``b`` the newer, both open
documents; ``tables`` names the tables or regions to compare (each by the same name in
both; ``table_b`` gives ``b``'s name for a single one), every table and declared region of
one name in both by default; ``key`` names the key column, else each region's declared or
only key column in both (none or several: refused, ``no_key_column``, the candidates as
``valid_options``); ``columns`` limits the cells compared.

**The result.**  ``tables``: each pair compared, its addresses and ranges in both, its key
column, its rows in each, how many were paired by key and of those unchanged, and its
``totals``.  ``totals``: every difference counted by kind, always, over all of them.
``differences``: the differences themselves, in order -- sheets (``sheet``: added, removed,
renamed, by id), tables in one workbook only (``table``), columns (``column``: added,
removed, renamed), keys two rows of one workbook share (``duplicate_key``: never paired,
each workbook's rows by position, :data:`POSITIONS_SHOWN` at most, the rest counted in
``more``), rows without a key (``unkeyed``), rows added and removed (``row``, with its
key, its address by position and its cells), and cells changed (``cell``: ``changed``,
``type`` -- a number and the same number as text --, ``added``, ``removed``; with the
key, the column, and the cell in ``a`` and in ``b``, typed with their addresses).

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

from ooxml_edit.tools import CORE, Result, ToolError, array, string, tool

from ..address import TableAddress, parse_address
from ..compare import Diff, compare
from ..read.types import Value
from .docformat import KIND, handler
from .query import KEPT
from .read import _MARGIN, CURSOR, TEXT_SHOWN, _cursor, _read_cursor, _size

#: The most row positions a ``duplicate_key`` difference lists per workbook.
POSITIONS_SHOWN = 8


@tool("xlsx_compare",
      "Compare two open workbooks, a (older) and b (newer): sheets by id, tables by name, "
      "rows by key (added, removed), cells by type and value (old and new, typed, with "
      "addresses; number vs text is a type change), columns added, removed, renamed. A key "
      "two rows share is reported, never paired. totals always; paged.",
      {"a": string("Older document id."),
       "b": string("Newer document id."),
       "tables": array(string(), "Tables or regions, same name in both (t:Products). "
                       "Default every table.", optional=True),
       "table_b": string("b's name for the one table given, if it differs.", optional=True),
       "key": string("Key column. Default the declared or only key column in both.",
                     optional=True),
       "columns": array(string(), "Compare only these columns' cells.", optional=True),
       "cursor": CURSOR},
      group=CORE, kind=KIND, documents=("a", "b"))
@handler
def xlsx_compare(call: Any, a: str, b: str, tables: list[str] | None = None,
                 table_b: str | None = None, key: str | None = None,
                 columns: list[str] | None = None, cursor: str | None = None) -> Result:
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
    pairs = None if not tables else \
        ([(tables[0], table_b)] if table_b is not None else list(dict.fromkeys(tables)))
    # A cursor belongs to one comparison: the two documents as they are, and the arguments.
    versions = [call.entry.version, entry.version]
    ident = json.dumps(["compare", a, b, versions, pairs, key, columns], ensure_ascii=False)
    diff = _diff(old, new, ident, pairs, key, columns)

    items = list(_items(diff))
    total = len(items)
    start, _ = _read_cursor(cursor, ident, total, 1, f"{a} and {b}")
    chars = call.limits.max_result_chars
    limit = call.limits.max_list_items
    totals = diff.totals()
    data: dict[str, Any] = {"a": a, "b": b}
    shown_tables = [_table(t) for t in diff.tables[:limit]]
    data["tables"] = shown_tables
    data["totals"] = totals
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
          columns: list[str] | None) -> Diff:
    """The comparison, kept with the older workbook's last queries for its next pages."""
    kept = old._queries
    found = kept.get(ident)
    if found is None:
        found = compare(old, new, tables=pairs, key=key, columns=columns)
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
    out.update({"range_a": table.range_a, "range_b": table.range_b, "key": table.key,
                "rows": [table.rows_a, table.rows_b], "paired": table.paired,
                "unchanged": table.unchanged})
    totals = table.totals()
    if totals:
        out["totals"] = totals
    return out


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


def _row(where: str, row: Any, cap: int, cells: int | None) -> tuple[dict, int]:
    key, cut = _text(row.key, cap)
    out: dict[str, Any] = {"table": where, "row": row.change, "key": key}
    if cut:
        out["key_length"] = len(row.key)
    out["at"] = row.row
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
    key, cut = _text(change.key, cap)
    out: dict[str, Any] = {"table": where, "cell": change.change, "key": key}
    if cut:
        out["key_length"] = len(change.key)
    out["column"] = change.column
    for side in ("a", "b"):
        shaped, n = _value(getattr(change, side), cap)
        out[side] = shaped
        cut += n
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
_CUT_ROOM = {"differences": 10 ** 9, "chars": True, "text": 10 ** 6, "tables": 10 ** 5}


def _summary(a: str, b: str, diff: Diff, totals: dict[str, int], first: int, last: int,
             total: int, more: bool) -> str:
    text = f"{a} and {b}: {len(diff.tables)} tables compared"
    if not total:
        return text + "; no differences"
    counts = ", ".join(f"{n} {name.replace('_', ' ')}" for name, n in totals.items())
    text += f"; {counts}; differences {first}-{last} of {total}"
    return text + "; more with next_cursor" if more else text


TOOLS = [xlsx_compare]
