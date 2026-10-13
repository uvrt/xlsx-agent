"""Excel tables (ListObjects): the table parts each worksheet relates to.

A table part (ECMA-376 Part 1, §18.5: ``table``, §18.5.1.2) says where the table lies
(``ref``, header and totals rows included), how many header rows it shows
(``headerRowCount``, 1 when absent, 0 for a table without one) and how many totals rows
(``totalsRowCount``, 0 when absent), its columns in order (``tableColumn``, §18.5.1.3, each
with the unique ``name`` Excel shows in the header row, which a headerless table keeps too)
and its AutoFilter (``autoFilter``, §18.3.1.2).  Its ``displayName`` is "the name that shall
be used in formula references" and so is the name an address uses; ``name`` is the object
model's, by default the same.  Both are ST_Xstrings, as the column names are: their
``_xHHHH_`` escapes are decoded.

The table's own column names are used, never the header cells' text (ROADMAP.md, "Messy
workbooks").  What is odd about a table is said in its ``flags``, by code: a ``ref`` that
is not a range (``bad_ref``: the table is listed and cannot be read), a column count other
than the width (``columns_mismatch``: missing columns are named by their letter, as a
detected region's are), two columns with one name, case aside (``duplicate_column``: both
named by their letter).
"""

from __future__ import annotations

from dataclasses import dataclass

from ooxml_edit.charts.workbook import Area, column_letters

from .stream import _area
from .strings import decode_xstring
from .types import Problem

NS_MAIN = ("http://schemas.openxmlformats.org/spreadsheetml/2006/main",
           "http://purl.oclc.org/ooxml/spreadsheetml/main")


@dataclass(frozen=True)
class TablePart:
    """One table part as read: what :class:`~xlsx_agent.read.regions.Region` is made from.

    ``name`` is the ``displayName`` (the ``name`` when there is none); ``object_name``
    the ``name``; ``area`` the ``ref`` (``None`` when it is not a range); ``columns`` the
    column names, one per column of ``area``; ``totals`` each column's
    ``totalsRowFunction`` (``None`` when absent or ``none``); ``auto_filter`` the
    ``autoFilter``'s ``ref``; ``kind`` the ``tableType`` (``worksheet``, ``xml``,
    ``queryTable``); ``flags`` what is odd, by code."""

    name: str
    object_name: str
    id: int | None
    sheet: str
    part: str
    area: Area | None
    header_rows: int
    totals_rows: int
    columns: tuple[str, ...]
    totals: tuple[str | None, ...]
    auto_filter: Area | None
    kind: str
    flags: tuple[Problem, ...] = ()


def read_table(root, *, sheet: str, written_sheet: str, part: str) -> TablePart | None:
    """The table part whose root element is ``root`` (``None`` when it is not a
    ``table``).  ``sheet`` is the sheet's name, ``written_sheet`` the sheet as an address
    writes it (for the flags' addresses)."""
    if not isinstance(root.tag, str) or root.tag.rpartition("}")[2] != "table":
        return None
    object_name = decode_xstring(root.get("name") or "")
    name = decode_xstring(root.get("displayName") or "") or object_name
    area = _area(root.get("ref"))
    header_rows = _count(root.get("headerRowCount"), 1)
    totals_rows = _count(root.get("totalsRowCount"), 0)
    flags: list[Problem] = []
    where = f"{written_sheet}!{area.ref}" if area is not None else written_sheet
    if area is None:
        flags.append(Problem("bad_ref", written_sheet))
    names: list[str] = []
    totals: list[str | None] = []
    auto_filter = None
    for child in root:
        if not isinstance(child.tag, str):
            continue
        local = child.tag.rpartition("}")[2]
        if local == "autoFilter":
            auto_filter = _area(child.get("ref"))
        elif local == "tableColumns":
            for column in child:
                if isinstance(column.tag, str) and column.tag.rpartition("}")[2] == "tableColumn":
                    names.append(decode_xstring(column.get("name") or ""))
                    function = (column.get("totalsRowFunction") or "").strip()
                    totals.append(None if function in ("", "none") else function)
    columns: tuple[str, ...] = ()
    if area is not None:
        width = area.last.column - area.first.column + 1
        if len(names) != width:
            flags.append(Problem("columns_mismatch", where))
        names = (names + [""] * width)[:width]
        totals = (totals + [None] * width)[:width]
        columns = tuple(name_columns(names, area.first.column, flags, where,
                                     blank="columns_mismatch", duplicate="duplicate_column"))
        if header_rows + totals_rows > area.last.row - area.first.row + 1:
            flags.append(Problem("bad_ref", where))
    return TablePart(name=name, object_name=object_name, id=_int(root.get("id")),
                     sheet=sheet, part=part, area=area, header_rows=header_rows,
                     totals_rows=totals_rows, columns=columns, totals=tuple(totals),
                     auto_filter=auto_filter, kind=(root.get("tableType") or "worksheet"),
                     flags=tuple(dict.fromkeys(flags)))


def name_columns(names: list[str], first_column: int, flags: list[Problem], where: str, *,
                 blank: str, duplicate: str, addresses: list[str] | None = None) -> list[str]:
    """Column names as a region uses them: a blank name, or one that two columns share
    (case aside), becomes the column's letter (``Column F``), and a flag (``blank`` or
    ``duplicate``) is added, at ``addresses[i]`` when given, else at ``where``."""
    seen: dict[str, int] = {}
    for name in names:
        if name:
            seen[name.casefold()] = seen.get(name.casefold(), 0) + 1
    out = []
    for offset, name in enumerate(names):
        at = addresses[offset] if addresses else where
        if not name:
            flags.append(Problem(blank, at))
        elif seen[name.casefold()] > 1:
            flags.append(Problem(duplicate, at))
        else:
            out.append(name)
            continue
        out.append(f"Column {column_letters(first_column + offset)}")
    return out


def _count(value: str | None, default: int) -> int:
    number = _int(value)
    return default if number is None or number < 0 else number


def _int(value: str | None) -> int | None:
    try:
        return int(value) if value is not None else None
    except ValueError:
        return None
