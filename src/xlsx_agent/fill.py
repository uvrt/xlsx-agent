"""Fills: a column's blank cells given the value of another column of the same row, each
filled value with the cells it came from and marked as filled (round 35, R22;
:mod:`xlsx_agent.derive`'s ``fill``)::

    ds = wb.derive("t:Items", fill=[{"column": "Limit", "from": "Default limit",
                                     "coerce": "number"}], id="limits")
    ds.rows[1]["Limit"]                     # 40.0: the row's Default limit, Limit blank
    ds.provenance[1]["Limit"].filled        # "Default limit": marked, with its column
    ds.sources[...]["Limit"]                # the Default limit cell it came from
    ds.filled[0].as_dict()                  # {"column": "Limit", "from": "Default limit",
                                            #  "filled": {"count": 3, ...}, "unfilled": ...}

**Which cells.**  A cell of ``column`` is filled when it is blank -- no value, or text of
no characters (a formula's ``""``), which a person sees as blank too -- and the same row's
``from`` cell has a value.  Every other cell of ``column`` keeps its own value, whatever
``from`` holds: a fill never overwrites.  A ``from`` cell that is blank, an error or a
formula without a cached value fills nothing: the cell stays blank and is counted as
``unfilled`` by that reason (``blank``, ``error``, ``no_cached_value``), with its first
addresses (the cells of ``column`` left blank).

**coerce** (``"number"`` or ``"date"``) asks for values of that type, as a derive's
``coerce`` reads them (:mod:`xlsx_agent.derive`): a number (or percent) cell is taken as it
is, and text that reads as a number in the workbook's locale as that number
(:func:`~xlsx_agent.read.types.text_number`: ``number_as_text``, what a declared column's
cleaned text is flagged as too, :mod:`xlsx_agent.clean`); a date or datetime cell as it is,
and ISO 8601 text, or a date in the locale's day order, as that date
(:func:`~xlsx_agent.read.types.text_date`).  The values read so are counted (``coerced``,
with their first addresses).  A ``from`` value of another type fills nothing: counted as
``unfilled`` ``wrong_type``, never guessed.  Without ``coerce`` a value is filled as it is
(text stays text: a later ``coerce`` of the derive reads it as a number, as any text cell
that reads as one).  ``number_as_text`` and ``date_as_text`` are taken for ``number`` and
``date``, as the derive's ``coerce`` writes them.

**Several fills** apply in order, each to the rows as the ones before left them: two
fills of one column from two others take the first column's value, else the second's
(``[{"column": "Limit", "from": "Team limit"}, {"column": "Limit", "from": "Default
limit"}]``).  A fill from a column that an earlier fill filled carries that cell's sources.

**From a looked-up table.**  A ``join`` comes first (:mod:`xlsx_agent.join`): a fill's
``from`` (or ``column``) may be a joined column, so a blank is filled from another table's
row with the same key -- join that table's column, then fill from it.  A joined column of
several matches is a list, and fills as one (no ``coerce`` reads a list).

**Provenance.**  A filled value's sources are the ``from`` cell's (by key when the table
has a unique key column, as a derive writes them; a joined value's, the joined table's
cells), never the blank cell's; it is marked with the column it came from
(:attr:`Provenance.filled <xlsx_agent.derive.Provenance.filled>`, ``filled`` on its value
in a tool's result and in a read of the dataset).  A value derived from filled values
(a sum of a filled column) lists the ``from`` cells among its sources; :attr:`Filled`
counts what each fill did, over every row (before ``where``).

**Then the pipeline.**  Fills come after the join and before everything else: ``where``
filters the filled rows, ``group_by``, ``bins``, ``pivot`` and the measures read them.
Without ``group_by``, ``measures``, ``bins`` or ``pivot`` the result is the rows themselves,
as a join's (``where``, ``sort`` and ``top`` apply; ``round`` is refused).

**Bounded.**  The rows are read once, whole, as a join's: at most
:data:`~xlsx_agent.join.MAX_JOINED` values (rows times columns), else
:class:`~xlsx_agent.derive.DeriveError` ``too_many_cells``.

**Content.**  Errors name columns and keys of the fill, never a cell's value.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, Mapping, Sequence, cast

from ooxml_common.locales import EN_US

from .address import NameNotFound, nearest
from .dataset import DatasetTable
from .join import MAX_JOINED, JoinedTable, JoinedValue, _date1904, _keyed, _moved, _read
from .profile import FIRST_ADDRESSES, Count
from .read.regions import Region
from .read.types import Value, text_date, text_number

if TYPE_CHECKING:  # pragma: no cover
    from .workbook import Workbook

#: What a fill takes.
FILL_KEYS = ("column", "from", "coerce")

#: The types a fill's ``coerce`` asks for.
FILL_TYPES = ("number", "date")

#: The derive's ``coerce`` names, taken for the type they read.
_ALIASES = {"number_as_text": "number", "date_as_text": "date"}

#: Why a blank cell was left blank: its ``from`` cell's.
UNFILLED = ("blank", "error", "no_cached_value", "wrong_type")

#: The most fills in one derive.
MAX_FILLS = 20

_TYPES_OF = {"number": ("number", "percent"), "date": ("date", "datetime")}


class _Counter:
    __slots__ = ("count", "first")

    def __init__(self) -> None:
        self.count = 0
        self.first: list[str] = []

    def add(self, address: str) -> None:
        self.count += 1
        if len(self.first) < FIRST_ADDRESSES:
            self.first.append(address)

    def done(self) -> Count:
        return Count(self.count, tuple(self.first))


@dataclass(frozen=True)
class Filled:
    """What one fill did: ``column`` the column filled, ``source`` the column it was filled
    from, ``coerce`` the type asked for (``None``: as they are); ``filled`` the cells given
    a value, ``coerced`` those of them read from text by ``coerce``, ``unfilled`` the blank
    cells left blank, by reason (their ``from`` cell's: ``blank``, ``error``,
    ``no_cached_value``, ``wrong_type``), each with its first addresses (the cells of
    ``column``); ``note`` what made it do nothing, if anything did."""

    column: str
    source: str
    coerce: str | None
    filled: Count
    coerced: Count = field(default_factory=lambda: Count(0))
    unfilled: dict[str, Count] = field(default_factory=dict)
    note: str | None = None

    def as_dict(self) -> dict:
        out: dict[str, Any] = {"column": self.column, "from": self.source}
        if self.coerce is not None:
            out["coerce"] = self.coerce
        out["filled"] = self.filled.as_dict()
        if self.coerced.count:
            out["coerced"] = self.coerced.as_dict()
        if self.unfilled:
            out["unfilled"] = {k: c.as_dict() for k, c in self.unfilled.items()}
        if self.note is not None:
            out["note"] = self.note
        return out


class FilledTable(JoinedTable):
    """A table's (or a joined table's, or a dataset's) rows with blank cells filled: what
    the derive engine reads, as it reads a join (:class:`~xlsx_agent.join.JoinedTable`).
    ``filled`` what each fill did; ``marks`` per row position (from 1) and column offset
    (from 0), the column a filled value came from."""

    def __init__(self, columns: Sequence[str], values: list[list[JoinedValue]],
                 sources: list[list[tuple[str, ...]]], address: Any, keys: int,
                 date1904: bool, joined: Any, filled: list[Filled],
                 marks: dict[tuple[int, int], str]) -> None:
        super().__init__(columns, values, sources, address, keys, date1904, joined)
        self.filled = filled
        self.marks = marks


def _bad(message: str, **options: Any) -> Exception:
    from .derive import DeriveError

    return DeriveError("bad_fill", message, field="fill", **options)


def fills(fill: Any) -> list[dict[str, Any]]:
    """The caller's fills, checked for their form: a list of ``{column, from, coerce}``
    (one such mapping is taken as a list of one); placeholders a strict call sends (an
    empty list, an entry with neither column nor from) are dropped, ``coerce`` ``""`` is
    none, and the derive's ``number_as_text`` and ``date_as_text`` are ``number`` and
    ``date``."""
    if fill is None:
        return []
    if isinstance(fill, Mapping):
        fill = [fill]
    if isinstance(fill, (str, bytes)) or not isinstance(fill, Sequence):
        raise _bad("fill is a list of {column: the column whose blanks are filled, from: the "
                   "column of the same row they take, coerce: number or date}")
    out = []
    for n, item in enumerate(fill):
        if not isinstance(item, Mapping) or set(item) - set(FILL_KEYS):
            raise _bad(f"fill[{n}] is {{column, from, coerce}}: column the one whose blanks "
                       "are filled, from the column of the same row they take",
                       valid_options=FILL_KEYS)
        column, source = item.get("column") or None, item.get("from") or None
        if column is None and source is None:
            continue                                   # a strict call's placeholder
        if not isinstance(column, str) or not isinstance(source, str):
            raise _bad(f"fill[{n}] needs both column (whose blanks are filled) and from "
                       "(the column of the same row they take), as names: "
                       "{column: \"Limit\", from: \"Default limit\"}")
        coerce = item.get("coerce") or None
        coerce = _ALIASES.get(coerce, coerce) if isinstance(coerce, str) else coerce
        if coerce is not None and coerce not in FILL_TYPES:
            raise _bad(f"fill[{n}]'s coerce is {' or '.join(FILL_TYPES)} (the type the "
                       "filled values are read as), or none: as they are",
                       valid_options=FILL_TYPES)
        out.append({"column": column, "from": source, "coerce": coerce})
    if len(out) > MAX_FILLS:
        raise _bad(f"at most {MAX_FILLS} fills in one derive")
    return out


def _offset(table: "Region | DatasetTable", name: str, n: int, key: str) -> int:
    try:
        return table.column_index(name)
    except NameNotFound:
        raise _bad(f"fill[{n}]'s {key} names a column of the rows filled (after a join, its "
                   f"columns too); there are {len(table.columns)}: "
                   f"{', '.join(table.columns[:50])}",
                   valid_options=nearest(name, table.columns) or list(table.columns)[:50]
                   ) from None


def _is_blank(value: Value) -> bool:
    return (value.t == "blank" and value.issue != "no_cached_value") \
        or (value.t == "text" and value.v == "")


def filled_table(workbook: "Workbook | None", table: "Region | DatasetTable",
                 fill: Any) -> "FilledTable | Region | DatasetTable":
    """``table``'s rows with the blank cells ``fill`` names filled (see the module's notes);
    ``table`` itself when ``fill`` holds no fill."""
    wanted = fills(fill)
    if not wanted:
        return table
    plan = [(f, _offset(table, f["column"], n, "column"), _offset(table, f["from"], n, "from"))
            for n, f in enumerate(wanted)]
    rows, sources = _read(workbook, table)
    held = len(rows) * len(table.columns)
    if held > MAX_JOINED:
        from .derive import DeriveError

        raise DeriveError("too_many_cells", f"a fill of {held} values, more than "
                          f"{MAX_JOINED}: fill a smaller table (a derived dataset)",
                          field="fill")
    if workbook is not None:
        locale, grouping = workbook.locale, workbook.grouping
    else:
        locale, grouping = EN_US, True
    values: list[list[JoinedValue]] = [
        [v if isinstance(v, JoinedValue) else _moved(v, p, o) for o, v in enumerate(row)]
        for p, row in enumerate(rows, 1)]
    where = [list(row) for row in sources]
    marks: dict[tuple[int, int], str] = {}
    done: list[Filled] = []
    for item, target, source in plan:
        column, origin = table.columns[target], table.columns[source]
        coerce = item["coerce"]
        filled, coerced = _Counter(), _Counter()
        unfilled: dict[str, _Counter] = {}
        note = None
        if target == source:
            note = f"{column} is filled from itself (names match case aside): nothing to fill"
            joined = getattr(table, "joined", None)
            renamed = [n for n, c in (joined.columns.items() if joined is not None else ())
                       if c.casefold() == item["from"].casefold() and n != c]
            if renamed:
                note += f"; the joined {item['from']} is {renamed[0]}"
        for position, row in enumerate(values, 1):
            cell = row[target]
            if note is not None or not _is_blank(cell):
                continue
            given = row[source]
            reason, v, t = _taken(given, coerce, locale, grouping, workbook is None)
            if reason is not None:
                unfilled.setdefault(reason, _Counter()).add(cell.a)
                continue
            if t != given.t:
                coerced.add(cell.a)
            row[target] = JoinedValue(
                sheet="", row=position, column=target + 1, v=v, t=t, formula=given.formula,
                issue=given.issue if t == given.t else None,
                hidden=given.hidden or cell.hidden, cache=given.cache,
                external=given.external, cleaned=given.cleaned, at=cell.a)
            where[position - 1][target] = tuple(where[position - 1][source])
            marks[(position, target)] = marks.get((position, source), origin)
            filled.add(cell.a)
        done.append(Filled(column, origin, coerce, filled.done(), coerced.done(),
                           {r: unfilled[r].done() for r in UNFILLED if r in unfilled}, note))
    if isinstance(table, JoinedTable):
        keys, joined = table.left_keys, table.joined
    elif isinstance(table, DatasetTable):
        keys, joined = table.keys or 0, None
    else:
        key = _keyed(workbook, table)
        keys, joined = (1 if key is not None and table.columns[0] == key else 0), None
    out = FilledTable(list(table.columns), values, where, table.address, keys,
                       _date1904(workbook, table), joined, done, marks)
    # Round 36: a filled column keeps its own number format (each cell its column's).
    from .dataset import column_formats

    out.column_formats = column_formats(workbook, table, table.columns)
    return out


def _taken(value: Value, coerce: str | None, locale: Any, grouping: bool,
           workbook_less: bool) -> tuple[str | None, Any, str]:
    """``(reason, v, t)``: why ``value`` fills nothing (``None``: it fills), else the value
    and type it fills with."""
    if value.issue == "no_cached_value":
        return "no_cached_value", None, "blank"
    if _is_blank(value):
        return "blank", None, "blank"
    if value.t == "error":
        return "error", None, "blank"
    if coerce is None or value.t in _TYPES_OF[coerce]:
        return None, value.v, value.t
    if value.t == "text":
        text = cast(str, value.v)
        # a cell's text as the reader flags it; a dataset's (no workbook) as it reads
        if coerce == "number" and (value.issue == "number_as_text" or workbook_less):
            number = text_number(text, locale, grouping=grouping)
            if number is not None:
                return None, number, "number"
        if coerce == "date" and value.issue == "date_as_text":
            shown = text_date(text, locale, local_dates=True)
            if shown is not None:
                return None, shown, "datetime" if "T" in shown else "date"
    return "wrong_type", None, "blank"


__all__ = ["FILL_KEYS", "FILL_TYPES", "MAX_FILLS", "UNFILLED", "Filled", "FilledTable",
           "filled_table", "fills"]
