"""Rows traced to the rows they rest on (round 31, R7): what ``xlsx_trace`` returns::

    found = wb.trace(["t:issues/Q-1"], depth=2)
    found.traces[0].links      # [Link(t:Items/I-1, names, column Items), ...]
    found.values("t:Items/I-1")   # {"Ref": "I-1", "Name": "pump", ...}: every value in full

Checking a row -- of a table, or of a dataset a derive made -- needs every row
it rests on, each with why, in one call instead of a read per reference.  A row's links,
each a :class:`Link` with its ``why``:

* ``source``: a dataset row's values came from it -- the sources the derive kept beside
  each value (:mod:`xlsx_agent.dataset`): a key set's keys each the row holding the key, a
  range the table rows holding it (:meth:`Workbook.rows_holding
  <xlsx_agent.workbook.Workbook.rows_holding>`), else the cells themselves (``row`` is
  then their range).  ``columns`` names the dataset's columns whose values came from it.
  A table row has no sources: it is one.
* ``names``: its text names the row's key -- the traced row's own text (``by`` is then
  ``None``), a source row's, or, ``depth`` levels deep, a named row's (``by`` the row whose
  text names it, ``column`` that row's column holding the key, ``depth`` how many steps
  from the traced row).
* ``lists``: the row's text names the key of the traced row (``by`` ``None``) or of one of
  its source rows (``by`` that row); ``column`` is the listing row's column.

**Keys and text.**  The rows linked are those of the workbook's tables -- its Excel tables
and the tables the application declared (:mod:`xlsx_agent.declare`), as ``find_text``
searches them -- each by key when its table's key column (declared, or the one column that
can be one) has a key no other row shares, else by position (``t:Items/#3``), as
:meth:`~xlsx_agent.workbook.Workbook.rows_holding` names them.  A table where several
columns could be the key (none declared) has the first of them looked for in text, its
rows named by position; one where none can has no key, and its rows only name others.  A text value names a key
when the key appears in it whole: case counts, as a key's match does, and neither end
touches a letter, a digit or ``_`` (``I-1`` is in ``"I-1; I-3"``, not in ``"I-10"``).
Only keys holding a letter and a digit are looked for in text (``I-1``, ``R2``, ``2026-A``):
a key that is a number, or a word, would be found in prose and figures everywhere.  A
declared split column's items are each such a text.  A row never links to itself.

**Bounded.**  ``depth`` (0 to :data:`MAX_DEPTH`, :data:`DEPTH` unless given) is how far
names are followed; at most :data:`MAX_LINKS` links per traced row, the rest counted in
:attr:`Trace.more`.  Values are every value in full (blank cells left out), each row's
once, as a cell read gives them (:meth:`Workbook.read <xlsx_agent.workbook.Workbook.read>`:
a column hidden from the model is never read, redacted terms replaced): what a reply has
room for is the tool's to decide (:mod:`xlsx_agent.tools.trace`).

**Cost.**  The first trace of a workbook state reads every table's rows once and keeps the
keys their text names (:class:`TraceIndex`); an edit, an undo or another declaration makes
the next trace read them again.  Measured on the test machine (a made-up workbook of 5,000
rows naming two keys each and 2,000 rows naming one, 7,000 keys): about 3.4 s for the
first trace of a workbook just opened, 1.2 s after ``describe`` (which finds the key
columns), then about 0.2 ms a row traced.  Keys are found in text by dictionary look-ups
per word, not one pattern of every key (:class:`_Finder`: 0.05 s over 7,000 texts, where a
:mod:`re` alternation of the 7,000 keys took 1.6 s).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field, replace
from typing import TYPE_CHECKING, Any, Iterable, Sequence

from .address import (DatasetAddress, RangeAddress, TableAddress, format_address,
                      parse_address)
from .read.regions import Region, RegionError

if TYPE_CHECKING:
    from .dataset import DatasetTable
    from .workbook import Workbook

#: How far ``names`` are followed unless asked: the rows the traced row names, and those
#: they name.
DEPTH = 2

#: The deepest ``names`` are followed.
MAX_DEPTH = 4

#: The most links one traced row lists; the rest are counted (:attr:`Trace.more`).
MAX_LINKS = 200

#: The most rows one trace visits while following names, past which it stops.
MAX_VISITS = 5000

#: The most keys looked for in text; past it, the first ones (in table order) only, and
#: :attr:`TraceIndex.keys_cut` says how many were not.
MAX_KEYS = 20_000

#: The most cells a source outside every table gives its values for.
MAX_SOURCE_CELLS = 50

#: Where a key may be found in text: whole, not touching a letter, digit or ``_``.
_EDGE_BEFORE = r"(?<!\w)"
_EDGE_AFTER = r"(?!\w)"
_WORD = re.compile(r"\w+")
_WORD_CHAR = re.compile(r"\w")
_LETTER = re.compile(r"[^\W\d_]")
_DIGIT = re.compile(r"\d")


@dataclass(frozen=True)
class Link:
    """One row a traced row rests on: ``row`` its address, ``why`` ``source``, ``names``
    or ``lists`` (see the module's notes), ``by`` the row the link comes through (``None``:
    the traced row), ``column`` the column whose text holds the key (``names``: ``by``'s;
    ``lists``: ``row``'s), ``columns`` (``source``) the traced row's columns whose values
    came from it, ``depth`` (``names``) its steps from the traced row."""

    row: str
    why: str
    by: str | None = None
    column: str | None = None
    columns: tuple[str, ...] = ()
    depth: int = 0

    def as_dict(self) -> dict[str, Any]:
        out: dict[str, Any] = {"row": self.row, "why": self.why}
        if self.by is not None:
            out["by"] = self.by
        if self.column is not None:
            out["column"] = self.column
        if self.columns:
            out["columns"] = list(self.columns)
        if self.why == "names":
            out["depth"] = self.depth
        return out


@dataclass
class Trace:
    """One traced row: ``row`` its address, ``links`` the rows it rests on in order --
    sources, then names by depth, then lists -- and ``more`` the links past
    :data:`MAX_LINKS`."""

    row: str
    links: list[Link] = field(default_factory=list)
    more: int = 0


@dataclass
class TraceResult:
    """What :meth:`Workbook.trace <xlsx_agent.workbook.Workbook.trace>` found: the
    ``traces`` in the order asked, and every row's values by address (:meth:`values`)."""

    traces: list[Trace]
    rows: dict[str, dict[str, Any]]
    keys_cut: int = 0

    def values(self, address: str) -> dict[str, Any]:
        """A row's values, column by column, every one in full, blanks left out."""
        return self.rows[address]


@dataclass
class _Row:
    """A table row the index holds: its address, the whole table's, its position, its key
    (``None``: none, or not one to look for), its values and the keys its text names."""

    address: str
    table: str
    position: int
    key: str | None
    values: dict[str, Any]
    names: list[tuple[str, str]] = field(default_factory=list)   # (column, key)


class TraceIndex:
    """Every table row of a workbook state, with the keys its text names
    (:meth:`Workbook.trace <xlsx_agent.workbook.Workbook.trace>` keeps one per state)."""

    def __init__(self, workbook: "Workbook") -> None:
        self.workbook = workbook
        self.rows: dict[str, _Row] = {}
        #: (table, position) -> row; (table, key) -> row, for a key one row has.
        self.at: dict[tuple[str, int], _Row] = {}
        self.keyed: dict[tuple[str, str], _Row] = {}
        #: key -> the rows holding it, and the rows (with the column) whose text names it.
        self.holding: dict[str, list[_Row]] = {}
        self.listing: dict[str, list[tuple[_Row, str]]] = {}
        self.keys_cut = 0
        regions = workbook.regions
        tables = [regions.with_key(r) for r in [*workbook.tables, *regions.declared]]
        read: list[tuple[Region, list[_Row]]] = []
        for region in tables:
            read.append((region, self._read(region)))
        looked = [k for k in self.holding if _findable(k)]
        if len(looked) > MAX_KEYS:
            self.keys_cut = len(looked) - MAX_KEYS
            looked = looked[:MAX_KEYS]
        self.finder = _Finder(looked)
        for _, rows in read:
            for row in rows:
                row.names = self.names_in(row.values, own=row.key)
                for column, key in row.names:
                    self.listing.setdefault(key, []).append((row, column))

    def _read(self, region: Region) -> list[_Row]:
        whole = region.address
        table = whole.text
        regions = self.workbook.regions
        named = True        # its rows named by key (its key column is known)
        try:
            index = regions.key_index(region)
        except (RegionError, KeyError):
            index = None
            # Several columns could be its key: the first is looked for in text, its rows
            # named by position (an address by key would need key=).
            candidates = regions.key_columns(region)
            if candidates:
                index, named = regions.key_index(region, candidates[0]), False
        out: list[_Row] = []
        if region.data is None:
            return out
        columns = list(region.columns)
        for position, cells in enumerate(self.workbook.read(whole), 1):
            key = index.keys[position - 1] if index is not None and position <= len(
                index.keys) else None
            by_key = named and key is not None and index is not None and index.unique
            address = (replace(whole, key=key) if by_key
                       else replace(whole, position=position)).text
            values = {name: _plain(value) for name, value in zip(columns, cells)
                      if value.t != "blank" and value.v is not None}
            row = _Row(address, table, position, key, values)
            out.append(row)
            self.rows[address] = row
            self.at[(table, position)] = row
            if key is not None:
                self.holding.setdefault(key, []).append(row)
                if index is not None and key not in index.duplicates and named:
                    self.keyed[(table, key)] = row
        return out

    def names_in(self, values: dict[str, Any], own: str | None = None) -> list[tuple[str, str]]:
        """The keys the text of ``values`` names, each with its column, in column order
        and each once per column; ``own`` (the row's key) is never one."""
        out: list[tuple[str, str]] = []
        if not self.finder:
            return out
        for column, value in values.items():
            for text in (value if isinstance(value, list) else [value]):
                if not isinstance(text, str):
                    continue
                for key in self.finder.keys_in(text):
                    if key != own and (column, key) not in out:
                        out.append((column, key))
        return out


def _findable(key: str) -> bool:
    return len(key) >= 2 and bool(_LETTER.search(key)) and bool(_DIGIT.search(key))


class _Finder:
    """Finds keys in text, whole (neither end touching a letter, digit or ``_``).  A key
    that starts with a word character is looked up by the word it starts with and its
    length: where that word starts in the text, the text's next characters of each length
    the word's keys have are looked up, then no word character may follow -- a few
    dictionary look-ups per word of text, however many keys (an alternation of thousands
    of keys is slow in :mod:`re`).  The few that start otherwise are one pattern."""

    def __init__(self, keys: Sequence[str]) -> None:
        #: first word -> (lengths, longest first; the keys)
        self.by_word: dict[str, tuple[list[int], set[str]]] = {}
        other: list[str] = []
        for key in set(keys):
            first = _WORD.match(key)
            if first is None:
                other.append(key)
                continue
            lengths, held = self.by_word.setdefault(first.group(0), ([], set()))
            held.add(key)
            if len(key) not in lengths:
                lengths.append(len(key))
        for lengths, _ in self.by_word.values():
            lengths.sort(reverse=True)
        other.sort(key=lambda k: (-len(k), k))
        self.other = (re.compile(_EDGE_BEFORE + "(?:" + "|".join(map(re.escape, other)) + ")"
                                 + _EDGE_AFTER) if other else None)

    def __bool__(self) -> bool:
        return bool(self.by_word) or self.other is not None

    def keys_in(self, text: str) -> list[str]:
        """The keys in ``text``, in order (the longest one where two start alike)."""
        found: list[tuple[int, str]] = []
        size = len(text)
        for word in _WORD.finditer(text):
            held = self.by_word.get(word.group(0))
            if held is None:
                continue
            start = word.start()
            for length in held[0]:
                end = start + length
                if end > size or (end < size and _WORD_CHAR.match(text, end)):
                    continue
                if text[start:end] in held[1]:
                    found.append((start, text[start:end]))
                    break
        if self.other is not None:
            found += [(m.start(), m.group(0)) for m in self.other.finditer(text)]
            found.sort()
        return [key for _, key in found]


def _plain(value: Any) -> Any:
    """A cell's value as a trace gives it: a declared split's items as a list."""
    return list(value.v) if value.t == "list" else value.v


def index_of(workbook: "Workbook") -> TraceIndex:
    """The workbook's :class:`TraceIndex` for its state now, made once per state."""
    stamp = (tuple(id(r) for r in workbook.regions.declared),
             tuple(workbook.regions.declared_keys.items()), id(workbook.regions))
    kept = workbook._traces
    if kept is not None and kept[0] == stamp:
        return kept[1]
    made = TraceIndex(workbook)
    workbook._traces = (stamp, made)
    return made


class Tracer:
    """Traces rows of one workbook state (see the module's notes)."""

    def __init__(self, workbook: "Workbook", depth: int = DEPTH) -> None:
        if not 0 <= depth <= MAX_DEPTH:
            raise ValueError(f"depth is 0 to {MAX_DEPTH}, not {depth}")
        self.workbook = workbook
        self.depth = depth
        self.index = index_of(workbook)
        #: Every row's values the traces name, by address (cells outside tables too).
        self.values: dict[str, dict[str, Any]] = {}

    # -- the rows asked for ------------------------------------------------------------------

    def table_rows(self, address: str) -> list[str]:
        """The addresses of the table rows ``address`` names -- a table, rows by key or
        position, a key set, a row range, or a sheet's cells inside one table -- in order.
        :class:`RegionError` ``not_rows`` for a column or a table outside the index."""
        target = parse_address(address)
        if isinstance(target, RangeAddress):
            held = self.workbook.rows_holding(target)
            if held is None:
                raise RegionError("not_rows", f"{address} lies in no one table's data rows: "
                                  "name a table's rows (t:Items/I-4)", address=address)
            target = held
        if not isinstance(target, TableAddress):
            raise RegionError("not_rows", f"{format_address(target)} names no rows: name a "
                              "table's (t:Items/I-4) or a dataset's (d:by_owner/#2)",
                              address=address)
        selection = self.workbook.resolve(target)
        if selection.kind in ("column", "value") and selection.column is not None \
                and selection.row is None:
            raise RegionError("not_rows", f"{selection.address} is a column: name rows "
                              "(the table, a key, a key set or #2:#9)", address=address)
        region = self.workbook.regions.with_key(selection.region)
        table = region.address.text
        if selection.positions:
            positions = list(selection.positions)
        elif selection.row is not None:
            positions = [selection.row]
        else:
            positions = list(range(1, region.rows + 1))
        out = []
        for position in positions:
            row = self.index.at.get((table, position))
            if row is None:
                raise RegionError("not_rows", f"{table} is not a table rows are traced in: "
                                  "an Excel table or a declared one", address=address)
            out.append(row.address)
        return out

    # -- tracing -----------------------------------------------------------------------------

    def trace_row(self, address: str) -> Trace:
        """A table row's trace (an address :meth:`table_rows` gave)."""
        row = self.index.rows[address]
        self.values.setdefault(address, row.values)
        trace = Trace(address)
        links = _Links(trace, {address})
        self._names(links, [(None, row.names)])
        self._lists(links, [(None, row.key)])
        return links.done()

    def trace_dataset_row(self, table: "DatasetTable", position: int) -> Trace:
        """Row ``position`` (from 1) of a dataset: its values, its sources' rows, and the
        names and lists of its text and of its source rows."""
        from .dataset import source_units

        address = DatasetAddress(table.id, position=position).text
        values = {}
        for offset, name in enumerate(table.columns):
            value = table.value(position, offset)
            if value.v is not None and value.t != "blank":
                values[name] = value.v
        self.values.setdefault(address, values)
        trace = Trace(address)
        links = _Links(trace, {address})
        sources: dict[str, list[str]] = {}
        for offset, name in enumerate(table.columns):
            for column, text in source_units(table.cell_sources(position, offset)):
                for found in self._source_rows(column, text):
                    columns = sources.setdefault(found, [])
                    if name not in columns:
                        columns.append(name)
        for found, columns in sources.items():
            links.add(Link(found, "source", columns=tuple(columns)))
            links.visited.add(found)
            if found in self.index.rows:
                self.values.setdefault(found, self.index.rows[found].values)
        rows = [self.index.rows[a] for a in sources if a in self.index.rows]
        own = None
        if table.keys:
            own = key_text_of(values.get(table.columns[0]))
        self._names(links, [(None, self.index.names_in(values, own=own))]
                    + [(r.address, r.names) for r in rows])
        self._lists(links, [(None, own)] + [(r.address, r.key) for r in rows])
        return links.done()

    def _source_rows(self, column: Any, text: str) -> list[str]:
        """The rows (or cells) one source unit names: a key of a table's column, or an
        address as a derive wrote it."""
        if column is not None:      # a key set's key: the row holding it
            region = self.workbook.regions.with_key(self.workbook.regions.region(
                replace(column, column=None, any_column=False, key=None, keys=None)))
            row = self.index.keyed.get((region.address.text, text))
            return [row.address] if row is not None else [self._cells(
                replace(column, keys=None, key=text).text)]
        try:
            target = parse_address(text)
        except Exception:  # noqa: BLE001 -- a source the reader cannot parse is kept as text
            return [text]
        if isinstance(target, TableAddress):
            try:
                return self.table_rows(replace(target, column=None, any_column=False).text)
            except (RegionError, KeyError):
                return [text]
        if isinstance(target, RangeAddress):
            held = self.workbook.rows_holding(target)
            if held is not None:
                try:
                    return self.table_rows(held.text)
                except (RegionError, KeyError):
                    pass
            return [self._cells(text)]
        return [text]

    def _cells(self, text: str) -> str:
        """A source outside every table: its cells' values, by address (the first
        :data:`MAX_SOURCE_CELLS`)."""
        if text not in self.values:
            values: dict[str, Any] = {}
            try:
                for cells in self.workbook.read(text):
                    for value in cells:
                        if len(values) >= MAX_SOURCE_CELLS:
                            break
                        if value.t != "blank" and value.v is not None:
                            values[value.a] = _plain(value)
            except (RegionError, KeyError, ValueError):
                pass
            self.values[text] = values
        return text

    def _names(self, links: "_Links", starts: list[tuple[str | None, list[tuple[str, str]]]]
               ) -> None:
        frontier = starts
        for depth in range(1, self.depth + 1):
            reached: list[tuple[str | None, list[tuple[str, str]]]] = []
            for by, names in frontier:
                for column, key in names:
                    for row in self.index.holding.get(key, ()):
                        if row.address in links.visited:
                            continue
                        if not links.add(Link(row.address, "names", by=by, column=column,
                                              depth=depth)):
                            return
                        links.visited.add(row.address)
                        self.values.setdefault(row.address, row.values)
                        reached.append((row.address, row.names))
            frontier = reached
            if not frontier:
                return

    def _lists(self, links: "_Links", owners: list[tuple[str | None, str | None]]) -> None:
        own = links.trace.row
        for by, key in owners:
            if key is None:
                continue
            for row, column in self.index.listing.get(key, ()):
                if row.address in (own, by):
                    continue
                if links.add(Link(row.address, "lists", by=by, column=column)):
                    self.values.setdefault(row.address, row.values)


def key_text_of(value: Any) -> str | None:
    """A dataset value as a key: its text, or a whole number's digits."""
    if isinstance(value, str):
        return value or None
    if isinstance(value, bool):
        return None
    if isinstance(value, int) or (isinstance(value, float) and value.is_integer()):
        return str(int(value))
    return None


class _Links:
    """A trace's links as they are found: each once, at most :data:`MAX_LINKS`, the rest
    counted; ``visited`` the rows names reached (and the traced row)."""

    def __init__(self, trace: Trace, visited: set[str]) -> None:
        self.trace = trace
        self.visited = visited
        self.seen: set[tuple[Any, ...]] = set()
        self.found = 0

    def add(self, link: Link) -> bool:
        """Add ``link`` (once); ``False`` once the trace has visited :data:`MAX_VISITS`
        rows, to stop."""
        key = (link.row, link.why, link.by, link.column)
        if key in self.seen:
            return True
        self.seen.add(key)
        self.found += 1
        if len(self.trace.links) < MAX_LINKS:
            self.trace.links.append(link)
        else:
            self.trace.more += 1
        return self.found < MAX_VISITS

    def done(self) -> Trace:
        return self.trace


def trace(workbook: "Workbook", rows: Iterable[str], *, depth: int = DEPTH) -> TraceResult:
    """:meth:`Workbook.trace <xlsx_agent.workbook.Workbook.trace>`: the table rows
    ``rows`` names, each traced."""
    tracer = Tracer(workbook, depth)
    traces = [tracer.trace_row(address) for given in rows
              for address in tracer.table_rows(given)]
    return TraceResult(traces, tracer.values, tracer.index.keys_cut)


__all__ = ["DEPTH", "MAX_DEPTH", "MAX_LINKS", "Link", "Trace", "TraceIndex", "TraceResult",
           "Tracer", "trace"]
