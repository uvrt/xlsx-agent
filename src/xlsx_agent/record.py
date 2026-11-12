"""Values a caller judged, one per row, kept as a dataset traced to the cells judged (round 53,
R108): what ``xlsx_record_values`` keeps as ``d:<id>``::

    recorded = wb.record_values("t:Shelf", key="Ref", source="notes", column="outcome",
                                values=[{"key": "S-01", "value": "done"},
                                        {"key": "S-02", "value": "partly done",
                                         "note": "says 'en partie'"}])
    recorded.dataset            # columns (Ref, outcome, note), a row per key recorded
    recorded.dataset.sources[0][1]    # ("t:Shelf/S-01/notes",): the cell judged
    recorded.not_recorded       # rows with text in notes and no value given

An agent reads a column of free text (``notes``, worded many ways), decides a value for
each row (``done``, ``partly done``) and wants it kept like any figure: in a dataset whose
every value names where it came from.  The library cannot judge text; it can keep the
judgement honest: one row per key, each value tied to the source cell it was judged from,
and every row whose text was not judged listed.

**What is kept.**  A row per recorded key, in the table's row order (not the order given):
the key column's value (its name and type as in the table), ``column`` (the value as given,
text; with ``type="number"`` a number) and ``note`` (the caller's short reason, blank when
none).  ``keys`` is 1, so a row is read by key (``d:outcome/S-01``) as a derive's grouped
rows are.  Sources: the key's own cell for the key, the source cell for the value and the
note -- by key where the table has a key column (``t:Shelf/S-01/notes``), else the cell's
address, as :func:`~xlsx_agent.dataset.table_dataset` writes them -- so ``xlsx_trace`` of
a row reaches the row judged, and a value placed in a deck or document cites the cell.
``Dataset.detail["recorded"]`` says how it was recorded (table, key, source, column,
type), which ``add_to`` must match.

**Repeated keys** (round 54, R113).  Workbooks repeat a reference on a section heading, a
subtotal or a spacer row beside the item row that holds the text.  A key is matched among
the rows with text in ``source`` only: rows without text that share it are left out (counted
in :attr:`Recorded.notes`, listed in :attr:`Recorded.ignored`); a key named that two rows
with text hold is refused (``key_shared``, both rows named), and those rows are listed in
:attr:`Recorded.shared` -- the rest of the call is never refused for them.

**Numbers** (round 54).  ``type="number"`` takes JSON numbers and strings that are plain
numbers as code writes them (``3``, ``2.5``, ``-1e3``: no grouping, no decimal comma), kept
as numbers in a ``number`` column, so a derive averages them; anything else is refused
(``not_number``).

**Adding** (round 54).  ``add_to``, a dataset this function made with the same table, key,
source, column and type (else ``add_mismatch``), is extended: the union of its rows and the
call's.  A key it holds sent again with the same value is no error (its note replaced when
one is given); with another value it is refused (``already_recorded``), naming the key.  A
key it holds that no longer names one row with text (the workbook changed) refuses the
call (``add_stale``): record every value again without it.

**Refused before anything is kept** (:class:`RecordError`, every problem of the call at
once, each named by its place in ``values`` and the row by position): a key no row has
(``unknown_key``; the rows whose key is loosely equal -- ooxml-edit's
:func:`~ooxml_edit.tools.loose_key`: case, white space, separators, leading zeros -- are
suggested as ``{"row": "t:Shelf/#4/Ref", "key": "S-04"}``, the key as its cell holds it), a
key given twice (``key_twice``), a key two rows with text hold (``key_shared``), a value
whose source cell is empty (``empty_source``: nothing was there to judge), an empty value
(``empty_value``: leave the key out instead), a value that is not a number
(``not_number``), a key recorded earlier with another value (``already_recorded``).  The
problems carry key cells' text (the model reads them); the message names places and rows
only, never a key, a value or source text.  The output ``column`` may not be the key
column's name or ``note``.

**Listed, never refused.**  :attr:`Recorded.not_recorded`: the rows whose source cell holds
text (anything but blank or white space) and got no value, by key, in row order
(:attr:`Recorded.repeated`: those of their keys that another row holds too);
:attr:`Recorded.no_key`: rows with source text but no key (a blank key cell), by position.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field, replace
from typing import TYPE_CHECKING, Any, Mapping, Sequence

from ooxml_edit.tools import Dataset, loose_key

from .address import TableAddress, parse_address
from .dataset import table_dataset

if TYPE_CHECKING:
    from .workbook import Workbook

#: The name of the column holding each value's note.
NOTE = "note"

#: The most problems a refusal names; the rest are counted.
NAMED = 20

#: The most rows suggested for one unknown key, and named for a key held by several.
SUGGESTED = 3

#: The types of value recorded: text as given, or numbers only.
TYPES = ("text", "number")

#: The key of :attr:`Dataset.detail` saying how a dataset was recorded (what ``add_to``
#: must match): table, key, source, column, type.
DETAIL = "recorded"


class RecordError(ValueError):
    """A ``record_values`` call refused before anything was kept: ``code`` the first
    problem's kind (``unknown_key``, ``key_twice``, ``key_shared``, ``empty_source``,
    ``empty_value``, ``not_number``, ``already_recorded``, ``add_mismatch``, ``add_stale``,
    ``bad_column``, ``bad_type``), ``field`` the argument at fault, ``problems`` every
    problem as ``{"code", "item", "row"}`` (``item`` the place in ``values``, from 0;
    ``row`` the table row by position) with ``like`` (``[{"row", "key"}]``, rows whose key
    is loosely equal and that key) for an unknown key, ``rows`` for a shared key and
    ``key`` for a shared or already recorded one; ``valid_options`` the rows to read.  The
    message names places and rows, never a key or a value."""

    def __init__(self, code: str, message: str, *, field: str = "values",
                 problems: Sequence[Mapping[str, Any]] = (),
                 valid_options: Sequence[str] = ()) -> None:
        super().__init__(message)
        self.code = code
        self.message = message
        self.field = field
        self.problems = [dict(p) for p in problems]
        self.valid_options = list(valid_options)


@dataclass(frozen=True)
class Recorded:
    """What :func:`record_values` kept: ``dataset`` (ooxml-edit's :class:`Dataset`),
    ``address`` the table's, ``key``/``source``/``column`` the column names,
    ``not_recorded`` the rows with source text and no value as ``(key, source cell)``,
    ``no_key`` the rows with source text and no key (``t:Shelf/#7``), ``counts`` how many
    rows got each value, in the order first met; ``notes`` what was read as meant;
    ``added`` the keys new to the dataset and ``kept`` those kept from ``add_to``;
    ``shared`` the rows with text whose key another such row holds, ``ignored`` the rows
    without text that share a recorded key (both by position), ``repeated`` the keys of
    ``not_recorded`` another row holds too."""

    dataset: Dataset
    address: str
    key: str
    source: str
    column: str
    not_recorded: tuple[tuple[str, str], ...] = ()
    no_key: tuple[str, ...] = ()
    counts: Mapping[Any, int] = field(default_factory=dict)
    notes: tuple[str, ...] = ()
    added: int = 0
    kept: int = 0
    shared: tuple[str, ...] = ()
    ignored: tuple[str, ...] = ()
    repeated: frozenset[str] = frozenset()

    @property
    def recorded(self) -> int:
        return len(self.dataset.rows)


def _text(value: Any) -> str | None:
    """A given key or value as text: a whole number's digits, as a key cell's are
    (:func:`~xlsx_agent.read.regions.key_text`); ``None`` for nothing."""
    if value is None:
        return None
    if isinstance(value, float) and not isinstance(value, bool) and value.is_integer():
        return str(int(value))
    return value if isinstance(value, str) else str(value)


def _blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _number(value: Any) -> int | float | None:
    """A value given for ``type="number"`` as a number: a JSON number, or a string that is a
    plain number as written in code (``3``, ``-2.5``, ``1e3``; no grouping, no comma
    decimal); ``None`` for anything else."""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return value if math.isfinite(value) else None
    if isinstance(value, str) and _PLAIN_NUMBER.fullmatch(value.strip()):
        text = value.strip()
        number = int(text) if _WHOLE.fullmatch(text) else float(text)
        return number if math.isfinite(number) else None
    return None


_PLAIN_NUMBER = re.compile(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?")
_WHOLE = re.compile(r"[+-]?\d+")


def record_values(workbook: "Workbook", table: "str | TableAddress", *, key: str, source: str,
                  column: str, values: Sequence[Mapping[str, Any]], type: str = "text",
                  add_to: Dataset | None = None) -> Recorded:
    """Keep ``values`` (``[{"key", "value", "note"?}]``), judged by the caller from the
    ``source`` column of ``table``'s rows named by ``key``, as a dataset (see the module's
    notes); ``type`` ``"number"`` takes numbers only; ``add_to`` a dataset this function
    made, extended rather than replaced.  :class:`RecordError` for a call refused,
    :class:`~xlsx_agent.NameNotFound` for a table or column there is not."""
    notes: list[str] = []
    if type not in TYPES:
        raise RecordError("bad_type", f"type is one of {', '.join(TYPES)}", field="type",
                          valid_options=list(TYPES))
    target = parse_address(table) if isinstance(table, str) else table
    if not isinstance(target, TableAddress):
        raise RecordError("bad_table", "table names a table or region (t:Shelf)", field="table")
    whole = replace(target, column=None, key=None, keys=None, position=None, any_row=False,
                    any_column=False, row_range=None)
    if whole != target:
        notes.append(f"table: {target.text} read as the whole table {whole.text}; key and "
                     "source name its columns")
    region = workbook.region(whole)
    key_name = region.columns[region.column_index(key)]
    source_name = region.columns[region.column_index(source)]
    address = region.address
    if not isinstance(column, str) or not column.strip():
        raise RecordError("bad_column", "column names the output column (outcome)",
                          field="column")
    if column.casefold() in (key_name.casefold(), NOTE):
        raise RecordError("bad_column", f"column {column!r} is taken in the dataset: its "
                          f"columns are the key ({key_name!r}), yours and {NOTE!r}; name "
                          "it otherwise", field="column")
    made = {"table": address.text, "key": key_name, "source": source_name, "column": column,
            "type": type}
    earlier: dict[str, tuple[Any, str | None]] = {}
    if add_to is not None:
        detail = add_to.detail.get(DETAIL) if isinstance(add_to.detail, Mapping) else None
        if detail != made:
            raise RecordError("add_mismatch", _mismatch(detail, made), field="add")
        for k, was, why in add_to.rows:
            earlier[_text(k) or ""] = (was, why)
    index = workbook.regions.key_index(region, key_name)
    keys_ds = table_dataset(workbook, replace(address, column=key_name))
    source_ds = table_dataset(workbook, replace(address, column=source_name))
    assert keys_ds.sources is not None and source_ds.sources is not None
    # Duplicates are judged among the rows with text to judge only (round 54, R113): a
    # section heading, a subtotal or a spacer row may repeat an item's key.
    has_text = [not _blank(row[0]) for row in source_ds.rows]
    with_text: dict[str, list[int]] = {}
    every: dict[str, list[int]] = {}
    for position, k in enumerate(index.keys, 1):
        if k is not None:
            every.setdefault(k, []).append(position)
            if has_text[position - 1]:
                with_text.setdefault(k, []).append(position)

    def row(position: int, name: str) -> str:
        return replace(address, position=position, column=name).text

    problems: list[dict[str, Any]] = []
    given: dict[int, tuple[Any, str | None]] = {}       # position -> (value, note)
    item_of: dict[int, int] = {}
    ignored: list[int] = []
    loose: dict[tuple, list[int]] | None = None
    for n, item in enumerate(values):
        if not isinstance(item, Mapping):
            problems.append({"code": "bad_item", "item": n})
            continue
        wanted = _text(item.get("key"))
        raw = item.get("value")
        note = _text(item.get("note"))
        found = with_text.get(wanted, []) if wanted is not None else []
        if len(found) > 1:
            problems.append({"code": "key_shared", "item": n, "key": wanted,
                             "rows": [row(p, key_name) for p in found[:SUGGESTED]]})
            continue
        if not found:
            if wanted is not None and wanted in every:
                problems.append({"code": "empty_source", "item": n,
                                 "row": row(every[wanted][0], source_name)})
                continue
            if loose is None:
                loose = {}
                for p, k in enumerate(index.keys, 1):
                    if k is not None and (has_text[p - 1] or k not in with_text):
                        loose.setdefault(loose_key(k), []).append(p)
            like = loose.get(loose_key(wanted), []) if wanted else []
            seen: set[str] = set()
            near: list[dict[str, str]] = []
            for p in like:
                k = index.keys[p - 1]
                if k is not None and k not in seen and len(near) < SUGGESTED:
                    seen.add(k)
                    near.append({"row": row(p, key_name), "key": k})
            problems.append({"code": "unknown_key", "item": n, "like": near})
            continue
        assert wanted is not None
        position = found[0]
        if position in item_of:
            problems.append({"code": "key_twice", "item": n, "row": row(position, key_name),
                             "first": item_of[position]})
            continue
        item_of[position] = n
        ignored.extend(p for p in every[wanted] if p != position)
        value: Any
        if type == "number":
            if _blank(raw):
                problems.append({"code": "empty_value", "item": n,
                                 "row": row(position, key_name)})
                continue
            value = _number(raw)
            if value is None:
                problems.append({"code": "not_number", "item": n,
                                 "row": row(position, key_name)})
                continue
        else:
            value = _text(raw)
            if _blank(value):
                problems.append({"code": "empty_value", "item": n,
                                 "row": row(position, key_name)})
                continue
        if wanted in earlier and earlier[wanted][0] != value:
            problems.append({"code": "already_recorded", "item": n, "key": wanted,
                             "row": row(position, key_name)})
            continue
        given[position] = (value, None if _blank(note) else note)
    if problems:
        raise _refusal(problems, address.text, key_name, source_name)
    added = len(given)
    kept = 0
    if earlier:
        lost = 0
        for k, (value, note) in earlier.items():
            found = with_text.get(k, [])
            if len(found) != 1:
                lost += 1
            elif found[0] in given:
                added -= 1
                if given[found[0]][1] is None:
                    given[found[0]] = (given[found[0]][0], note)
            else:
                given[found[0]] = (value, note)
                kept += 1
        if lost:
            raise RecordError("add_stale", f"{lost} key(s) recorded earlier no longer name one "
                              f"row of {address.text} with text in {source_name!r} (the "
                              "workbook changed): record every value again without add",
                              field="add")
    rows: list[tuple[Any, ...]] = []
    sources: list[tuple[tuple[str, ...], ...]] = []
    not_recorded: list[tuple[str, str]] = []
    no_key: list[str] = []
    counts: dict[Any, int] = {}
    for position in range(1, len(index.keys) + 1):
        k = index.keys[position - 1]
        cell = source_ds.sources[position - 1][0]
        if position in given:
            value, note = given[position]
            rows.append((keys_ds.rows[position - 1][0], value, note))
            sources.append((keys_ds.sources[position - 1][0], cell, cell))
            counts[value] = counts.get(value, 0) + 1
        elif has_text[position - 1]:
            if k is None:
                no_key.append(replace(address, position=position).text)
            else:
                not_recorded.append((k, cell[0] if cell else row(position, source_name)))
    if ignored:
        notes.append(f"{len(ignored)} row(s) without text in {source_name} share a key with "
                     "a row recorded: left out (keys are matched among rows with text)")
    key_type = keys_ds.types[0] if keys_ds.types else None
    dataset = Dataset((key_name, column, NOTE), tuple(rows), sources=tuple(sources), keys=1,
                      types=(key_type, type, "text"), cache=_cache(keys_ds, source_ds),
                      detail={DETAIL: made})
    shared = tuple(replace(address, position=p).text for k, ps in with_text.items()
                   if len(ps) > 1 for p in ps)
    return Recorded(dataset, address.text, key_name, source_name, column,
                    tuple(not_recorded), tuple(no_key), counts, tuple(notes),
                    added=added, kept=kept, shared=shared,
                    ignored=tuple(replace(address, position=p).text for p in sorted(ignored)),
                    repeated=frozenset(k for k, _ in not_recorded if len(every[k]) > 1))


def _mismatch(detail: Mapping[str, Any] | None, made: Mapping[str, Any]) -> str:
    """Why a dataset cannot be added to, and how to go on (rule 2)."""
    if detail is None:
        return ("add extends a dataset kept by recording values; this id holds another kind: "
                "choose another id, or leave add out to replace it")
    differ = [f"{name} {detail.get(name)!r} (this call {made[name]!r})" for name in made
              if detail.get(name) != made[name]]
    return ("add extends a dataset recorded with the same table, key, source, column and "
            "type; it was recorded with " + ", ".join(differ) + ": send those, choose "
            "another id, or leave add out to replace it")


def _cache(*read: Dataset) -> dict[str, int] | None:
    """The formula cells the key and source columns were read from, by cache status."""
    out: dict[str, int] = {}
    for dataset in read:
        for status, count in (dataset.cache or {}).items():
            out[status] = out.get(status, 0) + count
    return out or None


_SAID = {"unknown_key": "unknown key", "key_twice": "key given twice",
         "key_shared": "key held by two or more rows with text to judge",
         "not_number": "not a number (type number)",
         "already_recorded": "key recorded earlier with another value",
         "empty_source": "source cell empty (nothing there was judged)",
         "empty_value": "empty value (leave the key out to record nothing for it)",
         "bad_item": "not an object {key, value, note}"}


def _refusal(problems: list[dict[str, Any]], table: str, key: str,
             source: str) -> RecordError:
    """One refusal naming every problem by kind and place (rule 2), the way out each."""
    kinds: dict[str, list[int]] = {}
    for p in problems:
        kinds.setdefault(p["code"], []).append(p["item"])
    parts = []
    for code, items in kinds.items():
        shown = ", ".join(f"values[{i}]" for i in items[:NAMED])
        more = f" and {len(items) - NAMED} more" if len(items) > NAMED else ""
        parts.append(f"{_SAID[code]} ({len(items)}): {shown}{more}")
    ways = []
    if "unknown_key" in kinds:
        ways.append(f"keys are matched exactly as column {key!r} holds them; problems' like "
                    "gives the rows whose key is loosely equal (case, spaces, separators, "
                    "leading zeros) and that key as held: send it so")
    if "empty_source" in kinds:
        ways.append(f"leave out the keys whose {source!r} is empty")
    if "key_shared" in kinds:
        ways.append(f"problems' rows are the rows holding the key with text in {source!r}: "
                    "give each its own key in the workbook, or leave the key out")
    if "not_number" in kinds:
        ways.append("type number takes numbers: 3, 2.5, or text written so (\"2.5\")")
    if "already_recorded" in kinds:
        ways.append("with add, a key recorded earlier keeps its value: leave it out, or "
                    "record every value again without add")
    options = [r["row"] for p in problems for r in p.get("like", [])][:50]
    return RecordError(problems[0]["code"], f"nothing was recorded from {table}: "
                       + "; ".join(parts) + (". " + "; ".join(ways) if ways else ""),
                       problems=problems[:NAMED * 2], valid_options=options)
