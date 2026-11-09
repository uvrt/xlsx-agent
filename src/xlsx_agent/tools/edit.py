"""The ``xlsx_edit`` group (0.2, deferred): ``xlsx_set_cells`` and (round 44, E4)
``xlsx_write_table``.

**xlsx_set_cells** writes cells of an open workbook, typed and guarded, in one step
(:meth:`Workbook.set_cells <xlsx_agent.Workbook.set_cells>`, :mod:`xlsx_agent.edit.cells`):
each item of ``cells`` is an ``address`` (``Stock!D10``, ``#1!D10``) and exactly one of

* ``text`` -- always text, ``"=..."`` too, never read as a number or a date;
* ``number``;
* ``bool``;
* ``date`` -- ``YYYY-MM-DD`` (or the session locale's date order with a four-digit
  year), into a cell whose number format shows a date; a General cell (a new one) and one
  of a format Excel replaces for a typed date (percentage, currency, time...) gets the
  short date as in Excel, listed in ``short_date``; a cell of a number format (``0.00``)
  needs ``format`` (a date format code, ``yyyy-mm-dd``), which the cell is given;
* ``formula`` -- ``=D10*2``: the only way to write one, with no cached value (it reads
  ``cache: absent`` until Excel calculates it);
* ``blank: true`` -- the value goes, the cell's format stays; never a 0;
* ``markdown`` -- rich text: ``**bold**``, ``*italic*``, ``~~struck~~``, each line a line
  of the cell, ``-`` and ``1.`` items shown as ``•`` and numbers
  (:mod:`xlsx_agent.markdown`); ``text`` never reads Markdown.

The model's types are separate fields because a type union is outside the definitions'
common subset: a value is never inferred from its spelling.  ``replace_formulas: true``
writes over cells that hold a formula; without it such a cell refuses the call.  So does
any value refused (text over 32,767 characters, a date into a cell of a number format, a
time zone...), a cell of an array formula or a merged area other than its first, a
table's header cell, a cell named twice, and a cell of a column the application withholds
from the model (:mod:`xlsx_agent.declare`): every refusal before anything changes, the
error's ``details`` giving each refusal by ``code`` and address (``problems``).  At most
:data:`MAX_CELLS` cells a call.

**Guards** (X4, :mod:`xlsx_agent.edit.policy`, :mod:`xlsx_agent.edit.keys`): an
``address`` may name a value by its row's key, ``t:stock/P-7/Quantity`` -- the row holding
the key now, wherever it is; a key two rows share, a key no row has, a row by place
(``#3``), a table with no key column declared and one whose key column holds formulas
(``formula_key``: keys are values the application supplies) refuse.  A key no row has
starts a new row only in the row an item names for it, ``row`` (the sheet's row number:
a free row of the table's declared capacity, continuing its data rows with no blank row
between; :mod:`xlsx_agent.edit.keys`): the application allocates the rows -- in its
instructions, or by the write policy over the capacity rows it opens -- and there is no
"next free row" form here.  A row outside the capacity, one of the data rows, a row not
empty, a row or key allocated twice, a gap and a key a row already has refuse.  Each
item's ``expect`` -- the cell's type now (``empty`` or ``blank``: no value and no formula;
``text``, ``number``, ``bool``, ``error``; ``formula``: any
formula cell, whatever it caches) -- and ``expect_value`` -- its value now as a read gave
it, as text (a text's own, a number's digits, ``true``/``false``, a date ``YYYY-MM-DD``; a
formula's cached value) -- refuse the call when the cell holds something else, the
error's ``details`` giving each such cell's type as it is (``held``), never its value.
Both are flat fields of the item: a nested object made the strict schema too complex for
Claude (HTTP 400 "Schema is too complex", measured online by the main session).
``revision`` (the version a result gave) refuses it when the workbook has changed since.
The application's write policy (``xlsx_agent.tools.policy``; ``describe``'s ``writable``)
and a protected sheet refuse the cells outside it, by key too.  All of them before
anything changes.

**The result**: ``written`` and ``changed`` (cells that held something else), ``cells``
each with its ``old`` type (``text``, ``number``, ``bool``, ``error``, ``blank`` or
``formula``; at most ``Limits.max_list_items``, the rest counted), ``released`` (shared
formula groups whose master was written, handed to their next member), ``calculation``
(``invalidated``: every formula's cached value in the workbook may now be stale -- reads
say ``cache: invalidated`` -- until Excel recalculates; there is no calculator here),
``fixed_height`` (cells given line breaks in a row of a fixed height, which Excel does not
grow: a warning says so), ``new_rows`` (the key cells of the rows started).  Never a value: the summary and the log carry counts, codes and
addresses only.
"""

from __future__ import annotations

from dataclasses import replace
from typing import Any

from ooxml_edit.tools import (Result, ToolError, array, boolean, integer, number, obj, string,
                              tool)

from ..address import AddressError, RangeAddress, TableAddress, parse_address
from ..edit import EditError
from .docformat import KIND, handler
from .read import DOC

#: The most cells one call writes.
MAX_CELLS = 5_000

#: The value fields of a cell, exactly one of which an item gives.
VALUE_FIELDS = ("text", "number", "bool", "date", "formula", "blank", "markdown")

#: The types an item's ``expect`` names (:data:`xlsx_agent.edit.cells.EXPECT_TYPES`).
EXPECT_TYPES = ("empty", "blank", "text", "number", "bool", "error", "formula")


GROUP = "xlsx_edit"


@tool("xlsx_set_cells",
      "Write cells, typed: each item an address and exactly one of text (always literal, "
      "=x too), number, bool, date (YYYY-MM-DD; a General cell gets a short date, a number "
      "format needs format), "
      "formula (=D10*2), blank or markdown. One step; any refusal changes nothing. "
      "Afterwards every formula value reads cache invalidated: Excel recalculates on open. "
      "describe's writable is all you may write.",
      {"doc": DOC,
       "cells": array(obj({
           "address": string("A cell with its sheet (Stock!D10), or a value by row key "
                             "(t:stock/P-7/Quantity)."),
           "text": string(optional=True),
           "number": number(optional=True),
           "bool": boolean(optional=True),
           "date": string(optional=True),
           "format": string("A date format to give the cell: yyyy-mm-dd.", optional=True),
           "formula": string(optional=True),
           "blank": boolean("true: remove the value, keep the format.", optional=True),
           "markdown": string("**bold**, *italic*, ~~strike~~, lines, - and 1. lists.",
                              optional=True),
           "expect": string("The cell's type now, else refused.", enum=EXPECT_TYPES,
                            optional=True),
           "expect_value": string("Its value now as read (dates YYYY-MM-DD), else refused.",
                                  optional=True),
           "row": integer("New key (t:...): the free capacity row it starts in.",
                          optional=True)}),
           f"At most {MAX_CELLS:,}.", min_items=1),
       "replace_formulas": boolean("Write over cells holding a formula. Default false.",
                                   optional=True),
       "revision": integer("The doc's version you read; refused if it changed.",
                           optional=True)},
      group=GROUP, kind=KIND, mutates=True)
@handler
def xlsx_set_cells(call: Any, doc: str, cells: list[dict],
                   replace_formulas: bool | None = None,
                   revision: int | None = None) -> Result:
    """``xlsx_set_cells``'s handler (see the module's notes)."""
    if len(cells) > MAX_CELLS:
        raise ToolError("limit", f"{len(cells):,} cells; at most {MAX_CELLS:,} a call: send "
                        "the rest in another", field="cells")
    workbook = call.document
    items = [(_address(item, index), _expected(item, index, _value(item, index)))
             for index, item in enumerate(cells)]
    _withheld(workbook, items)
    allocate = _allocated(cells, items)
    try:
        written = workbook.set_cells(items, replace_formulas=bool(replace_formulas),
                                     allocate=allocate or None, revision=revision)
    except EditError as error:
        details: dict[str, Any] = {"code": error.code, "addresses": list(error.addresses)}
        if len(error.problems) > 1:
            details["problems"] = [{"code": code, "a": address}
                                   for code, address in error.problems]
        if error.held:
            details["held"] = dict(error.held)
        field = "revision" if error.code == "revision" else "cells"
        raise ToolError("refused", str(error), field=field, details=details) from None
    data = written.as_dict(limit=call.limits.max_list_items)
    summary = f"Wrote {len(written.cells)} cell(s) ({written.changed} changed)"
    if written.changed:
        summary += "; formula caches invalidated"
    warnings = []
    if written.fixed_height:
        warnings.append(f"{len(written.fixed_height)} cell(s) with line breaks are in rows of "
                        "a fixed height (fixed_height): Excel keeps the height, so lines past "
                        "it are hidden until the row is fitted")
    if written.invalid:
        warnings.append(f"{len(written.invalid)} cell(s) now hold a value their data "
                        "validation does not allow (invalid): written as asked; Excel marks "
                        "such cells only when asked to circle invalid data")
    return Result(summary=summary, data=data, changed=[doc] if written.changed else [],
                  warnings=warnings)


def _address(item: dict, index: int) -> str:
    address = item.get("address")
    if not isinstance(address, str) or not address.strip():
        raise ToolError("invalid_arguments", "each cell needs an address, as Stock!D10",
                        field=f"cells[{index}].address")
    return address


def _value(item: dict, index: int) -> Any:
    """The item's value as :meth:`Workbook.set_cells` takes it: exactly one value field."""
    given = [name for name in VALUE_FIELDS if item.get(name) is not None]
    if item.get("blank") is False:
        given.remove("blank")
    if len(given) != 1:
        raise ToolError("invalid_arguments", f"give exactly one of {', '.join(VALUE_FIELDS)}"
                        f" ({len(given)} given)", field=f"cells[{index}]",
                        valid_options=list(VALUE_FIELDS))
    (name,) = given
    if item.get("format") is not None and name != "date":
        raise ToolError("invalid_arguments", "format is a date's format: give it with date",
                        field=f"cells[{index}].format")
    if name == "blank":
        return None
    if name == "formula":
        return {"formula": item["formula"]}
    if name == "markdown":
        return {"markdown": item["markdown"]}
    if name == "date":
        value = {"date": item["date"]}
        if item.get("format") is not None:
            value["format"] = item["format"]
        return value
    if name == "number" and isinstance(item["number"], float) and item["number"].is_integer() \
            and abs(item["number"]) <= 2 ** 53:
        return int(item["number"])
    return item[name]


def _expected(item: dict, index: int, value: Any) -> Any:
    """``value`` with the item's ``expect`` (a type) and ``expect_value`` (the value as
    text), as :meth:`Workbook.set_cells` takes them: ``{"type": ..., "shown": ...}``."""
    kind, shown = item.get("expect"), item.get("expect_value")
    if kind is None and shown is None:
        return value
    out: dict[str, Any] = {}
    if kind is not None:
        out["type"] = kind
    if shown is not None:
        out["shown"] = shown
    if isinstance(value, dict):
        return {**value, "expect": out}
    return {"value": value, "expect": out}


def _allocated(cells: list[dict], items: list) -> list[tuple[str, int]]:
    """The rows the items allocate (``row``), as :meth:`Workbook.set_cells` takes them:
    ``(t:<table>/<key>, row)`` pairs, each once."""
    out: list[tuple[str, int]] = []
    for index, (item, (address, _)) in enumerate(zip(cells, items)):
        row = item.get("row")
        if row is None:
            continue
        try:
            parsed = TableAddress.parse(address.strip())
        except (AddressError, ValueError):
            parsed = None
        if parsed is None or parsed.kind != "value" or parsed.key is None:
            raise ToolError("invalid_arguments", "row goes with a value by a new key, "
                            "t:<table>/<key>/<column>", field=f"cells[{index}].row")
        pair = (replace(parsed, column=None).text, row)
        if pair not in out:
            out.append(pair)
    return out


def _withheld(workbook: Any, items: list) -> None:
    """Refuse a cell of a column the application withholds from the model."""
    refused: list[str] = []
    for address, _ in items:
        try:
            parsed = parse_address(address)
        except (AddressError, ValueError):
            continue        # set_cells refuses it, with the others
        if not isinstance(parsed, RangeAddress) or parsed.sheet is None:
            continue
        try:
            sheet = workbook.sheet_of(parsed)
        except (KeyError, ValueError):
            continue
        cell = parsed.area.first
        for area in workbook.regions.withheld(sheet.name) if sheet.name else ():
            if area.contains(cell):
                refused.append(address)
                break
    if refused:
        raise ToolError("refused", f"{len(refused)} cell(s) are in columns withheld from you: "
                        "they are not written here", field="cells",
                        details={"code": "withheld", "addresses": refused[:20]})


# -- E4 xlsx_write_table (round 44) ---------------------------------------------------------------

_CELL = obj({
    "text": string(optional=True),
    "number": number(optional=True),
    "bool": boolean(optional=True),
    "date": string(optional=True),
    "formula": string(optional=True),
    "markdown": string(optional=True),
    "blank": boolean(optional=True)})


@tool("xlsx_write_table",
      "Write headers and rows from a first cell (Stock!B2), or append rows to a table "
      "(t:Orders: it grows, calculated columns and totals follow). One step; refused over "
      "non-empty cells (unless overwrite), merges or tables. Returns the range written.",
      {"doc": DOC,
       "at": string("Stock!B2, or t:<table> to append."),
       "headers": array(string(), "Column names: the first row.", optional=True),
       "rows": array(array(_CELL), "Cells as in xlsx_set_cells; {} blank.", optional=True),
       "from": string("d:<id>: a dataset's columns and rows instead.", optional=True),
       "table": boolean("Make an Excel table: filters, banding.", optional=True),
       "name": string("Table name.", optional=True),
       "style": string("Table style; default the workbook's.", optional=True),
       "overwrite": boolean("Write over non-empty cells.", optional=True)},
      group=GROUP, kind=KIND, mutates=True, strict=False)
@handler
def xlsx_write_table(call: Any, doc: str, at: str, headers: list | None = None,
                     rows: list | None = None, table: bool | None = None,
                     name: str | None = None, style: str | None = None,
                     overwrite: bool | None = None, **given: Any) -> Result:
    """``xlsx_write_table``'s handler: :meth:`Workbook.write_table
    <xlsx_agent.Workbook.write_table>` at a cell, :meth:`Workbook.append_rows
    <xlsx_agent.Workbook.append_rows>` at ``t:<table>``
    (:mod:`xlsx_agent.edit.tables`)."""
    workbook = call.document
    source = given.get("from")
    notes: list[str] = []
    if source:
        if rows:
            raise ToolError("invalid_arguments", "give rows or from, not both", field="from")
        dataset = call.dataset(source, field="from").dataset
        if headers and list(headers) != list(dataset.columns):
            notes.append("headers replaced by the dataset's columns")
        headers = list(dataset.columns)
        values = [list(row) for row in dataset.rows]
    else:
        values = [[_cell(item, f"rows[{r}][{c}]") for c, item in enumerate(row or [])]
                  for r, row in enumerate(rows or [])]
    count = sum(len(row) for row in values) + len(headers or [])
    if count > MAX_CELLS:
        raise ToolError("limit", f"{count:,} cells; at most {MAX_CELLS:,} a call: send the "
                        "rest in another (at t:<table> appends)", field="rows")
    target = at.strip() if isinstance(at, str) else ""
    try:
        if target.startswith("t:"):
            if headers:
                notes.append("headers ignored: the table's columns name its cells")
            if table or name or style or overwrite:
                notes.append("table, name, style and overwrite are for a new block; "
                             "ignored when appending")
            _withheld_table(workbook, target)
            done = workbook.append_rows(target, values)
            summary = f"Appended {done.rows} row(s) to {done.table}: {done.range}"
        else:
            _withheld_block(workbook, target, len(headers or []) or max(
                (len(row) for row in values), default=0), len(values) + bool(headers))
            done = workbook.write_table(None, target, headers or None, values,
                                        as_table=bool(table) or bool(name), name=name,
                                        style=style, overwrite=bool(overwrite))
            summary = f"Wrote {done.range}" + (f" as table {done.table}" if done.table else "")
    except EditError as error:
        details: dict[str, Any] = {"code": error.code, "addresses": list(error.addresses)}
        if len(error.problems) > 1:
            details["problems"] = [{"code": code, "a": address}
                                   for code, address in error.problems]
        raise ToolError("refused", str(error), field="at" if error.code in (
            "bad_address", "not_empty", "table_overlap", "merged_cell") else "rows",
            details=details) from None
    data = done.as_dict(limit=call.limits.max_list_items)
    written = done.written
    return Result(summary=summary, data=data,
                  changed=[doc] if written is None or written.changed or getattr(
                      done, "table", None) else [], warnings=notes)


def _cell(item: Any, where: str) -> Any:
    """A row's cell as :meth:`Workbook.set_cells` takes it: one value field (none, or
    ``blank``: a blank)."""
    if not isinstance(item, dict):
        raise ToolError("invalid_arguments", "a cell is an object with one of "
                        f"{', '.join(VALUE_FIELDS)}", field=where)
    given = [name for name in VALUE_FIELDS if item.get(name) is not None
             and not (name == "blank" and item[name] is False)]
    if not given or given == ["blank"]:
        return None
    if len(given) != 1:
        raise ToolError("invalid_arguments", f"give one of {', '.join(VALUE_FIELDS)} "
                        f"({len(given)} given)", field=where, valid_options=list(VALUE_FIELDS))
    (name,) = given
    if name in ("formula", "markdown", "date"):
        return {name: item[name]}
    if name == "number" and isinstance(item["number"], float) \
            and item["number"].is_integer() and abs(item["number"]) <= 2 ** 53:
        return int(item["number"])
    return item[name]


def _withheld_block(workbook: Any, at: str, width: int, height: int) -> None:
    """Refuse a block that meets a column the application withholds from the model."""
    try:
        parsed = parse_address(at)
        sheet = workbook.sheet_of(parsed)
    except (AddressError, KeyError, ValueError):
        return          # write_table refuses it, with its reason
    if not isinstance(parsed, RangeAddress) or not width or not height:
        return
    first = parsed.area.first
    for area in workbook.regions.withheld(sheet.name) if sheet.name else ():
        if not (area.last.column < first.column or area.first.column > first.column + width - 1
                or area.last.row < first.row or area.first.row > first.row + height - 1):
            raise ToolError("refused", "the block meets columns withheld from you: they are "
                            "not written here", field="at", details={"code": "withheld"})


def _withheld_table(workbook: Any, target: str) -> None:
    """Refuse appending to a table with a column the application withholds."""
    for table in workbook._table_parts():
        if table.name.casefold() != target[2:].strip().casefold() or table.area is None:
            continue
        for hidden in workbook.regions.withheld(table.sheet):
            if not (hidden.last.column < table.area.first.column
                    or hidden.first.column > table.area.last.column):
                raise ToolError("refused", "the table has columns withheld from you: rows "
                                "are not appended to it here", field="at",
                                details={"code": "withheld"})


TOOLS = [xlsx_set_cells, xlsx_write_table]

__all__ = ["GROUP", "MAX_CELLS", "TOOLS", "VALUE_FIELDS", "xlsx_set_cells", "xlsx_write_table"]
