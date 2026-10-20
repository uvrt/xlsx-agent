"""The ``xlsx_edit`` group (0.2, deferred): ``xlsx_set_cells``.

**xlsx_set_cells** writes cells of an open workbook, typed and guarded, in one step
(:meth:`Workbook.set_cells <xlsx_agent.Workbook.set_cells>`, :mod:`xlsx_agent.edit.cells`):
each item of ``cells`` is an ``address`` (``Stock!D10``, ``#1!D10``) and exactly one of

* ``text`` -- always text, ``"=..."`` too, never read as a number or a date;
* ``number``;
* ``bool``;
* ``date`` -- ``YYYY-MM-DD`` (or the session locale's date order with a four-digit
  year), into a cell whose number format shows a date, or with ``format`` (a date format
  code, ``yyyy-mm-dd``) that the cell is given;
* ``formula`` -- ``=D10*2``: the only way to write one, with no cached value (it reads
  ``cache: absent`` until Excel calculates it);
* ``blank: true`` -- the value goes, the cell's format stays; never a 0;
* ``markdown`` -- rich text: ``**bold**``, ``*italic*``, ``~~struck~~``, each line a line
  of the cell, ``-`` and ``1.`` items shown as ``•`` and numbers
  (:mod:`xlsx_agent.markdown`); ``text`` never reads Markdown.

The model's types are separate fields because a type union is outside the definitions'
common subset: a value is never inferred from its spelling.  ``replace_formulas: true``
writes over cells that hold a formula; without it such a cell refuses the call.  So does
any value refused (text over 32,767 characters, a date into a cell that shows no date, a
time zone...), a cell of an array formula or a merged area other than its first, a
table's header cell, a cell named twice, and a cell of a column the application withholds
from the model (:mod:`xlsx_agent.declare`): every refusal before anything changes, the
error's ``details`` giving each refusal by ``code`` and address (``problems``).  At most
:data:`MAX_CELLS` cells a call.

**The result**: ``written`` and ``changed`` (cells that held something else), ``cells``
each with its ``old`` type (``text``, ``number``, ``bool``, ``error``, ``blank`` or
``formula``; at most ``Limits.max_list_items``, the rest counted), ``released`` (shared
formula groups whose master was written, handed to their next member), ``calculation``
(``invalidated``: every formula's cached value in the workbook may now be stale -- reads
say ``cache: invalidated`` -- until Excel recalculates; there is no calculator here),
``fixed_height`` (cells given line breaks in a row of a fixed height, which Excel does not
grow: a warning says so).  Never a value: the summary and the log carry counts, codes and
addresses only.
"""

from __future__ import annotations

from typing import Any

from ooxml_edit.tools import Result, ToolError, array, boolean, number, obj, string, tool
from ooxml_edit.sheets import Formula

from ..address import AddressError, RangeAddress, parse_address
from ..edit import EditError
from .docformat import KIND, handler
from .read import DOC

#: The most cells one call writes.
MAX_CELLS = 5_000

#: The value fields of a cell, exactly one of which an item gives.
VALUE_FIELDS = ("text", "number", "bool", "date", "formula", "blank", "markdown")

GROUP = "xlsx_edit"


@tool("xlsx_set_cells",
      "Write cells, typed: each item an address and exactly one of text (always literal, "
      "=x too), number, bool, date (YYYY-MM-DD; the cell must show dates, or give format), "
      "formula (=D10*2), blank or markdown. One step; any refusal changes nothing. "
      "Afterwards every formula value reads cache invalidated: Excel recalculates on open.",
      {"doc": DOC,
       "cells": array(obj({
           "address": string("A cell with its sheet: Stock!D10."),
           "text": string(optional=True),
           "number": number(optional=True),
           "bool": boolean(optional=True),
           "date": string(optional=True),
           "format": string("A date format to give the cell: yyyy-mm-dd.", optional=True),
           "formula": string(optional=True),
           "blank": boolean("true: remove the value, keep the format.", optional=True),
           "markdown": string("**bold**, *italic*, ~~strike~~, lines, - and 1. lists.",
                              optional=True)}),
           f"At most {MAX_CELLS:,}.", min_items=1),
       "replace_formulas": boolean("Write over cells holding a formula. Default false.",
                                   optional=True)},
      group=GROUP, kind=KIND, mutates=True)
@handler
def xlsx_set_cells(call: Any, doc: str, cells: list[dict],
                   replace_formulas: bool | None = None) -> Result:
    """``xlsx_set_cells``'s handler (see the module's notes)."""
    if len(cells) > MAX_CELLS:
        raise ToolError("limit", f"{len(cells):,} cells; at most {MAX_CELLS:,} a call: send "
                        "the rest in another", field="cells")
    workbook = call.document
    items = [(_address(item, index), _value(item, index)) for index, item in enumerate(cells)]
    _withheld(workbook, items)
    try:
        written = workbook.set_cells(items, replace_formulas=bool(replace_formulas))
    except EditError as error:
        details: dict[str, Any] = {"code": error.code, "addresses": list(error.addresses)}
        if len(error.problems) > 1:
            details["problems"] = [{"code": code, "a": address}
                                   for code, address in error.problems]
        raise ToolError("refused", str(error), field="cells", details=details) from None
    data = written.as_dict(limit=call.limits.max_list_items)
    summary = f"Wrote {len(written.cells)} cell(s) ({written.changed} changed)"
    if written.changed:
        summary += "; formula caches invalidated"
    warnings = []
    if written.fixed_height:
        warnings.append(f"{len(written.fixed_height)} cell(s) with line breaks are in rows of "
                        "a fixed height (fixed_height): Excel keeps the height, so lines past "
                        "it are hidden until the row is fitted")
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


TOOLS = [xlsx_set_cells]

__all__ = ["GROUP", "MAX_CELLS", "TOOLS", "VALUE_FIELDS", "xlsx_set_cells"]
