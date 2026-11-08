"""The ``xlsx_format`` group (E7, deferred): ``xlsx_format``.

**xlsx_format** formats cells and their columns, rows and sheet in one step, through the
host API (:meth:`Workbook.format <xlsx_agent.Workbook.format>`,
:meth:`~xlsx_agent.Workbook.set_column_widths`, :meth:`~xlsx_agent.Workbook.set_row_heights`,
:meth:`~xlsx_agent.Workbook.freeze_panes`, :meth:`~xlsx_agent.Workbook.set_gridlines`):
``target`` the ranges (``Stock!A1:H1``, whole columns ``Stock!D:D``, rows ``Stock!1:1``);
flat options, each optional -- a font's ``bold``, ``italic``, ``color``, ``size``,
``font`` (its name); ``fill``; ``border`` (which sides, thin); ``horizontal`` (general,
left, center, right), ``vertical``, ``wrap``; ``number_format``; ``locked``; ``like``
(a cell whose whole format the cells take) or ``style`` (a named cell style); ``width`` or
``fit`` (Excel's AutoFit, measured) for the target's columns, ``height`` for its rows;
``freeze`` (the first cell not frozen, ``B2``; ``none``) and ``gridlines`` for its sheet.
Only what is given changes; equal formats are reused.  Kept small (a deferred tool's
definition is paid when loaded): underline, strike, indent, rotation, other border styles
and colours, and outline borders of other weights are the host API's.  Every refusal before anything
changes, the error's ``details`` listing each problem (``problems``) with what is valid.
Sheet protection is the host's (:meth:`~xlsx_agent.Workbook.protect_sheet`), not the
model's: a password is never in a tool call.  Not strict: the strict set stays
``xlsx_set_cells``.
"""

from __future__ import annotations

from typing import Any

from ooxml_edit.tools import Result, ToolError, array, boolean, number, string, tool

from ..edit import EditError
from ..edit.format import HORIZONTAL
from .docformat import KIND, handler
from .read import DOC

GROUP = "xlsx_format"

#: The sides ``border`` takes (``none`` takes every side's border away).
BORDER_SIDES = ("all", "outline", "top", "bottom", "left", "right", "none")


@tool("xlsx_format",
      "Format cells, their columns, rows and sheet in one step. Only what you give changes.",
      {"doc": DOC,
       "target": array(string(), "Stock!A1:H1, Stock!D:D or Stock!1:1.", min_items=1),
       "bold": boolean("Bold.", optional=True),
       "italic": boolean("Italic.", optional=True),
       "color": string("Font #RRGGBB.", optional=True),
       "size": number("Font points.", optional=True),
       "font": string("Font name.", optional=True),
       "fill": string("#RRGGBB or none.", optional=True),
       "border": string("Thin border; none removes.", enum=BORDER_SIDES, optional=True),
       "horizontal": string("Alignment.", enum=HORIZONTAL[:4], optional=True),
       "vertical": string("Alignment.", enum=("top", "center", "bottom"), optional=True),
       "wrap": boolean("Wrap text.", optional=True),
       "number_format": string("Code: 0.0, #,##0.00, 0%, yyyy-mm-dd.", optional=True),
       "locked": boolean("false: editable when protected.", optional=True),
       "like": string("Copy this cell's format (Stock!D2).", optional=True),
       "style": string("A named cell style.", optional=True),
       "width": number("Column width, characters.", optional=True),
       "fit": boolean("AutoFit the columns.", optional=True),
       "height": number("Row height, points.", optional=True),
       "freeze": string("First unfrozen cell (B2), or none.", optional=True),
       "gridlines": boolean("Show gridlines.", optional=True)},
      group=GROUP, kind=KIND, mutates=True, strict=False)
@handler
def xlsx_format(call: Any, doc: str, target: list[str], **given: Any) -> Result:
    """``xlsx_format``'s handler (see the module's notes)."""
    workbook = call.document
    given = {name: value for name, value in given.items() if value is not None}
    if not given:
        raise ToolError("invalid_arguments", "nothing to change: give a format, width, "
                        "fit, height, freeze or gridlines", field="target")
    if "width" in given and given.get("fit"):
        raise ToolError("invalid_arguments", "give width or fit, not both", field="fit")
    options = _cell_options(given)
    sheets = _sheets(workbook, target)
    if not sheets and ("freeze" in given or "gridlines" in given):
        raise ToolError("invalid_arguments", "freeze and gridlines need a target on a sheet, "
                        "as Stock!A1", field="target",
                        valid_options=[s.name for s in workbook.sheets][:20])
    changed = 0
    data: dict[str, Any] = {}
    notes: list[str] = []
    try:
        with workbook.batch():
            if options:
                done = workbook.format(target, **options)
                changed += done.changed
                data.update(done.as_dict())
            if given.get("fit") or "width" in given:
                done = workbook.set_column_widths(
                    target, "fit" if given.get("fit") else given["width"])
                changed += done.changed
                data["widths"] = done.widths
                if done.skipped:
                    data["skipped"] = done.skipped
                notes += done.notes
            if "height" in given:
                done = workbook.set_row_heights(target, given["height"])
                changed += done.changed
                data["heights"] = done.heights
            for sheet in sheets:
                if "freeze" in given:
                    at = given["freeze"].strip()
                    done = workbook.freeze_panes(sheet, rows=0, columns=0) \
                        if at.lower() in ("none", "", "a1") else workbook.freeze_panes(sheet, at=at)
                    changed += done.changed
                if "gridlines" in given:
                    changed += workbook.set_gridlines(sheet, given["gridlines"]).changed
    except EditError as error:
        details: dict[str, Any] = {"code": error.code, "addresses": list(error.addresses)}
        if error.problems:
            details["problems"] = [{"code": code, "a": where} for code, where in error.problems]
        raise ToolError("refused", str(error), field="target", details=details) from None
    data["changed"] = changed
    data.pop("notes", None)
    summary = f"Formatted {', '.join(target)} ({changed} change(s))"
    return Result(summary=summary, data=data, changed=[doc] if changed else [],
                  warnings=notes)


def _cell_options(given: dict[str, Any]) -> dict[str, Any]:
    """The host API's options from the flat fields."""
    out: dict[str, Any] = {}
    font = {key: given[name] for name, key in (("bold", "bold"), ("italic", "italic"),
                                               ("color", "color"),
                                               ("size", "size"), ("font", "name"))
            if name in given}
    if font:
        out["font"] = font
    if "fill" in given:
        out["fill"] = None if given["fill"].strip().lower() == "none" else given["fill"]
    if "border" in given:
        side = given["border"]
        out["border"] = ({"all": None} if side == "none" else {side: "thin"})
    alignment = {key: given[name] for name, key in (("horizontal", "horizontal"),
                                                    ("vertical", "vertical"), ("wrap", "wrap"))
                 if name in given}
    if alignment:
        out["alignment"] = alignment
    if "number_format" in given:
        out["number_format"] = given["number_format"]
    if "locked" in given:
        out["protection"] = {"locked": given["locked"]}
    for name in ("like", "style"):
        if name in given:
            out[name] = given[name]
    return out


def _sheets(workbook: Any, target: list[str]) -> list[str]:
    """The sheets the targets are on, each once (for freeze and gridlines)."""
    from ..address import AddressError, RangeAddress, parse_address

    out: list[str] = []
    for text in target:
        try:
            parsed = parse_address(text)
        except (AddressError, ValueError):
            continue        # the formatting call refuses it, with the others
        if isinstance(parsed, RangeAddress) and parsed.sheet is not None:
            try:
                name = workbook.sheet_of(parsed).name
            except (KeyError, ValueError):
                continue
            if name not in out:
                out.append(name)
    return out


TOOLS = [xlsx_format]

__all__ = ["GROUP", "TOOLS", "xlsx_format"]
