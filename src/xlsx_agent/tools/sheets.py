"""``xlsx_manage_sheets`` (0.2, E5; deferred, the ``xlsx_edit`` group): a workbook's sheets.

One tool, one ``action`` a call (:mod:`xlsx_agent.edit.sheets` does the work, one undoable
step each):

* ``add`` -- an empty worksheet called ``name`` at ``position`` (1 is the first tab; the
  end when left out), as Excel makes one;
* ``rename`` -- ``sheet`` called ``name``, every formula, defined name, rule, validation,
  chart series and pivot cache naming it rewritten as Excel rewrites them (``rewritten``);
  hyperlinks, which Excel leaves, are listed (``not_rewritten``);
* ``clear`` -- Excel's Clear All on every cell of ``sheet``: what refers to the sheet keeps
  referring to it (the way to regenerate a sheet);
* ``delete`` -- refused while anything refers to ``sheet``, each referent listed by
  address (``details.referents``), unless ``force``: each then reads ``#REF!``;
* ``hide`` and ``unhide``.

**Accepting intent** (tool design rule 1): an ``add`` given only ``sheet`` takes it as the
new sheet's name; an action already done (a sheet shown, hidden, called the name already,
a sheet cleared with nothing on it) is no error: ``changed: false`` with a ``note`` (rule
3).  A refusal says what would work (rule 2): the sheets there are as ``valid_options``,
the referents, the reason a name is refused.  Referents are listed up to the session's
``Limits.max_list_items``, the rest counted (rule 4).  Sheets withheld from the model are
neither named nor reached: a referent on one is counted, not listed.

Sent non-strict: the strict set stays ``xlsx_set_cells`` alone, within the complexity
measured to compile on Claude (X4: 13 optional parameters); this tool's four optional ones
would pass that.  Its arguments are checked here as strict grammar would.
"""

from __future__ import annotations

from typing import Any

from ooxml_edit.tools import Result, ToolError, boolean, integer, string, tool

from ..edit import EditError
from ..edit.sheets import SheetError, add_sheet, clear_sheet, delete_sheet, rename_sheet, \
    set_sheet_state
from .docformat import KIND, handler
from .read import DOC

GROUP = "xlsx_edit"

#: The actions ``xlsx_manage_sheets`` takes.
ACTIONS = ("add", "rename", "clear", "delete", "hide", "unhide")


@tool("xlsx_manage_sheets",
      "A workbook's sheets, one action a call: add (name, position), rename (sheet, name: "
      "every reference rewritten), clear (Clear All: references to it stay; regenerate a "
      "sheet this way), delete (refused while anything refers to it, listing what; force: "
      "#REF!), hide, unhide.",
      {"doc": DOC,
       "action": string("What to do.", enum=ACTIONS),
       "sheet": string("The sheet: its name or #id.", optional=True),
       "name": string("add, rename: the new name (31 characters at most).", optional=True),
       "position": integer("add: 1 is the first tab; default the end.", optional=True),
       "force": boolean("delete: write #REF! into what refers to it. clear: into "
                        "formulas naming its tables.", optional=True)},
      group=GROUP, kind=KIND, mutates=True, strict=False)
@handler
def xlsx_manage_sheets(call: Any, doc: str, action: str, sheet: str | None = None,
                       name: str | None = None, position: int | None = None,
                       force: bool | None = None) -> Result:
    """``xlsx_manage_sheets``'s handler (see the module's notes)."""
    workbook = call.document
    notes: list[str] = []
    if action == "add" and not name and sheet:
        name, sheet = sheet, None
        notes.append("sheet taken as the new sheet's name")
    if action in ("add", "rename") and not name:
        raise ToolError("invalid_arguments", f"{action} needs name: the sheet's new name",
                        field="name")
    target = None
    if action != "add":
        if not sheet:
            raise ToolError("invalid_arguments", f"{action} needs sheet: its name or #id",
                            field="sheet", valid_options=_shown(workbook))
        target = _reached(workbook, sheet)
    try:
        if action == "add":
            index = None
            if position is not None:
                if position < 1:
                    raise ToolError("invalid_arguments", "position is 1 for the first tab",
                                    field="position")
                index = min(position - 1, len(workbook._every_sheet))
            change = add_sheet(workbook, name or "", index)
        elif action == "rename":
            change = rename_sheet(workbook, target, name or "")
        elif action == "clear":
            change = clear_sheet(workbook, target, force=bool(force))
        elif action == "delete":
            change = delete_sheet(workbook, target, force=bool(force))
        elif action == "hide":
            change = set_sheet_state(workbook, target, "hidden")
        else:
            change = set_sheet_state(workbook, target, "visible")
    except SheetError as error:
        raise _refused(workbook, error, call.limits.max_list_items) from None
    except EditError as error:
        raise ToolError("refused", str(error), field="sheet",
                        details={"code": error.code,
                                 "addresses": list(error.addresses)}) from None
    data = change.as_dict(limit=call.limits.max_list_items)
    data["sheets"] = _shown(workbook)
    for key in ("rewritten", "not_rewritten"):
        if key in data:
            data[key], withheld = _visible(workbook, data[key])
            if withheld:
                data[f"{key}_withheld"] = withheld
    if change.sheet and target is not None and action != "rename":
        data["sheet"] = target.name
    if notes:
        data["notes"] = notes
    summary = _summary(action, change)
    warnings = []
    if change.not_rewritten:
        warnings.append(f"{len(change.not_rewritten)} hyperlink(s) or pivot cache(s) still "
                        "name the old sheet (not_rewritten): Excel leaves them so too")
    return Result(summary=summary, data=data, changed=[doc] if change.changed else [],
                  warnings=warnings)


def _reached(workbook: Any, text: str) -> Any:
    """The sheet the model names, among those it may see (withheld ones are not)."""
    from ..edit.sheets import find_sheet

    from ..formula.sheets import pick_sheet

    shown = workbook.sheets
    stripped = text.strip()
    if stripped.startswith("#") and stripped[1:].isdigit():
        found = [s for s in shown if s.id == int(stripped[1:])]
    else:
        found = [shown[i] for i in pick_sheet([s.name for s in shown], text)]
        if len(found) > 1:
            raise ToolError("not_found", f"no sheet {text!r}; "
                            + ", ".join(repr(s.name) for s in found)
                            + " differ from it only in their outer spaces: name one as "
                            "stored", field="sheet", valid_options=[s.name for s in found])
    if not found:
        raise ToolError("not_found", f"no sheet {text!r}", field="sheet",
                        valid_options=_shown(workbook))
    return find_sheet(workbook, f"#{found[0].id}" if found[0].id is not None
                      else found[0].stored_name)


def _shown(workbook: Any) -> list[str]:
    return [s.name for s in workbook.sheets]


def _visible(workbook: Any, listed: list[dict]) -> tuple[list[dict], int]:
    """The referents on sheets the model may see, redacted as it sees them; how many are
    on withheld ones."""
    withheld = {name.casefold() for name in workbook.withheld_sheets}
    out, hidden = [], 0
    for item in listed:
        address = item.get("a", "")
        sheet = address.rpartition("!")[0].strip("'").replace("''", "'") if "!" in address \
            else ""
        if sheet and sheet.casefold() in withheld:
            hidden += 1
            continue
        out.append({**item, "a": workbook._redacted(address)})
    return out, hidden


def _refused(workbook: Any, error: SheetError, limit: int) -> ToolError:
    details: dict[str, Any] = {"code": error.code}
    if error.referents:
        listed, withheld = _visible(workbook, [r.as_dict() for r in error.referents[:limit]])
        details["referents"] = listed
        if withheld:
            details["referents_withheld"] = withheld
        if error.referent_count > len(error.referents[:limit]):
            details["referents_cut"] = error.referent_count - len(error.referents[:limit])
    message = str(error)
    if error.code == "sheet_exists":
        message = "another sheet has that name, as Excel compares names (case aside)"
    if error.code == "no_sheet":
        return ToolError("not_found", message, field="sheet", details=details,
                         valid_options=_shown(workbook))
    field = "name" if error.code in ("bad_sheet_name", "sheet_exists") else "sheet"
    if error.code == "bad_index":
        field = "position"
    if error.code in ("referred",):
        message = (f"{error.referent_count} referent(s) refer to the sheet (details."
                   "referents). Regenerate it with action clear, then write it again; or "
                   "repeat with force to write #REF! into each")
    return ToolError("refused", message, field=field, details=details)


def _summary(action: str, change: Any) -> str:
    if not change.changed:
        return f"Nothing to do: {change.note}"
    verb = {"add": "Added", "rename": "Renamed to", "clear": "Cleared", "delete": "Deleted",
            "hide": "Hid", "unhide": "Showed"}.get(action, action)
    text = f"{verb} {change.sheet}"
    if change.rewritten:
        text += f"; {len(change.rewritten)} referent(s) rewritten"
    if change.cells:
        text += f"; {change.cells} cell(s) removed"
    return text


TOOLS = [xlsx_manage_sheets]

__all__ = ["ACTIONS", "GROUP", "TOOLS", "xlsx_manage_sheets"]
