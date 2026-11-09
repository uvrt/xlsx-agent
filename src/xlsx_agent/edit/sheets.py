"""Sheets (0.2, E5; production R52): add, rename, delete, clear and hide them, as Excel does.

::

    wb.add_sheet("Rota", index=1)          # a new, empty worksheet, as Excel makes one
    wb.rename_sheet("Lists", "Tax rates")  # every reference to it rewritten
    wb.clear_sheet("Results")              # Excel's Clear All on every cell: references stay
    wb.set_sheet_state("Meta", "hidden")   # "visible", "hidden" or "veryHidden"
    wb.sheet_referents("Lists")            # what refers to it, by address
    wb.delete_sheet("Old")                 # refused while anything refers to it
    wb.delete_sheet("Old", force=True)     # ... or deleted as Excel deletes it: #REF!

Each is one undoable step (or part of the open :meth:`Workbook.batch
<xlsx_agent.Workbook.batch>`), refused before anything changes -- every problem listed --
and returns a :class:`SheetChange`.  A sheet is named by its name as stored (case aside, as
Excel matches names), as an address writes it (``'Q3 sales'``), by id (``#3``), or as a
:class:`~xlsx_agent.Sheet`; the host's program reaches every sheet, those withheld from
the model too (the tool layer does not).

**Regenerating a sheet.**  Clear it and write it again: every formula, defined name, chart
series and validation that refers to it keeps referring to it.  Deleting it and adding one
of the same name breaks them -- Excel writes ``#REF!`` into each when the sheet goes, and a
new sheet of that name does not bring them back -- so :meth:`delete_sheet` refuses while
anything refers to the sheet, listing each referent by address.

**What Excel does, measured** (Excel for Mac 16.106, 2026-11-08, its object model on
made-up workbooks, each saved and read; the probes are lane S's of round 44):

* *Add* (``make new worksheet``): a worksheet part of a dimension ``A1``, one sheet view,
  the sheet format, empty sheet data and the page margins 0.7/0.7/0.75/0.75/0.3/0.3; the
  next ``sheetId`` (the largest plus one); defined names' ``localSheetId`` from the new
  sheet's place on moved up one.  Excel makes the new sheet the active, selected one; a
  program adding a sheet keeps the active sheet the one it was (its index moved with it).
  The sheet format is the workbook's first worksheet's ``baseColWidth`` and
  ``defaultRowHeight`` (Excel's new sheet had the same as the sheets it had made; they
  follow the workbook's default font, which is not measured here for every font), or none
  when it has none (Excel then derives them).
* *Rename*: every formula on every sheet, defined names, conditional formats and data
  validations (the ``x14`` ones in ``extLst`` too), chart series and titles (``c:f``), the
  pivot caches' source sheet and tables' calculated columns refer to the new name
  (:func:`xlsx_agent.formula.sheets.rename`).  A hyperlink's location is *not* rewritten
  by Excel: it is left as it is and listed in the result (``not_rewritten``), as is a
  sheet name inside a string (``INDIRECT("Lists!B2")``), which is no reference.
* *Delete*: every reference to the sheet becomes ``#REF!``, a 3-D run ending at it shrinks
  inward (:func:`xlsx_agent.formula.sheets.delete`); a structured reference to a table on
  it becomes ``#REF!``; a chart series' ``c:f`` becomes ``#REF!``; the defined names
  local to it go and the others' ``localSheetId`` move down; deleting the active sheet
  makes the sheet now at its place active (the one before it when it was the last).  A
  hyperlink's location and a pivot cache's source are left as they are (Excel turned the
  cache's source into a link to the file itself; it is listed in the result).
* *Hide*: ``state="hidden"`` (``veryHidden``: only a program unhides it); hiding the
  active sheet makes the next visible sheet to its right active and selected, else the
  one to its left.  Hiding the last visible sheet is refused (Excel refuses it).
* *Clear All* on every cell (``clear range`` of ``1:1048576``): the cells go -- values,
  formulas and formats, a row's and a column's format too -- with the merged areas,
  conditional formats, data validations, comments, hyperlinks, tables and pivot tables on
  the sheet.  What stays: column widths and custom row heights, the sheet's views, sheet
  format, print settings, header and footer, page breaks, its drawing (a chart stays),
  extensions Excel does not know, and every defined name and reference into the sheet.
  A formula elsewhere naming a table on the sheet is a referent (refused unless
  ``force``: it then reads ``#REF!``, as when the sheet is deleted).
* *Names* (``set name``): 1 to 31 characters; none of ``: \\ / ? * [ ]``; not starting or
  ending with ``'``; unique among the workbook's sheets as Excel compares names
  (:func:`~xlsx_agent.formula.sheets.fold_name`).  Excel's object model takes ``History``,
  a lone space and control characters; the first two are taken here, control characters
  are refused (a line break in a sheet's name is useless and no formula can hold one),
  and so is text that reads as an ``_xHHHH_`` escape (``_x0041_``: Excel writes it
  escaped again, and readers that do not decode it see another name).

A workbook whose structure is locked (``workbookProtection lockStructure``) refuses every
change of its sheets but :meth:`clear_sheet`, as Excel does; a protected sheet refuses a
clear; the host's write policy must cover a sheet whole for it to be cleared, deleted,
renamed or hidden.  ``docProps/app.xml`` (the sheet names a file browser shows) is brought
up to date on a rename; Excel rewrites it on every save.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Iterator, Sequence

from lxml import etree
from ooxml_edit.sheets import MAX_COLUMN, MAX_ROW, Area, Cell

from ..address import quote_sheet
from ..formula import sheets as refs
from ..read.strings import decode_xstring
from . import EditError

X = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_X = f"{{{X}}}"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_REL = R + "/"
REL_WORKSHEET = _REL + "worksheet"
CT_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"

#: CT_Worksheet's children in order (ECMA-376 Part 1, §18.3.1.99).
WORKSHEET_ORDER = ("sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData",
                   "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios",
                   "autoFilter", "sortState", "dataConsolidate", "customSheetViews",
                   "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations",
                   "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter",
                   "rowBreaks", "colBreaks", "customProperties", "cellWatches",
                   "ignoredErrors", "smartTags", "drawing", "legacyDrawing",
                   "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls",
                   "webPublishItems", "tableParts", "extLst")

#: ``ST_SheetState`` (ECMA-376 Part 1, §18.18.68).
STATES = ("visible", "hidden", "veryHidden")
#: The longest sheet name Excel takes (measured: 31 characters; 32 refused).
MAX_NAME = 31
#: Characters Excel refuses in a sheet name (measured, each refused by ``set name``).
FORBIDDEN = ":\\/?*[]"
#: What reads as an ``_xHHHH_`` escape in a name.
_ESCAPE_LIKE = re.compile(r"_x[0-9A-Fa-f]{4}_")
#: The extension elements Clear All removes with the sheet's rules (``x14``): conditional
#: formats and data validations.
_CLEARED_EXT = {"conditionalFormattings", "dataValidations"}
#: Most referents listed in a refusal or a result.
MAX_LISTED = 50


@dataclass(frozen=True)
class Referent:
    """Something that refers to a sheet: ``kind`` (``formula``, ``name``,
    ``conditional_format``, ``validation``, ``chart``, ``pivot_cache``,
    ``table_formula``, ``hyperlink``) and where, by address (``Totals!B6``, a defined
    name, ``Stock!C10:C200`` for a rule, the chart's sheet and part, a pivot cache's
    part).  Never the formula's text."""

    kind: str
    address: str

    def as_dict(self) -> dict[str, str]:
        return {"kind": self.kind, "a": self.address}


@dataclass(frozen=True)
class SheetChange:
    """What a sheet operation did.  ``sheet`` and ``id`` name the sheet (after a rename,
    its new name); ``changed`` is false for an operation that found it already done
    (``note`` says what); ``sheets`` the workbook's sheets after, in order;
    ``rewritten`` the referents written anew (a rename's, a forced delete's ``#REF!``),
    ``not_rewritten`` those Excel leaves as they were (hyperlinks, pivot caches) -- now
    naming a sheet that is not there; ``removed`` what went with it (defined names local
    to a deleted sheet; a clear's tables, merges, rules, comments, hyperlinks, pivot
    tables); ``cells`` how many cells a clear removed."""

    action: str
    sheet: str
    id: int | None
    changed: bool = True
    sheets: tuple[str, ...] = ()
    rewritten: tuple[Referent, ...] = ()
    not_rewritten: tuple[Referent, ...] = ()
    removed: tuple[str, ...] = ()
    cells: int = 0
    note: str | None = None

    def as_dict(self, limit: int = MAX_LISTED) -> dict[str, Any]:
        out: dict[str, Any] = {"action": self.action, "sheet": self.sheet, "id": self.id,
                               "changed": self.changed, "sheets": list(self.sheets)}
        for name in ("rewritten", "not_rewritten"):
            found = getattr(self, name)
            if found:
                out[name] = [r.as_dict() for r in found[:limit]]
                if len(found) > limit:
                    out[f"{name}_cut"] = len(found) - limit
        if self.removed:
            out["removed"] = list(self.removed[:limit])
            if len(self.removed) > limit:
                out["removed_cut"] = len(self.removed) - limit
        if self.cells:
            out["cells"] = self.cells
        if self.note:
            out["note"] = self.note
        return out


class SheetError(EditError):
    """A sheet operation refused: :class:`EditError` with :attr:`referents` (at most
    :data:`MAX_LISTED`; :attr:`referent_count` all of them) for a delete or clear that
    something refers to, and :attr:`options` (the sheets there are, the names to choose
    from) for a sheet not found or a name taken."""

    def __init__(self, code: str, detail: str | None = None, *,
                 referents: Sequence[Referent] = (), options: Sequence[str] = (),
                 addresses: Sequence[str] = ()) -> None:
        super().__init__(code, detail, addresses=list(addresses))
        self.referents = tuple(referents)[:MAX_LISTED]
        self.referent_count = len(referents)
        self.options = tuple(options)[:MAX_LISTED]


# -- the sheet named ----------------------------------------------------------------------------


def find_sheet(workbook: Any, which: Any) -> Any:
    """The sheet ``which`` names (see the module): every sheet the workbook lists, those
    withheld from the model too; :class:`SheetError` ``no_sheet`` with the sheets there
    are."""
    from ..workbook import Sheet

    every = workbook._every_sheet
    if isinstance(which, Sheet):
        which = f"#{which.id}" if which.id is not None else which.stored_name
    if not isinstance(which, str) or not which.strip():
        raise SheetError("no_sheet", "a sheet is named by its name or #id",
                         options=[s.stored_name for s in every])
    text = which.strip()
    if text.startswith("#") and text[1:].isdigit():
        for sheet in every:
            if sheet.id == int(text[1:]):
                return sheet
    names = [text]
    body = text[:-1] if text.endswith("!") else text
    if len(body) >= 2 and body.startswith("'") and body.endswith("'"):
        names.append(body[1:-1].replace("''", "'"))
    for name in names:
        key = refs.fold_name(name)
        for sheet in every:
            if refs.fold_name(sheet.stored_name) == key:
                return sheet
    raise SheetError("no_sheet", f"no sheet {text!r}",
                     options=[s.stored_name for s in every])


def name_problems(workbook: Any, name: Any, *, keep: Any = None,
                  taken: Sequence[str] = ()) -> list[str]:
    """Why ``name`` cannot name a sheet of ``workbook`` (see the module), each as a short
    reason; empty when it can.  ``keep``: the sheet being renamed, whose own name is no
    clash.  ``workbook`` ``None``: a workbook not made yet (:meth:`Workbook.new
    <xlsx_agent.Workbook.new>`), whose sheets are ``taken``."""
    if not isinstance(name, str):
        return ["a sheet's name is text"]
    problems = []
    if not name:
        problems.append("a sheet's name is never empty")
    if len(name) > MAX_NAME:
        problems.append(f"{len(name)} characters; a sheet's name has at most {MAX_NAME}")
    bad = sorted({c for c in name if c in FORBIDDEN})
    if bad:
        problems.append(f"Excel refuses {' '.join(bad)} in a sheet's name")
    if name.startswith("'") or name.endswith("'"):
        problems.append("a sheet's name does not start or end with '")
    if any(ord(c) < 0x20 or 0xD800 <= ord(c) <= 0xDFFF or c in "\ufffe\uffff" for c in name):
        problems.append("a sheet's name holds no control characters (line breaks, tabs)")
    if _ESCAPE_LIKE.search(name):
        problems.append("text that reads as an _xHHHH_ escape (as _x0041_) is not written "
                        "in a sheet's name")
    if not problems and workbook is None:
        key = refs.fold_name(name)
        for other in taken:
            if refs.fold_name(other) == key:
                problems.append(f"the workbook has a sheet {other!r} already, which Excel "
                                "takes for the same name")
                break
    elif not problems:
        key = refs.fold_name(name)
        for sheet in workbook._every_sheet:
            if keep is not None and sheet.part == keep.part and sheet.id == keep.id:
                continue
            if refs.fold_name(sheet.stored_name) == key:
                problems.append(f"the workbook has a sheet {sheet.stored_name!r} already, "
                                "which Excel takes for the same name")
                break
    return problems


def _check_name(workbook: Any, name: Any, *, keep: Any = None) -> None:
    problems = name_problems(workbook, name, keep=keep)
    if problems:
        code = "sheet_exists" if problems[0].startswith("the workbook has") else \
            "bad_sheet_name"
        raise SheetError(code, "; ".join(problems),
                         options=[s.stored_name for s in workbook._every_sheet])


# -- the workbook part ---------------------------------------------------------------------------


def _root(workbook: Any) -> Any:
    return workbook.package.tree(workbook.part)


def _child(parent: Any, local: str) -> Any:
    return parent.find(_X + local)


def _sheet_nodes(workbook: Any) -> list[Any]:
    holder = _child(_root(workbook), "sheets")
    return [] if holder is None else [n for n in holder if n.tag == _X + "sheet"]


def _views(workbook: Any) -> list[Any]:
    holder = _child(_root(workbook), "bookViews")
    return [] if holder is None else [n for n in holder if n.tag == _X + "workbookView"]


def _index_of(workbook: Any, sheet: Any) -> int:
    for index, node in enumerate(_sheet_nodes(workbook)):
        if node.get("name") == sheet.stored_name:
            return index
    raise SheetError("no_sheet", f"no sheet {sheet.stored_name!r}")


def _int_attr(node: Any, name: str, default: int = 0) -> int:
    try:
        return int(node.get(name, default))
    except (TypeError, ValueError):
        return default


def _guard(workbook: Any, sheet: Any = None, *, structure: bool = True,
           whole: bool = True) -> None:
    """Refusals every sheet operation shares: a strict workbook, a locked structure, the
    host's policy not covering the sheet whole."""
    if workbook.strict:
        raise EditError("strict")
    if structure and _structure_locked(workbook):
        raise SheetError("structure_locked", "the workbook's structure is protected "
                         "(workbookProtection lockStructure): its sheets are not added, "
                         "deleted, renamed or hidden")
    if sheet is not None and whole and workbook.policies:
        allowed = workbook._allowed()
        if allowed is not None and not _whole(allowed, sheet.part):
            raise SheetError("not_writable", f"the host's write policy does not cover the "
                             f"sheet {sheet.stored_name!r} whole",
                             addresses=[quote_sheet(sheet.stored_name) + "!"])


def _structure_locked(workbook: Any) -> bool:
    node = _child(_root(workbook), "workbookProtection")
    return node is not None and node.get("lockStructure", "").strip() in ("1", "true")


def _whole(allowed: Any, part: str | None) -> bool:
    every = Area(Cell(1, 1), Cell(MAX_ROW, MAX_COLUMN))
    return part is not None and all(any(a.first == every.first and a.last == every.last
                                        for a in areas.get(part, ()))
                                    for areas in allowed._areas)


def _names_after(workbook: Any) -> tuple[str, ...]:
    return tuple(node.get("name") or "" for node in _sheet_nodes(workbook))


def _refresh(workbook: Any) -> None:
    """Forget what was read: the sheets changed.  The cell writer's sheets by name are
    found again (a renamed or deleted sheet's name may come back on another part)."""
    workbook._forget()
    editor = workbook._editor
    if editor is not None and editor._cells is not None:
        editor._cells._sheets.clear()


# -- what refers to a sheet ------------------------------------------------------------------------


@dataclass
class _Site:
    """A formula held somewhere in the package: its kind, address, element (its text, or
    ``attr``) and part."""

    kind: str
    address: str
    node: Any
    part: str
    attr: str | None = None

    @property
    def text(self) -> str:
        return (self.node.get(self.attr) if self.attr else self.node.text) or ""

    def set(self, text: str, package: Any) -> None:
        if self.attr:
            self.node.set(self.attr, text)
        else:
            self.node.text = text
        package.mark_dirty(self.part)


def _local(node: Any) -> str:
    return etree.QName(node).localname if isinstance(node.tag, str) else ""


def _mentions(data: bytes | None, words: Sequence[str]) -> bool:
    """A cheap test before a part is parsed: whether its bytes hold any of ``words``
    (case aside, a ``'`` doubled or not)."""
    if not data:
        return False
    lowered = data.lower()
    for word in words:
        for form in {word.lower(), word.lower().replace("'", "''")}:
            if form.encode("utf-8") in lowered:
                return True
    return False


def _sites(workbook: Any, words: Sequence[str]) -> Iterator[_Site]:
    """Every formula in the package that may name one of ``words`` (sheet or table
    names): cells, rules and validations of every worksheet, tables' column formulas,
    defined names, chart references, pivot cache sources; and hyperlinks' locations."""
    package = workbook.package
    root = _root(workbook)
    names = _child(root, "definedNames")
    if names is not None:
        for node in names:
            if node.tag == _X + "definedName":
                label = node.get("name") or ""
                local = node.get("localSheetId")
                if local is not None:
                    nodes = _sheet_nodes(workbook)
                    index = _int_attr(node, "localSheetId", -1)
                    if 0 <= index < len(nodes):
                        label = quote_sheet(nodes[index].get("name") or "?") + "!" + label
                yield _Site("name", label, node, workbook.part)
    for sheet in workbook._every_sheet:
        if sheet.part is None or sheet.type not in ("worksheet", "macrosheet"):
            continue
        if _mentions(package.read(sheet.part), words):
            yield from _sheet_sites(workbook, sheet)
        for rel in package.relationships(sheet.part).values():
            target = rel.target_part
            if rel.is_external or target is None or not package.has_part(target):
                continue
            if rel.type.endswith("/table") and _mentions(package.read(target), words):
                yield from _table_sites(package, target)
    for part in package.part_names:
        content_type = package.content_type(part) or ""
        if content_type.endswith("drawingml.chart+xml") or \
                content_type.endswith("chartex+xml"):
            if _mentions(package.read(part), words):
                chart = package.tree(part)
                label = _chart_label(workbook, part)
                for node in chart.iter():
                    if _local(node) == "f" and isinstance(node.tag, str):
                        yield _Site("chart", label, node, part)
        elif content_type.endswith("pivotCacheDefinition+xml"):
            if _mentions(package.read(part), words):
                cache = package.tree(part)
                for node in cache.iter(_X + "worksheetSource"):
                    if node.get("sheet") is not None:
                        yield _Site("pivot_cache", part, node, part, attr="sheet")


def _sheet_sites(workbook: Any, sheet: Any) -> Iterator[_Site]:
    tree = workbook.package.tree(sheet.part)
    label = quote_sheet(sheet.stored_name)
    data = _child(tree, "sheetData")
    if data is not None:
        for f in data.iter(_X + "f"):
            if not f.text:
                continue
            cell = f.getparent()
            where = f.get("ref") if f.get("t") in ("shared", "array") and f.get("ref") \
                else cell.get("r") or "?"
            yield _Site("formula", f"{label}!{where}", f, sheet.part)
    for rules in tree.iterchildren(_X + "conditionalFormatting"):
        where = (rules.get("sqref") or "?").split()[0]
        for node in rules.iter(_X + "formula"):
            yield _Site("conditional_format", f"{label}!{where}", node, sheet.part)
    holder = _child(tree, "dataValidations")
    if holder is not None:
        for rule in holder:
            where = (rule.get("sqref") or "?").split()[0]
            for local in ("formula1", "formula2"):
                node = rule.find(_X + local)
                if node is not None:
                    yield _Site("validation", f"{label}!{where}", node, sheet.part)
    links = _child(tree, "hyperlinks")
    if links is not None:
        for link in links:
            if link.get("location"):
                yield _Site("hyperlink", f"{label}!{link.get('ref') or '?'}", link,
                            sheet.part, attr="location")
    extensions = _child(tree, "extLst")
    if extensions is not None:
        for node in extensions.iter():
            if _local(node) != "f" or node.tag == _X + "f":
                continue
            kind, where = _ext_kind(node)
            yield _Site(kind, f"{label}!{where}", node, sheet.part)


def _ext_kind(node: Any) -> tuple[str, str]:
    """An extension formula's kind (an ``x14`` rule, validation or sparkline) and the
    cells it applies to, from its ancestors."""
    kind, where = "formula", "?"
    parent = node.getparent()
    while parent is not None:
        local = _local(parent)
        if local in ("conditionalFormatting", "cfRule"):
            kind = "conditional_format"
        elif local == "dataValidation":
            kind = "validation"
        elif local in ("sparkline", "sparklineGroup"):
            kind = "sparkline"
        for child in parent:
            if _local(child) == "sqref" and child.text:
                where = child.text.split()[0]
        if where != "?" and kind != "formula":
            break
        parent = parent.getparent()
    return kind, where


def _table_sites(package: Any, part: str) -> Iterator[_Site]:
    table = package.tree(part)
    name = table.get("displayName") or table.get("name") or part
    for column in table.iter(_X + "tableColumn"):
        for local in ("calculatedColumnFormula", "totalsRowFormula"):
            node = column.find(_X + local)
            if node is not None and node.text:
                yield _Site("table_formula", f"t:{name}[{column.get('name') or '?'}]",
                            node, part)


def _chart_label(workbook: Any, part: str) -> str:
    """A chart's address: the sheet whose drawing holds it, and its part."""
    package = workbook.package
    by_part = {s.part: s for s in workbook._every_sheet if s.part}
    for source, _ in package.relationship_sources(part):
        for owner, _ in package.relationship_sources(source):
            if owner in by_part:
                return f"{quote_sheet(by_part[owner].stored_name)}!{part}"
        if source in by_part:
            return f"{quote_sheet(by_part[source].stored_name)}!{part}"
    return part


def _tables_on(workbook: Any, sheet: Any) -> dict[str, str]:
    """The tables on ``sheet``: their names case folded, to their parts."""
    package = workbook.package
    out: dict[str, str] = {}
    if sheet.part is None:
        return out
    for rel in package.relationships(sheet.part).values():
        target = rel.target_part
        if rel.type.endswith("/table") and not rel.is_external and target \
                and package.has_part(target):
            table = package.tree(target)
            for attr in ("displayName", "name"):
                if table.get(attr):
                    out[decode_xstring(table.get(attr)).casefold()] = target
    return out


def referents(workbook: Any, which: Any) -> list[Referent]:
    """Everything outside the sheet ``which`` that refers to it (see :class:`Referent`):
    formulas, defined names, rules, validations, chart series, pivot caches, tables'
    formulas and hyperlinks naming it, and formulas naming a table on it; each once, by
    address, in the package's order."""
    sheet = find_sheet(workbook, which)
    return _referents(workbook, sheet, tables=True)


def _referents(workbook: Any, sheet: Any, *, tables: bool, cells: bool = True
               ) -> list[Referent]:
    order = [s.stored_name for s in workbook._every_sheet]
    name = sheet.stored_name
    on_it = _tables_on(workbook, sheet) if tables else {}
    words = ([name] if cells else []) + list(on_it)
    found: list[Referent] = []
    seen = set()
    for site in _sites(workbook, words):
        if site.part == sheet.part or site.part in on_it.values():
            continue
        if site.kind == "name" and _local_to(workbook, site.node, sheet):
            continue
        hit = False
        if cells:
            if site.kind == "pivot_cache":
                hit = refs.fold_name(site.text) == refs.fold_name(name)
            elif site.kind == "hyperlink":
                hit = _location_names(site.text, name)
            else:
                hit = refs.refers_to(site.text, name, order)
        if not hit and on_it and site.kind not in ("pivot_cache", "hyperlink"):
            hit = refs.tables_named(site.text, set(on_it))
        if hit and (site.kind, site.address) not in seen:
            seen.add((site.kind, site.address))
            found.append(Referent(site.kind, site.address))
    return found


def _location_names(location: str, name: str) -> bool:
    """Whether a hyperlink's location (``Lists!A1``, ``'Q3 sales'!B2``) names the sheet."""
    prefix, sep, _ = location.rpartition("!")
    if not sep:
        return False
    read = refs.read_prefix(prefix + "!")
    return read is not None and refs.fold_name(read.first) == refs.fold_name(name)


def _local_to(workbook: Any, node: Any, sheet: Any) -> bool:
    if node.get("localSheetId") is None:
        return False
    nodes = _sheet_nodes(workbook)
    index = _int_attr(node, "localSheetId", -1)
    return 0 <= index < len(nodes) and nodes[index].get("name") == sheet.stored_name


# -- add -----------------------------------------------------------------------------------------


def add_sheet(workbook: Any, name: str, index: int | None = None) -> SheetChange:
    """Add an empty worksheet called ``name`` at ``index`` (0 is the first tab; ``None``
    the end), as Excel makes one (see the module); the active sheet stays the one it was.
    Refused: a name Excel refuses or the workbook has (``bad_sheet_name``,
    ``sheet_exists``), an index off the tabs (``bad_index``), a locked structure."""
    _guard(workbook)
    _check_name(workbook, name)
    nodes = _sheet_nodes(workbook)
    if index is None:
        index = len(nodes)
    if isinstance(index, bool) or not isinstance(index, int) or not 0 <= index <= len(nodes):
        raise SheetError("bad_index", f"a sheet's place is 0 to {len(nodes)} (the end)")
    package = workbook.package
    with workbook.editor.batch():
        part = package.unused_part_name("xl/worksheets/sheet{n}.xml")
        package.add_part(part, _new_sheet_xml(workbook), CT_WORKSHEET, override=True)
        rel_id = package.add_relationship(workbook.part, REL_WORKSHEET, part)
        root = _root(workbook)
        holder = _child(root, "sheets")
        ids = [_int_attr(n, "sheetId") for n in nodes]
        sheet_id = max(ids, default=0) + 1
        node = etree.Element(_X + "sheet")
        node.set("name", name)
        node.set("sheetId", str(sheet_id))
        node.set(f"{{{R}}}id", rel_id)
        if index < len(nodes):
            nodes[index].addprevious(node)
        else:
            holder.append(node)
        node.tail = nodes[0].tail if nodes else None
        _shift_locals(workbook, index, +1)
        for view in _views(workbook):
            for attr in ("activeTab", "firstSheet"):
                if view.get(attr) is not None and _int_attr(view, attr) >= index:
                    view.set(attr, str(_int_attr(view, attr) + 1))
        package.mark_dirty(workbook.part)
        _refresh(workbook)
    return SheetChange("add", name, sheet_id, sheets=_names_after(workbook))


def _new_sheet_xml(workbook: Any) -> bytes:
    """A new worksheet's part, as Excel writes one (see the module)."""
    fmt = ""
    for sheet in workbook._every_sheet:
        if sheet.type != "worksheet" or sheet.part is None:
            continue
        tree = workbook.package.tree(sheet.part)
        node = _child(tree, "sheetFormatPr")
        if node is not None and node.get("defaultRowHeight"):
            attrs = "".join(f' {a}="{node.get(a)}"' for a in ("baseColWidth", "defaultRowHeight")
                            if node.get(a) is not None)
            fmt = f"<sheetFormatPr{attrs}/>"
        break
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            f'<worksheet xmlns="{X}" xmlns:r="{R}"><dimension ref="A1"/>'
            '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
            f'{fmt}<sheetData/>'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" '
            'footer="0.3"/></worksheet>').encode("utf-8")


def _shift_locals(workbook: Any, index: int, delta: int) -> None:
    """Defined names' ``localSheetId`` from ``index`` on moved by ``delta``."""
    names = _child(_root(workbook), "definedNames")
    if names is None:
        return
    for node in names:
        if node.get("localSheetId") is not None and _int_attr(node, "localSheetId") >= index:
            node.set("localSheetId", str(_int_attr(node, "localSheetId") + delta))


# -- rename --------------------------------------------------------------------------------------


def rename_sheet(workbook: Any, which: Any, name: str) -> SheetChange:
    """Call the sheet ``which`` ``name``, every reference to it rewritten as Excel rewrites
    it (see the module).  The same name, as stored: nothing to do (``changed`` false).
    Refused: a name Excel refuses or another sheet has, a locked structure, a sheet the
    host's policy does not cover whole."""
    sheet = find_sheet(workbook, which)
    old = sheet.stored_name
    if name == old:
        return SheetChange("rename", old, sheet.id, changed=False,
                           sheets=_names_after(workbook), note="the sheet has that name")
    _guard(workbook, sheet)
    _check_name(workbook, name, keep=sheet)
    package = workbook.package
    rewritten: list[Referent] = []
    stale: list[Referent] = []
    with workbook.editor.batch():
        node = _sheet_nodes(workbook)[_index_of(workbook, sheet)]
        for site in list(_sites(workbook, [old])):
            if site.kind == "hyperlink":
                if _location_names(site.text, old):
                    stale.append(Referent(site.kind, site.address))
                continue
            if site.kind == "pivot_cache":
                if refs.fold_name(site.text) == refs.fold_name(old):
                    site.set(name, package)
                    rewritten.append(Referent(site.kind, site.address))
                continue
            text = refs.rename(site.text, old, name)
            if text is not None:
                site.set(text, package)
                rewritten.append(Referent(site.kind, _renamed(site.address, old, name)))
        node.set("name", name)
        package.mark_dirty(workbook.part)
        _rename_in_app(package, old, name)
        _refresh(workbook)
    return SheetChange("rename", name, sheet.id, sheets=_names_after(workbook),
                       rewritten=tuple(_unique(rewritten)), not_rewritten=tuple(stale))


def _renamed(address: str, old: str, new: str) -> str:
    """A referent's address on the renamed sheet, named as it is now."""
    prefix = quote_sheet(old) + "!"
    return quote_sheet(new) + "!" + address[len(prefix):] if address.startswith(prefix) \
        else address


def _unique(found: list[Referent]) -> list[Referent]:
    return list(dict.fromkeys(found))


_APP = "docProps/app.xml"
_VT = "{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}"


def _rename_in_app(package: Any, old: str, new: str) -> None:
    """``docProps/app.xml``'s titles brought up to date: the sheet's own entry, and the
    names local to it (``Lists!Local``), as Excel writes them on a rename (measured)."""
    if not package.has_part(_APP):
        return
    try:
        root = package.tree(_APP)
    except etree.XMLSyntaxError:
        return
    if root is None:
        return
    changed = False
    for node in root.iter(_VT + "lpstr"):
        text = node.text or ""
        if text == old:
            node.text = new
            changed = True
        else:
            prefix, sep, rest = text.rpartition("!")
            read = refs.read_prefix(prefix + "!") if sep else None
            if read is not None and read.last is None and read.first == old:
                node.text = quote_sheet(new) + "!" + rest
                changed = True
    if changed:
        package.mark_dirty(_APP)


# -- delete --------------------------------------------------------------------------------------


def delete_sheet(workbook: Any, which: Any, *, force: bool = False) -> SheetChange:
    """Delete the sheet ``which``.  Refused while anything outside it refers to it
    (``referred``: :attr:`SheetError.referents`, each by address) unless ``force``, which
    deletes it as Excel does (see the module): each reference ``#REF!``.  Refused too: the
    last sheet, the last visible sheet, a locked structure, a sheet the host's policy does
    not cover whole."""
    sheet = find_sheet(workbook, which)
    _guard(workbook, sheet)
    nodes = _sheet_nodes(workbook)
    if len(nodes) == 1:
        raise SheetError("last_sheet", "a workbook keeps at least one sheet: add another "
                         "first")
    index = _index_of(workbook, sheet)
    visible = [i for i, n in enumerate(nodes) if (n.get("state") or "visible") == "visible"]
    if visible == [index]:
        raise SheetError("last_visible", "it is the only visible sheet: show another "
                         "first (set_sheet_state)")
    found = _referents(workbook, sheet, tables=True)
    if found and not force:
        raise SheetError("referred", f"{len(found)} referent(s) refer to the sheet or a "
                         "table on it; Excel would write #REF! into each. To regenerate the "
                         "sheet, clear it (clear_sheet) and write it again; or delete with "
                         "force", referents=found,
                         addresses=[r.address for r in found])
    order = [n.get("name") or "" for n in nodes]
    name = sheet.stored_name
    tables = set(_tables_on(workbook, sheet))
    package = workbook.package
    rewritten: list[Referent] = []
    stale: list[Referent] = []
    removed: list[str] = []
    table_parts = set(_tables_on(workbook, sheet).values())
    with workbook.editor.batch():
        for site in list(_sites(workbook, [name] + sorted(tables))):
            if site.part == sheet.part or site.part in table_parts:
                continue
            if site.kind == "name" and _local_to(workbook, site.node, sheet):
                continue
            if site.kind == "hyperlink":
                if _location_names(site.text, name):
                    stale.append(Referent(site.kind, site.address))
                continue
            if site.kind == "pivot_cache":
                if refs.fold_name(site.text) == refs.fold_name(name):
                    stale.append(Referent(site.kind, site.address))
                continue
            text = refs.delete(site.text, name, order)
            if tables:
                dropped = refs.drop_tables(text if text is not None else site.text, tables)
                text = dropped if dropped is not None else text
            if text is not None:
                site.set(text, package)
                rewritten.append(Referent(site.kind, site.address))
        removed += _drop_local(workbook, index)
        _shift_locals(workbook, index + 1, -1)
        _move_views(workbook, index, removing=True)
        node = nodes[index]
        rel_id = node.get(f"{{{R}}}id")
        node.getparent().remove(node)
        package.mark_dirty(workbook.part)
        pivots = _pivot_tables(package, sheet.part) if sheet.part else []
        if rel_id and rel_id in package.relationships(workbook.part):
            package.release(workbook.part, [rel_id])
        removed += _drop_unused_caches(workbook, pivots)
        writer = workbook.editor.cells
        writer.drop_calc_chain()
        writer.request_full_calculation()
        _recount_strings(workbook)
        _refresh(workbook)
    return SheetChange("delete", name, sheet.id, sheets=_names_after(workbook),
                       rewritten=tuple(_unique(rewritten)), not_rewritten=tuple(stale),
                       removed=tuple(removed))


def _drop_local(workbook: Any, index: int) -> list[str]:
    """Remove the defined names local to the sheet at ``index``; their names."""
    names = _child(_root(workbook), "definedNames")
    removed: list[str] = []
    if names is None:
        return removed
    for node in list(names):
        if node.get("localSheetId") is not None and _int_attr(node, "localSheetId") == index:
            removed.append(f"name:{node.get('name')}")
            names.remove(node)
    if not len(names):
        names.getparent().remove(names)
    return removed


def _move_views(workbook: Any, index: int, *, removing: bool) -> None:
    """The workbook views' active and first tab after the sheet at ``index`` goes
    (``removing``) or is hidden, as Excel moves them (see the module): the sheet now
    active selected, the one at ``index`` no longer."""
    nodes = _sheet_nodes(workbook)
    states = [(n.get("state") or "visible") for n in nodes]
    target: int | None = None
    for view in _views(workbook):
        active = _int_attr(view, "activeTab")
        if active == index:
            target = _next_visible(states, index)
            new = target if target is not None else 0
            if removing and new > index:
                new -= 1
            if new:
                view.set("activeTab", str(new))
            else:
                view.attrib.pop("activeTab", None)
        elif removing and active > index:
            view.set("activeTab", str(active - 1))
        first = view.get("firstSheet")
        if first is not None and removing and _int_attr(view, "firstSheet") > index:
            view.set("firstSheet", str(_int_attr(view, "firstSheet") - 1))
        elif first is not None and _int_attr(view, "firstSheet") == index:
            view.attrib.pop("firstSheet", None)
    _select(workbook, nodes, target, exclude=index)


def _next_visible(states: list[str], index: int) -> int | None:
    """The visible sheet to the right of ``index``, else to its left (measured)."""
    for i in list(range(index + 1, len(states))) + list(range(index - 1, -1, -1)):
        if states[i] == "visible":
            return i
    return None


def _select(workbook: Any, nodes: list[Any], target: int | None, *, exclude: int) -> None:
    """Make the sheet at ``target`` the selected tab, and the one at ``exclude`` not."""
    package = workbook.package
    rels = package.relationships(workbook.part)
    for i, node in enumerate(nodes):
        if i not in (target, exclude):
            continue
        rel = rels.get(node.get(f"{{{R}}}id") or "")
        part = rel.target_part if rel is not None and not rel.is_external else None
        if part is None or not package.has_part(part):
            continue
        tree = package.tree(part)
        views = _child(tree, "sheetViews")
        view = views.find(_X + "sheetView") if views is not None else None
        if view is None:
            if i != target:
                continue
            if views is None:
                views = etree.Element(_X + "sheetViews")
                placed(tree, views)
            view = etree.SubElement(views, _X + "sheetView")
            view.set("workbookViewId", "0")
        if i == target:
            if view.get("tabSelected") not in ("1", "true"):
                view.set("tabSelected", "1")
                package.mark_dirty(part)
        elif view.get("tabSelected") is not None:
            view.attrib.pop("tabSelected", None)
            package.mark_dirty(part)


def _pivot_tables(package: Any, part: str) -> list[str]:
    return [rel.target_part for rel in package.relationships(part).values()
            if rel.type.endswith("/pivotTable") and not rel.is_external and rel.target_part]


def _drop_unused_caches(workbook: Any, pivots: list[str]) -> list[str]:
    """After pivot tables went: the caches they used that no pivot table uses now,
    removed with their ``pivotCache`` entries.  What was removed."""
    package = workbook.package
    removed = [f"pivot_table:{p}" for p in pivots]
    used = set()
    for part in package.part_names:
        if (package.content_type(part) or "").endswith("pivotTable+xml"):
            for rel in package.relationships(part).values():
                if rel.type.endswith("/pivotCacheDefinition") and rel.target_part:
                    used.add(rel.target_part)
    root = _root(workbook)
    holder = _child(root, "pivotCaches")
    rels = package.relationships(workbook.part)
    for rel in list(rels.values()):
        if not rel.type.endswith("/pivotCacheDefinition") or rel.target_part in used:
            continue
        if holder is not None:
            for entry in list(holder):
                if entry.get(f"{{{R}}}id") == rel.id:
                    holder.remove(entry)
            if not len(holder):
                root.remove(holder)
                holder = None
        package.mark_dirty(workbook.part)
        package.release(workbook.part, [rel.id])
        removed.append(f"pivot_cache:{rel.target_part}")
    return removed


def _recount_strings(workbook: Any) -> None:
    """The shared strings' counts brought up to date when the step closes: cells went."""
    writer = workbook.editor.cells
    part = writer._shared_strings_part()
    if part is not None and workbook.package.tree(part) is not None:
        workbook.package.mark_dirty(part)


# -- clear ---------------------------------------------------------------------------------------


def clear_sheet(workbook: Any, which: Any, *, force: bool = False) -> SheetChange:
    """Excel's Clear All on every cell of the sheet ``which`` (see the module): its cells,
    merged areas, rules, validations, comments, hyperlinks, tables and pivot tables go; its
    widths, heights, views, print settings and drawing stay, and so does every reference
    to the sheet.  Refused: a formula elsewhere naming a table on the sheet (``referred``)
    unless ``force`` (it then reads ``#REF!``), a protected sheet, a chart sheet, a sheet
    the host's policy does not cover whole.  An empty sheet: ``changed`` false."""
    sheet = find_sheet(workbook, which)
    if sheet.type != "worksheet" or sheet.part is None:
        raise EditError("not_a_worksheet", addresses=[quote_sheet(sheet.stored_name) + "!"])
    _guard(workbook, sheet, structure=False)
    if sheet.protected:
        raise EditError("protected", addresses=[quote_sheet(sheet.stored_name) + "!"])
    tables = set(_tables_on(workbook, sheet))
    found = _referents(workbook, sheet, tables=True, cells=False) if tables else []
    if found and not force:
        raise SheetError("referred", f"{len(found)} formula(s) elsewhere name a table on the "
                         "sheet, which Clear All removes (they would read #REF!): write the "
                         "sheet's cells instead, or clear with force", referents=found,
                         addresses=[r.address for r in found])
    package = workbook.package
    tree = package.tree(sheet.part)
    if _empty(tree, package, sheet.part):
        return SheetChange("clear", sheet.stored_name, sheet.id, changed=False,
                           sheets=_names_after(workbook), note="the sheet holds nothing to "
                           "clear")
    rewritten: list[Referent] = []
    removed: list[str] = []
    with workbook.editor.batch():
        if tables:
            on_it = _tables_on(workbook, sheet)
            for site in list(_sites(workbook, sorted(tables))):
                if site.part == sheet.part or site.part in on_it.values() \
                        or site.kind in ("pivot_cache", "hyperlink"):
                    continue
                text = refs.drop_tables(site.text, tables)
                if text is not None:
                    site.set(text, package)
                    rewritten.append(Referent(site.kind, site.address))
        cells = _clear_cells(tree)
        removed += _clear_parts(workbook, sheet, tree)
        dimension = _child(tree, "dimension")
        if dimension is not None:
            dimension.set("ref", "A1")
        package.mark_dirty(sheet.part)
        writer = workbook.editor.cells
        writer.drop_calc_chain()
        writer.request_full_calculation()
        _recount_strings(workbook)
        _refresh(workbook)
        writer.sheet(sheet.stored_name)
    return SheetChange("clear", sheet.stored_name, sheet.id, sheets=_names_after(workbook),
                       rewritten=tuple(_unique(rewritten)), removed=tuple(removed),
                       cells=cells)


def _empty(tree: Any, package: Any, part: str) -> bool:
    data = _child(tree, "sheetData")
    if data is not None and data.find(f"{_X}row/{_X}c") is not None:
        return False
    if any(_child(tree, local) is not None for local in
           ("mergeCells", "conditionalFormatting", "dataValidations", "hyperlinks",
            "tableParts", "legacyDrawing")):
        return False
    if data is not None and any(row.get("s") is not None for row in data):
        return False
    cols = _child(tree, "cols")
    if cols is not None and any(col.get("style") is not None for col in cols):
        return False
    return not _pivot_tables(package, part) and not _cleared_extensions(tree)


def _clear_cells(tree: Any) -> int:
    """Every cell removed; a row kept only for its height, hidden state or outline level
    (its format goes, as Clear All's does), a column's format gone, its width kept.  How
    many cells went."""
    data = _child(tree, "sheetData")
    count = 0
    if data is not None:
        for row in list(data):
            if row.tag != _X + "row":
                continue
            cells = [c for c in row if c.tag == _X + "c"]
            count += len(cells)
            for cell in cells:
                row.remove(cell)
            for attr in ("s", "customFormat", "spans"):
                row.attrib.pop(attr, None)
            if row.get("customHeight") not in ("1", "true"):
                for attr in ("ht",):
                    row.attrib.pop(attr, None)
            kept = {k for k in row.attrib if etree.QName(k).localname not in ("r", "dyDescent")}
            if not kept and not len(row):
                data.remove(row)
    cols = _child(tree, "cols")
    if cols is not None:
        for col in list(cols):
            col.attrib.pop("style", None)
            if set(col.attrib) <= {"min", "max"}:
                cols.remove(col)
        if not len(cols):
            tree.remove(cols)
    return count


def _cleared_extensions(tree: Any) -> list[Any]:
    extensions = _child(tree, "extLst")
    if extensions is None:
        return []
    return [ext for ext in extensions if any(_local(child) in _CLEARED_EXT for child in ext)]


def _clear_parts(workbook: Any, sheet: Any, tree: Any) -> list[str]:
    """The sheet's merged areas, rules, validations, hyperlinks, comments, tables and pivot
    tables removed, with the parts only they used.  What went."""
    package = workbook.package
    removed = []
    release: list[str] = []
    for local in ("mergeCells", "dataValidations"):
        node = _child(tree, local)
        if node is not None:
            removed.append(local)
            tree.remove(node)
    rules = list(tree.iterchildren(_X + "conditionalFormatting"))
    if rules:
        removed.append("conditionalFormatting")
        for node in rules:
            tree.remove(node)
    links = _child(tree, "hyperlinks")
    if links is not None:
        removed.append("hyperlinks")
        release += [link.get(f"{{{R}}}id") for link in links if link.get(f"{{{R}}}id")]
        tree.remove(links)
    for ext in _cleared_extensions(tree):
        for child in list(ext):
            if _local(child) in _CLEARED_EXT:
                removed.append(f"x14:{_local(child)}")
                ext.remove(child)
        if not any(isinstance(c.tag, str) for c in ext):
            ext.getparent().remove(ext)
    extensions = _child(tree, "extLst")
    if extensions is not None and not len(extensions):
        tree.remove(extensions)
    legacy = _child(tree, "legacyDrawing")
    rels = package.relationships(sheet.part)
    comments = [rel.id for rel in rels.values()
                if rel.type.rpartition("/")[2] in ("comments", "threadedComment")]
    if legacy is not None and comments:
        removed.append("comments")
        release += comments + [legacy.get(f"{{{R}}}id")]
        tree.remove(legacy)
    parts = _child(tree, "tableParts")
    if parts is not None:
        for table in parts:
            release.append(table.get(f"{{{R}}}id"))
            rel = rels.get(table.get(f"{{{R}}}id") or "")
            if rel is not None and rel.target_part and package.has_part(rel.target_part):
                root = package.tree(rel.target_part)
                removed.append(f"table:{root.get('displayName') or root.get('name')}")
        tree.remove(parts)
    pivots = _pivot_tables(package, sheet.part)
    release += [rel.id for rel in rels.values() if rel.target_part in pivots]
    package.mark_dirty(sheet.part)
    package.release(sheet.part, [r for r in release if r])
    removed += _drop_unused_caches(workbook, pivots) if pivots else []
    return removed


# -- hide and show -------------------------------------------------------------------------------


def set_sheet_state(workbook: Any, which: Any, state: str) -> SheetChange:
    """Show or hide the sheet ``which``: ``state`` ``"visible"``, ``"hidden"`` or
    ``"veryHidden"``, as Excel does (see the module).  Already so: ``changed`` false.
    Refused: hiding the last visible sheet (``last_visible``), a state not listed, a
    locked structure, a sheet the host's policy does not cover whole."""
    if state not in STATES:
        raise SheetError("bad_state", f"a sheet's state is one of {', '.join(STATES)}",
                         options=list(STATES))
    sheet = find_sheet(workbook, which)
    index = _index_of(workbook, sheet)
    nodes = _sheet_nodes(workbook)
    if (nodes[index].get("state") or "visible") == state:
        return SheetChange("state", sheet.stored_name, sheet.id, changed=False,
                           sheets=_names_after(workbook), note=f"the sheet is {state}")
    _guard(workbook, sheet)
    if state != "visible":
        visible = [i for i, n in enumerate(nodes) if (n.get("state") or "visible") == "visible"]
        if visible == [index]:
            raise SheetError("last_visible", "a workbook shows at least one sheet: show "
                             "another first")
    with workbook.editor.batch():
        node = nodes[index]
        if state == "visible":
            node.attrib.pop("state", None)
        else:
            was_visible = (node.get("state") or "visible") == "visible"
            if was_visible:
                _move_views(workbook, index, removing=False)
            node.set("state", state)
            # The attribute after name and sheetId, before r:id, as Excel writes it.
            rid = node.get(f"{{{R}}}id")
            if rid is not None:
                del node.attrib[f"{{{R}}}id"]
                node.set(f"{{{R}}}id", rid)
        workbook.package.mark_dirty(workbook.part)
        _refresh(workbook)
    return SheetChange("state", sheet.stored_name, sheet.id, sheets=_names_after(workbook))


def placed(tree: Any, child: Any) -> None:
    """Put ``child`` into the worksheet ``tree`` where CT_Worksheet's sequence has it."""
    rank = WORKSHEET_ORDER.index(_local(child))
    for index, existing in enumerate(tree):
        local = _local(existing)
        if local in WORKSHEET_ORDER and WORKSHEET_ORDER.index(local) > rank:
            tree.insert(index, child)
            return
    tree.append(child)


__all__ = ["MAX_NAME", "Referent", "STATES", "SheetChange", "SheetError", "WORKSHEET_ORDER",
           "add_sheet", "clear_sheet", "delete_sheet", "find_sheet", "name_problems",
           "placed", "referents", "rename_sheet", "set_sheet_state"]
