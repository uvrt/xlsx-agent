"""The read tools: ``xlsx_read`` (cells by address, typed, paged) and ``xlsx_profile``
(every column of a region in full: ``describe``'s drill-down).  Both are in the loaded core
(ROADMAP.md, "Tools": the read tools are loaded).

**xlsx_read** takes any address that names cells (ROADMAP.md, "Addresses"): a cell or range
(``Stock!A1:H40``; whole columns and rows, and any rectangle of more than one cell, are
clipped to the sheet's used range, outside which every cell is blank), a table, name or
region (its data rows), a column, a row by position or key, or a value.  Each cell comes
back typed with its address (:meth:`Value.as_dict <xlsx_agent.read.types.Value.as_dict>`),
row by row; on request with its number format's code (``formats``), its formula's text
(``formulas``: ``f``; a shared group's other cells say ``shared_from``), its comment
(``comments``) and its hyperlink (``hyperlinks``), and with the merged areas over the rows
(``merges``).  The blocks array formulas and dynamic arrays fill that cross the rows are
always listed (``arrays``): a cell inside one says nothing of it by itself.

A dataset (``d:by_region``, ``d:by_region[total]``, ``d:by_region/#2/total``...) is read by
:func:`.derive.read_dataset`: no ``doc`` is needed, since a dataset outlives its workbook,
and ``sources`` adds each value's source cells (R11).  Every other target needs ``doc``.

**Declared tables** (:mod:`xlsx_agent.declare`): a table over several sheets is read as one,
each cell with its own address; the result's ``range`` is each sheet's, joined by ``,``,
and it names no one ``sheet``.  A cell of a column hidden from the model is never given:
a read of the sheet's cells gives it as ``{"a": ..., "withheld": true}``, and a visible formula that reaches one
gives its cached value, but with ``formulas`` no ``f``: ``"formula_withheld": true``
instead (:meth:`Regions.formula_withheld
<xlsx_agent.read.regions.Regions.formula_withheld>`).  A declared
column's cleaning rules (:mod:`xlsx_agent.clean`) apply to a read through its table: a
value a rule changed says ``"cleaned": true``, a split one is a ``list``, and the page
counts them (``cleaned``: how many, the first addresses); a read of the sheet's cells
gives them as stored.

**Bounded.**  A page holds at most ``limit`` rows (by default ``Limits.max_list_items``, at
most :data:`READ_ROWS`) and its whole result at most ``max_chars`` characters (by default,
and at most, ``Limits.max_result_chars``); a row too wide for what is left is split, the
rest of it starting the next page.  Text longer than :data:`TEXT_SHOWN` characters (a
value, a formula, a comment, a link) is cut, with its ``length``; in a read of one cell, only
as far as the page needs.  ``cut`` says what the page left out: ``rows`` not yet given
(the split row among them), ``cells`` of the split row, ``chars`` when the page was full
before ``limit`` rows, ``text`` the values whose text was cut, ``column_names`` when a
region's names would take more than a quarter of the page (``describe`` and
``xlsx_profile`` give them); ``next_cursor``, passed back as
``cursor`` with the same target, gives the next page.  ``total`` is the target's rows.

**xlsx_profile** gives :mod:`xlsx_agent.profile`'s whole profile of a table's or region's
columns (all, the ``columns`` named, or the one a column address names), as
:meth:`ColumnStats.as_dict <xlsx_agent.profile.ColumnStats.as_dict>` writes it: one region
per call, :attr:`~ooxml_edit.tools.Limits.max_list_items` columns per page at most, within
``max_result_chars``; text in it cut at :data:`PROFILE_TEXT` characters and each allowed
list at :data:`ALLOWED_SHOWN` values, and counted in ``cut``.

**Content.**  Results carry cell content -- values, formulas, comments, links -- as asked;
summaries and errors carry addresses, counts and names only, never a key (ROADMAP.md,
"Design rules", 7).  A formula's text appears only in a result that asked for it.
"""

from __future__ import annotations

import json
import zlib
from typing import Any, Iterable

from ooxml_edit.charts.workbook import Area, Cell
from ooxml_edit.tools import (CORE, DATASET_PREFIX, Result, ToolError, array, boolean, integer,
                              string, tool)

from ..address import (RangeAddress, RefAddress, SheetRef, TableAddress, parse_address,
                       quote_sheet)
from ..read.types import Value
from .docformat import KIND, handler

#: The most rows a page of ``xlsx_read`` holds (ROADMAP.md, "Tools": "at most 500").
READ_ROWS = 500

#: Text longer than this, in a read of more than one cell, is cut, with its length.
TEXT_SHOWN = 2000

#: Text in a profile (a value, an allowed value) is cut at this many characters.
PROFILE_TEXT = 200

#: The most allowed values a profile shows per validation list; the rest are counted.
ALLOWED_SHOWN = 100

#: The fewest characters a page may be held to: room for the envelope and one cell.
MIN_CHARS = 1000

#: Room left in a page for what the dispatcher adds to the envelope (version, warnings).
_MARGIN = 400

#: The most addresses ``cut.text`` lists.
_FIRST = 3

DOC = string("Document id.")
CURSOR = string("next_cursor of the previous page, with the same target.", optional=True)


# -- xlsx_read ---------------------------------------------------------------------------------


@tool("xlsx_read",
      "Cells by address, typed, row by row, each with its address; paged. Options add "
      "number formats, formula text (f), comments, hyperlinks, merged areas. cut says what a "
      "page left out; pass next_cursor back for more.",
      {"doc": string("Document id; not needed for a dataset (d:).", optional=True),
       "target": string("A cell or range (Stock!A1:H40, Stock!A:C), table or region "
                        "(t:Products: its data rows), column (t:Products[Price]), row "
                        "(t:Products/#5, t:Products/SKU-0042) or value; * is every row "
                        "(t:Products/*/Price); or a dataset (d:by_region), as tables."),
       "formats": boolean("Add format, the number format code. Default false.", optional=True),
       "formulas": boolean("Add f, the formula text. Default false.", optional=True),
       "comments": boolean("Add comment (author, text). Default false.", optional=True),
       "hyperlinks": boolean("Add link (url or location). Default false.", optional=True),
       "merges": boolean("List the merged areas over the rows. Default false.", optional=True),
       "sources": boolean("Dataset: add each value's source cells (8, more counted); "
                          "for one value, all, paged.", optional=True),
       "key_column": string("The key column of a row or value by key, when describe gives "
                            "none or several.", optional=True),
       "limit": integer(f"Rows per page. Default 50, at most {READ_ROWS}.", minimum=1,
                        maximum=READ_ROWS, optional=True),
       "max_chars": integer("Characters per page. Default and at most 24000.",
                            minimum=MIN_CHARS, optional=True),
       "cursor": CURSOR},
      group=CORE, datasets=("target",))
@handler
def xlsx_read(call: Any, target: str, doc: str | None = None, formats: bool = False,
              formulas: bool = False, comments: bool = False, hyperlinks: bool = False,
              merges: bool = False, sources: bool = False, key_column: str | None = None,
              limit: int | None = None, max_chars: int | None = None,
              cursor: str | None = None) -> Result:
    """``xlsx_read``'s handler: a workbook's cells, or a dataset's values (see the
    module's notes)."""
    from .derive import read_dataset, workbook_of

    if isinstance(target, str) and target.startswith(DATASET_PREFIX):
        return read_dataset(call, target, sources=bool(sources), limit=limit,
                            max_chars=max_chars, cursor=cursor)
    if sources:
        raise ToolError("invalid_arguments", "sources is for a dataset (d:); a cell is its "
                        "own source", field="sources")
    workbook = workbook_of(call, doc)
    found = _cells(workbook, target, key_column)
    rows_limit = min(limit or call.limits.max_list_items, READ_ROWS)
    chars = min(max_chars or call.limits.max_result_chars, call.limits.max_result_chars)
    chars = max(chars, MIN_CHARS)
    area = found.area
    total = 0 if area is None else area.last.row - area.first.row + 1
    width = 0 if area is None else area.last.column - area.first.column + 1
    if found.rows is not None:          # a key set: its rows, in its order
        total = len(found.rows)
        width = found.rows[0].last.column - found.rows[0].first.column + 1
    start_row, start_column = _read_cursor(cursor, found.text, total, width)

    data: dict[str, Any] = {"target": found.text}
    if found.sheet is not None:
        data["sheet"] = found.sheet
    data["range"] = found.located(area)
    if found.requested is not None:
        data["requested"] = found.requested
    names_cut = 0
    if found.columns is not None:
        if _size(found.columns) <= chars // 4:
            data["columns"] = list(found.columns)
        else:  # a region too wide for its names on every page: describe and xlsx_profile
            names_cut = len(found.columns)
    data["rows"] = []
    if area is None and not found.rows:
        return Result(summary=f"{found.label}: no cells to read", data=data, total=0)

    if found.rows is not None:
        pages = found.rows[start_row:start_row + rows_limit]
        rows = workbook._read_rows([RangeAddress(found.ref, a) for a in pages],
                                   formats=formats)
        more = start_row + len(pages) < total
        single = False
    else:
        top = area.first.row + start_row
        bottom = min(area.last.row, top + rows_limit - 1)
        pages = [Area(Cell(top, area.first.column), Cell(bottom, area.last.column))]
        rows = workbook.read(RangeAddress(found.ref, pages[0]), formats=formats)
        more = bottom < area.last.row
        single = area.first == area.last
    annotate = _Annotations(workbook, found.sheets, comments=comments, hyperlinks=hyperlinks)
    # What lies over the page besides its cells, as wide as the page can be: its length is
    # kept out of the budget for the cells, and only shrinks once the page is known.
    extras = _extras(workbook, found.places(pages), merges)
    budget = chars - _MARGIN - _size({**data, **extras, "cut": _CUT_ROOM}) \
        - len(_summary(found.label, 1, total, total, True))
    cap = None if single else TEXT_SHOWN

    cleaned: list[str] = []
    shown, cut_text, stop, full = _page(rows, start_row, start_column, budget, annotate,
                                        formulas, cap, cleaned)
    if stop is None and more:
        stop = (start_row + (len(pages) if found.rows is not None
                             else bottom - top + 1), 0)

    data["rows"] = shown
    if cleaned:
        data["cleaned"] = cleaned_count(cleaned)
    # The rows given, each with one cell at least: from the page's first, one after another.
    given = max(len(shown), 1)
    if found.rows is not None:
        covered = pages[:given]
    else:
        covered = [Area(Cell(top, area.first.column),
                        Cell(top + given - 1, area.last.column))]
    data.update(_extras(workbook, found.places(covered), merges))
    cut: dict[str, Any] = {}
    next_cursor = None
    if stop is not None:
        cut["rows"] = total - stop[0]
        if stop[1]:
            cut["cells"] = width - stop[1]
        if full:
            cut["chars"] = True
        next_cursor = _cursor(found.text, *stop)
    if cut_text:
        cut["text"] = {"count": len(cut_text), "first": cut_text[:_FIRST]}
    if names_cut:
        cut["column_names"] = names_cut
    if cut:
        data["cut"] = cut
    summary = _summary(found.label, start_row + 1, start_row + len(shown), total,
                       next_cursor is not None)
    return Result(summary=summary, data=data, total=total, next_cursor=next_cursor)


def cleaned_count(addresses: list[str]) -> dict:
    """The cells of a page a declared rule changed: how many, and the first."""
    return {"count": len(addresses), "first": addresses[:_FIRST]}


def _page(rows: Iterable[list[Value]], start_row: int, start_column: int, budget: int,
          annotate: "_Annotations", formulas: bool, cap: int | None,
          cleaned: "list[str] | None" = None
          ) -> tuple[list[list[dict]], list[str], tuple[int, int] | None, bool]:
    """A page's rows, filled from ``rows`` (the first of them at ``start_row``, from its
    ``start_column``) within ``budget`` characters: the rows as cells, the addresses whose
    text was cut, where the next page starts when the budget ran out (row and column), and
    whether it did.  A first cell too big for the page by itself is cut to fit
    (:func:`_fit`); a row too wide for what is left is split.  ``cleaned`` collects the
    addresses of the cells given whose value a declared rule changed."""
    shown: list[list[dict]] = []
    used = 0
    cut_text: list[str] = []
    stop: tuple[int, int] | None = None
    full = False
    for index, values in enumerate(rows):
        row = start_row + index
        first = start_column if index == 0 else 0
        out: list[dict] = []
        used += 2                            # the row's brackets and the comma after it
        for offset, value in enumerate(values[first:], first):
            cell, texts = _cell(value, annotate, formulas, cap)
            size = _size(cell) + 1
            if used + size > budget:
                if shown or out:
                    stop, full = (row, offset), True
                    break
                cell, texts = _fit(value, annotate, formulas, budget - used)
                size = _size(cell) + 1
            out.append(cell)
            used += size
            cut_text.extend(texts)
            if cleaned is not None and value.cleaned:
                cleaned.append(value.a)
        if out:
            shown.append(out)
        if stop is not None:
            break
    return shown, cut_text, stop, full


class _Found:
    """What an address names, as cells: ``text`` its canonical form; ``label`` the same
    without a key (cell content: never in a summary, which the call log keeps); ``ref`` the
    sheet (by name); ``sheet`` both its forms, as a result names a sheet; ``area`` the
    cells to read (``None``: none); ``requested`` the rectangle asked for, when it was
    clipped; ``columns`` the region's column names, for a ``t:`` address; ``rows`` a key
    set's rows, one area each, in its order (``area`` is then ``None``)."""

    def __init__(self, text: str, ref: SheetRef, sheet: Any, sheet_object: Any,
                 area: Area | None, requested: str | None = None,
                 columns: list[str] | None = None, label: str | None = None,
                 rows: list[Area] | None = None, region: Any = None,
                 workbook: Any = None) -> None:
        self.rows = rows
        self.text = text
        self.label = label or text
        self.ref = ref
        self.sheet = sheet
        self.sheet_object = sheet_object
        self.area = area
        self.requested = requested
        self.columns = columns
        #: A declared table read as one (its region has a span): ``ref`` and ``area`` are
        #: its own, and :meth:`places` gives the sheets' cells.
        self.region = region if region is not None and region.span is not None else None
        if self.region is not None:
            self.sheets = [workbook.sheet(part.sheet) for part in self.region.span.parts]
        else:
            self.sheets = [sheet_object]

    def located(self, area: Area | None) -> str | None:
        """``area`` as the sheets' range(s), as text."""
        if area is None:
            return None
        if self.region is not None:
            return self.region.located(area)
        return RangeAddress(self.ref, area).text

    def places(self, areas: list[Area]) -> list[tuple[Any, SheetRef, Area]]:
        """``areas`` as each sheet's cells: the sheet, how to name it, the area."""
        if self.region is None:
            return [(self.sheet_object, self.ref, a) for a in areas]
        by_name = {sheet.name: sheet for sheet in self.sheets}
        return [(by_name[r.sheet.name], r.sheet, r.area)
                for a in areas for r in self.region.span.locate(a)]


def _cells(workbook: Any, target: str, key_column: str | None) -> _Found:
    address = parse_address(target)
    if isinstance(address, RefAddress):
        raise ToolError("not_found", f"no call has defined {address.text}", field="target")
    if key_column is not None and not (isinstance(address, TableAddress)
                                       and (address.key is not None
                                            or address.keys is not None)):
        raise ToolError("invalid_arguments", "key_column is for a row or value by key "
                        "(t:Products/SKU-0042)", field="key_column")
    if isinstance(address, TableAddress):
        selection = workbook.resolve(address, key=key_column)
        region = selection.region
        if region.span is not None:      # a declared table over several sheets
            columns = list(region.columns) if selection.column is None else [selection.column]
            label = selection.address
            if address.key is not None:
                label = f"a {selection.kind} by key of {region.address.text}"
            rows = None
            if address.keys is not None:
                label = f"{len(address.keys)} {selection.kind} by key of {region.address.text}"
                rows = [a.area for a in selection.areas]
            area = selection.range.area if selection.range is not None else None
            return _Found(selection.address, region.range.sheet, None, None, area,
                          columns=columns, label=label, rows=rows, region=region,
                          workbook=workbook)
        sheet = _sheet(workbook, workbook.sheet_of(region.range))
        columns = list(region.columns) if selection.column is None else [selection.column]
        area = selection.range.area if selection.range is not None else None
        label = selection.address
        if address.key is not None:
            label = f"a {selection.kind} by key of {region.address.text}"
        rows = None
        if address.keys is not None:
            label = f"{len(address.keys)} {selection.kind} by key of {region.address.text}"
            rows = [a.area for a in selection.areas]
        return _Found(selection.address, region.range.sheet, _forms(workbook, sheet), sheet,
                      area, columns=columns, label=label, rows=rows)
    sheet = _sheet(workbook, workbook.sheet_of(address))
    ref = SheetRef(name=sheet.name) if sheet.name else SheetRef(id=sheet.id)
    area = address.area
    requested = None
    if address.kind != "cell":
        area = address.clipped(workbook.used_range(sheet)) if address.span != "cells" else \
            _intersect(address.area, workbook.used_range(sheet))
        if area != address.area:
            requested = address.text
    return _Found(address.text, ref, _forms(workbook, sheet), sheet, area, requested)


def _sheet(workbook: Any, sheet: Any) -> Any:
    """``sheet`` when it holds cells; a chart sheet or dialog sheet is refused."""
    try:
        return workbook._cell_sheet(sheet)
    except ValueError as error:
        raise ToolError("refused", str(error), field="target") from None


def _forms(workbook: Any, sheet: Any) -> dict[str, Any]:
    """Both ways to name a sheet, as a result gives them the first time it names one."""
    out: dict[str, Any] = {"name": sheet.name}
    if workbook.sheet_refs(sheet)[1] is not None:
        out["id"] = sheet.id
    return out


def _intersect(area: Area, used: Area | None) -> Area | None:
    if used is None:
        return None
    first = Cell(max(area.first.row, used.first.row), max(area.first.column, used.first.column))
    last = Cell(min(area.last.row, used.last.row), min(area.last.column, used.last.column))
    if first.row > last.row or first.column > last.column:
        return None
    return Area(first, last)


class _Annotations:
    """The comments and hyperlinks of the sheets read, by cell (each the sheet as an address
    writes it, and the cell), read only when asked for; and the cells withheld, of columns
    a declared table hides from the model (:mod:`xlsx_agent.declare`)."""

    def __init__(self, workbook: Any, sheets: list, *, comments: bool,
                 hyperlinks: bool) -> None:
        self.comments: dict[tuple[str, Cell], Any] = {}
        self.links: dict[tuple[str, Cell], Any] = {}
        self.spans: list[tuple[str, Any]] = []
        self.withheld: dict[str, list[Area]] = {}
        for sheet in sheets:
            if sheet is None:
                continue
            written = _written(sheet)
            hidden = workbook.regions.withheld(sheet.name) if sheet.name else []
            if hidden:
                self.withheld[written] = hidden
            if comments:
                for comment in workbook.comments(sheet):
                    self.comments.setdefault((written, comment.cell), comment)
            if hyperlinks:
                for link in workbook.hyperlinks(sheet):
                    if link.area.first == link.area.last:
                        self.links.setdefault((written, link.area.first), link)
                    else:
                        self.spans.append((written, link))

    def hidden(self, place: tuple[str, int, int]) -> bool:
        """Whether the cell is withheld: a column hidden from the model."""
        areas = self.withheld.get(place[0])
        return bool(areas) and any(a.first.row <= place[1] <= a.last.row
                                   and a.first.column <= place[2] <= a.last.column
                                   for a in areas)

    def link(self, sheet: str, cell: Cell) -> Any:
        found = self.links.get((sheet, cell))
        if found is not None or not self.spans:
            return found
        for written, link in self.spans:
            area = link.area
            if written == sheet and area.first.row <= cell.row <= area.last.row \
                    and area.first.column <= cell.column <= area.last.column:
                return link
        return None


def _written(sheet: Any) -> str:
    return quote_sheet(sheet.name) if sheet.name else f"#{sheet.id}"


def _cell(value: Value, annotate: _Annotations, formulas: bool,
          cap: int | None) -> tuple[dict, list[str]]:
    """A value as the result gives it, its text cut at ``cap`` (``None``: not cut), and the
    address once if any of its text was cut."""
    place = value.place
    if annotate.withheld and annotate.hidden(place):
        return {"a": value.a, "withheld": True}, []
    out = value.as_dict()
    cut = False
    if isinstance(out["v"], str) and cap is not None and len(out["v"]) > cap:
        out["length"] = len(out["v"])
        out["v"] = out["v"][:cap]
        cut = True
    elif isinstance(out["v"], list) and cap is not None \
            and sum(len(item) for item in out["v"]) > cap:     # a list: its items cut
        out["length"] = sum(len(item) for item in out["v"])
        left, items = cap, []
        for item in out["v"]:
            items.append(item[:left])
            left = max(left - len(item), 0)
        out["v"] = items
        cut = True
    if formulas and value.formula_withheld:
        out["formula_withheld"] = True
    if formulas and value.formula_text is not None:
        text = value.formula_text
        if cap is not None and len(text) > cap:
            out["f_length"] = len(text)
            text, cut = text[:cap], True
        out["f"] = text
    position = Cell(place[1], place[2])
    comment = annotate.comments.get((place[0], position)) if annotate.comments else None
    if comment is not None:
        note: dict[str, Any] = {"author": comment.author, "text": comment.text}
        if cap is not None and len(comment.text) > cap:
            note["text"], note["length"], cut = comment.text[:cap], len(comment.text), True
        out["comment"] = note
    link = annotate.link(place[0], position) if (annotate.links or annotate.spans) else None
    if link is not None:
        target: dict[str, Any] = {}
        if link.url is not None:
            url = link.url
            if cap is not None and len(url) > cap:
                url, cut = url[:cap], True
                target["url_length"] = len(link.url)
            target["url"] = url
        if link.location is not None:
            target["location"] = link.location[:cap] if cap is not None else link.location
        out["link"] = target
    return out, [value.a] if cut else []


def _fit(value: Value, annotate: _Annotations, formulas: bool,
         room: int) -> tuple[dict, list[str]]:
    """A page's first cell, too big for the page as it is: its text cut until it fits
    ``room`` characters (at the least, to none)."""
    cap = max(room, 0)
    while True:
        cell, texts = _cell(value, annotate, formulas, cap)
        if _size(cell) + 1 <= room or cap == 0:
            return cell, texts
        cap = cap // 2 if cap > 64 else 0


def _extras(workbook: Any, places: list[tuple[Any, SheetRef, Area]], merges: bool) -> dict:
    """The merged areas (when asked) and array blocks that overlap any of ``places`` (each a
    sheet, how to name it, and an area of it)."""
    out: dict[str, Any] = {}
    sheets: dict[str, tuple[Any, SheetRef, list[Area]]] = {}
    for sheet, ref, area in places:
        sheets.setdefault(ref.text, (sheet, ref, []))[2].append(area)
    found: list[str] = []
    blocks: list[str] = []
    for sheet, ref, areas in sheets.values():
        def over(block: Area, areas=areas) -> bool:
            return any(_overlaps(block, area) for area in areas)

        if merges:
            found += [RangeAddress(ref, m).text for m in workbook.merges(sheet) if over(m)]
        blocks += [RangeAddress(ref, b).text for b in workbook.arrays(sheet) if over(b)]
    if found:
        out["merges"] = list(dict.fromkeys(found))
    if blocks:
        out["arrays"] = list(dict.fromkeys(blocks))
    return out


def _overlaps(one: Area, two: Area) -> bool:
    return not (one.last.row < two.first.row or two.last.row < one.first.row
                or one.last.column < two.first.column or two.last.column < one.first.column)


#: A ``cut`` at its longest, to keep room for it in a page.
_CUT_ROOM = {"rows": 1048576, "cells": 16384, "chars": True,
             "text": {"count": 100000, "first": ["'" + "x" * 31 + "'!XFD1048576"] * _FIRST}}


def _summary(target: str, first: int, last: int, total: int, more: bool) -> str:
    if total == 0 or last < first:
        return f"{target}: no rows"
    text = f"{target}: rows {first}-{last} of {total}"
    return text + "; more with next_cursor" if more else text


# -- xlsx_profile ------------------------------------------------------------------------------


@tool("xlsx_profile",
      "Full profile of a table's or region's columns, describe's drill-down: types and mix, "
      "fill, min/max/mean, distinct values, text lengths (chars, words, longest cell), "
      "formula share, allowed values, flags with first addresses. Paged by column.",
      {"doc": DOC,
       "target": string("A table or region (t:Products), or one column (t:Products[Price])."),
       "columns": array(string(), "Only these columns, by name. Default all.", optional=True),
       "cursor": CURSOR},
      group=CORE, kind=KIND)
@handler
def xlsx_profile(call: Any, doc: str, target: str, columns: list[str] | None = None,
                 cursor: str | None = None) -> Result:
    """``xlsx_profile``'s workbook handler (see the module's notes)."""
    workbook = call.document
    address = parse_address(target)
    if not isinstance(address, TableAddress) or address.kind not in ("table", "region",
                                                                     "column"):
        raise ToolError("invalid_arguments", "xlsx_profile takes a table or region (t:Products) "
                        "or one of its columns (t:Products[Price])", field="target")
    if address.column is not None and columns:
        raise ToolError("invalid_arguments", "give a column in the target or columns, not both",
                        field="columns")
    region = workbook.region(address.whole)
    wanted = [address.column] if address.column is not None else (columns or list(region.columns))
    names = list(dict.fromkeys(region.columns[region.column_index(n)] for n in wanted))
    text = region.address.text
    start = _profile_cursor(cursor, text, len(names))
    chars = call.limits.max_result_chars
    page = names[start:start + call.limits.max_list_items]
    stats = workbook.profile(region, page)

    data: dict[str, Any] = {"target": text, "range": region.where, "rows": region.rows,
                            "columns": []}
    budget = chars - _MARGIN - _size({**data, "cut": _PROFILE_CUT_ROOM}) \
        - len(_summary_profile(text, 1, len(names), len(names), True))
    used = 0
    counts = {"text": 0, "allowed_values": 0}
    full = False
    for item in stats:
        shaped, cut = _profile(item.as_dict())
        size = _size(shaped) + 1
        if used + size > budget:
            if data["columns"]:
                full = True
                break
            shaped, cut = _profile(item.as_dict(), compact=True)
            size = _size(shaped) + 1
        data["columns"].append(shaped)
        used += size
        for key in counts:
            counts[key] += cut.get(key, 0)
    given = len(data["columns"])
    end = start + given
    cut_out: dict[str, Any] = {}
    next_cursor = None
    if end < len(names):
        cut_out["columns"] = len(names) - end
        if full:
            cut_out["chars"] = True
        next_cursor = _cursor(text, end, 0)
    cut_out.update({key: n for key, n in counts.items() if n})
    if cut_out:
        data["cut"] = cut_out
    summary = _summary_profile(text, start + 1, end, len(names), next_cursor is not None)
    return Result(summary=summary, data=data, total=len(names), next_cursor=next_cursor)


_PROFILE_CUT_ROOM = {"columns": 16384, "chars": True, "text": 100000, "allowed_values": 100000}


def _summary_profile(target: str, first: int, last: int, total: int, more: bool) -> str:
    text = f"{target}: columns {first}-{last} of {total}" if last >= first else \
        f"{target}: no columns"
    return text + "; more with next_cursor" if more else text


def _profile(stats: dict, *, compact: bool = False) -> tuple[dict, dict[str, int]]:
    """A column's profile with its text cut at :data:`PROFILE_TEXT` characters and its
    allowed lists at :data:`ALLOWED_SHOWN` values; ``compact`` (a column too big for a page
    by itself) leaves the values and allowed values out, counted.  The counts of what was
    cut, by kind."""
    cut = {"text": 0, "allowed_values": 0}

    def shown(v: Any) -> Any:
        if isinstance(v, str) and len(v) > PROFILE_TEXT:
            cut["text"] += 1
            return v[:PROFILE_TEXT] + "…"
        return v

    for name in ("min", "max"):
        if name in stats:
            stats[name] = shown(stats[name])
    if "values" in stats:
        if compact:
            cut["text"] += len(stats["values"])
            stats["values"] = len(stats["values"])
        else:
            for item in stats["values"]:
                item["v"] = shown(item["v"])
    for rule in stats.get("validation", ()):
        values = rule.get("values")
        if values is None:
            continue
        keep = 0 if compact else ALLOWED_SHOWN
        if len(values) > keep:
            cut["allowed_values"] += len(values) - keep
            rule["more"] = len(values) - keep
        rule["values"] = [shown(v) for v in values[:keep]]
    return stats, cut


# -- cursors and sizes -------------------------------------------------------------------------


def _token(target: str) -> str:
    return f"{zlib.crc32(target.encode('utf-8')):08x}"


def _cursor(target: str, row: int, column: int) -> str:
    """A page's start, opaque to the model: a position and the target it belongs to."""
    return f"{row}.{column}.{_token(target)}"


def _parse_cursor(cursor: str, target: str, label: str | None = None) -> tuple[int, int]:
    """A cursor's position, refused unless it was given for ``target`` (``label``, when
    given, names it in the message: a query's ``target`` holds its filters' values)."""
    parts = cursor.split(".")
    if len(parts) != 3 or not parts[0].isdigit() or not parts[1].isdigit():
        raise ToolError("invalid_arguments", "cursor is not one this tool gave", field="cursor")
    if parts[2] != _token(target):
        raise ToolError("invalid_arguments", "cursor belongs to another "
                        + (f"query than this one of {label}" if label is not None
                           else f"target than {target}"), field="cursor")
    return int(parts[0]), int(parts[1])


def _read_cursor(cursor: str | None, target: str, rows: int, width: int,
                 label: str | None = None) -> tuple[int, int]:
    if not cursor:
        return 0, 0
    row, column = _parse_cursor(cursor, target, label)
    if row >= rows or column >= max(width, 1):
        raise ToolError("invalid_arguments", "cursor is past the end of the target",
                        field="cursor")
    return row, column


def _profile_cursor(cursor: str | None, target: str, count: int) -> int:
    if not cursor:
        return 0
    start, column = _parse_cursor(cursor, target)
    if column or start >= count:
        raise ToolError("invalid_arguments", "cursor is past the end of the columns",
                        field="cursor")
    return start


def _size(data: Any) -> int:
    """Characters of ``data`` as compact JSON, as a result's text writes it."""
    return len(json.dumps(data, ensure_ascii=False, separators=(",", ":"), default=str))


TOOLS = [xlsx_read, xlsx_profile]
