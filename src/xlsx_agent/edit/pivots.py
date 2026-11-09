"""Pivot caches refreshed when the workbook opens (production round 44, R61), opt-in.

::

    wb.refresh_pivots_on_load(True)                     # every pivot cache
    wb.refresh_pivots_on_load(True, pivots=["Totals"])  # the caches of these pivot tables
    wb.refresh_pivots_on_load(False)                    # the flag taken away again

``refreshOnLoad`` on a pivot cache definition (ECMA-376 Part 1, 18.10.1.67) asks the
application to refresh the cache, and the pivot tables built on it, when it opens the file.
It is never set by any other edit: a write that changes a pivot's source leaves its cache
and tables as they were saved, and this is how a host asks for them to be brought up to
date.  One undoable step.

**The parts are kept as bytes.**  A pivot cache definition (and its records, and the pivot
tables) are kept as the file has them; the call edits the one attribute in the definition's
start tag, as text -- ``refreshOnLoad="1"`` added (before the tag's end), or its value set,
or (``False``) the attribute taken away, as Excel writes a cache that does not refresh
(measured: no attribute) -- and every other byte of the part stays as it was.

**What Excel does with it, measured** (Excel for Mac 16.106, 2026-11-08, a made-up
workbook whose pivot showed 20 for source data summing to 10): Excel read the flag (its
object model's ``refresh on file open`` was true) and kept it when it saved, but did
**not** refresh the pivot on open -- not when opened through its object model, not when
opened as a person opens a file (Launch Services), not after 8 seconds, and not with the
cache's records left out (``saveData="0"``): the pivot still showed 20.  Refreshing the
table by hand showed 10.  Excel for Windows is not measured.  So the flag is what Excel
asks for, written as Excel writes it; whether the pivot is current when a person opens the
file is the application's doing -- on Excel for Mac, measured, it is not.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Sequence

from lxml import etree

from . import EditError

_R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
_START = re.compile(rb"<(?:[A-Za-z_][\w.-]*:)?pivotCacheDefinition\b")
_FLAG = re.compile(rb"""\srefreshOnLoad\s*=\s*(?:"[^"]*"|'[^']*')""")


@dataclass(frozen=True)
class PivotCacheChange:
    """What :func:`refresh_pivots_on_load` did: ``caches`` changed (each its ``cache_id``,
    ``part`` and the pivot tables built on it, ``pivots``), ``unchanged`` those already
    so, and ``notes``."""

    on: bool
    caches: tuple[dict[str, Any], ...] = ()
    unchanged: tuple[dict[str, Any], ...] = ()
    notes: tuple[str, ...] = field(default_factory=tuple)

    @property
    def changed(self) -> bool:
        return bool(self.caches)

    def as_dict(self) -> dict[str, Any]:
        out: dict[str, Any] = {"on": self.on, "changed": self.changed,
                               "caches": [dict(c) for c in self.caches]}
        if self.unchanged:
            out["unchanged"] = [dict(c) for c in self.unchanged]
        if self.notes:
            out["notes"] = list(self.notes)
        return out


def _caches(workbook: Any) -> list[dict[str, Any]]:
    """The workbook's pivot caches: ``cache_id``, ``part``, and the pivot tables (by name,
    with their sheet) built on each."""
    package = workbook.package
    root = package.tree(workbook.part)
    holder = root.find(root.tag.rpartition("}")[0] + "}pivotCaches") if root is not None \
        else None
    out: list[dict[str, Any]] = []
    by_part: dict[str, dict[str, Any]] = {}
    for node in (holder if holder is not None else []):
        if not isinstance(node.tag, str) or etree.QName(node).localname != "pivotCache":
            continue
        part = package.related_part(workbook.part, node.get(_R_ID))
        if part is None or not package.has_part(part):
            continue
        entry = {"cache_id": node.get("cacheId"), "part": part, "pivots": []}
        out.append(entry)
        by_part[part] = entry
    for sheet in workbook._every_sheet:
        if sheet.part is None:
            continue
        for rel in package.relationships(sheet.part).values():
            if rel.is_external or rel.type.rpartition("/")[2] != "pivotTable" \
                    or rel.target_part is None or not package.has_part(rel.target_part):
                continue
            table = rel.target_part
            name = None
            tree = package.tree(table)
            if tree is not None:
                name = tree.get("name")
            for inner in package.relationships(table).values():
                if inner.type.rpartition("/")[2] == "pivotCacheDefinition" \
                        and inner.target_part in by_part:
                    by_part[inner.target_part]["pivots"].append(
                        {"name": name, "sheet": sheet.stored_name})
    return out


def _flagged(data: bytes, on: bool) -> bytes | None:
    """The part's bytes with ``refreshOnLoad`` set (``on``) or taken away; ``None`` when
    its start tag cannot be found as text (the part is not UTF-8 XML as Excel writes)."""
    start = _START.search(data)
    if start is None:
        return None
    end = data.find(b">", start.end())
    if end < 0:
        return None
    close = end - 1 if data[end - 1:end] == b"/" else end
    tag = data[start.start():close]
    found = _FLAG.search(tag)
    if on:
        if found is not None:
            new_tag = tag[:found.start()] + b' refreshOnLoad="1"' + tag[found.end():]
        else:
            new_tag = tag + b' refreshOnLoad="1"'
    else:
        if found is None:
            return data
        new_tag = tag[:found.start()] + tag[found.end():]
    return data[:start.start()] + new_tag + data[close:]


def _is_on(data: bytes) -> bool:
    start = _START.search(data)
    if start is None:
        return False
    end = data.find(b">", start.end())
    found = _FLAG.search(data[start.start():end])
    if found is None:
        return False
    value = found.group(0).split(b"=", 1)[1].strip()[1:-1].strip().lower()
    return value in (b"1", b"true")


def refresh_pivots_on_load(workbook: Any, on: bool = True, *,
                           pivots: "Sequence[Any] | None" = None) -> PivotCacheChange:
    """Set (``on``) or take away ``refreshOnLoad`` on the pivot caches -- every one, or
    those of ``pivots`` (pivot table names, or cache ids) -- as a targeted edit of each
    definition's bytes (see the module)."""
    if workbook.strict:
        raise EditError("strict")
    if not isinstance(on, bool):
        raise EditError("bad_value", "on is True (refresh when the workbook opens) or False")
    caches = _caches(workbook)
    if not caches:
        raise EditError("no_pivot", "the workbook has no pivot caches")
    chosen = caches
    if pivots is not None:
        wanted = [pivots] if isinstance(pivots, (str, int)) else list(pivots)
        chosen = []
        missing = []
        for item in wanted:
            text = str(item).strip()
            hits = [c for c in caches if str(c["cache_id"]) == text
                    or any((p["name"] or "").casefold() == text.casefold()
                           for p in c["pivots"])]
            if not hits:
                missing.append(text)
            for hit in hits:
                if hit not in chosen:
                    chosen.append(hit)
        if missing or not wanted:
            names = sorted({p["name"] for c in caches for p in c["pivots"] if p["name"]})
            raise EditError("no_pivot", f"no pivot table or cache {missing[:10]}; the pivot "
                            f"tables are {names[:50]}, the cache ids "
                            f"{[c['cache_id'] for c in caches][:50]}",
                            addresses=missing[:20] or ["pivots"])
    package = workbook.package
    edits: list[tuple[dict[str, Any], bytes]] = []
    unchanged: list[dict[str, Any]] = []
    refused: list[str] = []
    for cache in chosen:
        data = package.read(cache["part"]) or b""
        if _start_tag(data) and _is_on(data) == on:
            unchanged.append(cache)
            continue
        new = _flagged(data, on)
        if new is None:
            refused.append(cache["part"])
            continue
        if new == data:
            unchanged.append(cache)
            continue
        edits.append((cache, new))
    if refused:
        raise EditError("pivot_part", "the cache definition's start tag could not be read as "
                        "text (not UTF-8 XML as Excel writes it): nothing was changed",
                        addresses=refused[:20])
    notes = []
    if edits:
        with workbook.editor.batch():
            for cache, new in edits:
                package.replace_part(cache["part"], new)
            workbook._forget()
        notes.append("Excel for Mac 16.106 kept the flag but did not refresh the pivot on "
                     "open (measured); Excel for Windows is not measured")
    else:
        notes.append("the caches are already so")
    return PivotCacheChange(on, tuple(c for c, _ in edits), tuple(unchanged), tuple(notes))


def _start_tag(data: bytes) -> bytes:
    start = _START.search(data)
    if start is None:
        return b""
    end = data.find(b">", start.end())
    return data[start.start():end + 1]


__all__ = ["PivotCacheChange", "refresh_pivots_on_load"]
