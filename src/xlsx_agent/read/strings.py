"""Strings: the shared-string table, indexed once, and a string item's text.

The table (``sharedStrings.xml``, ECMA-376 Part 1, §18.4) is read the first time a cell
needs it, streamed, into a list: one pass, and each lookup after it an index into that
list (ROADMAP.md: the chart reader's ``findall`` per lookup is what this replaces).

A string item (``si`` in the table, ``is`` in a cell, CT_Rst, §18.4.8) is plain text
(``t``) or rich runs (``r``, each with its ``t``); its text is theirs joined, the run
properties dropped.  Phonetic runs (``rPh``, §18.4.6) are a reading guide over the base
text, not part of it, and are dropped too.  Text is an ST_Xstring (§22.9.2.19), whose
``_xHHHH_`` escapes stand for the character with that code (a character XML cannot carry,
or ``_x005F_`` for an underscore); they are decoded.
"""

from __future__ import annotations

import re

from ooxml_edit.xml import Element

from ..xml import iterparse_xml

NS_MAIN = ("http://schemas.openxmlformats.org/spreadsheetml/2006/main",
           "http://purl.oclc.org/ooxml/spreadsheetml/main")

_T = frozenset(f"{{{ns}}}t" for ns in NS_MAIN)
_R = frozenset(f"{{{ns}}}r" for ns in NS_MAIN)
_SI = tuple(f"{{{ns}}}si" for ns in NS_MAIN)

_ESCAPE = re.compile(r"_x([0-9A-Fa-f]{4})_")


def decode_xstring(text: str) -> str:
    """An ST_Xstring's text with its ``_xHHHH_`` escapes decoded."""
    if "_x" not in text:
        return text
    return _ESCAPE.sub(lambda match: chr(int(match.group(1), 16)), text)


def item_text(node: Element) -> str:
    """The text of a string item (``si``, ``is``): its ``t``, or its runs' ``t`` joined;
    phonetic runs left out."""
    parts = []
    for child in node:
        tag = child.tag
        if tag in _T:
            parts.append(child.text or "")
        elif tag in _R:
            for grand in child:
                if grand.tag in _T:
                    parts.append(grand.text or "")
    return decode_xstring("".join(parts))


class SharedStrings:
    """The shared-string table of a workbook, read on first use.

    ``data`` is the part's bytes (``None`` when the workbook has no table).  ``strings[i]``
    is the ``i``-th item's text; an index past the end is an ``IndexError``.  The part is
    streamed: each item is read, its text kept, and the item dropped."""

    def __init__(self, data: bytes | None) -> None:
        self._data = data
        self._items: list[str] | None = None

    @property
    def items(self) -> list[str]:
        if self._items is None:
            self._items = self._read()
        return self._items

    def __getitem__(self, index: int) -> str:
        if index < 0:
            raise IndexError(index)
        return self.items[index]

    def __len__(self) -> int:
        return len(self.items)

    def _read(self) -> list[str]:
        if not self._data:
            return []
        items: list[str] = []
        for _, node in iterparse_xml(self._data, tag=_SI):
            items.append(item_text(node))
            node.clear()
            parent = node.getparent()
            if parent is not None:
                while node.getprevious() is not None:
                    del parent[0]
        return items

