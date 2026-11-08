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

On request (:meth:`SharedStrings.markdown`, a read with ``markdown``) an item is given as
Markdown instead (:func:`xlsx_agent.markdown.cell_markdown`): its runs' bold, italic and
strikethrough (``rPr`` ``b``, ``i``, ``strike``: :func:`ooxml_edit.sheets.rich_runs`) as
``**``, ``*`` and ``~~``, every other character escaped, its lines' bullets and numbers as
list items.  The table is read a second time for it, only then.  With a redaction, each
run is redacted; when that would not give the item's text redacted as a whole (a term
across runs), the item is given as its redacted text, without its marks.
"""

from __future__ import annotations

import re

from ooxml_edit.markdown import Run, merge, plain
from ooxml_edit.sheets import rich_runs
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


#: What XML 1.0 cannot carry (its ``Char`` production: everything but tab, line feed and
#: carriage return below U+0020, the surrogates, U+FFFE and U+FFFF), and the carriage
#: return, which a parser reads as a line feed: each is written ``_xHHHH_``.
_UNCARRIED = re.compile("[\x00-\x08\x0b\x0c\x0d\x0e-\x1f\ud800-\udfff\ufffe\uffff]")
#: An underscore that would be read as the start of an escape (``_x0041_`` as text).
_LOOKS_ESCAPED = re.compile(r"_(?=x[0-9A-Fa-f]{4}_)")


def needs_xstring(text: str) -> bool:
    """Whether ``text`` is written otherwise than as it is (:func:`encode_xstring`)."""
    return bool(_UNCARRIED.search(text)) or ("_x" in text and bool(_LOOKS_ESCAPED.search(text)))


def encode_xstring(text: str) -> str:
    """``text`` as an ST_Xstring (ECMA-376 Part 1, §22.9.2.19) holds it, as Excel writes it
    (measured, Excel for Mac 16.106, 2026-11-08: text typed into cells through its object
    model and saved): each character XML 1.0 cannot carry, and the carriage return, as
    ``_xHHHH_`` with upper-case hex (``a_x0001_b``, ``e_x000D_f``); the underscore of text
    that would read as an escape as ``_x005F_`` (``_x0041_`` is written ``_x005F_x0041_``;
    ``_xabcd_`` too, ``_X0041_`` and ``_x00`` are not).  Tab and line feed are written as
    they are.  (In a sheet's name Excel writes the underscore's escape in lower case,
    ``_x005f_``, measured the same day; a name that needs it is refused here.)
    :func:`decode_xstring` gives ``text`` back."""
    if not needs_xstring(text):
        return text
    text = _LOOKS_ESCAPED.sub("_x005F_", text)
    return _UNCARRIED.sub(lambda match: f"_x{ord(match.group(0)):04X}_", text)


def item_runs(node: Element) -> tuple[Run, ...]:
    """A string item's runs, decoded (:func:`decode_xstring`): its rich runs with their
    marks, or its text as one unmarked run."""
    runs = rich_runs(node)
    if runs is None:
        return merge([Run(item_text(node))])
    return merge(Run(decode_xstring(run.text), run.bold, run.italic, run.strike)
                 for run in runs)


def text_markdown(runs: "tuple[Run, ...]", redact=None) -> str:
    """Runs as Markdown (see the module), redacted by ``redact`` when given."""
    from ..markdown import cell_markdown

    if redact is not None:
        whole = redact(plain(runs))
        shown = merge(Run(redact(run.text) if run.text else run.text, run.bold, run.italic,
                          run.strike) for run in runs)
        runs = shown if plain(shown) == whole else merge([Run(whole)])
    return cell_markdown(runs)


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
    streamed: each item is read, its text kept, and the item dropped.  ``redact`` (a
    function of a text) is applied to each item as it is read: the workbook's redaction
    (:mod:`xlsx_agent.redact`), so that no reader sees the text as stored."""

    def __init__(self, data: bytes | None, redact=None) -> None:
        self._data = data
        self._redact = redact
        self._items: list[str] | None = None
        #: Each rich item's runs (``None`` for a plain one), read on the first Markdown
        #: asked for, and the Markdown given so far.
        self._rich: list[tuple[Run, ...] | None] | None = None
        self._markdown: dict[int, str] = {}

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

    def markdown(self, index: int) -> str:
        """The ``index``-th item as Markdown (see the module); ``IndexError`` past the end."""
        found = self._markdown.get(index)
        if found is not None:
            return found
        text = self[index]
        if self._rich is None:
            self._rich = self._read_rich()
        runs = self._rich[index] if index < len(self._rich) else None
        if runs is None:
            found = text_markdown(merge([Run(text)])) if text else ""
        else:
            found = text_markdown(runs, self._redact)
        self._markdown[index] = found
        return found

    def _read_rich(self) -> "list[tuple[Run, ...] | None]":
        out: list[tuple[Run, ...] | None] = []
        if not self._data:
            return out
        for _, node in iterparse_xml(self._data, tag=_SI):
            out.append(item_runs(node) if any(child.tag in _R for child in node) else None)
            node.clear()
            parent = node.getparent()
            if parent is not None:
                while node.getprevious() is not None:
                    del parent[0]
        return out

    def _read(self) -> list[str]:
        if not self._data:
            return []
        items: list[str] = []
        redact = self._redact
        for _, node in iterparse_xml(self._data, tag=_SI):
            text = item_text(node)
            items.append(redact(text) if redact is not None and text else text)
            node.clear()
            parent = node.getparent()
            if parent is not None:
                while node.getprevious() is not None:
                    del parent[0]
        return items

