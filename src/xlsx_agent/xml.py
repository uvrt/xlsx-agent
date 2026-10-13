"""Streaming XML with the safe options of ooxml-edit's ``parse_xml``.

A whole part is parsed through :func:`ooxml_edit.xml.parse_xml`; a worksheet, which may hold
a million rows, is streamed instead (ROADMAP.md, "Design rules", 6).  The stream takes the
same options: entities never resolved, ``huge_tree`` off, nothing fetched from the network.
Until ooxml-edit has its own ``iterparse_xml`` (ROADMAP.md, O6) they are kept here, and
``tests/test_xml.py`` checks that they match ``parse_xml``'s.
"""

from __future__ import annotations

import io
from typing import BinaryIO, Iterable, Iterator

from lxml import etree

#: ``parse_xml``'s parser options, as ``iterparse`` takes them.
SAFE_OPTIONS = {"remove_blank_text": False, "resolve_entities": False, "huge_tree": False,
                "no_network": True}


def iterparse_xml(data: "bytes | BinaryIO", events: Iterable[str] = ("end",),
                  tag: "str | Iterable[str] | None" = None) -> Iterator[tuple[str, etree._Element]]:
    """``lxml.etree.iterparse`` over a part's bytes (or a binary file of them, read as it
    goes) with :data:`SAFE_OPTIONS`."""
    source = io.BytesIO(data) if isinstance(data, (bytes, bytearray)) else data
    return etree.iterparse(source, events=tuple(events), tag=tag, **SAFE_OPTIONS)
