"""Comments: the notes a worksheet's comments part holds, by cell.

A worksheet relates to at most one comments part (ECMA-376 Part 1, §18.7: ``comments``,
§18.7.3), which lists its ``authors`` (§18.7.2) and its ``commentList``: each ``comment``
(§18.7.3) with the cell it is on (``ref``), its author (``authorId``, an index into the
authors) and its ``text``, a string item (CT_Rst) whose text is read as a shared string's
is (:func:`~xlsx_agent.read.strings.item_text`: rich runs joined, phonetic runs dropped).
The VML drawing that shows a comment is presentation only and is not read.

Excel's threaded comments (Microsoft's [MS-XLSX] extension, a ``threadedComments`` part)
are written beside a comment in this part too, which carries the thread as text; that is
what is read here.

A comment's author and text are content: returned by a read when asked for, never in a
message (ROADMAP.md, "Design rules", 7), and kept off the repr.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from ooxml_edit.sheets import Cell
from ooxml_edit.xml import Element

from .stream import _area
from .strings import item_text

NS_MAIN = ("http://schemas.openxmlformats.org/spreadsheetml/2006/main",
           "http://purl.oclc.org/ooxml/spreadsheetml/main")


def _tags(local: str) -> frozenset[str]:
    return frozenset(f"{{{ns}}}{local}" for ns in NS_MAIN)


_AUTHORS, _AUTHOR = _tags("authors"), _tags("author")
_LIST, _COMMENT, _TEXT = _tags("commentList"), _tags("comment"), _tags("text")


@dataclass(frozen=True)
class Comment:
    """One comment: ``cell`` the cell it is on, ``author`` its author's name (``None`` when
    the part names none), ``text`` its text, rich runs flattened."""

    cell: Cell
    author: str | None = field(default=None, repr=False)
    text: str = field(default="", repr=False)


def read_comments(root: Element | None) -> list[Comment]:
    """The comments of a comments part's root, in the file's order; a comment whose
    ``ref`` is not one cell (or a range's first cell) is left out."""
    if root is None:
        return []
    authors: list[str] = []
    out: list[Comment] = []
    for child in root:
        if child.tag in _AUTHORS:
            authors = [node.text or "" for node in child if node.tag in _AUTHOR]
        elif child.tag in _LIST:
            for node in child:
                if node.tag not in _COMMENT:
                    continue
                area = _area(node.get("ref"))
                if area is None:
                    continue
                author = _index(node.get("authorId"))
                text = next((item_text(t) for t in node if t.tag in _TEXT), "")
                out.append(Comment(area.first, authors[author]
                                   if author is not None and author < len(authors) else None,
                                   text))
    return out


def _index(value: str | None) -> int | None:
    try:
        number = int(value) if value is not None else None
    except ValueError:
        return None
    return number if number is not None and number >= 0 else None
