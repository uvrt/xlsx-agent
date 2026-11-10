"""Markdown in a cell: what ``{"markdown": ...}`` writes, and a cell's text read back as
Markdown.

The Markdown is ooxml-edit's subset (:mod:`ooxml_edit.markdown`: ``**bold**``,
``*italic*``, ``~~strike~~``, escapes and code spans in a line; links, HTML and every
other block are text), read line by line -- each line given is a line of the cell (a line
break in its text), deliberately not CommonMark's paragraphs, whose lines are joined; one
empty line is kept between lines, none at either end.  A cell has no list, so a list item
is written as a line of text that shows one, its marker a plain run:

* a bullet (``-``, ``*``, ``+``) as ``•`` and a space (U+2022, :data:`BULLET`);
* a numbered item (``1.``, ``1)``) as its number, renumbered from the list's first as
  CommonMark numbers it, its delimiter and a space;
* a nested item :data:`INDENT` spaces further in for each level (CommonMark's content
  columns decide the level), and a further line of an item (one indented to its content)
  as far in as the item's text.

Read back (:func:`cell_markdown`), a cell's lines are turned back into those items -- a
line of ``•`` and a space is a bullet (``-``), a number with ``.`` or ``)`` and a space a
numbered item, the indentation its level -- so that writing the Markdown read gives the
cell back.  What Markdown cannot hold so (numbers that do not follow on, a level skipped,
spaces at a line's start) is written as text lines, which read back as the same
characters; spaces at a line's ends and a second empty line in a row are not kept.  Every
text is escaped (:func:`ooxml_edit.markdown.escape`): ``**`` in a cell is ``\\*\\*``.

**A cell as a document's text** (:func:`field_markdown`, round 48).  A cell read as one
text field of a dataset and written into a document as blocks is given as CommonMark, as
the document libraries read Markdown (paragraphs, not lines): its bold, italic and
strikethrough runs as ``**``, ``*`` and ``~~``; each line that is a list item -- the
writer's (``•``, a number and ``.`` or ``)``) or a bullet typed in Excel (``◦``, ``▪``,
``-``; :data:`FIELD_BULLETS`), then a space -- an item (``-``, or its number), nested by its
indentation (a line further in than the item above is its child; tabs count to the next
multiple of four); a further line indented past an item's marker the same item, after a
line break; other lines a paragraph, its lines kept by hard breaks (a backslash at the
line's end) only where the cell broke them; an empty line a new paragraph; a list and a
paragraph next to each other parted by an empty line.  A numbered item that does not
follow on from the one before it (``1.`` then ``3.``) is text, so the document shows the
numbers the cell does.  Every other character is escaped, ``|`` too (a table's).  A cell
with no bold, italic or struck run and no list line is ``None``: its text is the field, as
before.
"""

from __future__ import annotations

import re
from dataclasses import replace

from ooxml_edit.markdown import Line, Run, inline_markdown, merge, parse, plain, to_markdown

#: How a bullet item shows in a cell: U+2022 BULLET and a space.
BULLET = "• "

#: The spaces a nested item's line is indented by, per level.
INDENT = 4

_SHOWN_ITEM = re.compile(r"((?: {%d})*)(• |([0-9]{1,9})([.)]) )" % INDENT)


def cell_runs(markdown: str) -> tuple[Run, ...]:
    """The runs a cell is written with for ``markdown`` (see the module): its lines joined
    by line breaks, list markers as plain runs, code spans as text."""
    out: list[Run] = []
    widths: dict[int, int] = {}         # per depth, how far the last item's text is in
    for k, line in enumerate(parse(markdown)):
        if k:
            out.append(Run("\n"))
        runs = [replace(run, code=False) for run in line.runs]
        if line.kind in ("bullet", "number"):
            marker = BULLET if line.kind == "bullet" else f"{line.number}{line.delimiter} "
            lead = " " * (INDENT * line.depth) + marker
            widths[line.depth] = len(lead)
            for deeper in [d for d in widths if d > line.depth]:
                del widths[deeper]
            out.append(Run(lead))
        elif line.kind == "more":
            out.append(Run(" " * widths.get(line.depth, INDENT * line.depth)))
        out.extend(runs)
    return merge(out)


def cell_markdown(runs: "tuple[Run, ...] | list[Run]") -> str:
    """A cell's text, given as its runs (plain text is one unmarked run), as Markdown that
    :func:`cell_runs` writes back as the same cell (see the module)."""
    cell_lines = _split(merge(runs))
    lines: list[Line] = []
    items: list[tuple[int, int]] = []       # open items: (depth, where their text starts)
    for line_runs in cell_lines:
        text = plain(line_runs)
        if not text.strip():
            lines.append(Line((), "blank"))
            continue
        match = _SHOWN_ITEM.match(text)
        if match is not None and not any(r.formatted for r in _head(line_runs, match.end())):
            depth = len(match.group(1)) // INDENT
            body = _tail(line_runs, match.end())
            del items[depth:]
            items.append((depth, match.end()))
            if match.group(3) is None:
                lines.append(Line(body, "bullet", depth))
            else:
                lines.append(Line(body, "number", depth, int(match.group(3)),
                                  match.group(4)))
            continue
        lead = len(text) - len(text.lstrip(" "))
        owner = next((depth for depth, column in reversed(items) if column == lead), None)
        if lead and owner is not None:
            while items and items[-1][0] > owner:
                items.pop()
            lines.append(Line(_tail(line_runs, lead), "more", owner))
            continue
        items.clear()
        lines.append(Line(tuple(line_runs)))
    return to_markdown(lines, fallback=[Line(tuple(r)) if plain(r).strip() else
                                        Line((), "blank") for r in cell_lines])


# -- a cell as a document's text (round 48) ------------------------------------------------


#: The bullets a line of a cell starts with to be a list item when the cell is written into
#: a document (:func:`field_markdown`): the writer's U+2022 BULLET, and U+25E6 WHITE
#: BULLET, U+25AA BLACK SMALL SQUARE and ``-`` as typed in Excel.
FIELD_BULLETS = "•◦▪-"

_FIELD_ITEM = re.compile(r"([ \t]*)(?:([•◦▪\-])|([0-9]{1,9})([.)]))[ \t]+(?=[^ \t])")
#: A line CommonMark would read as a heading's underline or a thematic break.
_UNDERLINE = re.compile(r"(?:=+|-+)[ \t]*")
_ITEM_KINDS = ("bullet", "number")


def field_markdown(runs: "tuple[Run, ...] | list[Run]") -> str | None:
    """A cell's text, given as its runs (plain text is one unmarked run), as the CommonMark a
    document's text is written from (see the module), or ``None`` when nothing in it is
    bold, italic, struck through or a list item: its text is then the field as it is."""
    merged = merge(runs)
    lines, items = _field_lines(_split(merged))
    if not items and not any(run.formatted for run in merged):
        return None
    return _commonmark(lines)


def _field_lines(cell_lines: list[tuple[Run, ...]]) -> tuple[list[Line], bool]:
    """The cell's lines as :class:`Line` s -- items, their further lines, text and blank
    lines (see the module) -- and whether any is an item."""
    out: list[Line] = []
    items: list[int] = []                      # open items, by depth: where their marker is
    lists: dict[int, tuple[str, int]] = {}     # by depth: the last item's kind and number
    found = False
    for line_runs in cell_lines:
        text = plain(line_runs)
        if not text.strip():
            out.append(Line((), "blank"))
            continue
        match = _FIELD_ITEM.match(text)
        if match is not None:
            indent = _width(match.group(1))
            while items and items[-1] > indent:
                items.pop()
            if items and items[-1] == indent:
                items.pop()
            depth = len(items)
            for deeper in [d for d in lists if d > depth]:
                del lists[deeper]
            last = lists.get(depth)
            number = int(match.group(3)) if match.group(3) is not None else None
            delimiter = match.group(4)
            if number is None or last is None or last[0] != delimiter or number == last[1] + 1:
                items.append(indent)
                lists[depth] = ("bullet", 0) if number is None else (delimiter, number)
                body = _tail(line_runs, match.end())
                if number is None:
                    out.append(Line(body, "bullet", depth))
                else:
                    out.append(Line(body, "number", depth, number, delimiter))
                found = True
                continue
        lead = len(text) - len(text.lstrip(" \t"))
        column = _width(text[:lead])
        while items and items[-1] >= column:
            items.pop()
        if lead and items:
            for deeper in [d for d in lists if d >= len(items)]:
                del lists[deeper]
            out.append(Line(_tail(line_runs, lead), "more", len(items) - 1))
            continue
        items.clear()
        lists.clear()
        out.append(Line(_tail(line_runs, lead)))
    return out, found


def _width(spaces: str) -> int:
    """The columns ``spaces`` (spaces and tabs) take, tabs to the next multiple of four."""
    column = 0
    for char in spaces:
        column += 4 - column % 4 if char == "\t" else 1
    return column


def _commonmark(lines: list[Line]) -> str:
    """``lines`` as CommonMark (see the module): items nested under their parent's content,
    a further line after a hard break, text lines of one paragraph joined by hard breaks,
    an empty line between a list and a paragraph."""
    out: list[str] = []
    columns: list[int] = []         # per depth, where the open item's content starts
    previous = "blank"
    for line in lines:
        kind = line.kind
        if kind == "more" and not columns:
            kind = "text"
        if kind == "blank":
            if previous != "blank" and out:
                out.append("")
            previous = "blank"
            continue
        if kind in _ITEM_KINDS:
            if previous == "text":
                out.append("")
            depth = min(line.depth, len(columns))
            del columns[depth:]
            indent = columns[-1] if columns else 0
            marker = "-" if kind == "bullet" else f"{line.number}{line.delimiter}"
            columns.append(indent + len(marker) + 1)
            out.append(" " * indent + marker + " " + _field_line(line.runs))
        elif kind == "more":
            depth = min(line.depth, len(columns) - 1)
            del columns[depth + 1:]
            if previous != "blank":
                out[-1] += "\\"
            out.append(" " * columns[depth] + _field_line(line.runs))
        else:
            if previous in ("more", *_ITEM_KINDS):
                out.append("")
            elif previous == "text":
                out[-1] += "\\"
            columns.clear()
            out.append(_field_line(line.runs))
        previous = kind
    while out and not out[-1]:
        out.pop()
    return "\n".join(out)


def _field_line(runs: tuple[Run, ...]) -> str:
    """One line's runs as CommonMark inline text: spaces at its ends dropped, a marker at
    its start escaped, ``|`` escaped, a line that would underline a heading escaped."""
    runs = merge(runs)
    if runs:
        runs = merge([replace(runs[0], text=runs[0].text.lstrip(" \t")), *runs[1:]])
    if runs:
        runs = merge([*runs[:-1], replace(runs[-1], text=runs[-1].text.rstrip(" \t"))])
    text = inline_markdown(runs).replace("|", "\\|")
    return "\\" + text if _UNDERLINE.fullmatch(text) else text


def _split(runs: tuple[Run, ...]) -> list[tuple[Run, ...]]:
    """The runs of each line of the text (split at its line breaks)."""
    lines: list[list[Run]] = [[]]
    for run in runs:
        parts = run.text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        for k, part in enumerate(parts):
            if k:
                lines.append([])
            if part:
                lines[-1].append(replace(run, text=part))
    return [merge(line) for line in lines]


def _head(runs: tuple[Run, ...], count: int) -> tuple[Run, ...]:
    """The runs of the first ``count`` characters."""
    out, left = [], count
    for run in runs:
        if left <= 0:
            break
        out.append(replace(run, text=run.text[:left]))
        left -= len(run.text)
    return tuple(out)


def _tail(runs: tuple[Run, ...], count: int) -> tuple[Run, ...]:
    """The runs after the first ``count`` characters."""
    out, skip = [], count
    for run in runs:
        if skip >= len(run.text):
            skip -= len(run.text)
            continue
        out.append(replace(run, text=run.text[skip:]))
        skip = 0
    return merge(out)


__all__ = ["BULLET", "FIELD_BULLETS", "INDENT", "cell_markdown", "cell_runs", "field_markdown"]
