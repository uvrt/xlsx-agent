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
"""

from __future__ import annotations

import re
from dataclasses import replace

from ooxml_edit.markdown import Line, Run, merge, parse, plain, to_markdown

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


__all__ = ["BULLET", "INDENT", "cell_markdown", "cell_runs"]
