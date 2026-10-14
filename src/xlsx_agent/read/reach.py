"""What a formula's text refers to, as far as a column hidden from the model is concerned.

Not a formula lexer (that is E2, which also rewrites): a scan of a formula's text for the
references it holds, so that :meth:`Regions.formula_withheld
<xlsx_agent.read.regions.Regions.formula_withheld>` can tell whether a visible cell's
formula reaches a cell a declared table hides (:mod:`xlsx_agent.declare`).  It errs on the
side of reaching: what it cannot place is taken to reach everything.

What it finds (ECMA-376 Part 1, §18.17, the formula grammar, as a file stores it):

* cell and range references, ``$`` or not, whole columns (``D:D``) and rows (``2:2``),
  each on its sheet (``Log!D2``, ``'Area 1'!D2``), on several (``North:South!D2``) or on
  the formula's own; a spill reference (``C2#``) is taken to run to the sheet's end;
* names (``NoteCells``, ``Log!NoteCells``), to be looked up as defined names or tables;
* structured references (``ChecksTable[Note]``, ``ChecksTable[[#This Row],[Note]]``,
  ``ChecksTable[[Item]:[Note]]``, a whole table by its name, and ``[@Note]`` inside a
  table), with the column names they give;
* ``INDIRECT`` and ``OFFSET``, which reach cells their text does not name: taken to reach
  everything (:data:`ANYWHERE`).

String literals are skipped (their text names nothing), as are error literals; a
reference into another workbook (``[1]Log!D2``) reaches none of this one's cells.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from ooxml_edit.charts.workbook import column_index

#: The functions whose result is a reference their arguments' text does not name.
ANYWHERE = ("INDIRECT", "OFFSET")

_WORD = re.compile(r"[\w.$\\?]+")
_CELL = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$")
_COLUMN = re.compile(r"(\$?)([A-Za-z]{1,3})$")
_ROW = re.compile(r"(\$?)(\d{1,7})$")
_THREE_D = re.compile(r":([\w.\\?]+)!")
_ERROR = re.compile(r"#[A-Za-z0-9/_]+[!?]?")
_NAME_START = re.compile(r"[^\W\d]|[_\\]")


@dataclass(frozen=True)
class Bound:
    """One end of a reference: its row and column (``None``: a whole column or row runs
    there), each with whether it moves when the formula is filled (no ``$``)."""

    row: int | None
    column: int | None
    row_moves: bool = False
    column_moves: bool = False


@dataclass(frozen=True)
class Reference:
    """One reference a formula holds.  ``kind``: ``"cells"`` (``first`` and ``last``),
    ``"name"`` (``name``), ``"table"`` (``name``, ``None`` for the table the formula's cell
    lies in; ``columns`` the column names given, empty for every column), ``"anywhere"``.
    ``sheets`` is the sheet (one name), the sheets from one to another (two), or ``None``:
    the formula's own."""

    kind: str
    sheets: tuple[str, ...] | None = None
    first: Bound | None = None
    last: Bound | None = None
    name: str | None = None
    columns: tuple[str, ...] = ()


def references(text: str) -> list[Reference]:
    """The references ``text`` (a formula, with or without its ``=``) holds, in order."""
    return _Scan(text).run()


class _Scan:
    def __init__(self, text: str) -> None:
        self.text = text
        self.out: list[Reference] = []

    def run(self) -> list[Reference]:
        text, n, i = self.text, len(self.text), 0
        while i < n:
            ch = text[i]
            if ch == '"':
                i = _string_end(text, i)
            elif ch == "#":
                found = _ERROR.match(text, i)
                i = found.end() if found else i + 1
            elif ch == "'":
                name, j = _quoted(text, i)
                if j < n and text[j] == "!":
                    i = self._after_sheet(_sheets(name), j + 1)
                else:
                    i = j
            elif ch == "[":
                content, j = _bracket(text, i)
                if content.strip().isdigit():          # another workbook: [1]Log!D2
                    i = self._external(j)
                else:                                  # [@Note]: the cell's own table
                    self.out.append(_table(None, content))
                    i = j
            else:
                word = _WORD.match(text, i)
                if word is None:
                    i += 1
                    continue
                j = word.end()
                three_d = _THREE_D.match(text, j)
                if j < n and text[j] == "!":
                    i = self._after_sheet(_sheets(word.group()), j + 1)
                elif three_d is not None and not _CELL.match(word.group()):
                    i = self._after_sheet((word.group(), three_d.group(1)), three_d.end())
                else:
                    i = self._reference(None, word.group(), j)
        return self.out

    def _after_sheet(self, sheets: "tuple[str, ...] | None", i: int) -> int:
        """What follows a sheet's ``!``: a reference or a name on it (``None``: another
        workbook's sheet, noted as nothing)."""
        word = _WORD.match(self.text, i)
        if word is None:
            return i
        if sheets is None:
            kept, self.out = self.out, []
            j = self._reference(None, word.group(), word.end())
            self.out = kept
            return j
        return self._reference(sheets, word.group(), word.end())

    def _external(self, i: int) -> int:
        """Past another workbook's reference, after its ``[n]``."""
        text = self.text
        if i < len(text) and text[i] == "'":
            _, j = _quoted(text, i)
            return self._after_sheet(None, j + 1) if text[j:j + 1] == "!" else j
        word = _WORD.match(text, i)
        if word is None:
            return self._after_sheet(None, i + 1) if text[i:i + 1] == "!" else i
        j = word.end()
        three_d = _THREE_D.match(text, j)
        if three_d is not None:
            return self._after_sheet(None, three_d.end())
        if text[j:j + 1] == "!":
            return self._after_sheet(None, j + 1)
        return j

    def _reference(self, sheets: "tuple[str, ...] | None", word: str, j: int) -> int:
        text, n = self.text, len(self.text)
        if j < n and text[j] == "(":
            function = word.upper()
            for prefix in ("_XLFN.", "_XLWS."):
                function = function.removeprefix(prefix)
            if function in ANYWHERE:
                self.out.append(Reference("anywhere"))
            return j
        if j < n and text[j] == "[":
            content, k = _bracket(text, j)
            self.out.append(_table(word, content, sheets))
            return k
        first = _bound(word)
        if first is not None and j < n and text[j] == ":":
            other = _WORD.match(text, j + 1)
            last = _bound(other.group()) if other is not None else None
            if last is not None and (last.row is None) == (first.row is None) \
                    and (last.column is None) == (first.column is None):
                self.out.append(Reference("cells", sheets, first, last))
                return _spill(self, other.end())
        if first is not None and first.row is not None and first.column is not None:
            if j < n and text[j] == "#":
                self.out.append(Reference("cells", sheets, first, Bound(None, None)))
                return j + 1
            self.out.append(Reference("cells", sheets, first, first))
            return j
        if _NAME_START.match(word) and word.upper() not in ("TRUE", "FALSE"):
            self.out.append(Reference("name", sheets, name=word))
        return j


def _spill(scan: _Scan, j: int) -> int:
    return j + 1 if scan.text[j:j + 1] == "#" else j


def _bound(word: str) -> Bound | None:
    """A cell, a column alone or a row alone, as one end of a reference."""
    found = _CELL.match(word)
    if found:
        try:
            column = column_index(found.group(2).upper())
        except ValueError:
            return None
        return Bound(int(found.group(4)), column, not found.group(3), not found.group(1))
    found = _COLUMN.match(word)
    if found:
        try:
            column = column_index(found.group(2).upper())
        except ValueError:
            return None
        return Bound(None, column, False, not found.group(1))
    found = _ROW.match(word)
    if found:
        return Bound(int(found.group(2)), None, not found.group(1), False)
    return None


def _table(name: str | None, content: str,
           sheets: "tuple[str, ...] | None" = None) -> Reference:
    """A structured reference: the table's name and the column names its specifier gives
    (special items such as ``#This Row`` and ``@`` aside)."""
    items, i = ([content], len(content)) if "[" not in content else ([], 0)
    while i < len(content):
        if content[i] == "[":
            inner, i = _bracket(content, i)
            items.append(inner)
        else:
            i += 1
    columns = []
    for item in items:
        item = item.strip()
        if item.startswith("@"):
            item = item[1:].strip()
        if item and not item.startswith("#"):
            columns.append(re.sub(r"'(.)", r"\1", item).strip())
    return Reference("table", sheets, name=name, columns=tuple(columns))


def _sheets(name: str) -> "tuple[str, ...] | None":
    """A quoted or bare sheet prefix as the sheets it names (``None``: another workbook's;
    ``[0]`` is this one)."""
    if name.startswith("["):
        close = name.find("]")
        if close < 0 or name[1:close].strip() != "0":
            return None
        name = name[close + 1:]
    first, colon, last = name.partition(":")
    return (first, last) if colon else (name,)


def _string_end(text: str, i: int) -> int:
    """Past the string literal opening at ``i`` (``""`` is a quote inside it)."""
    j = i + 1
    while j < len(text):
        if text[j] == '"':
            if text[j + 1:j + 2] == '"':
                j += 2
                continue
            return j + 1
        j += 1
    return j


def _quoted(text: str, i: int) -> tuple[str, int]:
    """A quoted sheet name opening at ``i`` (``''`` is a quote inside it), and past it."""
    parts, j = [], i + 1
    while j < len(text):
        k = text.find("'", j)
        if k < 0:
            parts.append(text[j:])
            return "".join(parts), len(text)
        parts.append(text[j:k])
        if text[k + 1:k + 2] == "'":
            parts.append("'")
            j = k + 2
            continue
        return "".join(parts), k + 1
    return "".join(parts), j


def _bracket(text: str, i: int) -> tuple[str, int]:
    """What the bracket opening at ``i`` holds (to its matching ``]``; ``'`` escapes the
    character after it), and past it."""
    depth, j = 0, i
    while j < len(text):
        ch = text[j]
        if ch == "'":
            j += 2
            continue
        if ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0:
                return text[i + 1:j], j + 1
        j += 1
    return text[i + 1:], len(text)
