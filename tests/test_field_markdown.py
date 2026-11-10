"""A cell read as one text field of a dataset comes back as Markdown (round 48).

A workbook cell written from Markdown holds its lists as lines (``•`` bullets, numbers,
indentation) and its emphasis as rich runs.  Read into a dataset -- a derive's rows, or a
table's rows a document tool reads by ``t:`` -- each such text value carries its Markdown
(ooxml-edit's ``Dataset.markdown``), so a document writing that one field as blocks gets
list items nested in its list styles and its bold and italic, not lines of characters.
The value itself stays the cell's text: a reference inside a sentence, a table cell and a
check of numbers read it.  Bullets typed in Excel (``•``, ``◦``, ``▪``, ``-``, indented
for nesting) are list items too; a cell with no marks and no list line has no Markdown.

Round trip (a property over made-up Markdown): Markdown -> the writer's rich text ->
dataset field gives the same Markdown, for everything the writer holds.  The data is made
up.
"""

from __future__ import annotations

import random

import pytest
from ooxml_edit.markdown import Line, Run, merge, plain

from xlsx_agent import Workbook
from xlsx_agent.markdown import _commonmark, cell_runs, field_markdown

BODY = ("Doors need work:\n\n- hinges\n- **locks**\n  - front\n  - *back*\n- paint\n\n"
        "1. first\n2. second")


def worked(markdown) -> dict:
    """A dataset's markdown, every value worked out: those that are Markdown."""
    return {place: markdown[place] for place in markdown or {} if markdown[place] is not None}


def field(text: str) -> str | None:
    return field_markdown((Run(text),))


# -- the cell's lines as Markdown -------------------------------------------------------------


def test_the_writers_lists_and_marks_come_back_as_the_same_markdown():
    runs = cell_runs(BODY)
    assert plain(runs).splitlines()[2:6] == ["• hinges", "• locks", "    • front", "    • back"]
    assert field_markdown(runs) == BODY


@pytest.mark.parametrize("cell, markdown", [
    # Bullets typed in Excel, nested by indentation (any amount, tabs to four).
    ("• one\n  ◦ two\n      ▪ three\n• four", "- one\n  - two\n    - three\n- four"),
    ("- one\n\t- two\n- three", "- one\n  - two\n- three"),
    ("◦ a\n◦ b", "- a\n- b"),
    # A list after a line of text, and text after it: parted by an empty line.
    ("Steps:\n1. cut\n2. sand\nDone.", "Steps:\n\n1. cut\n2. sand\n\nDone."),
    # A further line of an item, indented past its marker: a hard break inside it.
    ("• a\n  more of a\n• b", "- a\\\n  more of a\n- b"),
    # Lines of a paragraph keep their breaks (where the cell broke them), empty lines part.
    ("• x\n\nOne\nTwo\n\nThree", "- x\n\nOne\\\nTwo\n\nThree"),
    # A number that does not follow on stays text, so the document shows the same numbers.
    ("1. a\n3. b", "1. a\n\n3\\. b"),
    # Characters that would be Markdown are escaped; | (a table) and an underline too.
    ("• a*b_c | d\n===", "- a\\*b\\_c \\| d\n\n\\==="),
    ("• 5 - 3\n# heading-like", "- 5 - 3\n\n\\# heading-like"),
])
def test_typed_lists_nest_by_indentation(cell, markdown):
    assert field(cell) == markdown


@pytest.mark.parametrize("cell", ["plain text", "two\nlines", "a*b*c", "-5 degrees",
                                  "1.5 litres", "---", "  indented"])
def test_a_plain_cell_has_no_markdown(cell):
    """Nothing bold, italic, struck or a list item: the text is the field, as before."""
    assert field(cell) is None


def test_emphasis_alone_is_markdown_and_breaks_stay_hard():
    runs = (Run("Note", bold=True), Run(": read "), Run("this", italic=True), Run("\nnext line"))
    assert field_markdown(runs) == "**Note**: read *this*\\\nnext line"
    assert field_markdown((Run("struck", strike=True),)) == "~~struck~~"


# -- the round trip (property) -------------------------------------------------------------------

#: Words a line may hold after its first: Markdown's characters among them, escaped.
WORDS = ["alpha", "beta", "gamma", "a*b", "snake_case", "x|y", "#3", "1.", "<tag>", "`t`",
         "[ref]", "~", "back\\slash", "&amp;", "50%", "=", "-", "café", "•"]
#: Words a line starts with: nothing a line of text could be read as a list marker from.
FIRST = ["Alpha", "Beta", "Gamma", "Delta", "Note", "Item", "Été"]


def _runs(rng: random.Random) -> tuple[Run, ...]:
    words = [rng.choice(FIRST)] + [rng.choice(WORDS) for _ in range(rng.randint(0, 5))]
    out: list[Run] = []
    for k, word in enumerate(words):
        marks = rng.choice([(), (), ("bold",), ("italic",), ("bold", "italic"), ("strike",)])
        out.append(Run(("" if k == 0 else " ") + word, **{m: True for m in marks}))
    return merge(out)


def _lines(rng: random.Random) -> list[Line]:
    """Made-up content the writer holds: paragraphs of one line, and lists nested a level
    at a time, bullets and numbers (numbered on from their first), between them empty
    lines -- as the writer's Markdown reads."""
    out: list[Line] = []
    depth = -1                      # the first item is a list's own
    lists: dict[int, tuple[str, int]] = {}
    for block in range(rng.randint(1, 4)):
        if block:
            out.append(Line((), "blank"))
        if rng.random() < 0.4:      # a paragraph ends the lists
            out.append(Line(_runs(rng)))
            depth, lists = -1, {}
            continue
        for _ in range(rng.randint(1, 6)):   # an empty line before does not: on they go
            depth = rng.randint(0, min(depth + 1, 3))
            for deeper in [d for d in lists if d > depth]:
                del lists[deeper]
            kind = rng.choice(["bullet", "bullet", "number"])
            if kind == "bullet":
                lists[depth] = ("bullet", 0)
                out.append(Line(_runs(rng), "bullet", depth))
                continue
            last = lists.get(depth)
            number = last[1] + 1 if last is not None and last[0] == "." else rng.choice([1, 1, 4])
            lists[depth] = (".", number)
            out.append(Line(_runs(rng), "number", depth, number, "."))
    return out


@pytest.mark.parametrize("seed", range(300))
def test_markdown_written_to_a_cell_reads_back_as_the_same_markdown(seed):
    rng = random.Random(seed)
    lines = _lines(rng)
    markdown = _commonmark(lines)
    runs = cell_runs(markdown)
    has_list = any(line.kind in ("bullet", "number") for line in lines)
    if has_list or any(r.formatted for r in runs):
        assert field_markdown(runs) == markdown, (seed, markdown)
    else:
        assert field_markdown(runs) is None


def test_the_round_trip_holds_through_a_saved_workbook_and_a_dataset():
    """The same, through the file: cells written from Markdown, the workbook saved and
    opened again, its table read as a dataset (what a document's ``t:`` reads) and derived
    (what ``xlsx_derive`` keeps as ``d:<id>``)."""
    rng = random.Random(48)
    texts = [_commonmark(_lines(rng)) for _ in range(40)] + [BODY, "Plain cell", "**Bold**"]
    rows = [[f"N-{k}", {"markdown": text}] for k, text in enumerate(texts, 1)]
    data = Workbook.from_tables({"Notes": {"columns": ["Ref", "Body"], "table": "Notes",
                                           "rows": rows}}).to_bytes()
    wb = Workbook.open(data)
    found = wb.dataset("t:Notes")
    derived = wb.derive("t:Notes", group_by=["Ref"],
                        measures=[{"fn": "first", "column": "Body", "as": "Body"}],
                        id="notes").dataset()
    for k, text in enumerate(texts):
        expected = field_markdown(cell_runs(text))
        assert found.markdown_at(k, 1) == expected == derived.markdown_at(k, 1), text
        assert found.rows[k][1] == plain(cell_runs(text))      # the value is the cell's text
    assert found.markdown_at(len(texts) - 2, 1) is None          # a plain cell has none
    assert found.markdown_at(0, 0) is None                       # nor a key


# -- what carries it ---------------------------------------------------------------------------------


def _book() -> Workbook:
    return Workbook.open(Workbook.from_tables({
        "Notes": {"columns": ["Ref", "Body", "Size"], "table": "Notes", "rows": [
            ["N-1", {"markdown": BODY}, 3], ["N-2", "• typed\n    ◦ nested", 1],
            ["N-3", "plain", 2], ["N-4", {"markdown": "*only* this"}, 5]]},
        "Owners": {"columns": ["Ref", "Note"], "table": "Owners", "rows": [
            ["N-1", {"markdown": "- one\n- two"}], ["N-3", "plain too"]]}}).to_bytes())


def test_a_derive_keeps_the_markdown_of_values_from_one_cell():
    wb = _book()
    one = wb.derive("t:Notes", group_by=["Ref"],
                    measures=[{"fn": "first", "column": "Body", "as": "Body"}], id="one")
    assert worked(one.markdown) == {(0, 1): field_markdown(cell_runs(BODY)), (1, 1): "- typed\n  - nested",
                            (3, 1): "*only* this"}
    # Several rows to one group: the value is one of several cells -- no Markdown read.
    many = wb.derive("t:Notes", measures=[{"fn": "first", "column": "Body", "as": "Body"}])
    assert many.provenance[0]["Body"].count == 4 and worked(many.markdown) == {}
    # From a dataset: carried through by position.
    again = wb.derive(one, where={"Ref": {"in": ["N-2", "N-4"]}}, group_by=["Ref"],
                      measures=[{"fn": "first", "column": "Body", "as": "Body"}])
    assert worked(again.dataset().markdown) == {(0, 1): "- typed\n  - nested", (1, 1): "*only* this"}


def test_a_join_keeps_the_markdown_of_both_sides():
    wb = _book()
    joined = wb.derive("t:Notes", join={"from": "t:Owners", "on": ["Ref", "Ref"],
                                        "columns": ["Note"]}, id="joined")
    data = joined.dataset()
    body, note = data.columns.index("Body"), data.columns.index("Note")
    assert data.markdown_at(0, body) == field_markdown(cell_runs(BODY))
    assert data.markdown_at(0, note) == "- one\n- two"
    assert data.markdown_at(2, note) is None and data.rows[2][note] == "plain too"


def test_a_workbook_without_rich_text_or_lists_is_not_read_again():
    """The extra pass over a sheet happens only when a value could have Markdown."""
    wb = Workbook.open(Workbook.from_tables({"T": {"columns": ["k", "v"], "table": "T", "rows": [
        [f"K{n}", f"plain {n}"] for n in range(50)]}}).to_bytes())
    calls = []
    original = wb._read

    def counted(*args, **kwargs):
        calls.append(kwargs.get("markdown"))
        return original(*args, **kwargs)

    wb._read = counted
    derived = wb.derive("t:T", group_by=["k"], measures=[{"fn": "first", "column": "v", "as": "v"}])
    assert derived.markdown == {} and "raw" not in calls        # its own passes only
    table = wb.dataset("t:T")
    assert table.markdown is None and "raw" not in calls


def test_markdown_is_worked_out_only_when_a_value_is_asked_for():
    """Making the dataset reads nothing more; the first value asked for reads the named
    cells' strings once (no Markdown worked out), and each value's Markdown is worked out
    on its own, once.  Pickled, it is every value's, worked out."""
    import pickle

    from xlsx_agent.read import strings as strings_module

    wb = _book()
    calls, worked_out = [], []
    original_read, original_field = wb._read, strings_module.text_field

    def counted(*args, **kwargs):
        calls.append(kwargs.get("markdown"))
        return original_read(*args, **kwargs)

    def converting(runs, redact=None):
        worked_out.append(plain(runs))
        return original_field(runs, redact)

    wb._read = counted
    strings_module.text_field = converting
    try:
        data = wb.derive("t:Notes", group_by=["Ref"],
                         measures=[{"fn": "first", "column": "Body", "as": "Body"}]).dataset()
        table = wb.dataset("t:Notes")
        assert "raw" not in calls and worked_out == []
        assert data.markdown_at(1, 1) == "- typed\n  - nested"
        assert calls.count("raw") == 1 and len(worked_out) == 1
        assert data.markdown_at(3, 1) == "*only* this" and data.markdown_at(1, 1)
        assert calls.count("raw") == 1 and len(worked_out) == 2
        assert table.row_slice(2, 2).markdown_at(0, 1) == "- typed\n  - nested"
        assert table.markdown_at(2, 1) is None                     # a plain cell: its text
        again = pickle.loads(pickle.dumps(data))
        assert isinstance(again.markdown, dict) and again.markdown == worked(data.markdown)
    finally:
        strings_module.text_field = original_field


def test_a_cell_changed_since_gives_no_markdown():
    wb = _book()
    data = wb.dataset("t:Notes")
    wb.set_cells({"Notes!B3": "• another\n• list"})
    assert data.rows[1][1] == "• typed\n    ◦ nested" and data.markdown_at(1, 1) is None
