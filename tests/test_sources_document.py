"""A workbook's rich-text cells in a Word document, by reference (round 48), across
formats: one session holds a made-up workbook and a document; ``xlsx_derive`` keeps the
workbook's notes as ``d:notes``, and docx-agent's ``word_insert_markdown`` writes one note
by reference -- ``from: d:notes/#1/Body``, or ``{{d:notes/N-2/Body}}`` standing alone in its
Markdown -- as list items in the document's list styles, nested, with the cell's bold and
italic; a table's note named by ``t:`` alike.  A note inside a sentence is its text on one
line.  The workbook is only read.

docx-agent is a test-only dependency (the ``test`` extra; CI installs it from its
repository, as it does pptx-agent).
"""

from __future__ import annotations

import io
import json
import zipfile

import pytest
from ooxml_edit.tools import Toolbox

docx_agent = pytest.importorskip("docx_agent")
docx_tools = pytest.importorskip("docx_agent.tools")

from xlsx_agent import Workbook  # noqa: E402
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

NOTE = ("Hinges and locks:\n\n- hinges, *oiled*\n- **locks**\n  - front door\n  - back door\n"
        "    1. latch\n    2. bolt\n- paint")


def book() -> bytes:
    return Workbook.from_tables({"Notes": {"columns": ["Ref", "Body"], "table": "Notes", "rows": [
        ["N-1", {"markdown": NOTE}],
        ["N-2", "• typed in Excel\n    ◦ nested by spaces\n• last"],
        ["N-3", "plain note"]]}}).to_bytes()


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox([*docx_tools.TOOLS, *TOOLS], formats=[docx_tools.FORMAT, FORMAT],
                 groups=[*docx_tools.GROUPS, *GROUPS], workers=0) as box:
        yield box


@pytest.fixture
def session(toolbox):
    session = toolbox.session()
    session.open(docx_agent.Document.new().to_bytes(), "report.docx")      # d1
    session.open(book(), "notes.xlsx")                                     # d2
    ok(toolbox.dispatch(session, "xlsx_derive", {
        "doc": "d2", "from": "t:Notes", "id": "notes", "group_by": ["Ref"],
        "measures": [{"fn": "first", "column": "Body", "as": "Body"}]}))
    return session


def ok(result):
    assert result.ok, json.dumps(result.to_json(), indent=1, default=str)
    return result


def insert(toolbox, session, **arguments):
    result = ok(toolbox.dispatch(session, "word_insert_markdown",
                                 {"doc": "d1", "at": "end", **arguments}))
    document = session.entry("d1").document
    return result, [(document.paragraph(b).style, document.paragraph(b).text)
                    for b in result.data["blocks"]]


def test_a_rich_text_cell_goes_in_as_nested_list_items(toolbox, session):
    workbook = session.entry("d2")
    version, before = workbook.version, workbook.document.to_bytes()
    result, made = insert(toolbox, session, **{"from": "d:notes/#1/Body"})
    assert made == [(None, "Hinges and locks:"), ("ListBullet", "hinges, oiled"),
                    ("ListBullet", "locks"), ("ListBullet2", "front door"),
                    ("ListBullet2", "back door"), ("ListNumber3", "latch"),
                    ("ListNumber3", "bolt"), ("ListBullet", "paint")]
    xml = zipfile.ZipFile(io.BytesIO(session.entry("d1").document.to_bytes())).read(
        "word/document.xml").decode()
    assert 'w:val="Strong"' in xml and 'w:val="Emphasis"' in xml
    assert result.data["facts"][0]["sources"] == ["Notes!B2"]
    # The same as the note's Markdown typed in the call.
    other = toolbox.session()
    other.open(docx_agent.Document.new().to_bytes(), "typed.docx")
    typed = ok(toolbox.dispatch(other, "word_insert_markdown",
                                {"doc": "d1", "at": "end", "markdown": NOTE}))
    document = other.entry("d1").document
    assert [(document.paragraph(b).style, document.paragraph(b).text)
            for b in typed.data["blocks"]] == made
    assert workbook.version == version and workbook.document.to_bytes() == before


def test_typed_bullets_a_reference_alone_and_one_in_a_sentence(toolbox, session):
    _, made = insert(toolbox, session, markdown="## Typed\n\n{{d:notes/N-2/Body}}\n\n"
                                                "In short: {{d:notes/N-2/Body}}.\n\n{{d:notes/N-3/Body}}")
    assert made == [("Heading2", "Typed"), ("ListBullet", "typed in Excel"),
                    ("ListBullet2", "nested by spaces"), ("ListBullet", "last"),
                    (None, "In short: • typed in Excel ◦ nested by spaces • last."),
                    (None, "plain note")]


def test_a_tables_cell_named_by_t_alike(toolbox, session):
    _, made = insert(toolbox, session, **{"from": "t:Notes/#2/Body"})
    assert made == [("ListBullet", "typed in Excel"), ("ListBullet2", "nested by spaces"),
                    ("ListBullet", "last")]
