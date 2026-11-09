"""The workbook's document properties (production R62): ``wb.properties()`` and
``wb.set_properties(...)``, through ooxml-edit's shared reader and writer
(:mod:`xlsx_agent.edit.properties`).

Tool design rules: every problem is listed with the names that would work (2), properties
already so are ``changed`` false (3), core and custom ones are one call and one step (5).
"""

from __future__ import annotations

import io
import zipfile
from datetime import date, datetime, timezone

import pytest

from xlsx_agent import Workbook
from xlsx_agent.edit import EditError

import notes_book


def part(data: bytes, name: str) -> bytes | None:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return archive.read(name) if name in archive.namelist() else None


def test_a_new_workbook_s_properties():
    wb = Workbook.new(["Stock"], author="Ann", title="Stock list")
    props = wb.properties()
    assert props.core["creator"] == props.core["last_modified_by"] == "Ann"
    assert props.core["title"] == "Stock list"
    assert isinstance(props.core["created"], datetime) and props.custom == {}
    assert props.as_dict()["core"]["created"].endswith("Z")


def test_core_and_custom_set_in_one_step_and_undone():
    wb = Workbook.open(notes_book.plain())
    before = wb.to_bytes()
    change = wb.set_properties(title="Stock list", subject="Counts", author="Bea",
                               keywords="stock; count", comments="Made up", category="Lists",
                               custom={"Project": "Alpha", "Run": 42, "Ratio": 2.5,
                                       "Checked": True, "Due": date(2026, 3, 14)})
    assert change.changed and change.core == ("title", "subject", "creator", "keywords",
                                              "description", "category")
    assert change.custom == ("Project", "Run", "Ratio", "Checked", "Due")
    props = wb.properties()
    assert props.core["creator"] == "Bea" and props.core["description"] == "Made up"
    assert {name: (p.value, p.type) for name, p in props.custom.items()} == {
        "Project": ("Alpha", "text"), "Run": (42, "integer"), "Ratio": (2.5, "number"),
        "Checked": (True, "bool"), "Due": (datetime(2026, 3, 14, tzinfo=timezone.utc), "date")}
    core = part(wb.to_bytes(), "docProps/core.xml").decode()
    # Excel's order (measured): title, subject, creator, keywords, description, ...
    assert core.index("<dc:title>") < core.index("<dc:subject>") < core.index(
        "<dc:creator>") < core.index("<cp:keywords>") < core.index("<dc:description>") \
        < core.index("<dcterms:created") < core.index("<cp:category>")
    assert wb.undo() and wb.to_bytes() == before


def test_properties_already_so_change_nothing():
    wb = Workbook.open(notes_book.plain())
    wb.set_properties(title="Stock list", custom={"Run": 42})
    version = wb.version
    change = wb.set_properties(title="Stock list", custom={"run": 42, "Gone": None})
    assert not change.changed and change.note == "the properties are already so"
    assert wb.version == version


def test_removing_properties():
    wb = Workbook.open(notes_book.plain())
    wb.set_properties(title="Stock list", custom={"Run": 42})
    change = wb.set_properties(title=None, custom={"Run": None})
    assert change.core == ("title",) and change.custom == ("Run",)
    assert "title" not in wb.properties().core and wb.properties().custom == {}
    assert part(wb.to_bytes(), "docProps/custom.xml") is None


def test_the_modified_date_stays_unless_set():
    wb = Workbook.open(notes_book.plain())
    core = part(wb.to_bytes(), "docProps/core.xml")
    wb.set_cells({"Stock!C1": "edited"})
    wb.set_comment("Stock!C1", "and noted")
    out = io.BytesIO()
    wb.save(out)
    assert part(out.getvalue(), "docProps/core.xml") == core     # a save writes it as it is
    wb.set_properties(modified=datetime(2026, 11, 9, 8, 30, tzinfo=timezone.utc),
                      last_modified_by="Cas")
    props = wb.properties().core
    assert props["modified"] == datetime(2026, 11, 9, 8, 30, tzinfo=timezone.utc)
    assert props["last_modified_by"] == "Cas"


def test_every_problem_listed_and_nothing_changed():
    wb = Workbook.open(notes_book.plain())
    before = wb.to_bytes()
    with pytest.raises(EditError) as raised:
        wb.set_properties(titel="x", created="yesterday", subject=5,
                          custom={"Odd": object(), "": 1})
    assert raised.value.code == "bad_property"
    assert raised.value.addresses == ("titel", "created", "subject", "custom:Odd", "custom:")
    assert "one of title, subject, creator" in str(raised.value)   # the way out
    with pytest.raises(EditError):
        wb.set_properties(custom=["Run"])
    assert wb.to_bytes() == before


def test_redacted_properties_read_redacted():
    wb = Workbook.open(notes_book.plain())
    wb.set_properties(title="Stock for ACME", custom={"Client": "ACME"})
    wb.redact([{"redact": ["ACME"], "with": "[client]"}])
    props = wb.properties()
    assert props.core["title"] == "Stock for [client]"
    assert props.custom["Client"].value == "[client]"


def test_in_a_batch_with_cells():
    wb = Workbook.open(notes_book.plain())
    before = wb.to_bytes()
    with pytest.raises(EditError):
        with wb.batch():
            wb.set_properties(title="Half")
            wb.set_cells({"Nowhere!A1": 1})
    assert wb.to_bytes() == before and "title" not in wb.properties().core
