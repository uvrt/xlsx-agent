"""Round 5, T15: terms redacted from every text a workbook gives (``{"redact": [...],
"with": "[redacted]"}``, :mod:`xlsx_agent.redact`) -- cell values however stored, headers,
sheet names, an Excel table's column names, comments, links, drop-down lists, number
formats; formula text withheld when it holds a term; tables and defined names whose name
holds one not listed; ``find_text`` never finding one, nor saying it was there; and every
tool's result free of them.  Every workbook is made up here: a client list naming two
made-up firms, Acme Corp and Globex, and a code name, Project Zeta."""

from __future__ import annotations

import json
import time

import pytest
from build import Formula, Inline, Rich, WorkbookBuilder
from ooxml_edit.tools import Toolbox

from xlsx_agent import DeclarationError, NameNotFound, Workbook
from xlsx_agent.declare import parse_declarations
from xlsx_agent.redact import WINDOW, Redaction, RedactError, parse_rule
from xlsx_agent.tools import GROUPS, TOOLS, format as xlsx_format

REDACT = {"redact": ["Acme Corp", "Globex", r"re:Project\s+\w+"], "with": "[redacted]"}

#: Words only the terms hold (case aside).
SECRETS = ("acme", "globex", "zeta")

NOTES = "Acme Corp notes"


def book() -> WorkbookBuilder:
    built = WorkbookBuilder()
    clients = built.sheet("Clients")
    clients.rows("A1", [["Ref", "Client", "Globex share", "Amount"],
                        ["C-1", "Acme Corp", "low", 10],
                        ["C-2", Inline("call ACME   corp. today"), "high", 20],
                        ["C-3", Rich(("Glo", "bex")), "mid", 30],
                        ["C-4", "Initech", "low", 40]])
    clients.table("ClientTable", "A1:D5", ["Ref", "Client", "Globex share", "Amount"],
                  write_header=False)
    clients.cell("F1", Formula('"Acme Corp"&"!"', cached="Acme Corp!"))     # literal, str result
    clients.cell("F2", Formula("SUM(D2:D5)", cached=100))                   # nothing to hide
    clients.cell("F3", Formula(f"'{NOTES}'!A2", cached="Initech"))         # names the sheet
    clients.cell("F4", Formula("Globex_cells", cached=3))                  # a name with a term
    clients.cell("F5", 5, style=built.style(num_fmt='0 "Globex"'))
    clients.cell("G1", "Project Zeta")
    clients.comment("B2", "check with Acme Corp first", author="Globex desk")
    clients.hyperlink("B3", url="https://globex.example/files")
    clients.validation_list("C2:C5", ["low", "Globex", "high"])
    clients.validation_list("B2:B5", source=f"'{NOTES}'!$A$1:$A$3")
    notes = built.sheet(NOTES)
    notes.rows("A1", [["Who", "Score"], ["Initech", 1], ["Acme Corp", 2]])
    other = built.sheet("Other")
    other.rows("A1", [["Code", "Count"], ["x", 1], ["y", 2]])
    other.table("GlobexTable", "A1:B3", ["Code", "Count"], write_header=False)
    built.define_name("Globex_cells", "Clients!$D$2:$D$4")
    built.define_name("Scores", f"'{NOTES}'!$B$2:$B$3")
    return built


def opened(*declarations, redact=REDACT) -> Workbook:
    wb = Workbook.open(book().to_bytes())
    wb.declare_tables([redact, *declarations])
    return wb


def shown(thing) -> str:
    return json.dumps(thing, ensure_ascii=False, default=str)


def clean(text: str) -> None:
    for secret in SECRETS:
        assert secret not in text.casefold(), secret


# -- cells ---------------------------------------------------------------------------------------


def test_every_kind_of_text_cell_is_redacted():
    wb = opened()
    values = [v.v for row in wb.read("Clients!B1:C5") for v in row]
    assert values == ["Client", "[redacted] share", "[redacted]", "low",
                      "call [redacted]. today", "high", "[redacted]", "mid", "Initech", "low"]
    assert wb.cell("Clients!F1").v == "[redacted]!"
    assert wb.cell("Clients!G1").v == "[redacted]"
    assert wb.cell("Clients!F5", formats=True).format == '0 "[redacted]"'
    assert wb.redaction.redacted >= 5


def test_without_the_declaration_nothing_changes():
    wb = Workbook.open(book().to_bytes())
    assert wb.cell("Clients!B2").v == "Acme Corp"
    assert [s.name for s in wb.sheets] == ["Clients", NOTES, "Other"]
    assert wb.redaction is None


def test_a_formula_holding_a_term_keeps_its_value_not_its_text():
    wb = opened()
    for ref in ("F1", "F3", "F4"):
        value = wb.cell(f"Clients!{ref}")
        assert value.formula_withheld and value.formula_text is None, ref
    assert wb.cell("Clients!F3").v == "Initech"
    plain = wb.cell("Clients!F2")
    assert plain.formula_text == "SUM(D2:D5)" and not plain.formula_withheld


# -- names ---------------------------------------------------------------------------------------


def test_a_sheet_is_shown_and_addressed_by_its_name_redacted():
    wb = opened()
    assert [s.name for s in wb.sheets] == ["Clients", "[redacted] notes", "Other"]
    assert wb.sheets[1].stored_name == NOTES
    assert wb.cell("'[redacted] notes'!A3").v == "[redacted]"
    with pytest.raises(NameNotFound) as caught:
        wb.cell(f"'{NOTES}'!A1")
    clean(shown(caught.value.valid_options) + shown(caught.value.candidates))


def test_two_sheet_names_that_read_alike_are_made_unique():
    built = WorkbookBuilder()
    for name in ("Acme Corp", "Globex", "[redacted]"):
        built.sheet(name).rows("A1", [["x"]])
    wb = Workbook.open(built.to_bytes())
    wb.declare_tables([REDACT])
    assert [s.name for s in wb.sheets] == ["[redacted] (2)", "[redacted] (3)", "[redacted]"]
    assert wb.cell("'[redacted] (3)'!A1").v == "x"


def test_names_and_tables_holding_a_term_are_not_listed():
    wb = opened()
    assert [t.name for t in wb.tables] == ["ClientTable"]
    assert [n.name for n in wb.names] == ["Scores"]
    for address in ("t:GlobexTable", "t:Globex_cells"):
        with pytest.raises(NameNotFound) as caught:
            wb.region(address)
        clean(shown(caught.value.valid_options) + shown(caught.value.candidates))
    # The table's cells stay: found as a block of its sheet.
    assert [r.address.text for r in wb.find_regions("Other")] == ["t:Other!A1"]
    # A name over the redacted sheet stays, on its shown name.
    assert wb.region("t:Scores").sheet == "[redacted] notes"
    assert [v.v for row in wb.read("t:Scores") for v in row] == [1, 2]


def test_an_excel_tables_column_names_are_redacted():
    wb = opened()
    table = wb.region("t:ClientTable")
    assert table.columns == ("Ref", "Client", "[redacted] share", "Amount")
    assert [v.v for row in wb.read("t:ClientTable['[redacted'] share]") for v in row][:1] == ["low"]
    wb.declare_key("t:ClientTable", "Ref")
    assert wb.cell("t:ClientTable/C-1/Client").v == "[redacted]"


def test_keys_read_as_redacted_and_two_alike_are_no_key():
    built = book()
    built.sheets[0].cell("A4", "Acme Corp").cell("A5", "Globex")   # keys that are terms
    wb = Workbook.open(built.to_bytes())
    wb.declare_tables([REDACT])
    wb.declare_key("t:ClientTable", "Ref")
    assert wb.cell("Clients!A5").v == "[redacted]"
    assert wb.cell("t:ClientTable/C-2/Amount").v == 20
    with pytest.raises(Exception) as caught:
        wb.cell("t:ClientTable/[redacted]/Amount")
    assert getattr(caught.value, "code", None) == "duplicate_key"
    clean(shown(getattr(caught.value, "valid_options", [])) + str(caught.value))


# -- comments, links, lists ----------------------------------------------------------------------


def test_comments_links_and_drop_downs_are_redacted():
    wb = opened()
    (comment,) = wb.comments("Clients")
    assert (comment.text, comment.author) == ("check with [redacted] first", "[redacted] desk")
    (link,) = wb.hyperlinks("Clients")
    assert link.url == "https://[redacted].example/files"
    lists = {stats.column: stats.validation for stats in wb.profile("t:ClientTable")}
    assert lists["[redacted] share"][0].values == ("low", "[redacted]", "high")
    assert lists["Client"][0].values == ("Who", "Initech", "[redacted]")
    assert lists["Client"][0].source == "'[redacted] notes'!A1:A3"


# -- finding -------------------------------------------------------------------------------------


@pytest.mark.parametrize("needle", ["Acme", "acme corp", "Globex", "Glo", "Zeta", "Project"])
def test_find_text_never_finds_a_term(needle):
    wb = opened()
    assert list(wb.find_text(needle)) == []


def test_find_text_finds_the_replacement_only():
    import re
    wb = opened()
    assert list(wb.find_text(re.compile("(?i)acme|glob"))) == []
    found = list(wb.find_text("[redacted]"))
    assert found and all("[redacted]" in f.text for f in found)
    clean(shown([f.as_dict() for f in found]))


# -- the rules -----------------------------------------------------------------------------------


def test_literal_terms_case_and_white_space_aside_and_whole_words():
    rule = Redaction([parse_rule({"redact": ["Acme Corp", "Ann"]})])
    assert rule.text("ACME\n  corp; Annual Ann") == "[redacted]; [redacted]ual [redacted]"
    whole = Redaction([parse_rule({"redact": ["Ann"], "whole_words": True, "with": "X"})])
    assert whole.text("Annual Ann, ann_x (Ann)") == "Annual X, ann_x (X)"


def test_rules_apply_in_order():
    rules = Redaction([parse_rule({"redact": ["Acme"], "with": "Firm"}),
                       parse_rule({"redact": ["re:Firm\\s+Corp"], "with": "[client]"})])
    assert rules.text("Acme Corp and Acme") == "[client] and Firm"


def test_a_long_text_is_searched_a_window_at_a_time():
    rule = Redaction([parse_rule({"redact": [r"re:Project\s+\w+", r"re:\bZed\b"]})])
    for at in (0, WINDOW - 5, 2 * WINDOW - 7, 2 * WINDOW - 3, 5 * WINDOW + 1):
        text = "x " * (at // 2) + "Project Zeta " + "y" * 3000 + " Zed"
        out = rule.text(text)
        assert "Zeta" not in out and "Zed" not in out and out.count("[redacted]") == 2, at
    assert rule.text("a" * 4000 + "Zedd") == "a" * 4000 + "Zedd"     # no false \b at a cut
    assert rule.text("Project " + "a" * 2000) == "[redacted]"        # a match longer than one


def test_a_worst_pattern_costs_a_fraction_of_a_second_on_the_longest_cell():
    rule = Redaction([parse_rule({"redact": ["re:[a-z]+[0-9]"]})])
    start = time.perf_counter()
    rule.text("a" * 32767)
    assert time.perf_counter() - start < 2.0


@pytest.mark.parametrize("item, code, key", [
    ({"redact": []}, "bad_declaration", "redact"),
    ({"redact": [""]}, "bad_declaration", "redact[0]"),
    ({"redact": ["x"], "with": ""}, "bad_declaration", "with"),
    ({"redact": ["x"], "with": "a\nb"}, "bad_declaration", "with"),
    ({"redact": ["x"], "whole_words": "yes"}, "bad_declaration", "whole_words"),
    ({"redact": ["x"], "also": 1}, "bad_declaration", "also"),
    ({"redact": ["re:(a+)+"]}, "bad_pattern", "redact[0]"),
    ({"redact": ["re:.*;"]}, "bad_pattern", "redact[0]"),
    ({"redact": ["re:x*"]}, "bad_pattern", "redact[0]"),
    ({"redact": ["re:("]}, "bad_pattern", "redact[0]"),
    ({"redact": ["Acme"], "with": "[Acme]"}, "bad_declaration", "with"),
])
def test_the_form_is_checked(item, code, key):
    with pytest.raises(RedactError) as caught:
        parse_rule(item)
    assert (caught.value.code, caught.value.key) == (code, key)
    with pytest.raises(DeclarationError) as caught:
        parse_declarations([item])
    assert caught.value.code == code and caught.value.path == f"tables[0].{key}"
    assert "Acme" not in str(caught.value)


# -- declarations --------------------------------------------------------------------------------


def test_declarations_name_sheets_and_headers_as_stored():
    wb = opened({"name": "marks", "sheets": NOTES, "key": "Who"},
                {"name": "clients", "sheets": "Clients", "columns": "A:D", "key": "Ref",
                 "fields": [{"header": "Globex share", "id": "share"},
                            {"header": "Client", "description": "who the work is for"}]},
                {"sheets": "Other", "hidden": True})
    assert wb.region("t:marks").sheet == "[redacted] notes"
    assert wb.region("t:clients").columns == ("Ref", "Client", "share", "Amount")
    assert wb.cell("t:clients/C-2/share").v == "high"
    assert [s.name for s in wb.sheets] == ["Clients", "[redacted] notes"]


def test_a_header_found_by_its_text_holding_a_term_is_found():
    wb = opened({"name": "clients", "sheets": "Clients", "columns": "A:D",
                 "header_row": {"match": ["Ref", "Globex share"]}})
    assert wb.region("t:clients").columns[2] == "[redacted] share"


def test_a_sheet_glob_matches_the_stored_name():
    wb = opened({"name": "marks", "sheets": "Acme*", "key": "Who"})
    assert wb.region("t:marks").sheet == "[redacted] notes"


@pytest.mark.parametrize("declaration, path", [
    ({"name": "globex_list", "sheets": "Clients"}, "tables[1].name"),
    ({"name": "list", "sheets": "Clients", "description": "for Acme Corp"},
     "tables[1].description"),
    ({"name": "list", "sheets": "Clients",
      "fields": [{"header": "Client", "id": "globex_id"}]}, "tables[1].fields[0].id"),
    ({"name": "list", "sheets": "Clients",
      "fields": [{"header": "Client", "description": "Globex"}]},
     "tables[1].fields[0].description"),
])
def test_the_callers_own_words_holding_a_term_are_refused(declaration, path):
    wb = Workbook.open(book().to_bytes())
    with pytest.raises(DeclarationError) as caught:
        wb.declare_tables([REDACT, declaration])
    assert (caught.value.code, caught.value.path) == ("redacted_term", path)
    clean(str(caught.value))
    assert wb.redaction is None and wb.cell("Clients!B2").v == "Acme Corp"   # whole or not


def test_an_unknown_sheet_holding_a_term_is_refused_without_it():
    wb = Workbook.open(book().to_bytes())
    with pytest.raises(DeclarationError) as caught:
        wb.declare_tables([REDACT, {"name": "list", "sheets": "Acme Corp plans"}])
    assert caught.value.code == "unknown_sheet"
    clean(str(caught.value))


def test_terms_are_declared_first_or_not_at_all():
    wb = Workbook.open(book().to_bytes())
    wb.declare_tables([{"name": "list", "sheets": "Clients"}])
    with pytest.raises(DeclarationError) as caught:
        wb.declare_tables([REDACT])
    assert caught.value.code == "redact_late"
    again = opened()
    again.declare_tables([REDACT, {"name": "list", "sheets": "Other"}])   # the same: fine
    assert again.cell("Clients!B2").v == "[redacted]"
    with pytest.raises(DeclarationError) as caught:
        again.declare_tables([{"redact": ["Initech"]}])
    assert caught.value.code == "redact_late"


def test_redact_in_python():
    wb = Workbook.open(book().to_bytes())
    wb.redact([{"redact": ["Initech"], "with": "***"}])
    assert wb.cell("Clients!B5").v == "***"
    assert [v.v for row in wb.read(f"'{NOTES}'!A2:A2") for v in row] == ["***"]


# -- datasets, derive, compare -------------------------------------------------------------------


def test_datasets_derive_and_compare_see_the_text_redacted():
    wb = opened()
    data = wb.dataset("t:ClientTable")
    clean(shown(data.columns) + shown(data.rows))
    derived = wb.derive("t:ClientTable", group_by=["Client"],
                        measures=[{"fn": "sum", "column": "Amount", "as": "total"}])
    clean(shown(derived.rows) + shown(derived.columns))
    assert {"Client": "[redacted]", "total": 40} in derived.rows    # two firms, one group
    changed = book()
    changed.sheets[0].cell("D2", 11)
    other = Workbook.open(changed.to_bytes())
    other.declare_tables([REDACT])
    diff = Workbook.compare(wb, other, tables=["t:ClientTable"], key="Ref")
    clean(shown(diff))


# -- through the tools ---------------------------------------------------------------------------


def toolbox():
    return Toolbox(TOOLS, formats=[xlsx_format(tables=[REDACT])], groups=GROUPS)


def test_no_tool_shows_a_term():
    tools = toolbox()
    session = tools.session()
    changed = book()
    changed.sheets[0].cell("D2", 11)
    one = session.add_blob(book().to_bytes(), name="one.xlsx")
    two = session.add_blob(changed.to_bytes(), name="two.xlsx")
    texts = []
    for blob in (one, two):
        result = tools.dispatch(session, "open_document", {"blob": blob})
        assert result.ok
        texts.append(shown(result.to_json()))
    calls = [
        ("describe", {"doc": "d1"}),
        ("xlsx_read", {"doc": "d1", "target": "Clients!A1:G5", "formulas": True,
                       "hyperlinks": True, "comments": True, "formats": True}),
        ("xlsx_read", {"doc": "d1", "target": "'[redacted] notes'!A1:B3"}),
        ("xlsx_read", {"doc": "d1", "target": "t:ClientTable"}),
        ("xlsx_read", {"doc": "d1", "target": "t:Scores"}),
        ("xlsx_query", {"doc": "d1", "table": "t:ClientTable"}),
        ("xlsx_profile", {"doc": "d1", "target": "t:ClientTable"}),
        ("xlsx_derive", {"doc": "d1", "id": "by_client", "from": "t:ClientTable",
                         "group_by": ["Client"],
                         "measures": [{"fn": "count", "as": "n"}]}),
        ("find_text", {"doc": "d1", "text": "redacted"}),
        ("find_text", {"doc": "d1", "text": "e"}),
        ("xlsx_compare", {"a": "d1", "b": "d2", "key": "Ref"}),
        ("xlsx_validate_data", {"doc": "d1", "schema": {"sheets": ["Clients"], "tables": [
            {"table": "t:ClientTable", "key": ["Ref"],
             "columns": [{"name": "Client", "type": "text", "enum": ["Initech"]}]}]}}),
    ]
    for name, arguments in calls:
        result = tools.dispatch(session, name, arguments)
        assert result.ok, (name, result.to_json())
        texts.append(shown(result.to_json()))
    clean("\n".join(texts))
    assert "Initech" in texts[3] and "[redacted]" in texts[3]


def test_a_term_searched_for_is_not_found_nor_said_to_be_there():
    tools = toolbox()
    session = tools.session()
    blob = session.add_blob(book().to_bytes(), name="one.xlsx")
    assert tools.dispatch(session, "open_document", {"blob": blob}).ok
    results = [tools.dispatch(session, "find_text", {"doc": "d1", "text": text}).to_json()
               for text in ("Acme Corp", "Initechz")]
    for result in results:
        result.pop("meta", None)
    for result in results:
        assert not result["data"].get("matches")
    assert shown(results[0]).replace("Acme Corp", "Q") == shown(results[1]).replace(
        "Initechz", "Q")


def test_redacting_changes_no_tool_definition():
    plain = Toolbox(TOOLS, groups=GROUPS).definitions("anthropic")
    assert toolbox().definitions("anthropic") == plain


def test_a_declaration_error_never_holds_a_term():
    tools = Toolbox(TOOLS, formats=[xlsx_format(tables=[
        REDACT, {"name": "list", "sheets": "Acme Corp plans"}])], groups=GROUPS)
    session = tools.session()
    blob = session.add_blob(book().to_bytes(), name="one.xlsx")
    result = tools.dispatch(session, "open_document", {"blob": blob})
    assert not result.ok and result.error.code == "refused"
    clean(shown(result.to_json()))


def test_a_deck_or_document_source_and_the_problems_see_the_text_redacted():
    from xlsx_agent.tools.docformat import resolve_table
    wb = opened()
    placed = resolve_table(wb, "t:ClientTable")
    clean(shown(placed.columns) + shown(placed.rows) + shown(placed.sources))
    assert "[redacted]" in shown(placed.rows)
    built = book()
    built.sheets[1].cell("C1", Formula("NOW()"))                 # a problem on the sheet
    flagged = Workbook.open(built.to_bytes())
    flagged.declare_tables([REDACT])
    problems = shown(flagged.problems)
    clean(problems)
    assert "'[redacted] notes'!C1" in problems


def test_a_shared_formula_group_holding_a_term_withholds_its_text():
    built = book()
    sheet = built.sheets[0]
    sheet.cell("H2", Formula('B2&"Globex"', cached="x", kind="shared", ref="H2:H3", si=0))
    sheet.cell("H3", Formula("", cached="y", kind="shared", si=0))
    wb = Workbook.open(built.to_bytes())
    wb.declare_tables([REDACT])
    first, member = wb.cell("Clients!H2"), wb.cell("Clients!H3")
    assert first.formula_withheld and first.formula_text is None
    assert member.formula_text is None and member.v == "y"


def test_a_redacted_sheet_is_hidden_by_its_stored_name_too():
    wb = opened({"sheets": NOTES, "hidden": True})
    assert [s.name for s in wb.sheets] == ["Clients", "Other"]
    assert wb.withheld_sheets == ["[redacted] notes"]
    assert [n.name for n in wb.names] == []           # Scores lies on it
    with pytest.raises(NameNotFound):
        wb.cell("'[redacted] notes'!A1")
    other = Workbook.open(book().to_bytes())
    other.declare_tables([REDACT])
    other.withhold_sheets([NOTES])
    assert [s.name for s in other.sheets] == ["Clients", "Other"]
