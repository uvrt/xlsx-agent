"""Round 5, T14: a sheet withheld from the model (``{"sheets": ..., "hidden": true}``) is no
sheet of the workbook, whichever way it is reached -- by name or id, through its Excel
tables, the defined names on it or local to it, ``find_text``, the problems, a comparison,
a drop-down drawn from it or a link to it -- and a formula elsewhere that reaches it keeps
its value, not its text (:mod:`xlsx_agent.declare`, "Sheets hidden from the model").  A
withheld sheet's name refused looks exactly like a name no sheet has.  Every workbook is
made up here: a log sheet, a sheet of made-up wages to withhold, and a spare sheet after
it."""

from __future__ import annotations

import json

import pytest
from build import Formula, WorkbookBuilder
from ooxml_edit.tools import Toolbox

from xlsx_agent import Workbook
from xlsx_agent.declare import DeclarationError, parse_declarations
from xlsx_agent.tools import GROUPS, TOOLS, format as xlsx_format

SECRET_SHEET = "Payroll Zeta"
HIDE = {"sheets": SECRET_SHEET, "hidden": True}

#: Words only the withheld sheet holds or names: its name, its table's, its names', its
#: header and its cells.
SECRETS = ("Payroll", "Zeta", "PayTable", "PayCells", "PayTotal", "LocalPay", "Wage", "omega")


def book() -> WorkbookBuilder:
    built = WorkbookBuilder()
    log = built.sheet("Log")
    log.rows("A1", [["Ref", "Item", "Count"], ["R-1", "pump", 3], ["R-2", "valve", 5],
                    ["R-3", "hose", 2]])
    log.cell("E1", Formula("SUM(C2:C4)", cached=10))                         # its own cells
    log.cell("E2", Formula("'Payroll Zeta'!B2", cached=100))                 # a cell there
    log.cell("E3", Formula("SUM('Payroll Zeta'!B:B)", cached=300))           # a column there
    log.cell("E4", Formula('INDIRECT("\'Payroll Zeta\'!B2")', cached=100))   # a string
    log.cell("E5", Formula("SUM(PayTable[Wage])", cached=300))              # its table
    log.cell("E6", Formula("SUM(PayCells)", cached=300))                    # a name there
    log.cell("E7", Formula("SUM(Log:Spare!C2)", cached=3))                  # a run over it
    log.cell("E8", Formula("PayTotal*2", cached=600))                       # a name's formula
    log.cell("E9", Formula('HYPERLINK("#Notes1","go")', cached="go"))       # names no sheet
    log.hyperlink("B2", location="'Payroll Zeta'!A1")                       # a link there
    log.hyperlink("B3", location="Log!A1")                                  # one here
    log.validation_list("D2:D4", source="'Payroll Zeta'!$A$2:$A$3")         # a drop-down
    secret = built.sheet(SECRET_SHEET)
    secret.rows("A1", [["Who", "Wage"], ["omega one", 100], ["omega two", 200]])
    secret.table("PayTable", "A1:B3", ["Who", "Wage"], write_header=False)
    secret.cell("D1", Formula("NOW()"))                                     # a problem there
    secret.comment("A2", "omega remark")
    built.sheet("Spare").rows("A1", [["Ref", "Item", "Count"], ["S-1", "x", 1], ["S-2", "y", 2]])
    built.define_name("PayCells", "'Payroll Zeta'!$B$2:$B$3")
    built.define_name("PayTotal", "SUM('Payroll Zeta'!$B$2:$B$3)")
    built.define_name("LocalPay", "'Payroll Zeta'!$A$2", sheet=SECRET_SHEET)
    built.define_name("Items", "Log!$B$2:$B$4")
    return built


def opened(*declarations) -> Workbook:
    wb = Workbook.open(book().to_bytes())
    wb.declare_tables([HIDE, *declarations])
    return wb


def shown(thing) -> str:
    return json.dumps(thing, ensure_ascii=False, default=str)


def clean(text: str) -> None:
    for secret in SECRETS:
        assert secret.casefold() not in text.casefold(), secret


# -- the workbook as a program sees it -----------------------------------------------------------


def test_the_sheet_is_no_sheet_of_the_workbook():
    wb = opened()
    assert [s.name for s in wb.sheets] == ["Log", "Spare"]
    assert wb.withheld_sheets == [SECRET_SHEET]     # for the program that withheld it
    assert [t.name for t in wb.tables] == []
    assert [n.name for n in wb.names] == ["Items"]
    assert "t:PayTable" not in wb.regions.known()


def test_without_the_declaration_the_sheet_is_as_before():
    wb = Workbook.open(book().to_bytes())
    assert [s.name for s in wb.sheets] == ["Log", SECRET_SHEET, "Spare"]
    assert [t.name for t in wb.tables] == ["PayTable"]
    assert wb.cell("'Payroll Zeta'!A2").v == "omega one"


@pytest.mark.parametrize("address", ["'Payroll Zeta'!A1", "'payroll zeta'!B2:B3",
                                     "#2!A1", "t:PayTable", "t:PayTable[Wage]",
                                     "t:PayTable/#1", "t:PayCells", "t:PayTotal",
                                     "t:'Payroll Zeta'!LocalPay", "t:'Payroll Zeta'!A1"])
def test_no_address_reaches_its_cells(address):
    wb = opened()
    with pytest.raises(KeyError) as caught:
        list(wb.read(address))
    options = getattr(caught.value, "valid_options", [])
    clean(shown(options))
    clean(shown(getattr(caught.value, "candidates", [])))


@pytest.mark.parametrize("hidden, absent", [
    ("'Payroll Zeta'!A1", "'Payroll Zetb'!A1"),
    ("#2!A1", "#7!A1"),
    ("t:PayTable", "t:PayTablf"),
    ("t:'Payroll Zeta'!A1", "t:'Payroll Zetb'!A1"),
])
def test_refused_exactly_as_a_name_no_sheet_has(hidden, absent):
    wb = opened()
    errors = []
    for address in (hidden, absent):
        with pytest.raises(KeyError) as caught:
            list(wb.read(address))
        error = caught.value
        errors.append((type(error), str(error).replace(address[2:] if address.startswith("t:")
                                                       else address.split("!")[0], "X"),
                       getattr(error, "candidates", None)))
    assert errors[0][0] is errors[1][0] and errors[0][2] == errors[1][2]
    assert errors[0][1].replace("Zeta", "?").replace("id 2", "id ?") \
        == errors[1][1].replace("Zetb", "?").replace("id 7", "id ?")


def test_the_sheet_and_its_parts_are_refused_by_every_method():
    wb = opened()
    for call in (lambda: wb.sheet(SECRET_SHEET), lambda: wb.sheet_by_id(2),
                 lambda: list(wb.cells(SECRET_SHEET)), lambda: wb.comments(SECRET_SHEET),
                 lambda: wb.used_range(SECRET_SHEET), lambda: wb.scan(SECRET_SHEET),
                 lambda: wb.find_regions(SECRET_SHEET),
                 lambda: list(wb.find_text("omega", within="'Payroll Zeta'!")),
                 lambda: wb.profile("t:PayTable"), lambda: wb.dataset("t:PayCells"),
                 lambda: wb.derive("t:PayTable", measures=[{"fn": "sum", "column": "Wage"}]),
                 lambda: wb.declare_region("pay", "'Payroll Zeta'!A1:B3")):
        with pytest.raises(KeyError):
            call()


def test_the_catalogue_says_nothing_of_it():
    wb = opened()
    data = wb.catalogue(detail="full")
    assert [s["name"] for s in data["sheets"]] == ["Log", "Spare"]
    clean(shown(data))
    clean(wb.describe())


def test_find_text_searches_none_of_it():
    wb = opened()
    for needle in ("omega", "Wage", "Pay", "Zeta"):
        assert list(wb.find_text(needle)) == [], needle
    assert [f.address for f in wb.find_text("pump")] == ["Log!B2"]


def test_the_problems_mention_none_of_it():
    wb = opened()
    clean(shown([(p.code, p.address) for p in wb.problems]))
    every = Workbook.open(book().to_bytes()).problems
    assert "'Payroll Zeta'!D1" in [p.address for p in every]     # NOW() there: volatile


# -- formulas elsewhere ---------------------------------------------------------------------------


def test_a_formula_that_reaches_it_keeps_its_value_not_its_text():
    wb = opened()
    row = {v.a: v for r in wb.read("Log!E1:E9") for v in r}
    assert row["Log!E1"].formula_text == "SUM(C2:C4)" and not row["Log!E1"].formula_withheld
    for cell in ("E2", "E3", "E4", "E5", "E6", "E7", "E8"):
        value = row[f"Log!{cell}"]
        assert value.formula_withheld and value.formula_text is None, cell
        assert value.v is not None, cell                      # the value is kept
    assert row["Log!E9"].formula_text == 'HYPERLINK("#Notes1","go")'


def test_a_link_to_it_is_left_out_and_a_drop_down_from_it_reads_as_no_sheet():
    wb = opened()
    assert [link.location for link in wb.hyperlinks("Log")] == ["Log!A1"]
    wb.declare_region("drops", "Log!D1:D4", headers=["Pick"])
    hidden = wb.profile("t:drops")[0].validation
    absent = Workbook.open(book().to_bytes())
    absent.declare_tables([{"sheets": "Spare", "hidden": True}])   # the list's sheet shown
    gone = _with_list_from("'Gone Sheet'!$A$2:$A$3")
    gone.declare_region("drops", "Log!D1:D4", headers=["Pick"])
    expected = gone.profile("t:drops")[0].validation
    assert [(a.values, a.reason, a.source) for a in hidden] \
        == [(a.values, a.reason, a.source) for a in expected] == [(None, "unknown_sheet", None)]
    shown_list = absent
    shown_list.declare_region("drops", "Log!D1:D4", headers=["Pick"])
    assert shown_list.profile("t:drops")[0].validation[0].values == ("omega one", "omega two")


def _with_list_from(source: str) -> Workbook:
    built = WorkbookBuilder()
    built.sheet("Log").rows("A1", [["Ref", "Item"], ["R-1", "pump"]]).validation_list(
        "D2:D4", source=source)
    return Workbook.open(built.to_bytes())


# -- declaring -----------------------------------------------------------------------------------


def test_declared_tables_never_take_it():
    wb = opened({"name": "every", "sheets": "*", "columns": "A:C"})
    assert [p.sheet for p in wb.region("t:every").span.parts] == ["Log", "Spare"]
    with pytest.raises(DeclarationError) as caught:
        opened({"name": "pay", "sheets": SECRET_SHEET})
    hidden = caught.value
    with pytest.raises(DeclarationError) as caught:
        opened({"name": "pay", "sheets": "Payroll Zetb"})
    absent = caught.value
    assert (hidden.code, hidden.valid_options) == (absent.code, absent.valid_options) \
        == ("unknown_sheet", ["Log", "Spare"][:len(absent.valid_options)])


def test_a_table_kept_already_on_the_sheet_is_refused_and_nothing_is_withheld():
    wb = Workbook.open(book().to_bytes())
    wb.declare_tables([{"name": "pay", "sheets": SECRET_SHEET}])
    with pytest.raises(DeclarationError) as caught:
        wb.declare_tables([HIDE])
    assert caught.value.code == "on_hidden_sheet" and caught.value.path == "tables[0].sheets"
    assert "t:pay" in str(caught.value)
    assert SECRET_SHEET not in str(caught.value)
    assert wb.withheld_sheets == []
    wb.declare_tables([HIDE, {"name": "pay", "sheets": "Spare"}])   # declared again, elsewhere
    assert wb.withheld_sheets == [SECRET_SHEET]


def test_a_failing_declaration_withholds_nothing():
    wb = Workbook.open(book().to_bytes())
    with pytest.raises(DeclarationError):
        wb.declare_tables([HIDE, {"name": "log", "sheets": "Log", "key": "Nope"}])
    assert wb.withheld_sheets == [] and len(wb.sheets) == 3


def test_globs_lists_optional_and_more_than_once():
    wb = Workbook.open(book().to_bytes())
    wb.declare_tables([{"sheets": ["Payroll *", "Spare"], "hidden": True}])
    assert [s.name for s in wb.sheets] == ["Log"]
    wb.declare_tables([{"sheets": "Payroll Zeta", "hidden": True}])       # again: no error
    wb.declare_tables([{"sheets": "Gone", "hidden": True, "optional": True}])
    assert [s.name for s in wb.sheets] == ["Log"]


@pytest.mark.parametrize("declaration, code, path", [
    ({"sheets": "Gone", "hidden": True}, "unknown_sheet", "tables[0].sheets"),
    ({"sheets": "Gone *", "hidden": True}, "no_sheets", "tables[0].sheets"),
    ({"sheets": ["Log", "Gone"], "hidden": True}, "unknown_sheet", "tables[0].sheets[1]"),
])
def test_a_sheet_that_is_not_there_is_refused_without_its_name(declaration, code, path):
    wb = Workbook.open(book().to_bytes())
    with pytest.raises(DeclarationError) as caught:
        wb.declare_tables([declaration])
    assert (caught.value.code, caught.value.path) == (code, path)
    assert "Gone" not in str(caught.value)
    assert wb.withheld_sheets == []


@pytest.mark.parametrize("declaration, path", [
    ({"sheets": "Notes", "hidden": False}, "tables[0].hidden"),
    ({"sheets": "Notes", "hidden": "yes"}, "tables[0].hidden"),
    ({"sheets": [], "hidden": True}, "tables[0].sheets"),
    ({"hidden": True}, "tables[0].sheets"),
    ({"name": "notes", "sheets": "Notes", "hidden": True}, "tables[0].name"),
    ({"sheets": "Notes", "hidden": True, "optional": 1}, "tables[0].optional"),
])
def test_the_form_is_checked(declaration, path):
    with pytest.raises(DeclarationError) as caught:
        parse_declarations([declaration])
    assert caught.value.code == "bad_declaration" and caught.value.path == path


def test_every_column_hidden_is_still_refused_and_a_hidden_sheet_is_the_way():
    wb = Workbook.open(book().to_bytes())
    with pytest.raises(DeclarationError) as caught:
        wb.declare_tables([{"name": "pay", "sheets": SECRET_SHEET, "fields": [
            {"header": "Who", "hidden": True}, {"header": "Wage", "hidden": True}]}])
    assert caught.value.code == "no_columns"
    wb.declare_tables([HIDE])
    assert wb.withheld_sheets == [SECRET_SHEET]


# -- through the tools ---------------------------------------------------------------------------


def test_no_tool_shows_the_sheet():
    toolbox = Toolbox(TOOLS, formats=[xlsx_format(tables=[HIDE])], groups=GROUPS)
    session = toolbox.session()
    one = session.add_blob(book().to_bytes(), name="one.xlsx")
    two = session.add_blob(book().to_bytes(), name="two.xlsx")
    assert toolbox.dispatch(session, "open_document", {"blob": one}).ok
    assert toolbox.dispatch(session, "open_document", {"blob": two}).ok
    calls = [
        ("describe", {"doc": "d1"}),
        ("xlsx_read", {"doc": "d1", "target": "Log!A1:E9", "formulas": True,
                       "hyperlinks": True, "comments": True}),
        ("xlsx_read", {"doc": "d1", "target": "'Payroll Zeta'!A1:B3"}),
        ("xlsx_read", {"doc": "d1", "target": "#2!A1"}),
        ("xlsx_read", {"doc": "d1", "target": "t:PayTable"}),
        ("xlsx_read", {"doc": "d1", "target": "t:PayCells"}),
        ("xlsx_read", {"doc": "d1", "target": "t:PayTot"}),
        ("xlsx_query", {"doc": "d1", "table": "t:PayTable"}),
        ("xlsx_profile", {"doc": "d1", "target": "t:Items"}),
        ("xlsx_profile", {"doc": "d1", "target": "t:PayTable"}),
        ("xlsx_derive", {"doc": "d1", "from": "t:PayTable",
                         "measures": [{"fn": "count", "as": "n"}]}),
        ("find_text", {"doc": "d1", "text": "omega"}),
        ("find_text", {"doc": "d1", "text": "a"}),
        ("xlsx_compare", {"a": "d1", "b": "d2"}),
        ("xlsx_validate_data", {"doc": "d1", "schema": {"sheets": ["Log"]}}),
    ]
    texts = []
    for name, arguments in calls:
        result = toolbox.dispatch(session, name, arguments)
        text = json.dumps(result.to_json(), ensure_ascii=False, default=str)
        for given in arguments.values():   # what the model wrote it may be told again
            if isinstance(given, str):
                text = text.replace(given, "")
                if SECRET_SHEET in given:
                    text = text.replace(SECRET_SHEET, "")
        texts.append(text)
    clean("\n".join(texts))
    assert "Spare" in texts[0] and "pump" in texts[1]


def test_open_document_counts_the_sheets_shown():
    toolbox = Toolbox(TOOLS, formats=[xlsx_format(tables=[HIDE])], groups=GROUPS)
    session = toolbox.session()
    blob = session.add_blob(book().to_bytes(), name="one.xlsx")
    result = toolbox.dispatch(session, "open_document", {"blob": blob})
    assert result.ok
    clean(json.dumps(result.to_json(), default=str))
    assert '"sheets": 2' in json.dumps(result.to_json(), default=str)


def test_withholding_changes_no_tool_definition():
    plain = Toolbox(TOOLS, groups=GROUPS).definitions("anthropic", groups="core")
    hiding = Toolbox(TOOLS, formats=[xlsx_format(tables=[HIDE])],
                     groups=GROUPS).definitions("anthropic", groups="core")
    assert hiding == plain


def test_the_sheet_name_is_never_in_a_declaration_error():
    toolbox = Toolbox(TOOLS, formats=[xlsx_format(tables=[
        {"sheets": "Secret plans", "hidden": True}])], groups=GROUPS)
    session = toolbox.session()
    blob = session.add_blob(book().to_bytes(), name="one.xlsx")
    result = toolbox.dispatch(session, "open_document", {"blob": blob})
    assert not result.ok and result.error.code == "refused"
    assert "Secret" not in json.dumps(result.to_json(), default=str)


def test_a_schema_requiring_it_finds_it_missing_as_any_sheet_not_there():
    wb = opened()
    found = [[(p.code, p.address.replace(name, "X")) for p in
              wb.validate_data({"sheets": [name]}).problems]
             for name in (SECRET_SHEET, "Payroll Zetb")]
    assert found[0] == found[1] == [("missing_sheet", "'X'!")]
