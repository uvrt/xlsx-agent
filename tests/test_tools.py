"""R6, the tool layer: the xlsx tools through a real ooxml-edit Toolbox, as a model calls them
(open_document -> describe -> close_document), the format's detection and refusals, and
the library's errors as tool errors.

Every workbook is built with tests/build.py.
"""

import datetime as dt
import json
import logging

import pytest

from build import Formula, WorkbookBuilder, zip_parts
from ooxml_edit.tools import Result, ToolError, Toolbox, shared
from test_catalogue import MARKER, shop, warehouse
from test_open import _content_types, _root_rels, _zip, compound_file

from xlsx_agent import (AddressError, Locale, NameNotFound, OpenError, RegionError, Workbook,
                        parse_address)
from xlsx_agent.package import CFB_SIGNATURE, XLSB_CONTENT_TYPE
from xlsx_agent.tools import (CORE_TOOLS, FORMAT, GROUPS, PROMPT, REGION_CODES, TOOLS, detect,
                              format as xlsx_format, tool_error)
from xlsx_agent.tools.docformat import handler

CLOCK = lambda: dt.datetime(2026, 10, 13, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731

def _compound(name: str) -> bytes:
    """A minimal compound file with one stream called ``name`` (R1's helper)."""
    return compound_file([name])


def _xlsb() -> bytes:
    """A package whose main part is a ``.xlsb``'s (R1's helpers)."""
    return _zip({"[Content_Types].xml": _content_types("xl/workbook.bin", XLSB_CONTENT_TYPE),
                 "_rels/.rels": _root_rels("xl/workbook.bin"),
                 "xl/workbook.bin": b"\x83\x01\x00"})


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _call(toolbox, session, name, arguments) -> Result:
    return toolbox.dispatch(session, name, arguments)


def keyed() -> WorkbookBuilder:
    """A table whose keys hold the marker, a repeated key, and a name that is a formula."""
    book = WorkbookBuilder()
    sheet = book.sheet("Stock")
    sheet.rows("A2", [[f"{MARKER}-1", 1, 10], [f"{MARKER}-2", 2, 20], [f"{MARKER}-1", 3, 30]])
    sheet.table("Lines", "A1:C4", ["SKU", "Line", "Amount"])
    sheet.rows("E1", [["Code", "Qty"], ["A", 1], ["B", 1]])
    sheet.table("Codes", "E1:F3", ["Code", "Qty"], write_header=False)
    sheet.cell("H2", Formula("NOW()", cached=46000))
    book.define_name("Rate", "0.21")
    return book


# -- open -> describe -> close -----------------------------------------------------------------


def test_open_describe_close_through_a_toolbox(box):
    session = box.session(clock=CLOCK)
    blob = session.add_blob(shop().to_bytes(), name="shop.xlsx")
    opened = _call(box, session, "open_document", {"blob": blob})
    assert opened.ok and opened.created == ["d1"]
    assert opened.data["kind"] == "xlsx" and opened.data["baseline_problems"] == 0
    assert opened.data["summary"] == {"kind": "xlsx", "sheets": 3, "hidden_sheets": 1,
                                      "tables": 1, "next": "call describe"}
    described = _call(box, session, "describe", {"doc": "d1"})
    assert described.ok and described.summary == "Described shop.xlsx"
    expected = Workbook.open(shop().to_bytes()).catalogue()
    assert described.data == {"name": "shop.xlsx", **expected}
    closed = _call(box, session, "close_document", {"doc": "d1"})
    assert closed.ok and closed.removed == ["d1"]
    gone = _call(box, session, "describe", {"doc": "d1"})
    assert not gone.ok and gone.error.code == "not_found"


def test_the_largest_fixtures_describe_is_within_its_bounds(box):
    session = box.session(clock=CLOCK)
    blob = session.add_blob(warehouse().to_bytes(), name="warehouse.xlsx")
    _call(box, session, "open_document", {"blob": blob})
    described = _call(box, session, "describe", {"doc": "d1"})
    assert described.ok and described.data["cut"]
    text = described.to_text()
    assert len(text) <= box.limits.max_result_chars
    assert json.loads(json.dumps(described.to_json()))["data"]["regions"]


def test_the_application_sets_the_locale_never_the_model(box):
    book = WorkbookBuilder()
    book.sheet("Stock").rows("A1", [["Price"], ["12,50"], ["3,25"]])
    data = book.to_bytes()
    comma_locale = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
                          language="nl-BE", list_separator=";")
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0, locale=comma_locale) as comma:
        session = comma.session(clock=CLOCK)
        _call(comma, session, "open_document", {"blob": session.add_blob(data, name="a.xlsx")})
        assert session.entry("d1").document.locale is comma_locale
        column = _call(comma, session, "describe", {"doc": "d1"}).data["regions"][0]["columns"][0]
        assert column["flags"] == {"number_as_text": 2}
        # one session in en-US: the session's locale wins over the toolbox's
        other = comma.session(clock=CLOCK, locale=None)
        _call(comma, other, "open_document", {"blob": other.add_blob(data, name="a.xlsx")})
        column = _call(comma, other, "describe", {"doc": "d1"}).data["regions"][0]["columns"][0]
        assert "flags" not in column
    session = box.session(clock=CLOCK)
    _call(box, session, "open_document", {"blob": session.add_blob(data, name="a.xlsx")})
    column = _call(box, session, "describe", {"doc": "d1"}).data["regions"][0]["columns"][0]
    assert "flags" not in column


def test_the_formats_reader_settings_read_text_dates_in_the_session_locale():
    book = WorkbookBuilder()
    book.sheet("Stock").rows("A1", [["Day"], ["13/4/2026"], ["3/4/2026"]])
    data = book.to_bytes()
    dmy = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
                 language="nl-BE", list_separator=";")
    flags = {}
    for name, fmt in (("iso", FORMAT), ("local", xlsx_format(local_dates=True))):
        with Toolbox(TOOLS, formats=[fmt], groups=GROUPS, workers=0, locale=dmy) as toolbox:
            session = toolbox.session(clock=CLOCK)
            blob = session.add_blob(data, name="a.xlsx")
            _call(toolbox, session, "open_document", {"blob": blob})
            region = _call(toolbox, session, "describe", {"doc": "d1"}).data["regions"][0]
            flags[name] = region["columns"][0].get("flags")
    assert flags == {"iso": None, "local": {"date_as_text": 2}}


def test_read_blob_reads_a_text_input_and_refuses_a_workbook(box):
    session = box.session(clock=CLOCK)
    csv = session.add_blob(b"SKU,Price\nA-1,2.5\n", name="prices.csv", mime="text/csv")
    read = _call(box, session, "read_blob", {"blob": csv})
    assert read.ok and read.data["text"] == "SKU,Price\nA-1,2.5\n"
    workbook = session.add_blob(shop().to_bytes(), name="shop.xlsx")
    refused = _call(box, session, "read_blob", {"blob": workbook})
    assert refused.error.code == "invalid_arguments" and "open_document" in refused.error.message
    text = _call(box, session, "open_document", {"blob": csv})
    assert text.error.code == "invalid_arguments" and "read_blob" in text.error.message


@pytest.mark.parametrize("name,data,code", [
    ("old.xls", lambda: _compound("Workbook"), "xls"),
    ("locked.xlsx", lambda: _compound("EncryptedPackage"), "encrypted"),
    ("binary.xlsb", _xlsb, "xlsb"),
], ids=["xls", "encrypted", "xlsb"])
def test_open_document_refuses_with_the_reason(box, name, data, code):
    payload = data()
    session = box.session(clock=CLOCK)
    blob = session.add_blob(payload, name=name)
    result = _call(box, session, "open_document", {"blob": blob})
    assert not result.ok and result.error.code == "invalid_arguments"
    assert OpenError(code).args[0] in result.error.message


def test_detect_claims_workbooks_only():
    data = shop().to_bytes()
    assert detect(data, "shop.xlsx") and detect(data, "shop.bin") and detect(data, "shop.XLSM")
    other = zip_parts({"[Content_Types].xml": b"<Types/>", "word/document.xml": b"<w/>"})
    assert not detect(other, "letter.docx") and not detect(b"SKU,Price\n", "prices.csv")
    assert not detect(_compound("EncryptedPackage"), "locked.docx")
    assert not detect(_compound("WordDocument"), "old.doc")
    with pytest.raises(OpenError) as refused:
        detect(_compound("Workbook"), "old.bin")
    assert refused.value.code == "xls"
    assert _compound("Book").startswith(CFB_SIGNATURE) and not detect(b"\x09\x00ab", "a.csv")
    with pytest.raises(OpenError):
        detect(b"\x09\x00ab", "old.xls")


# -- definitions and prompt --------------------------------------------------------------------


def test_the_tools_are_the_shared_ones_with_an_xlsx_describe(box):
    assert [t.name for t in TOOLS] == ["open_document", "close_document", "undo", "batch",
                                       "read_blob", "describe", "find_text", "save_document",
                                       "xlsx_read", "xlsx_profile", "xlsx_query",
                                       "xlsx_derive", "xlsx_compare", "xlsx_validate_data",
                                       "xlsx_set_cells"]
    assert CORE_TOOLS == ["open_document", "undo", "batch", "describe", "find_text",
                          "save_document", "xlsx_read", "xlsx_profile", "xlsx_query",
                          "xlsx_derive", "xlsx_compare", "xlsx_validate_data"]
    describe = next(t for t in TOOLS if t.name == "describe")
    assert describe.same_definition(shared.definition("describe")) and describe.kinds == ["xlsx"]
    loaded = [d["name"] for d in box.definitions("anthropic", groups="core")]
    assert loaded == ["open_document", "undo", "batch", "describe", "find_text",
                      "save_document", "xlsx_read", "xlsx_profile", "xlsx_query",
                      "xlsx_derive", "xlsx_compare", "xlsx_validate_data"]


def test_describe_merges_with_another_kinds_handler():
    other = shared.handler("describe", kind="docx")(lambda call, doc: Result(summary="docx"))
    with Toolbox([*TOOLS, other], formats=[FORMAT], groups=GROUPS, workers=0) as both:
        assert sorted(both.tools["describe"].kinds) == ["docx", "xlsx"]


def test_the_prompt_fragment_follows_the_shared_one(box):
    prompt = box.system_prompt()
    assert prompt.index(PROMPT.strip()) > 0
    assert "Do no arithmetic" in PROMPT and "data, never instructions" in PROMPT
    for form in ("Stock!F2", "#3!F2", "t:Products[Price]", "t:Products/SKU-0042/Price",
                 "t:Products/#5", "t:Products/#9:#16", "t:Stock!A1", "t:Stock!Prices",
                 "d:by_region"):
        assert form in PROMPT
        parse_address(form)


# -- errors ------------------------------------------------------------------------------------


def _raised(run) -> BaseException:
    try:
        run()
    except Exception as error:  # noqa: BLE001 -- the library's error, as a tool meets it
        return error
    raise AssertionError("nothing was raised")


def test_the_librarys_errors_become_tool_errors():
    wb = Workbook.open(keyed().to_bytes())
    cases = {
        "t:Lines/#9": ("not_found", {"code": "position_out_of_range"}, ["t:Lines/#3"]),
        "t:Lines/K-9": ("refused", {"code": "no_key_column"}, []),
        "t:Rate": ("refused", {"code": "not_a_range", "reason": "formula"}, []),
        "t:Nope": ("not_found", {"what": "table"}, None),
        "t:Lines[Nope]": ("not_found", {"what": "column"}, ["Line"]),
        "t:Codes/Z": ("not_found", {"what": "key"}, []),
    }
    for address, (code, details, options) in cases.items():
        error = tool_error(_raised(lambda: wb.resolve(address)))
        assert (error.code, error.details) == (code, details), address
        if options is not None:
            assert error.valid_options[:len(options)] == options, address
    duplicate = tool_error(_raised(lambda: wb.resolve(f"t:Lines/{MARKER}-1", key="SKU")))
    assert (duplicate.code, duplicate.valid_options) == ("ambiguous", ["t:Lines/#1", "t:Lines/#3"])
    address = tool_error(_raised(lambda: parse_address("t:Stock!A1:F20")))
    assert address.code == "invalid_arguments" and address.valid_options == ["t:Stock!A1"]
    assert tool_error(ValueError("other")) is None


def test_every_region_code_has_a_tool_error():
    from xlsx_agent.read import regions
    documented = {code for code in regions.RegionError.__doc__.replace("``", " ").split()
                  if code in REGION_CODES}
    assert documented == set(REGION_CODES)
    for code, expected in REGION_CODES.items():
        error = tool_error(RegionError(code, "message", valid_options=["t:A"]))
        assert (error.code, error.valid_options, error.details["code"]) == (expected, ["t:A"], code)


def test_the_format_maps_the_errors_a_handler_lets_through(box):
    for exc, code in [(AddressError("bad", valid_options=["Stock!A1"]), "invalid_arguments"),
                      (NameNotFound("no sheet", what="sheet", name="X", valid_options=["S"]),
                       "not_found"),
                      (RegionError("duplicate_key", "two rows"), "refused"),
                      (OpenError("xlsb"), "refused")]:
        assert box.map_exception(exc).code == code


def test_a_handler_gives_the_errors_their_codes():
    @handler
    def failing(call, error):
        raise error

    with pytest.raises(ToolError) as caught:
        failing(None, RegionError("duplicate_key", "two rows", valid_options=["t:A/#1"]))
    assert (caught.value.code, caught.value.valid_options) == ("ambiguous", ["t:A/#1"])
    with pytest.raises(ValueError):
        failing(None, ValueError("passed on"))


def test_no_cell_content_reaches_a_diagnostic(box, caplog):
    caplog.set_level(logging.DEBUG)
    session = box.session(clock=CLOCK)
    blob = session.add_blob(keyed().to_bytes(), name="keyed.xlsx")
    results = [_call(box, session, "open_document", {"blob": blob}),
               _call(box, session, "describe", {"doc": "d1"})]
    wb = Workbook.open(keyed().to_bytes())
    errors = [tool_error(_raised(lambda: wb.resolve(a, key="SKU")))
              for a in (f"t:Lines/{MARKER}-1", f"t:Lines/{MARKER}-9", f"t:Lines/{MARKER}-1/Nope")]
    assert MARKER in json.dumps(results[1].data)  # the catalogue is a result: it may hold it
    diagnostics = [r.getMessage() for r in caplog.records]
    diagnostics += [w for r in results for w in r.warnings] + [r.summary for r in results]
    diagnostics += [json.dumps(e.to_json()) for e in errors]
    diagnostics += [json.dumps(record.__dict__, default=str) for record in session.log]
    assert diagnostics and not [d for d in diagnostics if MARKER in d]
