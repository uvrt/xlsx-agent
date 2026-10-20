# Using the workbook tools with a model provider

The tools in `xlsx_agent.tools` are defined once, in a provider-neutral canonical schema
([`ooxml_edit.tools`](https://github.com/uvrt/ooxml-edit/blob/main/docs/TOOLS.md)). A
`Toolbox` turns them into each provider's definitions, runs the model's calls with
`toolbox.dispatch(session, name, arguments)`, and turns the results back into that
provider's messages. What each tool takes and returns: [ROADMAP.md](../../../ROADMAP.md),
"Tools"; the library underneath: [docs/api-tour.md](../../../docs/api-tour.md).

| Call | Sends |
| --- | --- |
| `definitions("anthropic")` | every tool; the core loaded, the rest deferred behind the BM25 tool-search tool, `cache_control` on the last core tool. Only `xlsx_set_cells` is sent strict: Claude's strict plan is for writing tools, and the others read |
| `definitions("openai-responses")` | the Responses API (GPT-6): the core as functions, each other group a deferred `namespace`, and `tool_search` |
| `definitions("openai-chat", groups=[...])` | Chat Completions, which has no tool search: the core and the groups named |

**Status (0.1.0, and E1 and E3 of 0.2).** Read and derive, the edit foundation -- the
shared `save_document`, `undo` and `batch` serve workbooks (E1) -- and typed cell writes:
`xlsx_set_cells` in the deferred `xlsx_edit` group (E3). `save_document` hands the
application the workbook as `xlsx`, `xlsm` or `xltx`, reopened and every written cell read
back first. The
definitions pass Claude's and OpenAI's documented rules offline
(`ooxml_edit.tools.adapters.anthropic_problems`, `openai_problems`), and on Claude online
(`tests/test_tools_online.py`, marked `provider`): Claude accepts them as sent, and the
loaded core counts **4,958 tokens** (claude-sonnet-5-5's count-tokens, 2026-10-13, 9
definitions with the tool-use prompt; the budget is 6,500). Three tasks were run on Claude
(claude-haiku-5-5) and are kept as golden transcripts (`tests/golden/`). The tools have not
been run live on OpenAI yet; the loop below is the siblings', which a production user ran
on Azure OpenAI's Responses API.

## Groups

| Group | Tools |
| --- | --- |
| `core` (always loaded) | `open_document`, `undo`, `batch`, `describe`, `find_text`, `save_document`, `xlsx_read`, `xlsx_profile`, `xlsx_query`, `xlsx_derive`, `xlsx_compare`, `xlsx_validate_data` |
| `shared_misc` | `close_document`, `read_blob` |
| `xlsx_edit` (E3) | `xlsx_set_cells` |

`describe` and `find_text` are the shared tools with the workbook's handlers: one
definition for decks, documents and workbooks alike. `describe` of any document also
lists the session's datasets (ooxml-edit 0.30.0): id, columns, types, row count and
`from` -- `d1/t:Sales` for a derive from a workbook, `d:by_rep` for one from a dataset --
never their values.

`save_document`, `undo` and `batch` are the shared ones too (E1). A workbook's undo step is
one call (or one `batch`), and its version is the workbook history's: undoing every step
gives the bytes that were opened, and a dataset derived before a change is current again
once the change is undone. `undo`'s `scope` takes a sheet's name: that sheet's latest step
is taken back while no later step changed what it changed -- a step that touched what
every sheet shares (the workbook part, shared strings, styles, the calculation chain,
which the first write of a session always does) is taken back by a plain undo only; a
caller's own undo (`Session.view`) follows the same rule. `save_document`'s validation
report is the structure check (`Workbook.validate`) against the one the workbook opened
with: a new issue refuses the save; `open_document`'s `baseline_problems` counts those
issues (cell facts such as numbers stored as text are `describe`'s).

**Writing cells** (`xlsx_set_cells`, E3). Each item of `cells` is an `address` and
exactly one of `text` (always literal, `"=x"` too), `number`, `bool`, `date` (`YYYY-MM-DD`,
or the session locale's order with a four-digit year, into a cell whose format shows a
date, or with `format`), `formula` (`=D10*2`, the only way to write one), `blank: true`
(the value goes, the format stays) or `markdown` (rich text: `**bold**`, `*italic*`,
`~~strike~~`, each line a line of the cell, `-` and `1.` items shown as `•` and numbers;
`xlsx_agent.markdown`). `xlsx_read`'s `markdown: true` gives text as that Markdown --
rich runs marked, list lines as items, the rest escaped -- so that writing it back gives
the same cell; a cell given line breaks in a row of a fixed height is listed in
`fixed_height`, with a warning. `replace_formulas: true` writes over formula cells.
Every refusal -- a value refused, a formula cell without `replace_formulas`, an array
formula or merged cell, a table's header, a column the application withholds from the
model -- comes before anything changes, each by code and address in the error's
`details`; at most 5,000 cells a call, one undo step. Afterwards every formula value the
workbook gives reads `cache: invalidated` (Excel recalculates when it opens the file;
nothing here does), and a dataset derived from those values is stale; a deck or document
table placed from them says `stale` and the counts in its `from`. The application's write
policy and expected-value preconditions are X4's.

## Setup (every provider)

```python
from ooxml_edit.tools import Toolbox
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS

toolbox = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS)   # toolbox.close() when done
session = toolbox.session()                                 # clock=... for a fixed clock
session.open(xlsx_bytes, name="sales-july.xlsx")            # the model calls it "d1"
system = toolbox.system_prompt(extra=HOUSE_RULES)           # your guidance after the tools'
task = ("In d1, total the revenue by region as the dataset by_region, "
        "then report each figure with its address.")
```

A workbook can also come in as an input the model opens itself:
`b1 = session.add_blob(xlsx_bytes, name="sales.xlsx")`, then `open_document` with
`{"blob": "b1"}`. `.xlsx`, `.xlsm`, `.xltx` and `.xltm` open; `.xls`, `.xlsb` and encrypted
workbooks are refused with the reason. A macro is never run.

**Roles.** Every workbook tool but `undo` and `xlsx_set_cells` reads, so those are the
`review` role's (ooxml-edit 0.45.0,
[TOOLS.md, "Roles"](https://github.com/uvrt/ooxml-edit/blob/main/docs/TOOLS.md)); `undo`
and `xlsx_set_cells` are `edit`'s. A reviewer's toolbox, `Toolbox(..., roles=["review"])`,
sends the same definitions without those two (and without the `xlsx_edit` group), and in a
toolbox with the deck or Word tools a reviewer keeps every reading workbook tool.

| Role | Workbook tools |
| --- | --- |
| `review` | `describe`, `find_text`, `xlsx_read`, `xlsx_profile`, `xlsx_query`, `xlsx_derive`, `xlsx_compare`, `xlsx_validate_data` |
| `edit` | `undo`, `xlsx_set_cells` |
| any | `open_document`, `close_document`, `read_blob`, `batch`, `save_document` |
Agents sharing one session (ooxml-edit 0.46.0) each work through a view of it,
`session.view(roles=[...], context={...})`: the workbook and the datasets derived from it
are the session's, so a deck builder's view uses the `d:` an analyst's view derived.

**The user's locale.** How text that looks like a number or a date is read (numbers stored
as text), and the language of the language-dependent built-in number formats, is the
application's setting -- never the model's -- per toolbox or per session, with the shared
`ooxml_common.locales.Locale`:

```python
from ooxml_common.locales import Locale

belgian = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
                 language="nl-BE", list_separator=";")
toolbox = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, locale=belgian)
other = toolbox.session(locale=None)                  # en-US
```

Without one it is en-US, never the server's. Two reader settings are not locale facts and
go on the format: `xlsx_agent.tools.format(grouping=True, local_dates=False)` in place of
`FORMAT`. `grouping` (on) reads `1,234.5` as a number in en-US; `local_dates` (off: ISO
8601 only) reads `3/4/2026` in the locale's date order, as Excel does. A format code's own
locale tag (`[$-409]`) wins for its cells. No tool definition or prompt changes.

**Declared tables.** The application often knows its workbooks better than any detection:
which sheets hold one table (the same header on every area's sheet), on which row the
header is, which column is the key, what a long header means, and which columns the model
must not see. It says so as plain data -- its own configuration mapped to this -- on the
format (every workbook it opens) or on one session (the workbooks it opens from then on,
in place of the format's); no tool takes it, so no definition grows:

```python
from xlsx_agent.tools import declare, format

tables = {"tables": [
    {"name": "checks", "sheets": "Area *", "header_row": 3, "key": "Ref",
     "description": "One row per check, every area",
     "fields": [{"header": "Rating (1 to 5, in halves)", "id": "rating",
                 "description": "The reviewer's rating"},
                {"header": "Remark", "hidden": True}]},
    {"name": "run", "type": "key_value", "sheets": "Run facts"}]}
toolbox = Toolbox(TOOLS, formats=[format(tables=tables)], groups=GROUPS)
declare(session, tables)              # or per session, in place of the format's
```

`t:checks` is then one table over every `Area ...` sheet, rows in the sheets' order, each
value with its own cell (`'Area South'!C5`); `t:checks/S-2/rating` finds the row's sheet;
`describe` shows `rating` with its description, and never `Remark` -- no tool shows it, an
address naming it is `not_found` without it among the options, and a read of the sheet's
cells gives its cells as `{"a": "'Area North'!E4", "withheld": true}`. A declared table
over an Excel table replaces it: `describe` lists `t:checks` alone, and the Excel table's
`t:` name is refused toward `t:checks`; a defined name over a hidden column is refused
and not listed. `t:run/author/value`
reads a key-value table. A workbook the declarations do not fit is refused at
`open_document` (`refused`, with the declaration's `code` and `path` in `details`); one
marked `optional` is skipped where its sheets are not there. Every rule is in
`xlsx_agent.declare` and docs/api-tour.md, "Declaring tables".

A field may also clean its column's text -- `"clean": {"strip": "^Lead:[^;]*;\\s*",
"collapse_space": true, "split": ";"}`: a leading label stripped, white space collapsed, a
cell split into a list (`"t": "list"`) -- applied where the table is read, never to the
key column. A value a rule changed says `"cleaned": true`; `xlsx_read` and `xlsx_query`
pages count them (`cleaned`: count and first addresses), `xlsx_profile` and `describe`
flag them, `xlsx_derive` counts them per column, and a slide's or document's table from
`t:` gets a list as one cell with a line per item and a warning saying how many values
were cleaned. A pattern that could backtrack without bound is refused when declared.
See `xlsx_agent.clean` and docs/api-tour.md, "Cleaning text".

Terms can be redacted from everything a workbook shows -- `{"redact": ["Acme Corp",
"re:Project\\s+\\w+"], "with": "[redacted]"}` among the declarations: every cell, header,
comment, link, drop-down list, sheet name and an Excel table's column names reads with the
match replaced, in every tool and in a deck's or document's `from`; a formula holding a
match keeps its value, not its text, and `find_text` never finds a term. The model sees
the replacement only -- no count, no flag. See `xlsx_agent.redact`.

**Dates stored as text.** `xlsx_derive`'s `min` and `max` take dates and give a date. On a
column whose dates are stored as text they are refused, and the error says so in one step:
its `details` give the `issue` (`date_as_text`) and its `cells` (the count, the first
addresses), and the message the `coerce: ["date_as_text"]` that reads them as dates. Text
in the locale's day order (`3/4/2026`) is `local_date_text` without `local_dates`: no
coercion reads it then. `first` and `last` take any column as it is. `xlsx_profile` gives
the range of a column's dates stored as text (`date_text`) beside its flag.

**Datasets (`d:`).** `xlsx_derive` keeps its figures in the session as a dataset,
`d:<id>`: columns, rows, and beside every value the cells it came from. `xlsx_read` reads
it back (`{"target": "d:by_region", "sources": true}`, no `doc` needed), `xlsx_derive`
derives from it again, and it outlives `close_document`. When the table has a key column
(the one `describe` reports, or one the caller declares with `Workbook.declare_key` on the
session's workbook), a value's sources are written by key, `t:Stock/{SKU-0001,SKU-0004}/Qty`:
they name the same cells after a consultant inserts or sorts rows, where ranges would not.
Without one they are ranges, as before. A row by key is a target (`t:Parts/BX-17`;
`key` names its key column when several could be, `columns` picks its columns, in order).
`xlsx_read` takes `*` for every row
(`t:Stock/*/Price`) and key sets (`t:Stock/{SKU-0001,SKU-0004}/Price`, read in that
order); `Workbook.expand` turns either into one address per cell, for a check that every
required cell was used. A dataset whose workbook has
changed since is marked stale, never refreshed. An application's own figures are a dataset
too: `Dataset.from_records(rows, types={"opened": "date"})` (ooxml-edit), kept with
`session.keep_datasets`. A Python `date`, `datetime` or `time` in it needs no label; ISO
text is a date only in a column labelled `date` (`datetime`, `time`), or with
`coerce: ["date_as_text"]`. The labels are `number`, `percent`, `date`, `time`,
`datetime`, `text`, `bool` and `error`. Datasets are ooxml-edit's, so one session
can hold a workbook beside a deck or a document -- one toolbox with each library's tools
and formats:

```python
from pptx_agent.tools import FORMAT as DECK, GROUPS as DECK_GROUPS, TOOLS as DECK_TOOLS
from xlsx_agent.tools import FORMAT as WORKBOOK, GROUPS as BOOK_GROUPS, TOOLS as BOOK_TOOLS

toolbox = Toolbox([*DECK_TOOLS, *BOOK_TOOLS], formats=[DECK, WORKBOOK],
                  groups=[*DECK_GROUPS, *BOOK_GROUPS])   # a shared tool merges its kinds
```

**Comparing two versions.** `xlsx_compare` pairs rows by key and opens on a summary per
table: for each column, the cells changed, and for numbers and dates how many went up and
down and the largest move each way, named by key --
`{"rating": {"changed": 4, "up": 2, "down": 2, "max_up": "t:checks/A3/rating +1.5", ...}}`
-- and the rows added and removed by key. The differences that follow name each cell by
key too (`"at": "t:checks/A3/rating"`, with its sheet address in each workbook), so an
answer cites the same cell after rows are inserted or sorted, and carries its row's
`label` -- the row's text in the table's first text column other than the key
(`"label": "Hex bolt M6"`) -- so an answer names the row by what it is, not only by its
key. The summary is on the first page; the differences page as before. A column the
application declared hidden is never compared, named or used as the label.

### Workbook tables into slides and documents

A deck or document tool that takes its rows from a source takes a workbook's table, or a
dataset, as they are: pptx-agent's `ppt_add_table` and `ppt_edit_table` (`from`,
`from_columns`) and the shared `edit_chart` (`from`, `columns`; ooxml-edit 0.22.0, O3).
docx-agent's tables follow. No figure is retyped, and the result names the cells every
placed value came from:

```python
session.entry("d2").document.declare_key("t:Stock", "SKU")   # the workbook, d2; or detected
toolbox.dispatch(session, "edit_chart", {
    "doc": "d1", "target": "s:256", "action": "add", "chart_type": "column",
    "box": {"x": 40, "y": 60, "w": 600, "h": 300},
    "from": "t:Stock/{SKU-0001,SKU-0004}", "columns": ["SKU", "Qty"]})
# result.data["from"]: {"source": "d2/t:Stock/{SKU-0001,SKU-0004}", "doc": "d2", ...,
#   "categories": {"column": "SKU", "rows": "1-2",
#                  "sources": ["t:Stock/{SKU-0001,SKU-0004}/SKU"]},
#   "series": [{"name": "Qty", "column": "Qty", "rows": "1-2",
#               "sources": ["t:Stock/{SKU-0001,SKU-0004}/Qty"]}]}
```

`from` is a table, name or region (`t:Stock`), one column of it (`t:Stock[Qty]`,
`t:Stock/*/Qty`), or a key set's rows in its order (`t:Stock/{SKU-0003,SKU-0001}`); one
value is refused (`invalid_arguments`), as is a table over `Limits.max_dataset_cells`
(`limit`, before a cell is read). Values come as `xlsx_read` types them: numbers as
numbers, dates as ISO 8601 text, blanks as blanks; nothing is coerced -- a number stored as
text stays text, and a chart refuses it, naming its cell. Each value's source is its cell
by key when the table has a key column (as a derive's are), else its cell address, and a
column's are merged (`t:Stock/{...}/Qty`, `Stock!C2:C9`). The workbook is only read: no
undo step, its version unchanged. `Workbook.dataset("t:Stock")` gives the same rows to a
script.

**Limits.** Every limit is the session's (`ooxml_edit.tools.Limits`, given to
`toolbox.session(limits=...)` or the toolbox), and every result says what it cut and gives a
cursor for the rest. The ones a workbook meets:

| Limit | Default | Bounds |
| --- | --- | --- |
| `max_package_bytes`, `max_parts`, `max_uncompressed_bytes`, `max_compression_ratio` | 50 MB, 2,000, 500 MB, 200 | a workbook, checked from the zip's directory before it is unpacked |
| `max_documents` | 4 | open workbooks (and decks and documents) in a session |
| `max_result_chars` | 24,000 | one result |
| `max_list_items` | 50 | rows, matches, differences or problems on one page |
| `max_datasets`, `max_dataset_cells` | 50, 50,000 | datasets in a session; values in one |

And the tools' own: `describe` targets about 3,000 tokens and names what it left for
`xlsx_profile`; `xlsx_read` gives 50 rows by default and at most 500; `xlsx_derive` keeps
at most 1,000 result rows and shows 50, with at most eight sources per value, counted in
units -- a range, or one key of a key set (the rest through `xlsx_read`'s `sources`); `xlsx_validate_data` keeps 10,000 problems and
counts the rest. The read tools do not check `Limits.soft_deadline` yet: a call on a very
large sheet runs to the end (ROADMAP.md, design rule 6).

**What is never logged.** Every call is recorded on the session and logged to the
`ooxml_edit.tools` logger at INFO with its tool, documents, outcome, version and duration;
its arguments appear only as a digest (a hash and their shape: names, types, lengths), so a
filter value or a search text is not in the log (`Toolbox(log_arguments=True)` is the
application's choice to log them). Warnings, errors, exception messages and summaries carry
addresses, types and codes only -- never a cell's value, a comment, a formula's text or a
key -- and the test suite holds them to it with marker strings. Tool *results* do return
content: that is what the model asked for.

The provider SDKs read their keys from the environment (`ANTHROPIC_API_KEY`,
`OPENAI_API_KEY`).

## Claude (Messages API)

```python
import anthropic

client = anthropic.Anthropic()
tools = toolbox.definitions("anthropic")
messages = [{"role": "user", "content": task}]
while True:
    response = client.messages.create(model="claude-sonnet-5-5", max_tokens=8000,
                                      system=system, tools=tools, messages=messages)
    messages.append({"role": "assistant", "content": response.content})
    calls = [block for block in response.content if block.type == "tool_use"]
    if not calls:
        break
    results = [(c.id, toolbox.dispatch(session, c.name, c.input)) for c in calls]
    messages.append(toolbox.render_results("anthropic", results))   # tool_result blocks
```

The core is cached on the last core definition's `cache_control`; a system prompt over
Claude's minimum is cached too when you give it a `cache_control` block of its own.

## OpenAI (Responses API)

```python
from openai import OpenAI

client = OpenAI()
tools = toolbox.definitions("openai-responses")
items = [{"role": "user", "content": task}]
while True:
    response = client.responses.create(model="gpt-6", instructions=system, tools=tools,
                                       input=items)
    items += response.output
    calls = [item for item in response.output if item.type == "function_call"]
    if not calls:
        break
    results = [(c.call_id, toolbox.dispatch(session, c.name, c.arguments)) for c in calls]
    items += toolbox.render_results("openai-responses", results)    # function_call_output items
```

`items += response.output` passes reasoning items and a hosted tool search's calls back as
they came; only `function_call` items are dispatched, by their bare `name`. `dispatch`
takes the arguments as a dict or as the JSON string OpenAI sends, and never raises for the
model's mistakes: an error comes back as a result the model can read, naming the valid
options where there are some.
