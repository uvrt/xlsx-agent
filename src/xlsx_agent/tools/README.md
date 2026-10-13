# Using the workbook tools with a model provider

The tools in `xlsx_agent.tools` are defined once, in a provider-neutral canonical schema
([`ooxml_edit.tools`](https://github.com/uvrt/ooxml-edit/blob/main/docs/TOOLS.md)). A
`Toolbox` turns them into each provider's definitions, runs the model's calls with
`toolbox.dispatch(session, name, arguments)`, and turns the results back into that
provider's messages. What each tool takes and returns: [ROADMAP.md](../../../ROADMAP.md),
"Tools"; the library underneath: [docs/api-tour.md](../../../docs/api-tour.md).

| Call | Sends |
| --- | --- |
| `definitions("anthropic")` | every tool; the core loaded, the rest deferred behind the BM25 tool-search tool, `cache_control` on the last core tool. No tool is sent strict: Claude's strict plan is for writing tools, and 0.1's tools all read |
| `definitions("openai-responses")` | the Responses API (GPT-6): the core as functions, each other group a deferred `namespace`, and `tool_search` |
| `definitions("openai-chat", groups=[...])` | Chat Completions, which has no tool search: the core and the groups named |

**Status (0.1.0).** Read and derive only: nothing a tool does changes a workbook, and no
tool saves one. The definitions pass Claude's and OpenAI's documented rules offline
(`ooxml_edit.tools.adapters.anthropic_problems`, `openai_problems`), and on Claude online
(`tests/test_tools_online.py`, marked `provider`): Claude accepts them as sent, and the
loaded core counts **4,958 tokens** (claude-sonnet-5-5's count-tokens, 2026-10-13, 9
definitions with the tool-use prompt; the budget is 6,000). Three tasks were run on Claude
(claude-haiku-5-5) and are kept as golden transcripts (`tests/golden/`). The tools have not
been run live on OpenAI yet; the loop below is the siblings', which a production user ran
on Azure OpenAI's Responses API.

## Groups

| Group | Tools |
| --- | --- |
| `core` (always loaded) | `open_document`, `describe`, `find_text`, `xlsx_read`, `xlsx_profile`, `xlsx_query`, `xlsx_derive`, `xlsx_compare`, `xlsx_validate_data` |
| `shared_misc` | `close_document`, `read_blob` |

`describe` and `find_text` are the shared tools with the workbook's handlers: one
definition for decks, documents and workbooks alike.

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
Without one they are ranges, as before. `xlsx_read` takes `*` for every row
(`t:Stock/*/Price`) and key sets (`t:Stock/{SKU-0001,SKU-0004}/Price`, read in that
order); `Workbook.expand` turns either into one address per cell, for a check that every
required cell was used. A dataset whose workbook has
changed since is marked stale, never refreshed. Datasets are ooxml-edit's, so one session
can hold a workbook beside a deck or a document -- one toolbox with each library's tools
and formats:

```python
from pptx_agent.tools import FORMAT as DECK, GROUPS as DECK_GROUPS, TOOLS as DECK_TOOLS
from xlsx_agent.tools import FORMAT as WORKBOOK, GROUPS as BOOK_GROUPS, TOOLS as BOOK_TOOLS

toolbox = Toolbox([*DECK_TOOLS, *BOOK_TOOLS], formats=[DECK, WORKBOOK],
                  groups=[*DECK_GROUPS, *BOOK_GROUPS])   # a shared tool merges its kinds
```

The deck and Word table and chart tools do not take `d:` yet: that is ooxml-edit's O3,
after 0.1 (ROADMAP.md, "ooxml-edit changes"). Until then a model reads the dataset with
`xlsx_read` and passes its values on, each with its `d:` address.

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
