# xlsx-agent

Workbooks (`.xlsx`) as a source an agent can describe, query and derive from without doing
arithmetic itself, every derived value traceable to its cells; later, lossless edits. The
third library beside `pptx-agent` and `docx-agent`, on `ooxml-edit` and `ooxml-common`.

## Status

**0.1.0: read and derive** (milestones R0-R14 of [ROADMAP.md](ROADMAP.md)). A workbook is
opened, catalogued, read, queried, profiled and searched by address; figures are derived
here, each with its source cells, and kept as datasets (`d:<id>`); two workbooks are
compared by key, and a workbook is checked against a declared schema. Since 0.1.0: the
application's own tables declared from its configuration (a table over several sheets,
column ids and descriptions, columns and whole sheets hidden from the model, headers found
by their text, key-value tables, text cleaned by the caller's rules), addresses by key, comparisons summarised per column, dates, tables
as sources for decks and documents
([CHANGELOG.md](CHANGELOG.md)).
Nothing is edited or saved yet: that is 0.2. Not released to PyPI; installed from source.

- The Python API: [docs/api-tour.md](docs/api-tour.md).
- The tools, for Claude and OpenAI: [src/xlsx_agent/tools/README.md](src/xlsx_agent/tools/README.md).
  The loaded core counts 4,958 tokens on Claude (the budget is 6,000).
- What changed: [CHANGELOG.md](CHANGELOG.md).

## Install from source

Python 3.11-3.15. ooxml-common and ooxml-edit are not on PyPI either; install them from
their checkouts first, then this one:

```bash
pip install -e ../ooxml-common -e ../ooxml-edit
pip install -e .            # or -e '.[test]' for the test suite: python -m pytest -q
```

The online check of the tool definitions calls Claude and runs only when asked:
`ANTHROPIC_API_KEY=... python -m pytest -m provider tests/test_tools_online.py`.
