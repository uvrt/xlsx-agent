# xlsx-agent

Workbooks (`.xlsx`) as a source an agent can describe, query and derive from without doing
arithmetic itself, every derived value traceable to its cells; later, lossless edits. The
third library beside `pptx-agent` and `docx-agent`, on `ooxml-edit` and `ooxml-common`.

## Status

**0.1.1** (0.1.0: read and derive, milestones R0-R14 of [ROADMAP.md](ROADMAP.md)). A workbook is
opened, catalogued, read, queried, profiled and searched by address; figures are derived
here, each with its source cells, and kept as datasets (`d:<id>`); two workbooks are
compared by key, and a workbook is checked against a declared schema. Since 0.1.0: the
application's own tables declared from its configuration (a table over several sheets,
column ids and descriptions, columns and whole sheets hidden from the model, headers found
by their text, key-value tables, text cleaned by the caller's rules), addresses by key, comparisons summarised per column, dates, tables
as sources for decks and documents
([CHANGELOG.md](CHANGELOG.md)).
0.2's edit foundation is in (E1: a history with atomic batches, undo and redo to the bytes,
a structure check, bounded edit reports, a checked save, and the Excel oracle); the cell
writes themselves come with E3. Not released to PyPI; installed from source.

- The Python API: [docs/api-tour.md](docs/api-tour.md).
- The tools, for Claude and OpenAI: [src/xlsx_agent/tools/README.md](src/xlsx_agent/tools/README.md).
  The loaded core counts 4,958 tokens on Claude (the budget is 6,500).
- What changed: [CHANGELOG.md](CHANGELOG.md).

## Addresses

Every address the library returns (a value's `a`, a problem's or an error's address, a
find result, a range, a tool's output) is in one canonical form, written by one formatter:

- a sheet name bare where Excel writes it bare: letters of any script, digits, `_` and a
  `.`, starting with a letter or `_`, and not read as a cell on the grid (`A1`), an R1C1
  reference or one at its start (`R1C1`, `R`, `C2`, `R1x`), a boolean, or a `.` Excel
  quotes (`Q3.sales`): `Sales!B4`, `Überblick!B4`; any other name quoted, each `'` inside
  doubled: `'Q3 Sales'!B4`, `'Bob''s'!B4`, `'2024'!B4`, `'A1'!B4`, `'R1x'!B4`; a sheet by
  id `#3!B4`;
- column letters upper case, no `$`, rows without leading zeros;
- a range top left to bottom right (`A1:C9`), one of one cell the cell (`B4`); whole
  columns `A:C`, rows `2:5`, a whole sheet `Sales!`;
- `t:`, `d:` and `$` addresses as in [ROADMAP.md](ROADMAP.md), "Addresses".

`canonical_address(text)` turns any spelling an address may be given in (`sales!$b$4`,
`'Sales'!B4:B4`) into that form; with `sheets=[s.name for s in wb.sheets]` a sheet's
name is spelled as the workbook spells it, matched case aside. Compare addresses through
it, never by copying the rules. The quoting is measured (Excel for Mac 16.106 on macOS,
2026-10-27, 253 made-up sheet names): bare where Excel's saved formulas write a name bare,
quoted where Excel quotes it; Excel read every one of those names quoted, so a name is
also quoted where Excel is irregular (it writes some non-ASCII symbols bare and quotes
others, so every symbol is quoted). The full rules: the docstring of
`src/xlsx_agent/address.py`.

## Install from source

Python 3.11-3.15. ooxml-common and ooxml-edit are not on PyPI either; install them from
their checkouts first, then this one:

```bash
pip install -e ../ooxml-common -e ../ooxml-edit
pip install -e .            # or -e '.[test]' for the test suite: python -m pytest -q
```

`-e '.[dev]'` adds the linter and the type checker at the versions CI's `build` job runs:
`ruff check .` (a few bug-finding rules, no formatting) and `mypy` (the package, not
strict); both are configured in `pyproject.toml`. The package ships `py.typed`, so its
annotations are its types for a checker in the code that imports it.


The online check of the tool definitions calls Claude and runs only when asked:
`ANTHROPIC_API_KEY=... python -m pytest -m provider tests/test_tools_online.py`.
