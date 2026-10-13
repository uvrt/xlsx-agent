"""The catalogue: what a workbook holds, in one bounded result (what ``describe`` returns).

An agent learns a workbook from this before it reads a row (ROADMAP.md, "Tools")::

    wb.catalogue()                       # the catalogue as data, within its token target
    wb.catalogue(detail="full")          # every column's whole profile, nothing cut
    print(wb.describe(max_tokens=3000))  # the catalogue as text (compact JSON)

**What it lists.**  The workbook: its kind, date system, protection and ``flags`` (cached
values that may be stale -- ``full_calc_on_load``, ``manual_calculation`` -- external
links, a VBA project, data connections: listed, never followed or run).  Every sheet: its
name and id (both forms of its address), its state when not visible, protection, type
when not a worksheet, and the used range.  Every region -- Excel tables, defined names over
a range, declared regions, and the regions header detection finds on each sheet -- with its
address (``t:Products``, ``t:Stock!Prices``, ``t:Stock!A1``), range, data rows, its key (the
one declared, or the one column that can be a key: unique, never blank, text or whole
numbers; ``key_candidates`` when several can -- the first :data:`KEYS_SHOWN`, the rest
counted -- and a value by key then names one), and per
column a summary of its profile (:mod:`xlsx_agent.profile`): type, fill,
min, max and mean, the values when there are :data:`~xlsx_agent.profile.FEW_VALUES` or fewer
(else how many distinct), text lengths, formula share, the types mixed in, allowed
values, and flags by count.  Up to :data:`SAMPLE_ROWS` sample rows per region.  The blocks
whose header is not detected (``header_unclear``, with the reason: declare them), the
defined names that are not ranges (with the reason, never evaluated), and the problems the
workbook has (by code, address and count).

**The token target.**  ``max_tokens`` (default :data:`DEFAULT_TOKENS`, about 3,000) bounds
the catalogue as the siblings' offline proxy counts it (:func:`estimate_tokens`: compact
JSON, characters / 3.5, times 1.48).  A workbook too big for it loses, in order: its second
sample row; then the columns past the first ``k`` of each region, ``k`` as large as fits,
with the rest listed by name -- with one sample row, else with none, else with the rest only
counted, else with only the first few problems, unclear blocks and names; then the first
column of as many regions as fit, from the first; then every column's summary; and, last,
regions past the first ones.  ``cut`` says what was
left out, by count, so the caller knows to ask for it (``xlsx_profile``, in a later
milestone, gives a region's every column).  Every region is listed whenever it fits.

**Content.**  The catalogue is a result: min, max, values, allowed values and sample rows
are cell content, as a read returns them (text cut at :data:`SAMPLE_TEXT` characters, with
``…``).  It is never a diagnostic: nothing here warns or raises with a cell's content
(ROADMAP.md, "Design rules", 7).
"""

from __future__ import annotations

import json
import math
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

from ooxml_edit.charts.workbook import Area, Cell

from .address import RangeAddress, quote_sheet
from .profile import ColumnStats
from .read.regions import Region, RegionError, UnclearRegion
from .read.types import Value

if TYPE_CHECKING:  # pragma: no cover
    from .workbook import Sheet, Workbook

#: The catalogue's token target by default (ROADMAP.md, "Tools": ``describe`` is "about
#: 3,000 tokens by default").
DEFAULT_TOKENS = 3000

#: The offline proxy for Anthropic's count of a JSON text: characters / 3.5 ...
CHARS_PER_TOKEN = 3.5
#: ... times 1.48, the ratio of counted to estimated tokens the siblings measured on
#: their tool definitions (pptx-agent's and docx-agent's T5 and T5b).  JSON data counts no
#: worse than a schema does, so the proxy errs on the side of a smaller catalogue.
COUNTED_PER_ESTIMATED = 1.48

#: The detail levels: ``catalogue`` (within the token target) and ``full``.
DETAILS = ("catalogue", "full")

#: The most sample rows per region.
SAMPLE_ROWS = 2

#: Text in a sample, a value list or an allowed list is cut at this many characters.
SAMPLE_TEXT = 40

#: The most allowed values a column's summary shows per validation list.
ALLOWED_SHOWN = 12

#: The most key candidates shown per region (the rest counted).
KEYS_SHOWN = 3

#: What is kept of the problems, unclear blocks and names when the target is tight.
FEW_ITEMS = 5

#: The relationship types (their last segment) of the workbook part that are flagged.
_LINKS = {"externalLink": "external_links", "vbaProject": "macros",
          "connections": "data_connections"}

#: The sheet types that hold cells.
_CELL_SHEETS = ("worksheet", "macrosheet")


def estimate_tokens(data: Any) -> int:
    """``data``'s tokens as the siblings' offline proxy counts them: its compact JSON's
    characters / :data:`CHARS_PER_TOKEN`, times :data:`COUNTED_PER_ESTIMATED`."""
    text = json.dumps(data, ensure_ascii=False, separators=(",", ":"), default=str)
    return math.ceil(len(text) / CHARS_PER_TOKEN * COUNTED_PER_ESTIMATED)


def catalogue(workbook: "Workbook", *, detail: str = "catalogue",
              max_tokens: int = DEFAULT_TOKENS) -> dict:
    """The catalogue of ``workbook`` (see the module's notes).  ``detail`` is
    ``"catalogue"`` (within ``max_tokens``) or ``"full"`` (every column's whole profile,
    every sample row, nothing cut)."""
    if detail not in DETAILS:
        raise ValueError(f"detail is one of {list(DETAILS)}")
    if max_tokens < 1:
        raise ValueError("max_tokens is a positive number")
    found = _gathered(workbook)
    if detail == "full":
        return _render(found, _Level(full=True))
    for level in _levels(found, max_tokens):
        data = _render(found, level)
        if estimate_tokens(data) <= max_tokens:
            return data
    return data


# ------------------------------------------------------------------------------------------
# What is gathered
# ------------------------------------------------------------------------------------------


@dataclass
class _Found:
    """One region, and what the catalogue says of it."""

    region: Region
    keys: list[str]
    stats: list[ColumnStats]
    samples: list[list[Value]]


@dataclass
class _Gathered:
    head: dict
    sheets: list[dict]
    regions: list[_Found]
    unclear: list[dict]
    names: list[dict]
    problems: list[dict]


def _gathered(workbook: "Workbook") -> _Gathered:
    """What the catalogue is made from, gathered once per workbook state: a workbook does
    not change in 0.1, but the regions declared on it do."""
    key = tuple(workbook.regions.declared)
    cached = workbook._catalogue
    if cached is None or cached[0] != key:
        cached = workbook._catalogue = (key, _gather(workbook))
    return cached[1]


def _gather(workbook: "Workbook") -> _Gathered:
    sheets = workbook.sheets
    stale = set(workbook.calculation)   # in the flags already
    problems = [_problem(p) for p in workbook.problems if p.code not in stale]
    regions: list[Region] = list(workbook.tables)   # sheet by sheet, in the workbook's order
    unclear: list[dict] = []
    for sheet in sheets:
        if sheet.part is None or sheet.type not in _CELL_SHEETS:
            continue
        try:
            blocks = workbook.find_regions(sheet)
        except RegionError as error:  # a block the reader cannot take: by code, never content
            problems.append({"code": error.code, "address": _written(sheet)})
            continue
        for block in blocks:
            if isinstance(block, UnclearRegion):
                unclear.append({"address": block.address, "code": block.code,
                                "reason": block.reason})
            else:
                regions.append(block)
    names: list[dict] = []
    for name in workbook.names:
        if name.builtin:
            continue
        address = name.address
        if name.range is None or address is None:
            item: dict = {"name": address.text if address is not None else name.name}
            if name.reason is not None:
                item["not_a_range"] = name.reason
            if name.hidden:
                item["hidden"] = True
            names.append(item)
            continue
        try:
            regions.append(workbook.region(address))
        except (RegionError, KeyError) as error:
            names.append({"name": address.text, "not_a_range": getattr(error, "code", "unknown")})
    regions.extend(workbook.regions.declared)
    found = [_found(workbook, region) for region in regions]
    return _Gathered(head=_head(workbook), sheets=[_sheet(workbook, s) for s in sheets],
                     regions=found, unclear=unclear, names=names, problems=problems)


def _found(workbook: "Workbook", region: Region) -> _Found:
    stats = workbook.profile(region)
    if region.key is not None:
        keys = [region.key]
    elif any(_may_be_key(s) for s in stats):
        keys = workbook.key_columns(region)
    else:  # no column can be one: no pass over the cells to find out
        keys = []
    samples: list[list[Value]] = []
    data = region.data
    if data is not None:
        last = Cell(min(data.first.row + SAMPLE_ROWS - 1, data.last.row), data.last.column)
        samples = list(workbook.read(RangeAddress(region.sheet_ref, Area(data.first, last))))
    return _Found(region, keys, stats, samples)


def _may_be_key(stats: ColumnStats) -> bool:
    """Whether the profile leaves a column able to be a key: every data cell a value, all
    different (or too many to count), text or numbers.
    :meth:`~xlsx_agent.read.regions.Regions.key_columns` decides."""
    return (stats.rows > 0 and stats.filled == stats.rows
            and stats.distinct in (stats.rows, None) and set(stats.types) <= {"text", "number"})


def _head(workbook: "Workbook") -> dict:
    out: dict = {"kind": workbook.kind, "date_system": 1904 if workbook.date1904 else 1900}
    if workbook.protected:
        out["protected"] = True
    flags: dict = {}
    for code in workbook.calculation:
        flags[code] = True
    for rel in workbook.package.relationships(workbook.part).values():
        flag = _LINKS.get(rel.type.rpartition("/")[2])
        if flag == "external_links":
            flags[flag] = flags.get(flag, 0) + 1
        elif flag is not None:
            flags[flag] = True
    if flags:
        out["flags"] = flags
    return out


def _sheet(workbook: "Workbook", sheet: "Sheet") -> dict:
    out: dict = {"name": sheet.name, "id": sheet.id}
    if sheet.state != "visible":
        out["state"] = sheet.state
    if sheet.protected:
        out["protected"] = True
    if sheet.type != "worksheet":
        out["type"] = sheet.type
    if sheet.part is not None and sheet.type in _CELL_SHEETS:
        used = workbook.used_range(sheet)
        out["used"] = used.ref if used is not None else None
    return out


def _written(sheet: "Sheet") -> str:
    return quote_sheet(sheet.name) if sheet.name else f"#{sheet.id}"


def _problem(problem: Any) -> dict:
    out: dict = {"code": problem.code}
    if problem.address is not None:
        out["address"] = problem.address
    if problem.count != 1:
        out["count"] = problem.count
    return out


# ------------------------------------------------------------------------------------------
# What is shown
# ------------------------------------------------------------------------------------------


@dataclass(frozen=True)
class _Level:
    """How much of what was gathered is shown."""

    full: bool = False
    #: Columns summarised per region (``None``: all).
    columns: int | None = None
    samples: int = SAMPLE_ROWS
    #: The names of the columns not summarised.
    column_names: bool = True
    #: Problems, unclear blocks and names shown (``None``: all).
    items: int | None = None
    #: Regions shown (``None``: all).
    regions: int | None = None
    #: The regions, from the first, that get ``columns`` summaries; the rest get none
    #: (``None``: all).
    summarised: int | None = None

    def shown(self, index: int, width: int) -> int:
        """How many columns of the region at ``index``, ``width`` wide, are summarised."""
        if self.full or self.columns is None:
            return width
        if self.summarised is not None and index >= self.summarised:
            return 0
        return min(self.columns, width)


def _levels(found: _Gathered, max_tokens: int):
    """The levels to try, from all of it to the least: the first that fits is kept."""
    widest = max((len(f.region.columns) for f in found.regions), default=0)
    yield _Level()
    yield _Level(samples=1)
    # The most columns per region that fits, with a sample row and the other columns'
    # names, then without the sample, then with the others counted, then with only the
    # first few problems, unclear blocks and names: the estimate grows with the columns, so
    # the largest that fits is found by halving.
    for samples, names, items in ((1, True, None), (0, True, None), (0, False, None),
                                  (0, False, FEW_ITEMS)):
        low, high, best = 1, widest, None
        while low <= high:
            middle = (low + high) // 2
            level = _Level(columns=middle, samples=samples, column_names=names, items=items)
            if estimate_tokens(_render(found, level)) <= max_tokens:
                best, low = level, middle + 1
            else:
                high = middle - 1
        if best is not None:
            yield best
    # The first column of as many regions as fit, from the first.
    low, high, best = 1, len(found.regions) - 1, None
    while low <= high:
        middle = (low + high) // 2
        level = _Level(columns=1, samples=0, column_names=False, items=FEW_ITEMS,
                       summarised=middle)
        if estimate_tokens(_render(found, level)) <= max_tokens:
            best, low = level, middle + 1
        else:
            high = middle - 1
    if best is not None:
        yield best
    yield _Level(columns=0, samples=0, column_names=False)
    yield _Level(columns=0, samples=0, column_names=False, items=FEW_ITEMS)
    count = len(found.regions)
    while count > 1:
        count //= 2
        yield _Level(columns=0, samples=0, column_names=False, items=FEW_ITEMS, regions=count)
    yield _Level(columns=0, samples=0, column_names=False, items=0, regions=0)


def _render(found: _Gathered, level: _Level) -> dict:
    out = dict(found.head)
    out["sheets"] = found.sheets
    regions = found.regions if level.regions is None else found.regions[:level.regions]
    out["regions"] = [_region(f, level, level.shown(k, len(f.stats)))
                      for k, f in enumerate(regions)]
    cut: dict = {}
    for name, items in (("unclear", found.unclear), ("names", found.names),
                        ("problems", found.problems)):
        if not items:
            continue
        shown = items if level.items is None else items[:level.items]
        out[name] = shown
        if len(shown) < len(items):
            cut[name] = len(items) - len(shown)
    if not level.full:
        columns = sum(len(f.stats) - level.shown(k, len(f.stats))
                      for k, f in enumerate(regions))
        if columns:
            cut["columns"] = columns
        if level.samples < SAMPLE_ROWS and any(f.samples for f in regions):
            cut["sample_rows"] = True
        if len(regions) < len(found.regions):
            cut["regions"] = len(found.regions) - len(regions)
    if cut:
        out["cut"] = cut
    return out


def _region(found: _Found, level: _Level, shown: int) -> dict:
    region = found.region
    out: dict = {"address": region.address.text, "kind": region.kind,
                 "range": region.range.text, "rows": region.rows}
    if region.hidden:
        out["hidden"] = True
    if region.header_rows != 1:
        out["header_rows"] = region.header_rows
    if region.totals_rows:
        out["totals_rows"] = region.totals_rows
    if len(found.keys) == 1:
        out["key"] = found.keys[0]
    elif found.keys:
        out["key_candidates"] = found.keys[:KEYS_SHOWN]
        if len(found.keys) > KEYS_SHOWN:
            out["more_key_candidates"] = len(found.keys) - KEYS_SHOWN
    if region.flags:
        out["flags"] = [_problem(p) for p in region.flags]
    if level.full:
        out["columns"] = [s.as_dict() for s in found.stats]
    else:
        out["columns"] = [_column(s) for s in found.stats[:shown]]
        rest = found.stats[shown:]
        if rest:
            out["more_columns"] = [s.column for s in rest] if level.column_names else len(rest)
    if found.samples and level.samples and shown:
        out["sample"] = [[_sample(v) for v in row[:shown]] for row in found.samples[:level.samples]]
    return out


def _column(stats: ColumnStats) -> dict:
    """A column's profile, summarised: what an agent needs to pick columns and filters."""
    out: dict = {"name": stats.column, "type": stats.type, "fill": round(stats.fill, 2)}
    if stats.measured:
        out["min"], out["max"] = _shown(stats.min), _shown(stats.max)
        if stats.mean is not None:
            out["mean"] = _shown(stats.mean)
    if stats.values is not None:
        out["values"] = [[_shown(d.v), d.count] for d in stats.values]
    elif stats.distinct is not None:
        out["distinct"] = stats.distinct
    else:
        out["distinct"] = "many"
    if stats.text_length_max is not None:
        out["text_length"] = [stats.text_length_median, stats.text_length_max]
    if stats.formulas:
        out["formula_share"] = round(stats.formula_share, 2)
    others = {t: c.count for t, c in stats.types.items() if t != stats.type}
    if others:
        out["also"] = others
    if stats.validation:
        out["allowed"] = [_allowed(a) for a in stats.validation]
    flags = {code: c.count for code, c in stats.flags.items() if code != "mixed_types"}
    if flags:
        out["flags"] = flags
    return out


def _allowed(allowed: Any) -> dict:
    if allowed.values is None:
        return {"reason": allowed.reason}
    values = list(allowed.values)
    out: dict = {"values": [_shown(v) for v in values[:ALLOWED_SHOWN]]}
    if len(values) > ALLOWED_SHOWN or allowed.cut:
        out["more"] = True
    if allowed.source is not None:
        out["source"] = allowed.source
    return out


def _sample(value: Value) -> Any:
    return _shown(value.v) if value.t != "blank" else None


def _shown(v: Any) -> Any:
    """A value as the catalogue shows it: text cut at :data:`SAMPLE_TEXT` characters; a
    float with at most 15 significant digits (what a cell shows of a double)."""
    if isinstance(v, str) and len(v) > SAMPLE_TEXT:
        return v[:SAMPLE_TEXT] + "…"
    if isinstance(v, float) and math.isfinite(v):
        return float(f"{v:.15g}")
    return v


__all__ = ["CHARS_PER_TOKEN", "COUNTED_PER_ESTIMATED", "DEFAULT_TOKENS", "DETAILS",
           "SAMPLE_ROWS", "SAMPLE_TEXT", "catalogue", "estimate_tokens"]
