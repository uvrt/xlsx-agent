"""A made-up "Shelf" workbook for recorded values (round 53, R108): a table of shelf items whose
``notes`` say, in many wordings -- English, French, rich text with a bold run -- how far each
item got, ending in a judgement a reader makes ("... Outcome: partly done"), plus rows whose
notes are blank, one written as white space only, and a key spelled the long way."""

from __future__ import annotations

from build import Rich, Run, WorkbookBuilder

#: Each row: Ref, Item, notes (text, Rich or None).
ROWS = [
    ["S-01", "brackets", "Fitted on all bays. Outcome: done"],
    ["S-02", "labels", "Half the labels printed; rest waiting on stock. Outcome: partly done"],
    ["S-03", "lighting", Rich((Run("Outcome: "), Run("not started", bold=True),
                               Run(" - supplier unknown")))],
    ["S-04", "dividers", "Livré et installé. Résultat : terminé"],
    ["S-05", "end caps", "Commandé, livraison partielle. Résultat : en partie"],
    ["S-06", "baskets", None],
    ["S-07", "price rails", "   "],
    ["S-08", "mirrors", Rich((Run("All done", bold=True), Run(", signed off 2026-03-02")))],
    ["S-09", "hooks", "no progress yet"],
    ["S-010", "shelf liners", "Mostly there: 8 of 10 bays."],
]


def shelf_book(rows: list | None = None) -> WorkbookBuilder:
    out = WorkbookBuilder()
    sheet = out.sheet("Shelf")
    every = rows if rows is not None else ROWS
    sheet.rows("A1", [["Ref", "Item", "notes"], *every])
    sheet.table("Shelf", f"A1:C{len(every) + 1}", ["Ref", "Item", "notes"])
    return out


#: The values an agent judged, one per row with notes (S-09 and S-010 left for later).
JUDGED = [
    {"key": "S-01", "value": "done"},
    {"key": "S-02", "value": "partly done", "note": "half printed"},
    {"key": "S-03", "value": "not started"},
    {"key": "S-04", "value": "done", "note": "French: terminé"},
    {"key": "S-05", "value": "partly done", "note": "French: en partie"},
    {"key": "S-08", "value": "done", "note": "bold 'All done'"},
]


#: Round 54, R113: a "section" heading row and an "item" row share each reference (only the
#: item row has notes), and a blank spacer row repeats one; a grade 1-5 is judged per item.
SECTIONS = [
    ["S-04", "Section: dividers", None],
    ["S-04", "dividers", "Livré et installé, all bays. Grade 4 of 5"],
    [None, None, None],
    ["S-05", "Section: end caps", None],
    ["S-05", "end caps", "Partial delivery; grade 3"],
    ["S-05", None, None],
    ["S-06", "Section: baskets", None],
    ["S-06", "baskets", Rich((Run("Not started", bold=True), Run(": grade 1")))],
    ["S-07", "hooks", "All fitted (5)"],
    ["S-08", "Section: mirrors", "   "],
    ["S-08", "mirrors", "Two of three hung, grade 3.5"],
]
