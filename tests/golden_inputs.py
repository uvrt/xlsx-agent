"""Build the inputs of the golden transcripts (test_golden.py): three made-up workbooks.

``g1-sales-by-region``: a small sales table (``t:Sales``), to summarise by region with
``xlsx_derive`` and read the dataset back.  ``g2-compare-review``: a generated inventory
and the copy a reviewer changed (a price corrected, a row removed, a row added, a
quantity typed as text), compared by SKU.  ``g3-validate-orders``: an order list with a
problem of each common kind, checked against a schema.

Every value is fixed here, dates included -- nothing depends on today's date or the
machine -- and :mod:`build` writes the same bytes for the same calls, so a transcript's
input hashes hold wherever the suite runs (deflate aside: see :mod:`zip_content`).  The
content is invented.  ``python tests/golden_inputs.py <root>`` writes every input to
``<root>/<task>/<file>``.
"""

from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

from build import WorkbookBuilder

# -- g1: sales by region -------------------------------------------------------------------

SALES = [
    # Date, Region, Product, Units, Revenue
    (date(2026, 7, 1), "North", "Kettle", 4, 159.6),
    (date(2026, 7, 2), "South", "Toaster", 2, 69.0),
    (date(2026, 7, 3), "East", "Kettle", 1, 39.9),
    (date(2026, 7, 6), "North", "Blender", 3, 177.0),
    (date(2026, 7, 7), "West", "Toaster", 5, 172.5),
    (date(2026, 7, 8), "South", "Kettle", 2, 79.8),
    (date(2026, 7, 9), "East", "Blender", 2, 118.0),
    (date(2026, 7, 10), "North", "Toaster", 1, 34.5),
    (date(2026, 7, 13), "West", "Blender", 1, 59.0),
    (date(2026, 7, 14), "South", "Blender", 4, 236.0),
    (date(2026, 7, 15), "East", "Toaster", 3, 103.5),
    (date(2026, 7, 16), "North", "Kettle", 2, 79.8),
    (date(2026, 7, 17), "West", "Kettle", 6, 239.4),
    (date(2026, 7, 20), "South", "Toaster", 1, 34.5),
    (date(2026, 7, 21), "East", "Kettle", 2, 79.8),
    (date(2026, 7, 22), "North", "Blender", 1, 59.0),
]


def sales() -> bytes:
    book = WorkbookBuilder()
    sheet = book.sheet("Sales")
    header = ["Date", "Region", "Product", "Units", "Revenue"]
    sheet.rows("A2", [list(row) for row in SALES])
    sheet.table("Sales", f"A1:E{len(SALES) + 1}", header)
    return book.to_bytes()


# -- g2: a generated inventory and its reviewed copy ---------------------------------------

STOCK = [
    # SKU, Item, Category, Qty, Price
    ["SKU-0001", "Hex bolt M6", "Hardware", 120, 0.35],
    ["SKU-0002", "Wood screw 4x40", "Hardware", 800, 0.08],
    ["SKU-0003", "Wall plug 6 mm", "Hardware", 450, 0.05],
    ["SKU-0004", "Claw hammer", "Tools", 14, 18.9],
    ["SKU-0005", "Spirit level 60 cm", "Tools", 9, 24.5],
    ["SKU-0006", "Masking tape 25 mm", "Supplies", 60, 2.75],
    ["SKU-0007", "Sandpaper P120", "Supplies", 200, 0.6],
    ["SKU-0008", "Utility knife", "Tools", 22, 7.95],
]


def inventory(reviewed: bool) -> bytes:
    rows = [list(row) for row in STOCK]
    if reviewed:
        rows[1][4] = 0.09                                       # a price corrected
        rows[4][3] = "9"                                        # a quantity typed as text
        del rows[5]                                             # a row removed (SKU-0006)
        rows.append(["SKU-0009", "Paint brush 50 mm", "Supplies", 35, 3.2])   # a row added
    book = WorkbookBuilder()
    sheet = book.sheet("Stock")
    sheet.rows("A2", rows)
    sheet.table("Stock", f"A1:E{len(rows) + 1}", ["SKU", "Item", "Category", "Qty", "Price"])
    return book.to_bytes()


# -- g3: orders to validate ------------------------------------------------------------------

ORDERS = [
    # Order, Customer, Status, Qty, Discount
    ["ORD-1001", "Alder & Co", "Open", 5, 0.1],
    ["ORD-1002", "Birchwood Ltd", "Shipped", 12, 0.0],
    ["ORD-1003", None, "Open", 3, 0.05],                      # required blank
    ["ORD-1004", "Cedar Works", "Pending", 7, 0.0],           # not in the enum
    ["ORD-1002", "Dunmore Farm", "Closed", 2, 0.15],          # duplicate key
    ["ORD-1006", "Elmfield School", "Open", 0, 0.0],          # out of range (min 1)
    ["ORD-1007", "Fernhill Cafe", "Shipped", "4", 0.2],       # wrong type (text)
    ["ORD-1008", "Greystone Inn", "Closed", 9, 0.6],          # out of range (max 0.5)
]


def orders() -> bytes:
    book = WorkbookBuilder()
    sheet = book.sheet("Orders")
    sheet.rows("A2", ORDERS)
    sheet.table("Orders", f"A1:E{len(ORDERS) + 1}",
                ["Order", "Customer", "Status", "Qty", "Discount"])
    return book.to_bytes()


#: Each task's inputs: file name -> builder.
BUILDERS = {
    "g1-sales-by-region": {"sales-july.xlsx": sales},
    "g2-compare-review": {"inventory-generated.xlsx": lambda: inventory(False),
                          "inventory-reviewed.xlsx": lambda: inventory(True)},
    "g3-validate-orders": {"orders.xlsx": orders},
}


def build(task: str, root: Path) -> None:
    folder = Path(root) / task
    folder.mkdir(parents=True, exist_ok=True)
    for name, make in BUILDERS[task].items():
        (folder / name).write_bytes(make())


if __name__ == "__main__":
    for task in BUILDERS:
        build(task, Path(sys.argv[1]))
