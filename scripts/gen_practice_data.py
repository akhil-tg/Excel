"""Generate the Practice Data workbooks (assets/files/practice-data/).

Run from the repo root:

    python scripts/gen_practice_data.py

Each workbook holds raw, realistic office data for learning lookups:
- Main data table (100 rows) on the first sheet.
- Lookup/decision tables pre-placed on their own sheet, exactly how the
  practice-data.html page teaches learners to organise a workbook.
Data is generated deterministically (fixed seed) so regenerating produces
identical files.
"""
import os
import random
import sys
from datetime import date, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlcommon import (
    new_workbook, put_headers, put_rows, put_title, set_widths,
)

OUT = "assets/files/practice-data"
random.seed(20261115)

FIRST = ["Ramesh", "Priya", "Anil", "Deepa", "Sunil", "Meena", "Karan", "Lata",
         "Vikram", "Sunita", "Rahul", "Kavita", "Arun", "Bhavna", "Chetan",
         "Divya", "Farhan", "Gita", "Harish", "Ishita", "Jayanth", "Kiran",
         "Lalit", "Manju", "Naveen", "Omkar", "Pooja", "Qureshi", "Ravi",
         "Sneha", "Tarun", "Uma", "Vinod", "Bela", "Yash"]
LAST = ["Kumar", "Sharma", "Verma", "Nair", "Patel", "Iyer", "Mehta", "Rao",
        "Singh", "Joshi", "Das", "Reddy", "Gupta", "Bose", "Menon", "Ali",
        "Pillai", "Sen", "Shah", "Kulkarni", "Bhatt", "Chowdhury", "Mishra",
        "Kapoor", "Naik"]
DEPTS = ["Accounts", "Sales", "Stores", "HR", "Production", "Purchase"]


def person_names(n):
    """n unique full names drawn deterministically."""
    seen, out = set(), []
    while len(out) < n:
        name = random.choice(FIRST) + " " + random.choice(LAST)
        if name not in seen:
            seen.add(name)
            out.append(name)
    return out


def style_table(ws, start_row, headers, rows, widths, num_formats=None):
    put_headers(ws, start_row, headers)
    put_rows(ws, start_row + 1, rows, num_formats=num_formats)
    set_widths(ws, widths)


def gen_employee():
    wb = new_workbook("Employees")
    ws = wb["Employees"]
    put_title(ws, "Employee master — 100 employees",
              "Practise VLOOKUP, XLOOKUP and INDEX + MATCH on this table.")
    names = person_names(100)
    start = 4
    headers = ["Emp ID", "Name", "Department", "Join Date", "Salary"]
    rows = []
    base = date(2022, 4, 1)
    for i in range(100):
        join = base + timedelta(days=random.randint(0, 1650))
        rows.append(["E%03d" % (101 + i), names[i], random.choice(DEPTS),
                     join, random.randrange(20000, 48000, 500)])
    style_table(ws, start, headers, rows, [10, 20, 14, 14, 12],
                num_formats={4: "DD-MMM-YYYY", 5: "#,##0"})
    wb.save(OUT + "/employee-master.xlsx")


def gen_products():
    wb = new_workbook("Order Log", "Price List")
    # --- Sheet1: order log (learner works here) ---
    ws = wb["Order Log"]
    put_title(ws, "Order log — 100 orders",
              "Fetch Product and Price from the Price List sheet with VLOOKUP, "
              "then Bill Value = Price x Qty. Wrap lookups in IFERROR — a few "
              "SKUs below are missing from the price list on purpose.")
    skus = ["NB-001", "NB-002", "PN-010", "PN-011", "CL-020", "ST-030",
            "GL-040", "MP-050", "KB-060", "CP-070", "PR-100", "PR-101",
            "MN-080", "SP-090", "CH-100", "HP-110", "MK-120", "TR-130",
            "FM-140", "BB-150"]
    start = 4
    headers = ["Order ID", "Date", "SKU", "Qty", "Product", "Price", "Bill Value"]
    rows = []
    base = date(2026, 11, 1)
    for i in range(100):
        sku = random.choice(skus) if i % 19 != 7 else random.choice(
            ["XX-999", "YY-888", "ZZ-777", "AA-666"])
        rows.append(["SO-%d" % (2001 + i),
                     base + timedelta(days=i // 3),
                     sku, random.choice([1, 2, 3, 4, 5, 6, 8, 10, 12, 15,
                                         20, 25, 30, 40, 50, 60]),
                     "", "", ""])
    style_table(ws, start, headers, rows, [11, 13, 9, 7, 20, 10, 12],
                num_formats={2: "DD-MMM-YYYY"})
    # --- Sheet2: price list (lookup table) ---
    ws2 = wb["Price List"]
    put_title(ws2, "Product price list — 20 products",
              "The lookup table. Keep it sorted by SKU for reading ease.")
    products = {
        "NB-001": ("Notebook A5", "Stationery", 45), "NB-002": ("Notebook B5", "Stationery", 60),
        "PN-010": ("Pen pack (10)", "Stationery", 72), "PN-011": ("Gel pen single", "Stationery", 15),
        "CL-020": ("A4 paper ream", "Stationery", 210), "ST-030": ("Stapler medium", "Stationery", 130),
        "GL-040": ("Glue stick", "Stationery", 35), "MP-050": ("Mouse wireless", "Electronics", 640),
        "KB-060": ("Keyboard wired", "Electronics", 890), "CP-070": ("CPU cabinet", "Electronics", 9200),
        "PR-100": ("Printer LaserJet", "Electronics", 12500), "PR-101": ("Ink cartridge", "Electronics", 850),
        "MN-080": ("Monitor 22in", "Electronics", 7800), "SP-090": ("Speakers USB", "Electronics", 1150),
        "CH-100": ("Charger 65W", "Electronics", 980), "HP-110": ("Headset with mic", "Electronics", 1450),
        "MK-120": ("Marker pack (4)", "Stationery", 55), "TR-130": ("Toner cartridge", "Electronics", 3200),
        "FM-140": ("File box", "Stationery", 260), "BB-150": ("Whiteboard marker", "Stationery", 28),
    }
    prows = [[k, products[k][0], products[k][1], products[k][2]] for k in skus]
    style_table(ws2, 4, ["SKU", "Product", "Category", "Price"], prows,
                [10, 20, 14, 10], num_formats={4: "#,##0"})
    wb.save(OUT + "/products-orders.xlsx")


def gen_twoway():
    wb = new_workbook("Sales Grid")
    ws = wb["Sales Grid"]
    put_title(ws, "Region x month sales grid — 12 branches, 12 months",
              "Practise the two-way lookup: MATCH the branch row and the month "
              "column, then INDEX to return the sales figure where they cross.")
    branches = ["Delhi", "Mumbai", "Kolkata", "Chennai", "Bengaluru", "Hyderabad",
                "Pune", "Ahmedabad", "Jaipur", "Lucknow", "Kochi", "Nagpur"]
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    start = 4
    headers = ["Branch"] + months
    rows = []
    for b in branches:
        base = random.randrange(90000, 260000, 1000)
        row = [b]
        for m in range(12):
            base = max(60000, base + random.randrange(-15000, 18000, 100))
            row.append(base)
        rows.append(row)
    style_table(ws, start, headers, rows, [14] + [10] * 12,
                num_formats={c: "#,##0" for c in range(2, 14)})
    wb.save(OUT + "/region-month-sales.xlsx")


def gen_grades():
    wb = new_workbook("Marks", "Grade Slabs")
    ws = wb["Marks"]
    put_title(ws, "Student marks — 100 students",
              "Fetch each Grade from the Grade Slabs sheet with an "
              "approximate-match VLOOKUP (range_lookup TRUE), or XLOOKUP "
              "with match_mode -1.")
    names = person_names(100)
    start = 4
    headers = ["Roll No", "Name", "Marks", "Grade"]
    rows = [[101 + i, names[i], random.randrange(10, 101), ""] for i in range(100)]
    style_table(ws, start, headers, rows, [9, 20, 8, 8])
    ws2 = wb["Grade Slabs"]
    put_title(ws2, "Grade slabs — keep sorted ascending",
              "The lookup table for approximate-match VLOOKUP.")
    slabs = [[0, "D"], [40, "C"], [60, "B"], [75, "A"], [90, "A+"]]
    style_table(ws2, 4, ["Marks From", "Grade"], slabs, [12, 8])
    wb.save(OUT + "/student-grades.xlsx")


def gen_invoice():
    wb = new_workbook("Invoice Lines", "Item Master")
    ws = wb["Invoice Lines"]
    put_title(ws, "Invoice lines — 100 lines across 25 invoices",
              "Fetch each Rate from the Item Master sheet with VLOOKUP, compute "
              "Amount = Rate x Qty, then total each invoice with SUMIF.")
    items = ["Cement bag 50kg", "Steel rod 12mm", "Bricks (per 1000)",
             "River sand", "Paint white 4L", "Wire 1.5mm", "PVC pipe 2in",
             "Wall tiles", "Door handle", "Hinge set", "Lock body",
             "Glass pane", "Wood plank 8ft", "Nails 2in", "Putty 1kg"]
    rates = {"Cement bag 50kg": 410, "Steel rod 12mm": 680,
             "Bricks (per 1000)": 7200, "River sand": 55, "Paint white 4L": 1850,
             "Wire 1.5mm": 320, "PVC pipe 2in": 240, "Wall tiles": 950,
             "Door handle": 430, "Hinge set": 120, "Lock body": 640,
             "Glass pane": 810, "Wood plank 8ft": 1150, "Nails 2in": 85,
             "Putty 1kg": 95}
    start = 4
    headers = ["Invoice", "Item", "Qty", "Rate", "Amount"]
    rows = []
    inv = 501
    lines = 0
    while lines < 100:
        n_items = random.choice([2, 3, 4, 5])
        chosen = random.sample(items, n_items)
        for it in chosen:
            if lines >= 100:
                break
            rows.append(["INV-%d" % inv, it,
                         random.choice([1, 2, 3, 4, 5, 6, 8, 10, 12, 15,
                                        20, 25, 40, 50, 60, 100, 120, 150,
                                        200, 250]),
                         "", ""])
            lines += 1
        inv += 1
    style_table(ws, start, headers, rows, [10, 18, 7, 9, 11])
    ws2 = wb["Item Master"]
    put_title(ws2, "Item master — the lookup table")
    style_table(ws2, 4, ["Item", "Rate"],
                [[it, rates[it]] for it in items], [18, 9],
                num_formats={2: "#,##0"})
    wb.save(OUT + "/invoice-items.xlsx")


if __name__ == "__main__":
    import os
    os.makedirs(OUT, exist_ok=True)
    gen_employee()
    gen_products()
    gen_twoway()
    gen_grades()
    gen_invoice()
    print("Written to", OUT)
