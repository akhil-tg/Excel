#!/usr/bin/env python3
"""Generate the Level 4 practice workbooks for Excel Lab.

Creates every .xlsx referenced in the Level 4 lessons' downloads entries
(assets/js/data/level4-1.js and level4-2.js):

  assets/files/level4/sales-data.xlsx       - lessons: charts, pivots, grouping, slicers
  assets/files/level4/kpi-report-data.xlsx  - lesson:  KPI reporting
  assets/files/level4/dashboard-data.xlsx   - lessons: monthly/yearly reporting,
                                              dashboard design, report design

Each workbook has an "Exercise" sheet (starter state) and a "Solution" sheet
(finished state mirroring the exact numbers quoted in the lesson solutions).

Run from the repo root:
    .venv/Scripts/python.exe scripts/gen_level4.py
"""
from collections import defaultdict
from datetime import date
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from openpyxl import load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

from xlcommon import (
    GRAY, TEAL_DARK,
    new_workbook, put_headers, put_note, put_rows, put_section, put_title, set_widths,
)

OUT_DIR = os.path.join("assets", "files", "level4")

# ---------------------------------------------------------------- number formats
RUPEE = '[>=10000000]"₹"##\\,##\\,##\\,##0;[>=100000]"₹"##\\,##\\,##0;"₹"#,##0'
RUPEE2 = '"₹"#,##0.00'
PCT = "0.0%"
PCT_SIGNED = "+0.0%;-0.0%;0.0%"
DATE_FMT = "dd-mmm-yyyy"

_THIN = Side(style="thin", color="cbd5e1")
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

GREEN_FILL = PatternFill("solid", fgColor="c6efce")
GREEN_FONT = Font(color="006100")
AMBER_FILL = PatternFill("solid", fgColor="ffeb9c")
AMBER_FONT = Font(color="9c6500")
RED_FILL = PatternFill("solid", fgColor="ffc7ce")
RED_FONT = Font(color="9c0006")

REGIONS = ["North", "South", "East", "West"]
PRODUCTS = ["Laptops", "Printers", "Monitors", "Staplers", "Pens"]
PEOPLE = ["Arjun Mehta", "Priya Nair", "Rahul Verma", "Kavya Iyer"]
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Region x Product revenue matrix (rupees). These figures reproduce every number
# the lessons quote: North 6,45,000 / South 5,90,000 / East 6,05,000 / West 6,40,000,
# grand total 24,80,000, Laptops 8,20,000 (about 33%) ahead of Printers.
CELL_AMOUNTS = {
    ("North", "Laptops"): 245000, ("North", "Printers"): 155000,
    ("North", "Monitors"): 120000, ("North", "Staplers"): 70000, ("North", "Pens"): 55000,
    ("South", "Laptops"): 175000, ("South", "Printers"): 140000,
    ("South", "Monitors"): 105000, ("South", "Staplers"): 95000, ("South", "Pens"): 75000,
    ("East", "Laptops"): 200000, ("East", "Printers"): 140000,
    ("East", "Monitors"): 120000, ("East", "Staplers"): 85000, ("East", "Pens"): 60000,
    ("West", "Laptops"): 200000, ("West", "Printers"): 125000,
    ("West", "Monitors"): 125000, ("West", "Staplers"): 80000, ("West", "Pens"): 110000,
}
REGION_ROWS = {"North": 95, "South": 52, "East": 51, "West": 52}  # 250 rows total
CELL_ROWS = {
    "North": {"Laptops": 15, "Printers": 18, "Monitors": 16, "Staplers": 23, "Pens": 23},
    "South": {"Laptops": 9, "Printers": 11, "Monitors": 10, "Staplers": 11, "Pens": 11},
    "East": {"Laptops": 9, "Printers": 11, "Monitors": 10, "Staplers": 10, "Pens": 11},
    "West": {"Laptops": 10, "Printers": 10, "Monitors": 10, "Staplers": 11, "Pens": 11},
}
UNIT_RANGES = {"Laptops": (1, 3), "Printers": (1, 4), "Monitors": (1, 3),
               "Staplers": (5, 25), "Pens": (10, 60)}


# ---------------------------------------------------------------- data builders
def split_amount(amount, n, unit=500):
    """Split `amount` into n positive multiples of `unit`, varied deterministically."""
    q, r = divmod(amount, unit)
    assert r == 0 and q >= n, (amount, n)
    q, extra = divmod(q, n)
    vals = [q] * n
    for i in range(extra):
        vals[i % n] += 1
    wig = [0, 2, -2, 4, -4, 1, -1, 3, -3, 2, 0, -2, 4, -4, 1, -1, 3, -3, 2, 0, 1, -1, 2]
    for i in range(n - 1):
        d = wig[i % len(wig)]
        j = (i + 1) % n
        if d > 0 and vals[j] - d >= 1:
            vals[i] += d
            vals[j] -= d
        elif d < 0 and vals[i] + d >= 1:
            vals[i] += d
            vals[j] -= d
    return [v * unit for v in vals]


def units_for(product, i, salt=0):
    lo, hi = UNIT_RANGES[product]
    return lo + (i * 7 + salt) % (hi - lo + 1)


def build_sales_rows():
    """250 Sunrise Traders rows for sales-data.xlsx, exact to the lesson figures."""
    rows = []
    for r_i, region in enumerate(REGIONS):
        region_rows = []
        for p_i, product in enumerate(PRODUCTS):
            count = CELL_ROWS[region][product]
            revenues = split_amount(CELL_AMOUNTS[(region, product)], count)
            for i, revenue in enumerate(revenues):
                person = PEOPLE[(p_i * 2 + i + r_i) % 4]
                units = units_for(product, i, salt=r_i)
                region_rows.append([None, region, person, product, units, revenue])
        for i, row in enumerate(region_rows):
            month = (i * 5 + r_i) % 12 + 1          # stride 5 coprime to 12: all 12 months
            day = 1 + (i * 7 + r_i * 3) % 28
            row[0] = date(2025, month, day)
        rows.extend(region_rows)
    assert len(rows) == 250
    rows.sort(key=lambda r: (r[0], r[1]))
    return rows


def aggregate(rows, month_of=None):
    """Aggregate row tuples (date, region, person, product, units, revenue)."""
    agg = {
        "region_rev": defaultdict(int), "region_units": defaultdict(int),
        "month_rev": defaultdict(int), "product_rev": defaultdict(int),
        "person_rev": defaultdict(int), "person_cnt": defaultdict(int),
        "group_rev": defaultdict(int), "total": 0,
    }
    for d, region, person, product, units, revenue in rows:
        m = month_of(d) if month_of else d.month
        agg["region_rev"][region] += revenue
        agg["region_units"][region] += units
        agg["month_rev"][m] += revenue
        agg["product_rev"][product] += revenue
        agg["person_rev"][person] += revenue
        agg["person_cnt"][person] += 1
        if units <= 9:
            b = "0-9"
        elif units <= 19:
            b = "10-19"
        elif units <= 29:
            b = "20-29"
        elif units <= 39:
            b = "30-39"
        elif units <= 49:
            b = "40-49"
        else:
            b = "50+"
        agg["group_rev"][b] += revenue
        agg["total"] += revenue
    return agg


def build_dashboard_rows():
    """312 rows (26 per month, all 12 months of 2025) for dashboard-data.xlsx."""
    factor = [0.82, 0.86, 0.94, 1.00, 1.06, 0.90, 0.96, 1.02, 1.10, 1.16, 1.24, 1.30]
    rows = []
    for idx in range(312):
        month = idx // 26 + 1
        day = 1 + (idx * 7) % 28
        region = REGIONS[idx % 4]
        product = PRODUCTS[(idx // 4) % 5]
        person = PEOPLE[(idx // 2) % 4]
        base = 14000 + ((idx * 5300) % 12000)
        revenue = int(round(base * factor[month - 1] / 250.0)) * 250
        units = units_for(product, idx, salt=idx % 4)
        rows.append([date(2025, month, day), region, person, product, units, revenue])
    assert len(rows) == 312
    rows.sort(key=lambda r: (r[0], r[1]))
    return rows


# ---------------------------------------------------------------- sheet helpers
def total_row(ws, row, values, num_formats=None):
    """Bold total row with borders."""
    for i, v in enumerate(values, 1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = Font(bold=True, color=TEAL_DARK)
        c.border = BORDER
        if num_formats and i in num_formats:
            c.number_format = num_formats[i]
        if isinstance(v, str) and v.startswith("="):
            c.alignment = Alignment(horizontal="right")
    return row + 1


def right_align_formulas(ws, r1, r2, c1, c2):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.alignment = Alignment(horizontal="right")


def note(ws, row, text, height=30):
    put_note(ws, row, text)
    ws.row_dimensions[row].height = height
    return row + 1


def text_cell(ws, row, col, text):
    """Store a string that starts with '=' as text, not as a formula."""
    c = ws.cell(row=row, column=col, value=text)
    c.data_type = "s"
    c.font = Font(italic=True, color=GRAY)
    return c


def add_table(ws, name, first_row, last_row, last_col="F"):
    table = Table(displayName=name, ref="A%d:%s%d" % (first_row, last_col, last_row))
    table.tableStyleInfo = TableStyleInfo(name="TableStyleLight15", showRowStripes=True)
    ws.add_table(table)


# ---------------------------------------------------------------- sales-data.xlsx
SALES_NOTE = ("Exercise — the 250-row table below is the Excel Table 'SalesTable' "
              "(Date, Region, Salesperson, Product, Units, Revenue). "
              "1) Build a region summary in H1:I5 with "
              "=SUMIFS(SalesTable[Revenue], SalesTable[Region], H2) and chart it as a "
              "clustered column titled 'Revenue by region, 2025'. "
              "2) Add a Month column with =MONTH([@Date]), build a 12-row month summary "
              "and chart revenue by month as a line. "
              "3) Add a Units column to the summary and make a combo chart "
              "(columns + line on the secondary axis). "
              "4) Build PivotTables from SalesTable (Region, Product, Month), group dates "
              "by Months and Years, try Show Values As (% of Grand Total, Running Total), "
              "add a calculated field =Revenue/Units, and practise slicers, a timeline and "
              "Report Connections. Check your numbers against the Solution sheet.")


def build_sales_data(agg):
    """sales-data.xlsx for the Charts, PivotTables, Grouping and Slicers lessons."""
    rows = build_sales_rows()
    sales_agg = aggregate(rows)
    checks = {}

    # sanity: the dataset must reproduce every figure the lessons quote
    assert sales_agg["region_rev"]["North"] == 645000
    assert sales_agg["region_rev"]["South"] == 590000
    assert sales_agg["region_rev"]["East"] == 605000
    assert sales_agg["region_rev"]["West"] == 640000
    assert sales_agg["total"] == 2480000
    assert sales_agg["product_rev"]["Laptops"] == 820000

    wb = new_workbook("Exercise", "Solution")

    # --- Exercise sheet -------------------------------------------------------
    ws = wb["Exercise"]
    put_title(ws, "Sunrise Traders — Sales Data 2025",
              "Level 4 practice workbook: charts, PivotTables, grouping, slicers")
    note(ws, 3, SALES_NOTE, height=90)
    headers = ["Date", "Region", "Salesperson", "Product", "Units", "Revenue"]
    put_headers(ws, 5, headers)
    put_rows(ws, 6, rows, num_formats={6: RUPEE})
    for r in range(6, 6 + len(rows)):
        ws.cell(row=r, column=1).number_format = DATE_FMT
    add_table(ws, "SalesTable", 5, 5 + len(rows))
    set_widths(ws, [12, 10, 16, 12, 9, 14, 4, 10, 12])

    # --- Solution sheet -------------------------------------------------------
    ws = wb["Solution"]
    put_title(ws, "Solution — Sunrise Traders Sales Data 2025",
              "Finished state for every Level 4 exercise built on sales-data.xlsx")

    r = 4
    put_section(ws, r, "1. Region summary (Charts / PivotTables lessons — SUMIFS in H1:I5)")
    r += 1
    hdr = r
    put_headers(ws, r, ["Region", "Revenue", "Units", "% of Grand Total",
                        "Running Total", "Formula used in the exercise"])
    r += 1
    first = r
    region_rows = []
    for region in REGIONS:
        region_rows.append([region, sales_agg["region_rev"][region],
                            sales_agg["region_units"][region], None, None, None])
        r += 1
    r = put_rows(ws, first, region_rows, num_formats={2: RUPEE, 4: PCT, 5: RUPEE})
    for row in range(first, first + 4):
        text_cell(ws, row, 6, "=SUMIFS(SalesTable[Revenue], SalesTable[Region], H%d)"
                  % (row - hdr + 1))
    for i, row in enumerate(range(first, first + 4)):
        ws.cell(row=row, column=4, value="=B%d/$B$%d" % (row, first + 4))
        ws.cell(row=row, column=5, value="=SUM($B$%d:B%d)" % (first, row))
        ws.cell(row=row, column=4).number_format = PCT
        ws.cell(row=row, column=5).number_format = RUPEE
    checks["north_rev"] = ("Solution", "B%d" % first, 645000)
    checks["south_rev"] = ("Solution", "B%d" % (first + 1), 590000)
    checks["east_rev"] = ("Solution", "B%d" % (first + 2), 605000)
    checks["west_rev"] = ("Solution", "B%d" % (first + 3), 640000)
    total_row(ws, r, ["Grand Total", "=SUM(B%d:B%d)" % (first, first + 3),
                      "=SUM(C%d:C%d)" % (first, first + 3), 1.0,
                      "=B%d" % (first + 4), None], num_formats={2: RUPEE, 4: PCT, 5: RUPEE})
    right_align_formulas(ws, first, r + 1, 2, 5)
    r += 2

    put_section(ws, r, "2. Monthly revenue 2025 (line chart, dates grouped by month)")
    r += 1
    put_headers(ws, r, ["Month no.", "Month", "Revenue", "", "",
                        "Formula used in the exercise"])
    r += 1
    first = r
    month_rows = [[m, MONTHS[m - 1], sales_agg["month_rev"][m], None, None, None]
                  for m in range(1, 13)]
    r = put_rows(ws, first, month_rows, num_formats={3: RUPEE})
    for row in range(first, first + 12):
        text_cell(ws, row, 6, "=SUMIFS(SalesTable[Revenue], SalesTable[Month], %d)" % (row - first + 1))
    checks["jan_rev"] = ("Solution", "C%d" % first, sales_agg["month_rev"][1])
    checks["dec_rev"] = ("Solution", "C%d" % (first + 11), sales_agg["month_rev"][12])
    total_row(ws, r, ["Total", None, "=SUM(C%d:C%d)" % (first, first + 11), None, None, None],
              num_formats={3: RUPEE})
    right_align_formulas(ws, first, r, 3, 3)
    r += 2

    put_section(ws, r, "3. Revenue by product (pie / Show Values As % of Grand Total)")
    r += 1
    put_headers(ws, r, ["Product", "Revenue", "% of Grand Total"])
    r += 1
    first = r
    prod_rows = []
    for product in sorted(PRODUCTS, key=lambda p: -sales_agg["product_rev"][p]):
        prod_rows.append([product, sales_agg["product_rev"][product], None])
        r += 1
    r = put_rows(ws, first, prod_rows, num_formats={2: RUPEE, 3: PCT})
    for row in range(first, first + len(prod_rows)):
        ws.cell(row=row, column=3, value="=B%d/$B$%d" % (row, first + len(prod_rows)))
        ws.cell(row=row, column=3).number_format = PCT
    checks["laptops_rev"] = ("Solution", "B%d" % first, 820000)
    total_row(ws, r, ["Total", "=SUM(B%d:B%d)" % (first, r - 1),
                      "=SUM(C%d:C%d)" % (first, r - 1)], num_formats={2: RUPEE, 3: PCT})
    right_align_formulas(ws, first, r, 2, 3)
    r += 2

    put_section(ws, r, "4. Revenue by region and product (pivot cross-tab — "
                       "Laptops lead every region)")
    r += 1
    put_headers(ws, r, ["Region"] + PRODUCTS + ["Total"])
    r += 1
    first = r
    cross = []
    for region in REGIONS:
        cross.append([region] + [CELL_AMOUNTS[(region, p)] for p in PRODUCTS] + [None])
        r += 1
    r = put_rows(ws, first, cross, num_formats={c: RUPEE for c in range(2, 8)})
    for row in range(first, first + 4):
        ws.cell(row=row, column=7, value="=SUM(B%d:F%d)" % (row, row))
        ws.cell(row=row, column=7).number_format = RUPEE
    total_row(ws, r, ["Grand Total"] +
              ["=SUM(%s%d:%s%d)" % (chr(65 + c), first, chr(65 + c), first + 3)
               for c in range(1, 7)],
              num_formats={c: RUPEE for c in range(2, 8)})
    right_align_formulas(ws, first, r, 2, 7)
    r += 2

    put_section(ws, r, "5. Revenue by salesperson (value filters / Count of transactions)")
    r += 1
    put_headers(ws, r, ["Salesperson", "Revenue", "Transactions"])
    r += 1
    first = r
    people = sorted(PEOPLE, key=lambda p: -sales_agg["person_rev"][p])
    person_rows = [[p, sales_agg["person_rev"][p], sales_agg["person_cnt"][p]] for p in people]
    r = put_rows(ws, first, person_rows, num_formats={2: RUPEE})
    r += 1

    put_section(ws, r, "6. Calculated field 'Avg Price' = Revenue/Units, by region")
    r += 1
    put_headers(ws, r, ["Region", "Revenue", "Units", "Avg Price"])
    r += 1
    first = r
    price_rows = []
    for region in REGIONS:
        price_rows.append([region, sales_agg["region_rev"][region],
                           sales_agg["region_units"][region], None])
        r += 1
    r = put_rows(ws, first, price_rows, num_formats={2: RUPEE, 4: RUPEE2})
    for row in range(first, first + 4):
        ws.cell(row=row, column=4, value="=B%d/C%d" % (row, row))
        ws.cell(row=row, column=4).number_format = RUPEE2
    right_align_formulas(ws, first, r - 1, 2, 4)
    r += 1

    put_section(ws, r, "7. Revenue by order-size group (Units grouped into ranges)")
    r += 1
    put_headers(ws, r, ["Units group", "Revenue"])
    r += 1
    first = r
    labels = ["0-9", "10-19", "20-29", "30-39", "40-49", "50+"]
    group_rows = [[lab, sales_agg["group_rev"][lab]] for lab in labels]
    r = put_rows(ws, first, group_rows, num_formats={2: RUPEE})
    total_row(ws, r, ["Total", "=SUM(B%d:B%d)" % (first, r - 1)], num_formats={2: RUPEE})
    r += 2
    note(ws, r, "Check: North + East = 6,05,000 + 6,45,000 = 12,50,000 — the total the "
                "slicer exercise expects when North and East are both selected.", height=30)

    set_widths(ws, [14, 14, 13, 15, 14, 46])
    ws.freeze_panes = None
    path = os.path.join(OUT_DIR, "sales-data.xlsx")
    wb.save(path)
    return path, checks, sales_agg


# ------------------------------------------------------------ kpi-report-data.xlsx
KPI_ACTUALS = {
    "Jan": 542000, "Feb": 486000, "Mar": 561000, "Apr": 528000,
    "May": 483500, "Jun": 438000, "Jul": 515000, "Aug": 469000,
    "Sep": 537000, "Oct": 445000, "Nov": 573000, "Dec": 610000,
}
KPI_TARGET = 500000
KPI_HEADERS = ["Month", "Target", "Actual", "Achievement %", "MoM %", "Status", "Trend"]

KPI_TASKS = [
    "1) Achievement %: type =C2/B2 in D2, format as a percentage with one decimal "
    "(Ctrl + 1), and fill down to D13.",
    "2) Conditional formatting on D2:D13 — three rules in this order: cell value >= 1 "
    "green; between 0.9 and 0.9999 amber; less than 0.9 red.",
    "3) MoM growth: in E3 use =IFERROR((C3-C2)/C2,\"\") and fill down; leave E2 as a "
    "dash. Custom-format the column +0.0%;-0.0%;0.0%.",
    "4) Trend: in G2 insert a Line sparkline over $C$2:C2, fill down to G13 and mark "
    "the High Point from the Sparkline tab.",
]


def kpi_status(achievement):
    if achievement >= 1:
        return "Green"
    if achievement >= 0.9:
        return "Amber"
    return "Red"


def build_kpi_data():
    """kpi-report-data.xlsx for the KPI Reporting lesson (months in rows 2-13)."""
    checks = {}

    wb = new_workbook("Exercise", "Solution")

    # --- Exercise sheet (lesson pins months to rows 2-13, headers row 1) -------
    ws = wb["Exercise"]
    put_headers(ws, 1, KPI_HEADERS)
    ws.freeze_panes = "A2"
    ex_rows = [[m, KPI_TARGET, KPI_ACTUALS[m], None, None, None, None] for m in MONTHS]
    put_rows(ws, 2, ex_rows, num_formats={2: RUPEE, 3: RUPEE})
    r = 15
    put_section(ws, r, "Your tasks — build the KPI page (compare with the Solution sheet)")
    r += 1
    for task in KPI_TASKS:
        r = note(ws, r, task, height=28)
    note(ws, r, "Status bands: green 100% or more, amber 90%-99.9%, red below 90%. "
                "The Status column should come from conditional formatting on the "
                "Achievement % column, never typed by hand.", height=30)
    set_widths(ws, [10, 13, 13, 14, 11, 10, 14])

    # --- Solution sheet --------------------------------------------------------
    ws = wb["Solution"]
    put_headers(ws, 1, KPI_HEADERS)
    ws.freeze_panes = "A2"
    first = 2
    sol_rows = []
    for i, m in enumerate(MONTHS):
        row = first + i
        actual = KPI_ACTUALS[m]
        achievement = actual / KPI_TARGET
        mom = None if i == 0 else (actual - KPI_ACTUALS[MONTHS[i - 1]]) / KPI_ACTUALS[MONTHS[i - 1]]
        sol_rows.append([m, KPI_TARGET, actual,
                         "=C%d/B%d" % (row, row),
                         "-" if i == 0 else '=IFERROR((C%d-C%d)/C%d,"")' % (row, row - 1, row - 1),
                         kpi_status(achievement), "[sparkline]"])
    put_rows(ws, first, sol_rows, num_formats={2: RUPEE, 3: RUPEE, 4: PCT, 5: PCT_SIGNED})
    last = first + 11
    for i, m in enumerate(MONTHS):
        row = first + i
        status = kpi_status(KPI_ACTUALS[m] / KPI_TARGET)
        cell = ws.cell(row=row, column=6)
        fill, font = {"Green": (GREEN_FILL, GREEN_FONT), "Amber": (AMBER_FILL, AMBER_FONT),
                      "Red": (RED_FILL, RED_FONT)}[status]
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=7).font = Font(italic=True, color=GRAY)
        if m in ("Jan", "Feb", "Jun"):
            checks["kpi_actual_" + m.lower()] = ("Solution", "C%d" % row, KPI_ACTUALS[m])
    checks["kpi_achievement_jan"] = ("Solution", "D%d" % first, "=C2/B2")
    checks["kpi_mom_feb"] = ("Solution", "E%d" % (first + 1), '=IFERROR((C3-C2)/C2,"")')

    # conditional formatting exactly as the lesson builds it (rules in order)
    rng = "D%d:D%d" % (first, last)
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="greaterThanOrEqual", formula=["1"], fill=GREEN_FILL, font=GREEN_FONT))
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="between", formula=["0.9", "0.9999"], fill=AMBER_FILL, font=AMBER_FONT))
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="lessThan", formula=["0.9"], fill=RED_FILL, font=RED_FONT))

    r = last + 2
    put_section(ws, r, "How this was built (Solution notes)")
    r += 1
    r = note(ws, r, "Achievement % = Actual/Target (=C2/B2), percentage format with one "
                    "decimal: Jan 108.4%, Feb 97.2%, Jun 87.6% — the exact figures in the "
                    "lesson. Variance in rupees would be =(C2-B2), a different question.", height=30)
    r = note(ws, r, "MoM % compares each month with the one before: Feb (486-542)/542 = "
                    "-10.3%, Jun (438-483.5)/483.5 = -9.4%. IFERROR keeps the sheet free "
                    "of #DIV/0!.", height=30)
    r = note(ws, r, "Status colors come from the three conditional-formatting rules on "
                    "D2:D13 (green >= 100%, amber 90-99.9%, red < 90%), so the page "
                    "recolors itself when next month's numbers arrive.", height=30)
    note(ws, r, "Trend column: Line sparklines over the anchored range $C$2:C2, filled "
                "down to G13, with High Point markers on. (Sparklines cannot be written "
                "by a generator — insert them per the lesson steps.)", height=30)
    set_widths(ws, [10, 13, 13, 14, 11, 10, 14])

    path = os.path.join(OUT_DIR, "kpi-report-data.xlsx")
    wb.save(path)
    return path, checks


# ------------------------------------------------------------ dashboard-data.xlsx
DASH_NOTE = ("Exercise — the raw 2025 transactions (312 rows) live on the Data sheet as "
             "the Excel Table 'DataTable' (Date, Region, Salesperson, Product, Units, "
             "Revenue). Keep raw data untouched on its own sheet. "
             "1) Add a Month helper column beside the table with =MONTH([@Date]), then "
             "fill this 12-row summary below: Revenue =SUMIFS(DataTable[Revenue], "
             "DataTable[Month], A2), MoM % =IFERROR((C3-C2)/C2,\"\") from row 3, "
             "YTD =SUM(C$2:C2) filled down. "
             "2) November region breakdown: =SUMIFS(DataTable[Revenue], DataTable[Month], "
             "11, DataTable[Region], region). "
             "3) Rename sheets Data / Report / Checks, add the title block "
             "('Sunrise Traders — Monthly Sales Report', 'Period: 1-30 November 2025', "
             "prepared-by line), then style, freeze panes, set print titles, scale to fit "
             "and export a PDF. "
             "4) On a Dashboard sheet build four KPI cards, two pivots with PivotCharts, "
             "slicers and a timeline; compare everything with the Solution sheet.")


def build_dashboard_data(rows, dash_agg):
    """dashboard-data.xlsx for the Monthly Reporting, Dashboard and Report Design lessons."""
    checks = {}
    nov_region = defaultdict(int)
    for d, region, person, product, units, revenue in rows:
        if d.month == 11:
            nov_region[region] += revenue

    wb = new_workbook("Exercise", "Data", "Solution", "Checks")

    # --- Exercise sheet (detail block pinned to rows 1-13, as in the lesson) ---
    ws = wb["Exercise"]
    put_headers(ws, 1, ["Month no.", "Month", "Revenue", "MoM %", "YTD"])
    ws.freeze_panes = "A2"
    ex_rows = [[m, MONTHS[m - 1], None, None, None] for m in range(1, 13)]
    put_rows(ws, 2, ex_rows)
    r = 15
    put_section(ws, r, "Your tasks — reporting, dashboard and report design")
    r += 1
    r = note(ws, r, "Fill C2:C13 with =SUMIFS(DataTable[Revenue], DataTable[Month], A2) "
                    "(add the Month helper column =MONTH([@Date]) beside DataTable first), "
                    "D3:D13 with =IFERROR((C3-C2)/C2,\"\"), and E2:E13 with =SUM(C$2:C2). "
                    "Cross-check on the Checks sheet.", height=44)
    put_section(ws, r, "November region breakdown (two-criteria SUMIFS)")
    r += 1
    put_headers(ws, r, ["Region", "Revenue (Nov)"])
    r += 1
    put_rows(ws, r, [[reg, None] for reg in REGIONS])
    r += 6
    put_section(ws, r, "KPI cards for the Dashboard sheet (labels above, formulas below)")
    r += 1
    put_headers(ws, r, ["Card", "Formula to use"])
    r += 1
    put_rows(ws, r, [
        ["TOTAL REVENUE 2025", None],
        ["BEST MONTH", None],
        ["AVG MONTH", None],
        ["MoM GROWTH (LATEST)", None],
    ])
    for i, formula in enumerate([
        "=SUM(DataTable[Revenue])",
        "=MAX(Report!C2:C13)",
        "=AVERAGE(Report!C2:C13)",
        "=Report!D13",
    ]):
        text_cell(ws, r + i, 2, formula)
    r += 6
    note(ws, r, "KPI cards are merged cells: small gray label above, big bold number "
                "below, always a live formula. On the Solution sheet the cards sit on "
                "the report page itself, so the formulas read =MAX(C2:C13) etc.; on a "
                "separate Dashboard sheet they point at Report!C2:C13.", height=44)
    note(ws, r + 1, "Title block: 16 pt bold title, period line 'Period: 1-30 November "
                    "2025', 'Prepared by: Arjun Mehta', source and status lines. Then "
                    "freeze panes, print titles $1:$1, scale to fit 1 page wide, page "
                    "numbers in the footer, and export 'Sunrise Monthly Report "
                    "2025-11.pdf'.", height=44)
    set_widths(ws, [10, 12, 14, 11, 12, 4, 22, 30])

    # --- Data sheet: raw DataTable --------------------------------------------
    ws = wb["Data"]
    put_title(ws, "Raw sales data — DataTable",
              "Source of truth: never edit these rows; add the Month helper column "
              "beside the table with =MONTH([@Date])")
    put_headers(ws, 4, ["Date", "Region", "Salesperson", "Product", "Units", "Revenue"])
    put_rows(ws, 5, rows, num_formats={6: RUPEE})
    for rr in range(5, 5 + len(rows)):
        ws.cell(row=rr, column=1).number_format = DATE_FMT
    add_table(ws, "DataTable", 4, 4 + len(rows))
    set_widths(ws, [12, 10, 16, 12, 9, 14])

    # --- Solution sheet: finished Report page ----------------------------------
    ws = wb["Solution"]
    put_headers(ws, 1, ["Month no.", "Month", "Revenue", "MoM %", "YTD"])
    ws.freeze_panes = "A2"
    first = 2
    sol_rows = []
    for m in range(1, 13):
        row = first + m - 1
        sol_rows.append([m, MONTHS[m - 1], dash_agg["month_rev"][m],
                         "-" if m == 1 else '=IFERROR((C%d-C%d)/C%d,"")' % (row, row - 1, row - 1),
                         "=SUM(C$2:C%d)" % row])
    put_rows(ws, first, sol_rows, num_formats={3: RUPEE, 4: PCT_SIGNED, 5: RUPEE})
    last = first + 11
    checks["dash_jan_rev"] = ("Solution", "C%d" % first, dash_agg["month_rev"][1])
    checks["dash_dec_rev"] = ("Solution", "C%d" % last, dash_agg["month_rev"][12])
    total_row(ws, last + 1, ["Total", None, "=SUM(C%d:C%d)" % (first, last), None, "=C%d" % last],
              num_formats={3: RUPEE, 5: RUPEE})
    right_align_formulas(ws, first, last + 1, 4, 5)

    r = last + 3
    put_section(ws, r, "Yearly summary line (the annual picture)")
    r += 1
    put_headers(ws, r, ["Metric", "Value"])
    r += 1
    put_rows(ws, r, [
        ["Total revenue", "=SUM(C%d:C%d)" % (first, last)],
        ["Best month", "=MAX(C%d:C%d)" % (first, last)],
        ["Average month", "=AVERAGE(C%d:C%d)" % (first, last)],
    ], num_formats={2: RUPEE})
    checks["dash_total_formula"] = ("Solution", "B%d" % r, "=SUM(C%d:C%d)" % (first, last))
    right_align_formulas(ws, r, r + 2, 2, 2)
    r += 4

    put_section(ws, r, "Title block (Report Design lesson)")
    r += 1
    c = ws.cell(row=r, column=1, value="Sunrise Traders — Monthly Sales Report")
    c.font = Font(bold=True, size=16, color=TEAL_DARK)
    r += 1
    put_rows(ws, r, [
        ["Period: 1-30 November 2025", "Prepared by: Arjun Mehta"],
        ["Source: DataTable, refreshed 01-Dec-2025", "Status: Final"],
    ])
    r += 3

    put_section(ws, r, "KPI cards (Dashboard lesson — live formulas, never typed numbers)")
    r += 1
    put_headers(ws, r, ["Card", "Value"])
    r += 1
    put_rows(ws, r, [
        ["TOTAL REVENUE 2025", "=SUM(DataTable[Revenue])"],
        ["BEST MONTH", "=MAX(C%d:C%d)" % (first, last)],
        ["AVG MONTH", "=AVERAGE(C%d:C%d)" % (first, last)],
        ["MoM GROWTH (LATEST)", "=D%d" % last],
    ], num_formats={2: RUPEE})
    ws.cell(row=r + 3, column=2).number_format = PCT_SIGNED
    right_align_formulas(ws, r, r + 3, 2, 2)
    r += 5

    put_section(ws, r, "November region breakdown (two-criteria SUMIFS)")
    r += 1
    put_headers(ws, r, ["Region", "Revenue (Nov)"])
    r += 1
    first_nov = r
    nov_rows = [[reg, nov_region[reg]] for reg in REGIONS]
    r = put_rows(ws, first_nov, nov_rows, num_formats={2: RUPEE})
    checks["dash_nov_total"] = ("Solution", "B%d" % first_nov, nov_region["North"])
    total_row(ws, r, ["Total", "=SUM(B%d:B%d)" % (first_nov, r - 1)], num_formats={2: RUPEE})
    right_align_formulas(ws, first_nov, r, 2, 2)
    r += 2
    note(ws, r, "The November total must equal November's row in the monthly summary "
                "above. Format revenue as currency with no decimals and MoM with the "
                "custom format +0.0%;-0.0%;0.0% before delivering.", height=30)
    set_widths(ws, [12, 26, 14, 12, 12])

    # --- Checks sheet -----------------------------------------------------------
    ws = wb["Checks"]
    put_title(ws, "Checks — sanity totals before delivery",
              "Both values below must match; if they do not, the report and the data disagree.")
    put_headers(ws, 4, ["Check", "Value"])
    put_rows(ws, 5, [
        ["Raw DataTable total", "=SUM(DataTable[Revenue])"],
        ["Report month rows total", "=SUM(Solution!C%d:C%d)" % (first, last)],
        ["Difference (must be 0)", "=B5-B6"],
    ], num_formats={2: RUPEE})
    right_align_formulas(ws, 5, 7, 2, 2)
    set_widths(ws, [28, 16])

    path = os.path.join(OUT_DIR, "dashboard-data.xlsx")
    wb.save(path)
    return path, checks


# ---------------------------------------------------------------- verification
def verify(paths, all_checks):
    ok = True
    for path in paths:
        assert os.path.exists(path), path
        wb = load_workbook(path)
        assert "Exercise" in wb.sheetnames, (path, wb.sheetnames)
        assert "Solution" in wb.sheetnames, (path, wb.sheetnames)
        # opens cleanly and every sheet is readable
        for name in wb.sheetnames:
            wb[name]["A1"].value
        checks = all_checks[path]
        for label, (sheet, coord, expected) in sorted(checks.items()):
            got = wb[sheet][coord].value
            if got != expected:
                print("FAIL %s: %s at %s!%s expected %r got %r"
                      % (os.path.basename(path), label, sheet, coord, expected, got))
                ok = False
        if path.endswith("sales-data.xlsx"):
            assert "SalesTable" in wb["Exercise"].tables
        if path.endswith("dashboard-data.xlsx"):
            assert "DataTable" in wb["Data"].tables
        print("OK  %-22s sheets=%s size=%d bytes"
              % (os.path.basename(path), ",".join(wb.sheetnames), os.path.getsize(path)))
    return ok


def main():
    os.makedirs(OUT_DIR, exist_ok=True)

    sales_path, sales_checks, _ = build_sales_data(None)
    kpi_path, kpi_checks = build_kpi_data()
    dash_rows = build_dashboard_rows()
    dash_agg = aggregate(dash_rows)
    dash_path, dash_checks = build_dashboard_data(dash_rows, dash_agg)

    paths = [sales_path, kpi_path, dash_path]
    all_checks = {p: c for p, c in zip(paths, [sales_checks, kpi_checks, dash_checks])}
    if not verify(paths, all_checks):
        sys.exit(1)
    print("All Level 4 workbooks generated and verified.")


if __name__ == "__main__":
    main()
