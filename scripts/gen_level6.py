"""Generate every practice file referenced by the Level 6 lessons.

Run from the repo root:

    .venv/Scripts/python.exe scripts/gen_level6.py

Creates exactly the files referenced in the Level 6 downloads entries:

  assets/files/level6/macro-practice.xlsx          (3 lessons: Recording Macros,
                                                    VBA Fundamentals, Automating Reports)
  assets/files/level6/power-pivot-data.xlsx        (2 lessons: Power Pivot and the Data Model,
                                                    Introduction to DAX)
  assets/files/level6/power-query-messy-data.xlsx  (1 lesson: Power Query - deliberately messy)
  assets/files/level6/macro-code.bas               (2 macro lessons: ready-to-import VBA module)

Every workbook follows the repo convention: an "Exercise" sheet (starter
state) and a "Solution" sheet (finished state, matching the numbers stated
in the lessons).
"""
import datetime as dt
from pathlib import Path

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from xlcommon import (
    GRAY,
    TEAL_DARK,
    blank_solution,
    new_workbook,
    put_headers,
    put_note,
    put_rows,
    put_section,
    put_title,
    set_widths,
)

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "assets" / "files" / "level6"

# Colours used by the lesson macros themselves
AMBER_HEADER = "FFD966"   # RGB(255, 217, 102) - what FormatReport applies to A1:E1
LAPTOP_BAND = "FFEBCD"    # RGB(255, 235, 205) - what MonthlyReport highlights laptop rows with

_THIN = Side(style="thin", color="cbd5e1")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

DDMMMYYYY = "dd-mmm-yyyy"
NUM0 = "#,##0"


# --------------------------------------------------------------------------
# small local helpers (built on top of xlcommon)
# --------------------------------------------------------------------------
def plain_table(ws, row, headers, rows, num_formats=None):
    """Write an UNstyled header row + bordered data rows (raw export look)."""
    for i, h in enumerate(headers, 1):
        ws.cell(row=row, column=i, value=h)
    row = put_rows(ws, row + 1, rows, num_formats=num_formats)
    return row


def style_row(ws, row, col_start, col_end, fill=None, bold=False, center=False,
              num_fmts=None):
    """Apply the formatting a recorded macro would produce."""
    for col in range(col_start, col_end + 1):
        c = ws.cell(row=row, column=col)
        if bold:
            c.font = Font(bold=True)
        if fill:
            c.fill = PatternFill("solid", fgColor=fill)
        if center:
            c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = _BORDER
        if num_fmts and col in num_fmts:
            c.number_format = num_fmts[col]


# ==========================================================================
# 1. macro-practice.xlsx  (serves l6-record-macros, l6-vba-fundamentals,
#                          l6-vba-automation)
# ==========================================================================
# Daily sales report from the "Recording Macros" lesson grid
DAILY_REPORT = [
    [dt.date(2026, 11, 2), "Arjun Mehta", "Staplers", 12, 2520],
    [dt.date(2026, 11, 2), "Priya Nair", "Notebooks", 48, 5760],
    [dt.date(2026, 11, 3), "Rahul Verma", "Pens", 100, 3500],
]

# Sales sheet for the automation lesson (A=Date, B=Salesperson, C=Category, D=Amount)
SALES_ROWS = [
    (dt.date(2026, 11, 2), "Arjun Mehta", "Laptops", 52000),
    (dt.date(2026, 11, 2), "Priya Nair", "Printers", 18000),
    (dt.date(2026, 11, 3), "Rahul Verma", "Stationery", 6200),
    (dt.date(2026, 11, 3), "Sneha Iyer", "Laptops", 46500),
    (dt.date(2026, 11, 4), "Arjun Mehta", "Stationery", 8900),
    (dt.date(2026, 11, 4), "Priya Nair", "Laptops", 38000),
    (dt.date(2026, 11, 5), "Rahul Verma", "Printers", 24000),
    (dt.date(2026, 11, 5), "Sneha Iyer", "Stationery", 5400),
    (dt.date(2026, 11, 6), "Arjun Mehta", "Laptops", 61500),
    (dt.date(2026, 11, 6), "Priya Nair", "Stationery", 7300),
    (dt.date(2026, 11, 7), "Rahul Verma", "Laptops", 29500),
    (dt.date(2026, 11, 7), "Sneha Iyer", "Printers", 20500),
    (dt.date(2026, 11, 10), "Arjun Mehta", "Pens", 1850),
    (dt.date(2026, 11, 10), "Priya Nair", "Pens", 2400),
]
TOTAL_SALES = sum(r[3] for r in SALES_ROWS)                      # 322050
LAPTOP_SALES = sum(r[3] for r in SALES_ROWS if r[2] == "Laptops")  # 227500
PEN_COUNT = sum(1 for r in SALES_ROWS if r[2] == "Pens")          # 2
LAST_ROW = 1 + len(SALES_ROWS)                                   # 15 (header is row 1)


def build_macro_practice():
    wb = new_workbook("Exercise", "Sales", "Solution")

    # ---------------- Exercise sheet ----------------
    # The macro lessons address A1:E1 and rows 2..lastRow literally, so the
    # daily report header MUST sit in row 1; title/instructions go below it.
    ws = wb["Exercise"]
    r = plain_table(ws, 1, ["Date", "Salesperson", "Product", "Qty", "Amount (₹)"],
                    DAILY_REPORT, num_formats={1: DDMMMYYYY, 5: NUM0})
    r += 2

    r = put_title(ws, "Macro Practice - Sunrise Traders",
                  "Level 6 - Recording Macros / VBA Fundamentals / Automating Reports with VBA",
                  row=r)
    r = put_note(ws, r,
        "Re-save this workbook as .xlsm before adding macros (File > Save As > "
        "Excel Macro-Enabled Workbook). Recording Macros: select A1, Developer > "
        "Record Macro, name it FormatReport, then apply bold, a fill colour and "
        "centre alignment to A1:E1, fit the column widths, freeze the top row and "
        "return to A1 before stopping. VBA Fundamentals: run FillReportTitle, "
        "CellTest and OrderSummary from the VBA editor (Alt+F11). Automating "
        "Reports: the Sales sheet holds the month's data for LoopTest, the "
        "Pen-count macro and MonthlyReport.", width=6)
    r += 1

    r = put_section(ws, r, "CellTest check area (VBA Fundamentals)")
    r = put_note(ws, r,
        "Run the CellTest macro: Range(\"C2\").Value and Cells(2, 3).Value both "
        "write to the same cell. Cell C2 should read Staplers afterwards.", width=6)
    r += 1

    r = put_section(ws, r, "Sales data for the automation macros")
    r = put_note(ws, r,
        "The Sales sheet uses the layout from the lessons: header in row 1, "
        "A = Date, B = Salesperson, C = Category, D = Amount (₹). Run LoopTest, "
        "the Pen-count macro, WriteSummary and MonthlyReport on it.", width=6)
    set_widths(ws, [12, 16, 14, 8, 12])

    # ---------------- Sales sheet ----------------
    # Header in row 1, data from row 2: exactly the layout MonthlyReport's
    # For r = 2 To lastRow loop (and End(xlUp)) expects.
    ws = wb["Sales"]
    r = put_headers(ws, 1, ["Date", "Salesperson", "Category", "Amount (₹)"])
    r = put_rows(ws, r, [list(row) for row in SALES_ROWS],
                 num_formats={1: DDMMMYYYY, 4: NUM0})
    note = ws.cell(row=1, column=6,
                   value="Sales - November 2026. Header row 1, data from row 2: "
                         "the layout the automation macros expect "
                         "(A=Date, B=Salesperson, C=Category, D=Amount).")
    note.font = Font(italic=True, size=10, color=GRAY)
    set_widths(ws, [12, 16, 14, 12, 4, 60])

    # ---------------- Solution sheet ----------------
    ws = wb["Solution"]
    r = put_title(ws, "Solution - what you should see after each macro")
    r += 1

    r = put_section(ws, r, "After FormatReport (or the hand-written With-block version)")
    for i, h in enumerate(["Date", "Salesperson", "Product", "Qty", "Amount (₹)"], 1):
        ws.cell(row=r, column=i, value=h)
    style_row(ws, r, 1, 5, fill=AMBER_HEADER, bold=True, center=True)
    r = put_rows(ws, r + 1, DAILY_REPORT, num_formats={1: DDMMMYYYY, 5: NUM0})
    r = put_note(ws, r, "Row 1 is bold, filled with RGB(255, 217, 102), centred, "
                        "the columns fit their contents and the top row is frozen.", width=5)
    r += 1

    r = put_section(ws, r, "After FillReportTitle (VBA Fundamentals)")
    ws.cell(row=r, column=1, value="Sunrise Traders - Daily Sales").font = Font(bold=True)
    r += 1
    ws.cell(row=r, column=1, value='="Report date: "&TEXT(TODAY(),"dd-mmm-yyyy")')
    r += 2

    r = put_section(ws, r, "After CellTest")
    ws.cell(row=r, column=1, value="C2 contains:")
    c = ws.cell(row=r, column=3, value="Staplers")
    c.border = _BORDER
    r += 2

    r = put_section(ws, r, "After MonthlyReport - the Sales sheet (laptop rows highlighted)")
    for i, h in enumerate(["Date", "Salesperson", "Category", "Amount (₹)"], 1):
        ws.cell(row=r, column=i, value=h)
    style_row(ws, r, 1, 4, fill=AMBER_HEADER, bold=True, center=True)
    r += 1
    for d, person, cat, amt in SALES_ROWS:
        ws.cell(row=r, column=1, value=d).number_format = DDMMMYYYY
        ws.cell(row=r, column=2, value=person)
        ws.cell(row=r, column=3, value=cat)
        c = ws.cell(row=r, column=4, value=amt)
        c.number_format = NUM0
        fill = LAPTOP_BAND if cat == "Laptops" else None
        style_row(ws, r, 1, 4, fill=fill)
        if fill is None:  # style_row overwrote nothing else; re-apply number formats
            ws.cell(row=r, column=1).number_format = DDMMMYYYY
            ws.cell(row=r, column=4).number_format = NUM0
        r += 1
    r = put_note(ws, r, "Every row whose Category is Laptops carries the "
                        "RGB(255, 235, 205) highlight applied by Rows(r).Interior.Color.", width=4)
    r += 1

    r = put_section(ws, r, "After MonthlyReport - the summary it writes to a new sheet")
    r = put_note(ws, r, "The macro creates a sheet named \"Report \" & Format(Date, "
                        "\"yyyy-mm\") - for example Report 2026-11 - and writes:", width=4)
    ws.cell(row=r, column=1, value="Sunrise Traders - Monthly Report").font = Font(bold=True)
    r += 1
    ws.cell(row=r, column=1, value="Total sales")
    ws.cell(row=r, column=2, value=TOTAL_SALES).number_format = NUM0
    r += 1
    ws.cell(row=r, column=1, value="Laptop sales")
    ws.cell(row=r, column=2, value=LAPTOP_SALES).number_format = NUM0
    r += 1
    r = put_note(ws, r, "Also from the exercises: LoopTest reports lastRow = "
                        f"{LAST_ROW} (it follows the data), and the Pen-count macro "
                        f"reports {PEN_COUNT} rows where Category is Pens.", width=4)
    set_widths(ws, [22, 16, 14, 12, 12])
    return wb


# ==========================================================================
# 2. power-pivot-data.xlsx  (serves l6-power-pivot and l6-intro-dax)
# ==========================================================================
PRODUCTS = [
    ("P101", "Notebook A4 96pg", "Stationery", 120),
    ("P102", "Gel Pen Blue", "Stationery", 35),
    ("P103", "Stapler Medium", "Stationery", 210),
    ("P104", "Laptop 15 i5", "Laptops", 52000),
    ("P105", "Laptop 14 i7", "Laptops", 61500),
    ("P106", "Laser Printer", "Printers", 18000),
    ("P107", "Inkjet Printer", "Printers", 12500),
    ("P108", "Marker Set", "Stationery", 240),
]
PRODUCT_CATEGORY = {p[0]: p[2] for p in PRODUCTS}

REGIONS = [
    ("R01", "North", "Delhi", "Meera Kapoor"),
    ("R02", "South", "Chennai", "Vikram Rao"),
    ("R03", "West", "Mumbai", "Anita Shah"),
]
REGION_NAME = {reg[0]: reg[1] for reg in REGIONS}

PP_SALES = [
    (dt.date(2026, 11, 2), "P104", "R01", "Arjun Mehta", 52000),
    (dt.date(2026, 11, 2), "P102", "R03", "Priya Nair", 1050),
    (dt.date(2026, 11, 3), "P106", "R02", "Rahul Verma", 18000),
    (dt.date(2026, 11, 3), "P101", "R01", "Sneha Iyer", 3600),
    (dt.date(2026, 11, 4), "P105", "R03", "Arjun Mehta", 123000),
    (dt.date(2026, 11, 4), "P103", "R02", "Priya Nair", 840),
    (dt.date(2026, 11, 5), "P107", "R01", "Rahul Verma", 25000),
    (dt.date(2026, 11, 5), "P102", "R02", "Sneha Iyer", 700),
    (dt.date(2026, 11, 6), "P104", "R02", "Arjun Mehta", 104000),
    (dt.date(2026, 11, 6), "P108", "R03", "Priya Nair", 4800),
    (dt.date(2026, 11, 7), "P106", "R01", "Rahul Verma", 36000),
    (dt.date(2026, 11, 7), "P101", "R03", "Sneha Iyer", 2400),
    (dt.date(2026, 11, 10), "P105", "R01", "Arjun Mehta", 184500),
    (dt.date(2026, 11, 10), "P102", "R01", "Priya Nair", 525),
    (dt.date(2026, 11, 11), "P104", "R03", "Rahul Verma", 156000),
    (dt.date(2026, 11, 11), "P103", "R01", "Sneha Iyer", 1050),
    (dt.date(2026, 11, 12), "P107", "R02", "Arjun Mehta", 37500),
    (dt.date(2026, 11, 12), "P108", "R01", "Priya Nair", 1200),
    (dt.date(2026, 11, 13), "P106", "R03", "Rahul Verma", 54000),
    (dt.date(2026, 11, 13), "P102", "R02", "Sneha Iyer", 1400),
    (dt.date(2026, 11, 14), "P103", "R03", "Arjun Mehta", 1680),
    (dt.date(2026, 11, 14), "P101", "R02", "Priya Nair", 4800),
    (dt.date(2026, 11, 14), "P107", "R01", "Rahul Verma", 37500),
    (dt.date(2026, 11, 14), "P104", "R01", "Sneha Iyer", 52000),
]
SALESPERSONS = ["Arjun Mehta", "Priya Nair", "Rahul Verma", "Sneha Iyer"]


def pp_aggregates():
    """Category x Region pivot matrix and per-salesperson DAX measures."""
    categories = ["Laptops", "Printers", "Stationery"]
    region_order = ["North", "South", "West"]
    matrix = {c: {reg: 0 for reg in region_order} for c in categories}
    per_person = {p: {"sales": 0, "orders": 0, "laptops": 0} for p in SALESPERSONS}
    for _date, pid, rid, person, amount in PP_SALES:
        cat = PRODUCT_CATEGORY[pid]
        region = REGION_NAME[rid]
        matrix[cat][region] += amount
        per_person[person]["sales"] += amount
        per_person[person]["orders"] += 1
        if cat == "Laptops":
            per_person[person]["laptops"] += amount
    return categories, region_order, matrix, per_person


def build_power_pivot_data():
    categories, region_order, matrix, per_person = pp_aggregates()
    wb = new_workbook("Exercise", "Solution")

    # ---------------- Exercise sheet ----------------
    ws = wb["Exercise"]
    r = put_title(ws, "Power Pivot & DAX - Data Model Tables",
                  "Level 6 - Power Pivot and the Data Model / Introduction to DAX")
    r = put_note(ws, r,
        "Three clean tables for the Data Model. For each block: select it, "
        "Ctrl+T, then Table Design > Table Name (Products, Regions, Sales). "
        "Add all three to the Data Model (Power Pivot > Add to Data Model), "
        "then in Diagram View relate Products[ProductID] to Sales[ProductID] "
        "and Regions[RegionID] to Sales[RegionID]. Build a PivotTable From "
        "Data Model with Products[Category] on Rows, Regions[Region] on "
        "Columns and Sales[Amount] on Values. Then, for the DAX lesson, add "
        "the four measures: Total Sales, Total Orders, Average Order and "
        "Laptop Sales (see the Solution sheet for the expected numbers).",
        width=8)
    r += 1

    r = put_section(ws, r, "Products (the \"one\" side - one row per ProductID)")
    r = put_headers(ws, r, ["ProductID", "Product", "Category", "Unit Price (₹)"])
    r = put_rows(ws, r, [list(p) for p in PRODUCTS], num_formats={4: NUM0})
    r += 1

    r = put_section(ws, r, "Regions (the \"one\" side - one row per RegionID)")
    r = put_headers(ws, r, ["RegionID", "Region", "City", "Manager"])
    r = put_rows(ws, r, [list(x) for x in REGIONS])
    r += 1

    r = put_section(ws, r, "Sales (the \"many\" side - one row per sale, keys only)")
    r = put_headers(ws, r, ["Date", "ProductID", "RegionID", "Salesperson", "Amount (₹)"])
    sales_start = r
    r = put_rows(ws, r, [list(x) for x in PP_SALES], num_formats={1: DDMMMYYYY, 5: NUM0})
    sales_end = r - 1
    ws.freeze_panes = ws.cell(row=8, column=1)  # below the first table's header
    set_widths(ws, [12, 18, 12, 16, 14])

    # ---------------- Solution sheet ----------------
    ws = wb["Solution"]
    r = put_title(ws, "Solution - expected results from the Data Model")
    r += 1

    r = put_section(ws, r, "Relationships to create in Diagram View")
    r = put_rows(ws, r, [
        ["Products[ProductID] -> Sales[ProductID]", "one-to-many (Products is the one side)"],
        ["Regions[RegionID] -> Sales[RegionID]", "one-to-many (Regions is the one side)"],
    ])
    r += 1

    r = put_section(ws, r, "Expected PivotTable - Sum of Amount by Category and Region")
    headers = ["Category"] + region_order + ["Total"]
    r = put_headers(ws, r, headers)
    rows = []
    for cat in categories:
        vals = [matrix[cat][reg] for reg in region_order]
        rows.append([cat] + vals + [sum(vals)])
    col_tot = [sum(matrix[c][reg] for c in categories) for reg in region_order]
    rows.append(["Total"] + col_tot + [sum(col_tot)])
    r = put_rows(ws, r, rows, num_formats={i: NUM0 for i in range(2, 6)})
    r += 1

    r = put_section(ws, r, "Expected DAX measures by salesperson")
    r = put_headers(ws, r, ["Salesperson", "Total Sales (₹)", "Total Orders",
                            "Average Order (₹)", "Laptop Sales (₹)"])
    for person in SALESPERSONS:
        agg = per_person[person]
        ws.cell(row=r, column=1, value=person)
        ws.cell(row=r, column=2,
                value=f"=SUMIFS(Exercise!$E${sales_start}:$E${sales_end},"
                      f"Exercise!$D${sales_start}:$D${sales_end},$A{r})"
                ).number_format = NUM0
        ws.cell(row=r, column=3,
                value=f"=COUNTIFS(Exercise!$D${sales_start}:$D${sales_end},$A{r})"
                ).number_format = NUM0
        ws.cell(row=r, column=4, value=f"=B{r}/C{r}").number_format = NUM0
        ws.cell(row=r, column=5, value=agg["laptops"]).number_format = NUM0
        for col in range(1, 6):
            ws.cell(row=r, column=col).border = _BORDER
        r += 1
    r = put_note(ws, r, "Laptop Sales comes from =CALCULATE(SUM(Sales[Amount]), "
                        "Products[Category]=\"Laptops\") - a model measure that follows "
                        "the Sales -> Products relationship, so it is shown as a value "
                        "here. The other three are plain-range equivalents of the "
                        "measures, so Excel computes them live.", width=5)
    r += 1

    r = put_section(ws, r, "The four measures (definition reference)")
    r = put_rows(ws, r, [
        ["Total Sales", "=SUM(Sales[Amount])"],
        ["Total Orders", "=COUNTROWS(Sales)"],
        ["Average Order", "=DIVIDE([Total Sales], [Total Orders])"],
        ["Laptop Sales", "=CALCULATE(SUM(Sales[Amount]), Products[Category]=\"Laptops\")"],
    ])
    set_widths(ws, [42, 18, 14, 16, 16])
    return wb


# ==========================================================================
# 3. power-query-messy-data.xlsx  (serves l6-power-query)
# ==========================================================================
# None marks a deliberately blank spacer row.
MESSY_ROWS = [
    ["1", "02/11/2026", "NOTEBOOKS", " Arjun Mehta ", " 2,520"],
    ["2", "02/11/2026", " pens ", "PRIYA NAIR", "5,760 "],
    ["3", "03/11/2026", "Staplers ", "Rahul Verma", "3,500"],
    None,
    ["4", "04/11/2026", " Printers", " Sneha Iyer", "18,000"],
    ["4", "04/11/2026", " Printers", " Sneha Iyer", "18,000"],   # exact duplicate
    ["5", "05/11/2026", "PRINTERS", "RAHUL VERMA", "12,500"],
    ["6", "06/11/2026", " notebooks", "ARJUN MEHTA", " 3,600"],
    ["7", "06/11/2026", "Laptops", "Sneha Iyer ", "46,500"],
    ["8", "07/11/2026", "staplers", "Priya Nair ", "840"],
    ["7", "06/11/2026", "Laptops", "Sneha Iyer ", "46,500"],     # exact duplicate
    ["10", "10/11/2026", "Pens", "Rahul Verma", "N/A"],
    ["11", "10/11/2026", "Markers", " Priya Nair", "1,200"],
    None,
    ["12", "11/11/2026", "laptops", "ARJUN MEHTA", "123000"],
    ["13", "12/11/2026", "Markers ", "SNEHA IYER", " 480"],
]

# Clean state: Row No. removed, blanks removed, duplicates removed, text trimmed
# and capitalised, Date/Amount retyped, N/A replaced with empty.
def parse_date(s):
    d, m, y = s.split("/")
    return dt.date(int(y), int(m), int(d))


def parse_amount(s):
    s = s.strip().replace(",", "")
    return int(s) if s.isdigit() else None


def title_case(s):
    return " ".join(w[:1].upper() + w[1:].lower() for w in s.split())


CLEAN_ROWS = []
SEEN_KEYS = set()
for row in MESSY_ROWS:
    if row is None:
        continue
    key = tuple(x.strip() for x in row)
    if key in SEEN_KEYS:
        continue  # exact duplicate (after trim)
    SEEN_KEYS.add(key)
    CLEAN_ROWS.append([
        parse_date(key[1]),
        title_case(key[2]),
        title_case(key[3]),
        parse_amount(key[4]),
    ])

CROSSTAB = [
    ["Notebooks", 120, 125],
    ["Pens", 35, 38],
    ["Staplers", 210, 215],
    ["Printers", 8500, 8600],
]
UNPIVOT = [[p, "Jan", jan] for p, jan, _ in CROSSTAB] + \
          [[p, "Feb", feb] for p, _, feb in CROSSTAB]

PROBLEMS = [
    "Dates stored as text (\"02/11/2026\") - they sort wrongly and fail date maths",
    "Amounts stored as text with commas and stray spaces (\" 2,520\", \"123000\")",
    "Leading and trailing spaces in Product and Salesperson (\" pens \", \" Sneha Iyer\")",
    "Mixed case in Product and Salesperson (NOTEBOOKS, PRIYA NAIR, staplers)",
    "Two exact duplicate rows (04-Nov Printers and 06-Nov Laptops)",
    "Two blank spacer rows inside the table",
    "A useless Row No. column that only numbers the export",
    "One amount entered as \"N/A\" instead of a number",
    "A merged banner header above the real column headers",
    "A small cross-tab block (Product x Jan/Feb) for the Unpivot exercise",
]


def build_power_query_messy():
    wb = new_workbook("Exercise", "Solution")

    # ---------------- Exercise sheet (the mess) ----------------
    ws = wb["Exercise"]
    r = put_title(ws, "Messy data for Power Query - Sunrise Traders export",
                  "Level 6 - Power Query: Import and Clean Data Automatically")
    r = put_note(ws, r,
        "Do NOT clean this by hand. Import with Data > Get Data > From File > "
        "From Workbook, pick this sheet and click Transform Data. In the editor: "
        "remove the useless column, remove the blank rows, trim the text columns, "
        "fix the Date and Amount types, replace \"N/A\" with empty, remove the "
        "duplicate rows, then Close & Load to a new worksheet. A small cross-tab "
        "sits below the export for the Unpivot exercise.", width=5)
    r += 1

    # merged banner header - part of the mess
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    c = ws.cell(row=r, column=1,
                value="SUNRISE TRADERS - RAW SYSTEM EXPORT - generated 06-Nov-2026 - DO NOT EDIT")
    c.font = Font(bold=True, size=10, color=GRAY)
    c.alignment = Alignment(horizontal="center")
    r += 1

    r = plain_table(ws, r, ["Row No.", "Date", "Product", "Salesperson", "Amount"], [])
    headers_done = r
    for row in MESSY_ROWS:
        if row is None:
            r += 1  # blank spacer row
        else:
            r = put_rows(ws, r, [row])
    r += 1

    r = put_section(ws, r, "Cross-tab for the Unpivot exercise")
    r = plain_table(ws, r, ["Product", "Jan", "Feb"], CROSSTAB, num_formats={2: NUM0, 3: NUM0})
    set_widths(ws, [10, 14, 14, 16, 12])
    _ = headers_done

    # ---------------- Solution sheet (the clean state) ----------------
    ws = wb["Solution"]
    r = put_title(ws, "Solution - the clean, refreshable table")
    r = put_note(ws, r,
        "One possible Applied Steps order: Removed Columns (Row No.) > Removed "
        "Blank Rows > Trimmed Text (Product, Salesperson) > Capitalized Each Word "
        "(optional) > Changed Type (Date -> Date, Amount -> Number) > Replaced "
        "Values (\"N/A\" -> null) > Removed Duplicates (trim before deduplicating, "
        "or padded duplicates survive). Then Close & Load to a new worksheet.",
        width=5)
    r += 1

    r = put_section(ws, r, "Cleaned table - 12 rows, correct types")
    r = put_headers(ws, r, ["Date", "Product", "Salesperson", "Amount (₹)"])
    r = put_rows(ws, r, CLEAN_ROWS, num_formats={1: DDMMMYYYY, 4: NUM0})
    r = put_note(ws, r, "Dates are real dates, amounts are real numbers, the N/A row "
                        "has an empty amount, and both exact duplicates are gone.", width=4)
    r += 1

    r = put_section(ws, r, "Unpivot result - the shape Power Pivot expects")
    r = put_headers(ws, r, ["Product", "Month", "Sales"])
    r = put_rows(ws, r, UNPIVOT, num_formats={3: NUM0})
    r += 1

    r = put_section(ws, r, "Problems you should have spotted")
    r = put_rows(ws, r, [[p] for p in PROBLEMS])
    set_widths(ws, [14, 14, 16, 14])
    ws.freeze_panes = None
    return wb


# ==========================================================================
# 4. macro-code.bas  (serves l6-vba-fundamentals and l6-vba-automation)
# ==========================================================================
BAS = '''Attribute VB_Name = "ExcelLabLevel6Macros"
'=====================================================================
' Excel Lab - Level 6: Advanced Excel
' Ready-to-import VBA module with the macros taught in the macro
' lessons: "Recording Macros", "VBA Fundamentals" and
' "Automating Reports with VBA".
'
' HOW TO USE THIS FILE
'   1. Open the macro practice workbook and re-save it as .xlsm
'      (File > Save As > Excel Macro-Enabled Workbook) - a plain
'      .xlsx file cannot store macros.
'   2. Press Alt+F11 to open the VBA editor.
'   3. Either: File > Import File... and choose this .bas file,
'      or:     Insert > Module, then copy-paste this code in.
'   4. Put the cursor inside a Sub and press F5 to run it, or F8 to
'      step through it line by line. From Excel, use Alt+F8.
'
' Contents:
'   FormatReport       recorded-style macro (uses .Select)
'   RelativeInitials   recorded with Use Relative References
'   FillReportTitle    first hand-written Sub (Range + Format)
'   CellTest           Range vs Cells
'   OrderSummary       Dim variables + MsgBox
'   WhoIsIt            InputBox with Cancel handling
'   FormatReportClean  the With-block rewrite (no .Select)
'   CountPens          For loop + If counter
'   FilterLaptops      bonus: AutoFilter the Sales sheet
'   MonthlyReport      full loop/summary/save macro with On Error
'=====================================================================
Option Explicit

Sub FormatReport()
' Formats the daily sales report header - exactly as the macro
' recorder writes it (record on the Exercise sheet of the practice
' workbook, with A1:E1 holding the report header).

  Range("A1:E1").Select
  Selection.Font.Bold = True
  With Selection.Interior
    .Color = RGB(255, 217, 102)
  End With
  Selection.HorizontalAlignment = xlCenter
  Range("A1").Select

End Sub

Sub RelativeInitials()
' Recorded with Use Relative References switched ON.
' From wherever the cursor is: go two rows down and type initials
' (replace "AB" with your own initials). Run it from different
' starting cells - the initials always land two rows below the
' ACTIVE cell, never at a fixed address.

  ActiveCell.Offset(2, 0).Range("A1").Select
  ActiveCell.FormulaR1C1 = "AB"

End Sub

Sub FillReportTitle()
' Puts the title in A1 and today's date in A2

  Range("A1").Value = "Sunrise Traders - Daily Sales"
  Range("A2").Value = "Report date: " & Format(Date, "dd-mmm-yyyy")

End Sub

Sub CellTest()
' Writes the same value to C2 two ways: Range (address) and
' Cells (row, column). The second form is the one loops use.

  Range("C2").Value = "Staplers"     ' the cell C2
  Cells(2, 3).Value = "Staplers"     ' row 2, column 3 = C2

End Sub

Sub OrderSummary()
' Shows a one-line summary of a sample order.
' Step through it with F8 and hover over the variables to watch
' product, qty, price and total fill in.

  Dim product As String
  Dim qty As Long
  Dim price As Double
  Dim total As Double

  product = "Staplers"
  qty = 12
  price = 210
  total = qty * price

  MsgBox "Sold " & qty & " " & product & " for " & total, vbInformation

End Sub

Sub WhoIsIt()
' Asks for a name and greets the user

  Dim who As String

  who = InputBox("Which salesperson is this report for?")
  If who = "" Then Exit Sub          ' user pressed Cancel - stop politely

  MsgBox "Preparing the report for " & who, vbInformation

End Sub

Sub FormatReportClean()
' Hand-written replacement for the recorded FormatReport: the same
' result with a With block and no .Select anywhere, so it works no
' matter which cell the cursor is in.

  With Range("A1:E1")
    .Font.Bold = True
    .Interior.Color = RGB(255, 217, 102)
    .HorizontalAlignment = xlCenter
  End With

End Sub

Sub CountPens()
' Loops the Sales sheet rows and counts every Pens sale.
' Reports ONCE at the end, never inside the loop.

  Dim ws As Worksheet
  Dim lastRow As Long, r As Long
  Dim penCount As Long

  Set ws = ThisWorkbook.Worksheets("Sales")
  lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

  For r = 2 To lastRow
    If ws.Cells(r, "C").Value = "Pens" Then
      penCount = penCount + 1
    End If
  Next r

  MsgBox "Pens sales found: " & penCount, vbInformation

End Sub

Sub FilterLaptops()
' Bonus: applies an AutoFilter to the Sales sheet and shows only
' the laptop rows. Reset with Data > Clear when done.

  With ThisWorkbook.Worksheets("Sales")
    .Range("A1:D1").AutoFilter Field:=3, Criteria1:="Laptops"
  End With

End Sub

Sub MonthlyReport()
' Scans the Sales sheet, highlights laptop sales,
' writes a summary sheet and saves a dated copy.

  Dim ws As Worksheet, rep As Worksheet
  Dim lastRow As Long, r As Long
  Dim total As Double, laptopTotal As Double

  On Error GoTo ErrHandler

  Set ws = ThisWorkbook.Worksheets("Sales")
  lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

  total = 0
  laptopTotal = 0
  For r = 2 To lastRow
    total = total + ws.Cells(r, "D").Value
    If ws.Cells(r, "C").Value = "Laptops" Then
      ws.Rows(r).Interior.Color = RGB(255, 235, 205)
      laptopTotal = laptopTotal + ws.Cells(r, "D").Value
    End If
  Next r

  Set rep = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count))
  On Error Resume Next           ' name may already exist this month
  rep.Name = "Report " & Format(Date, "yyyy-mm")
  On Error GoTo ErrHandler

  rep.Range("A1").Value = "Sunrise Traders - Monthly Report"
  rep.Range("A2").Value = "Total sales"
  rep.Range("B2").Value = total
  rep.Range("A3").Value = "Laptop sales"
  rep.Range("B3").Value = laptopTotal
  rep.Range("B2:B3").NumberFormat = "#,##0"
  rep.Range("A1").Font.Bold = True

  ThisWorkbook.SaveAs Filename:= _
    ThisWorkbook.Path & Application.PathSeparator & _
    "Sunrise Traders " & Format(Date, "yyyy-mm-dd") & ".xlsm", _
    FileFormat:=xlOpenXMLWorkbookMacroEnabled

  MsgBox "Report ready. Total " & Format(total, "#,##0"), vbInformation
  Exit Sub

ErrHandler:
  MsgBox "The macro stopped: " & Err.Description, vbExclamation
End Sub
'''


# ==========================================================================
# main
# ==========================================================================
def main():
    OUT.mkdir(parents=True, exist_ok=True)

    builders = {
        "macro-practice.xlsx": build_macro_practice,
        "power-pivot-data.xlsx": build_power_pivot_data,
        "power-query-messy-data.xlsx": build_power_query_messy,
    }
    for name, build in builders.items():
        wb = build()
        path = OUT / name
        wb.save(path)
        print(f"saved {path.relative_to(ROOT)}  ({path.stat().st_size:,} bytes)")

    bas_path = OUT / "macro-code.bas"
    bas_path.write_text(BAS, encoding="ascii")   # .bas must be plain ASCII
    print(f"saved {bas_path.relative_to(ROOT)}  ({bas_path.stat().st_size:,} bytes)")

    # sanity: the exact aggregates shown on the macro Solution sheet
    print(f"  macro totals: total={TOTAL_SALES:,} laptops={LAPTOP_SALES:,} "
          f"pens={PEN_COUNT} lastRow={LAST_ROW}")


if __name__ == "__main__":
    main()
