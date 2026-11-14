"""Generate the Level 3 practice workbooks for Excel Lab.

Creates exactly the three .xlsx files referenced by the Level 3 lessons'
downloads entries (no others):

    assets/files/level3/tables-validation.xlsx   5 lessons: sorting, filtering,
                                                 Excel Tables, data validation,
                                                 conditional formatting
    assets/files/level3/data-cleaning.xlsx       4 lessons: removing duplicates,
                                                 text to columns, find & replace,
                                                 complete cleaning workflow
    assets/files/level3/large-dataset.xlsx       1 lesson: working with large
                                                 datasets (1,000 rows)

Each workbook contains exercise sheets named exactly as the lessons refer to
them ("the Sorting sheet", "the Duplicates sheet", ...) plus one "Solution"
sheet whose sections mirror the exact numbers stated in the lessons' exercise
solutions. Exercise sheets hold the starter state (title, instructions, data
grids, empty answer areas); students do the work themselves.

Run from the repo root:  python scripts/gen_level3.py
"""
import random
import sys
from datetime import date, timedelta
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

from xlcommon import (  # noqa: E402
    TEAL_DARK,
    new_workbook,
    put_headers,
    put_note,
    put_rows,
    put_section,
    put_title,
    set_widths,
)

ROOT = SCRIPT_DIR.parent
OUT = ROOT / "assets" / "files" / "level3"

# ---------------------------------------------------------------------------
# Shared data — the 12-row Sunrise Traders order list used across the
# tables-validation lessons. Every number is consistent with the grids printed
# in the lessons (orders 1001-1004, the two West laptop orders 1007/1012 whose
# revenues sum to the Rs 2,64,000 quoted in the filtering lesson, order 1012).
# Columns: Order ID, Date, Salesperson, Region, Product, Units, Revenue
# ---------------------------------------------------------------------------
ORDERS = [
    [1001, date(2026, 1, 5), "Arjun Mehta", "North", "Notebooks", 120, 14400],
    [1002, date(2026, 1, 6), "Priya Nair", "South", "Staplers", 45, 6750],
    [1003, date(2026, 1, 8), "Rahul Verma", "North", "Laptops", 22, 352000],
    [1004, date(2026, 1, 9), "Sneha Iyer", "East", "Pens", 100, 6000],
    [1005, date(2026, 1, 12), "Arjun Mehta", "South", "Printers", 8, 19200],
    [1006, date(2026, 1, 15), "Priya Nair", "North", "Staplers", 60, 9000],
    [1007, date(2026, 1, 19), "Sneha Iyer", "West", "Laptops", 11, 176000],
    [1008, date(2026, 1, 21), "Rahul Verma", "East", "Notebooks", 200, 24000],
    [1009, date(2026, 1, 23), "Arjun Mehta", "West", "Pens", 150, 9000],
    [1010, date(2026, 1, 27), "Priya Nair", "East", "Laptops", 15, 240000],
    [1011, date(2026, 1, 29), "Rahul Verma", "South", "Printers", 22, 52800],
    [1012, date(2026, 2, 2), "Sneha Iyer", "West", "Laptops", 8, 88000],
]

# Priority values used by the sorting lesson's custom-list exercise.
PRIORITY = {
    1001: "Medium", 1002: "Low", 1003: "High", 1004: "Low",
    1005: "Medium", 1006: "Low", 1007: "High", 1008: "Medium",
    1009: "Low", 1010: "High", 1011: "Medium", 1012: "High",
}

ORDER_HEADERS = ["Order ID", "Date", "Salesperson", "Region", "Product", "Units", "Revenue"]
ORDER_WIDTHS = [10, 12, 16, 10, 13, 8, 12]

# Pre-computed sort results mirrored from the lesson exercises.
BY_REV_DESC = sorted(ORDERS, key=lambda r: -r[6])                      # ex 1
BY_REGION_REV_ASC = sorted(ORDERS, key=lambda r: (r[3], r[6]))         # ex 2
_PRIO_RANK = {"High": 0, "Medium": 1, "Low": 2}
BY_PRIORITY = sorted(ORDERS, key=lambda r: _PRIO_RANK[PRIORITY[r[0]]])  # ex 3

GRAND_TOTAL = sum(r[6] for r in ORDERS)                                # 997150
GT20000_ROWS = [r for r in ORDERS if r[6] > 20000]                     # 6 rows
GT20000_TOTAL = sum(r[6] for r in GT20000_ROWS)                        # 932800
TOP5_ROWS = sorted(ORDERS, key=lambda r: -r[6])[:5]
WEST_LAPTOP_TOTAL = 176000 + 88000                                     # 264000

# Validation lesson: Region must land in column D and Units in column F so the
# lesson's "dropdown on D2:D500" and "Units F2:F500" addresses stay meaningful.
VAL_ROWS = [
    [1001, date(2026, 1, 5), "Arjun Mehta", "North", "Notebooks", 120],
    [1002, date(2026, 1, 6), "Priya Nair", "South", "Staplers", 45],
    [1003, date(2026, 1, 8), "Rahul Verma", "north", "Laptops", 22],   # invalid
    [1004, date(2026, 1, 9), "Sneha Iyer", "East", "Pens", 0],         # invalid
    [1005, date(2026, 1, 12), "Arjun Mehta", "West", "Printers", 8],
    [1006, date(2026, 1, 15), "Priya Nair", "North", "Staplers", 60],
    [1007, date(2026, 1, 19), "Sneha Iyer", "South", "Notebooks", 200],
    [1008, date(2026, 1, 21), "Rahul Verma", "East", "Laptops", 5],
]

# Conditional-formatting lesson: payment tracker with Status in column F so the
# lesson's formula rule =$F2="Overdue" works as printed.
TRACKER_HEADERS = ["Invoice", "Customer", "Amount", "Due Date", "Terms", "Status"]
TRACKER = [
    ["INV-2041", "Arjun Mehta", 48500, date(2026, 2, 5), "Net 30", "Paid"],
    ["INV-2042", "Priya Nair", 12750, date(2026, 2, 12), "Net 30", "Overdue"],
    ["INV-2043", "Rahul Verma", 352000, date(2026, 2, 20), "Net 30", "Paid"],
    ["INV-2044", "Sneha Iyer", 8800, date(2026, 2, 28), "Net 30", "Overdue"],
]

# ---------------------------------------------------------------------------
# data-cleaning.xlsx data (duplicates, text-to-columns, find & replace, full
# cleaning workflow). Dates and units are deliberately stored as TEXT strings
# so the sheets start genuinely messy, exactly as the lessons describe.
# ---------------------------------------------------------------------------

# Duplicates lesson: Customer, Email, Region, City
DUP_ROWS = [
    ["Arjun Mehta", "arjun.mehta@sunrise.in", "North", "Delhi"],
    ["Arjun Mehta", "arjun.mehta@sunrise.in", "North", "Delhi"],    # exact copy
    ["Priya Nair", "priya.nair@sunrise.in", "South", "Chennai"],
    ["Rahul Verma", "rahul.verma@sunrise.in", "East", "Kolkata"],
    ["Sneha Iyer", "sneha.iyer@sunrise.in", "West", "Mumbai"],
    ["Arjun Mehta", "arjun.mehta@sunrise.in", "North", "Pune"],     # newer address
    ["Priya Nair", "priya.nair@sunrise.in", "South", "Chennai"],    # exact copy
    ["Meera Krishnan", "meera.k@sunrise.in", "South", "Kochi"],
    ["Rahul Verma", "rahul.verma@sunris.in", "East", "Kolkata"],    # near-typo
    ["Vikram Shah", "vikram.shah@sunrise.in", "North", "Jaipur"],
]

# Text-to-columns lesson: "Last, First" names, units stored as text, product
# codes split by fixed width at position 5.
SPLIT_ROWS = [
    ["Mehta, Arjun", "120", "NORTH0142"],
    ["Nair, Priya", "45", "SOUTH0207"],
    ["Verma, Rahul", "8", "EAST0113"],
    ["Iyer, Sneha", "200", "WEST0058"],
    ["Krishnan, Meera", "60", "NORTH0321"],
    ["Shah, Vikram", "30", "SOUTH0154"],
]
SPLIT_UNITS_SUM = 120 + 45 + 8 + 200 + 60 + 30  # 463 after conversion

# Find & replace lesson: supplier variants, lowercase regions, blank regions
# for the Go To Special > Blanks exercise.
FINDREPLACE_ROWS = [
    [1001, "Sunrise Traders", "North", 120],
    [1002, "Sunrise Traders - North", "north", 45],
    [1003, "Sunrise Traders", None, 200],          # blank region
    [1004, "Global Traders", "South", 60],
    [1005, "Sunrise Traders - South", None, 30],   # blank region
    [1006, "Sunrise Traders", "West", 90],
    [1007, "Sunrise Traders", "north", 150],
    [1008, "Metro Suppliers", "East", 75],
]

# Complete cleaning workflow: columns A=No., B=Customer, C=Email, D=Date(text),
# E=Units(text) so the lesson formula =PROPER(TRIM(B2)) matches the layout.
# Rows 1-4 are exactly the sample grid printed in the lesson.
CLEAN_ROWS = [
    [1, "arjun mehta ", "arjun.mehta@sunrise.in", "05-01-2026", "120"],
    [2, "PRIYA NAIR", "priya.nair@sunrise.in", "05012026", "45"],
    [3, "Rahul  Verma", "rahul.verma@sunrise.in", "5/1/26", "8"],
    [4, "arjun mehta", "arjun.mehta@sunrise.in", "06-01-2026", "200"],
    [5, "SNEHA IYER ", "sneha.iyer@sunrise.in", "07-01-2026", "60"],
    [6, "meera krishnan", "meera.k@sunrise.in", "08012026", "30"],
    [7, "Rahul Verma", "rahul.verma@sunrise.in", "09-01-2026", "75"],
    [8, "SHAH, Vikram", "vikram.shah@sunrise.in", "10-01-2026", "90"],
    [9, "PRIYA NAIR ", "priya.nair@sunrise.in", "11012026", "50"],
    [10, "sneha iyer", "sneha.iyer@sunrise.in", "12-01-2026", "110"],
    [11, "krishnan, meera", "meera.k@sunrise.in", "13012026", "25"],
    [12, "VIKRAM SHAH", "vikram.shah@sunrise.in", "14-01-2026", "80"],
]
CLEAN_UNITS_SUM = 120 + 45 + 8 + 200 + 60 + 30 + 75 + 90 + 50 + 110 + 25 + 80  # 893

CLEANED_CUSTOMERS = [
    "Arjun Mehta", "Priya Nair", "Rahul Verma", "Sneha Iyer",
    "Meera Krishnan", "Vikram Shah",
]


# ---------------------------------------------------------------------------
# Small style helpers on top of xlcommon
# ---------------------------------------------------------------------------

def note(ws, row, text, width=8, height=60):
    """Instruction note with an explicit height for long multi-line text."""
    put_note(ws, row, text, width=width)
    ws.row_dimensions[row].height = height
    return row + 1


def sol_lines(ws, row, items):
    """Solution text lines. Lines starting with '!' are rendered bold."""
    for item in items:
        bold = item.startswith("!")
        text = item.lstrip("!").strip()
        c = ws.cell(row=row, column=1, value=text)
        c.font = Font(bold=bold, size=10, color=TEAL_DARK if bold else "1e293b")
        c.alignment = Alignment(vertical="top")
        row += 1
    return row


def order_table(ws, row, rows, title=None, extra_header=None, widths=None):
    """Write a Sunrise Traders order table. Returns (next_row, header_row, data_start)."""
    headers = ORDER_HEADERS + ([extra_header] if extra_header else [])
    if title:
        row = put_section(ws, row, title, width=len(headers))
    header_row = row
    start = put_headers(ws, row, headers)  # put_headers returns the next row
    fmt = {2: "dd-mmm-yyyy", 7: "#,##0"}
    end = put_rows(ws, start, rows, num_formats=fmt)
    if widths:
        set_widths(ws, widths)
    return end, header_row, start


def answer_grid(ws, row, headers, prompts):
    """Empty bordered answer area with task prompts in column A."""
    row = put_section(ws, row, "Your answers — fill this in as you work", width=len(headers))
    row = put_headers(ws, row, headers)
    rows = [[p] + [None] * (len(headers) - 1) for p in prompts]
    return put_rows(ws, row, rows)


def with_priority(rows):
    return [r + [PRIORITY[r[0]]] for r in rows]


# ---------------------------------------------------------------------------
# tables-validation.xlsx — exercise sheets
# ---------------------------------------------------------------------------

def build_sorting(ws):
    set_widths(ws, ORDER_WIDTHS + [10])
    row = put_title(ws, "Level 3 · Sorting Data — Exercise",
                    "Sunrise Traders order list · tables-validation workbook")
    row = note(ws, row, (
        "Sort the order list below, keeping every row intact.\n"
        "1) Sort so the highest Revenue appears first: click ONE cell in the Revenue column, Data > Sort Z to A. "
        "If the Sort Warning appears, choose Expand the selection.\n"
        "2) Sort Region A to Z, and within each region Revenue smallest to largest: Data > Sort, Region A to Z, Add Level, Revenue Smallest to Largest.\n"
        "3) Sort Priority so High rows come first, then Medium, then Low: Data > Sort, Sort by Priority, Order = Custom List, type High, Medium, Low and click Add. "
        "(Alphabetical order would give High, Low, Medium — the custom list fixes it.)"
    ))
    row += 1
    end, header_row, start = order_table(
        ws, row, with_priority(ORDERS),
        title="Order list (entry order) — sort this",
        extra_header="Priority", widths=ORDER_WIDTHS + [10])
    row = end + 1
    row = answer_grid(ws, row,
                      ["Sort task", "Top row you see after sorting (Order ID · Salesperson · Region · Product · Revenue)"],
                      ["1 — Revenue largest to smallest",
                       "2 — Region A to Z, then Revenue smallest to largest",
                       "3 — Priority custom list High, Medium, Low"])
    ws.freeze_panes = ws.cell(row=start, column=1)
    return {"start": start, "last": end - 1}


def build_filtering(ws):
    set_widths(ws, ORDER_WIDTHS)
    row = put_title(ws, "Level 3 · Filtering Data — Exercise",
                    "Sunrise Traders order list · tables-validation workbook")
    row = note(ws, row, (
        "Click any cell in the list and press Ctrl + Shift + L to switch the filter on.\n"
        "1) Show only the orders from the West region (Region dropdown, tick West only).\n"
        "2) Show only rows with Revenue greater than 20,000 (Number Filters > Greater Than), then add a live total that follows the filter "
        "with =SUBTOTAL(9, ...) in the answer area — with the whole list visible it equals the SUM of all 12 rows.\n"
        "3) Show the orders due this month (Date Filters > This Month), then copy the visible rows including headers and paste them into a new blank sheet — "
        "a filtered copy pastes only visible rows.\n"
        "4) Finish with Data > Clear to reset every column, then Ctrl + Shift + L to remove the filter arrows completely."
    ))
    row += 1
    end, header_row, start = order_table(ws, row, ORDERS, title="Order list (12 rows) — filter this")
    row = end + 1
    row = answer_grid(ws, row,
                      ["Filter applied", "Visible Order IDs", "Live total cell =SUBTOTAL(9, G" + str(start) + ":G" + str(end - 1) + ")"],
                      ["Region = West", "Revenue > 20000", "Date = This Month", "All filters cleared"])
    ws.freeze_panes = ws.cell(row=start, column=1)
    return {"start": start, "last": end - 1}


def build_tables(ws):
    set_widths(ws, ORDER_WIDTHS + [14])
    row = put_title(ws, "Level 3 · Excel Tables — Exercise",
                    "Sunrise Traders order list · tables-validation workbook")
    row = note(ws, row, (
        "1) Convert the plain range below into an Excel Table and rename it Sales: click inside the list, Ctrl + T, "
        "tick My table has headers, OK; then Table Design > Table Name, type Sales, Enter.\n"
        "2) Add a Revenue-per-Unit calculated column: type =Sales[@Revenue]/Sales[@Units] in the first cell of the empty column H "
        "and press Enter — the Table fills the whole column by itself.\n"
        "3) Tick Total Row on the Table Design tab and choose Sum for the Revenue cell.\n"
        "4) In a free cell write a formula for total South revenue: =SUMIF(Sales[Region], \"South\", Sales[Revenue]). "
        "Then type a new South order directly below the Table (e.g. 1013 · Priya Nair · South · Staplers · 30 units · 4500) "
        "and confirm the total updates with no range editing."
    ))
    row += 1
    end, header_row, start = order_table(ws, row, ORDERS[:6], title="Plain range — convert this to a Table named Sales")
    row = end + 1
    row = answer_grid(ws, row,
                      ["Task", "Your result"],
                      ["Table renamed to Sales",
                       "Revenue-per-Unit column (first value)",
                       "Totals row Sum of Revenue",
                       "South revenue before / after the new row"])
    ws.freeze_panes = ws.cell(row=start, column=1)
    return {"start": start, "last": end - 1}


def build_validation(ws):
    set_widths(ws, [10, 12, 16, 10, 13, 8])
    row = put_title(ws, "Level 3 · Data Validation — Exercise",
                    "Order entry sheet · tables-validation workbook")
    row = note(ws, row, (
        "1) Add a dropdown to the Region column (column D, all data rows): select the cells, Data > Data Validation, "
        "Allow = List, Source = North,South,East,West, OK. Then improve it: put the four values in the list cells below the table "
        "and use them as the range source with $ anchors.\n"
        "2) Restrict Units (column F) to whole numbers between 1 and 10000: Allow = Whole number, between, Minimum 1, Maximum 10000. "
        "On Error Alert choose Stop and type \"Units must be a whole number between 1 and 10000\". Test by typing 0 — Excel refuses it.\n"
        "3) Add an input message to the Region cells: Data Validation > Input Message tab, title \"Region\", "
        "message \"Choose one of the four regions from the dropdown\".\n"
        "4) Some values were typed before validation existed: Data > Data Validation dropdown > Circle Invalid Data finds them. "
        "Fix each one to match the rules, then Clear Validation Circles."
    ))
    row += 1
    row = put_section(ws, row, "Order entry sheet — add validation to Region (D) and Units (F)", width=6)
    header_row = row
    start = put_headers(ws, row, ["Order ID", "Date", "Salesperson", "Region", "Product", "Units"])
    end = put_rows(ws, start, VAL_ROWS, num_formats={2: "dd-mmm-yyyy"})
    row = end + 1
    row = put_section(ws, row, "Region list cells — use these as the range source for an improved dropdown", width=6)
    header_row2 = put_headers(ws, row, ["Region"])
    put_rows(ws, header_row2 + 1, [["North"], ["South"], ["East"], ["West"]])
    ws.freeze_panes = ws.cell(row=start, column=1)
    return {"start": start, "last": end - 1}


def build_cond_formatting(ws):
    set_widths(ws, ORDER_WIDTHS + [12, 14])
    row = put_title(ws, "Level 3 · Conditional Formatting — Exercise",
                    "Sunrise Traders order list + payment tracker · tables-validation workbook")
    row = note(ws, row, (
        "1) Highlight every Revenue above 20,000: select the Revenue cells, Home > Conditional Formatting > Highlight Cells Rules > Greater Than, "
        "enter 20000, keep the light red fill, OK.\n"
        "2) Mark the top 5 orders with a green fill without sorting: Conditional Formatting > Top/Bottom Rules > Top 10 Items, "
        "change 10 to 5, choose a green fill, OK.\n"
        "3) Add Data Bars to the Revenue column so the biggest orders show as in-cell bars.\n"
        "4) In the payment tracker, highlight the entire row light red whenever Status says Overdue: select the tracker rows from A across, "
        "Conditional Formatting > New Rule > Use a formula, type =$F2=\"Overdue\" (lock only the column), set a light red fill, OK."
    ))
    row += 1
    end, header_row, start = order_table(ws, row, ORDERS, title="Order list — apply the Revenue rules here (Revenue is column G)")
    row = end + 1
    row = put_section(ws, row, "Payment tracker — apply the formula rule =$F2=\"Overdue\" here (Status is column F)", width=6)
    header_row2 = put_headers(ws, row, TRACKER_HEADERS)
    put_rows(ws, header_row2 + 1, TRACKER, num_formats={3: "#,##0", 4: "dd-mmm-yyyy"})
    ws.freeze_panes = ws.cell(row=start, column=1)
    return {"start": start, "last": end - 1}

# __APPEND_3__

def build_tv_solution(ws, facts):
    """Solution sheet covering all five tables-validation lessons."""
    set_widths(ws, [14, 16, 16, 10, 13, 10, 12, 10])
    row = put_title(ws, "Solutions — Tables, filtering and validation workbook",
                    "Finished states matching the lesson exercises")
    row += 1

    # --- Lesson: Sorting Data ---
    row = put_section(ws, row, "Lesson: Sorting Data", width=8)
    row = sol_lines(ws, row, [
        "!Exercise 1 — Revenue largest to smallest",
        "Sort Z to A from one cell in the Revenue column. The laptop order 1003 (Rahul Verma, North, Laptops, ₹3,52,000) is at the top",
        "with its own Order ID, salesperson and region still on the same row. Expected order of the whole list:",
    ])
    row, _, _ = order_table(ws, row, BY_REV_DESC, title=None)
    row = sol_lines(ws, row, [
        "!Exercise 2 — Region A to Z, Revenue smallest to largest",
        "Data > Sort: Level 1 = Region, Values, A to Z. Add Level: Level 2 = Revenue, Values, Smallest to Largest.",
        "Regions appear as East, North, South, West with ascending revenue inside each group:",
    ])
    row, _, _ = order_table(ws, row, BY_REGION_REV_ASC, title=None)
    row = sol_lines(ws, row, [
        "!Exercise 3 — Priority custom list",
        "Data > Sort > Sort by Priority, Order = Custom List, type High, Medium, Low on three lines, Add, OK. Expected order:",
    ])
    row, _, _ = order_table(ws, row, with_priority(BY_PRIORITY), title=None, extra_header="Priority")
    row += 1

    # --- Lesson: Filtering Data ---
    row = put_section(ws, row, "Lesson: Filtering Data", width=8)
    flt = facts["filtering"]
    row = sol_lines(ws, row, [
        "!Exercise 1 — Region = West",
        "After Ctrl + Shift + L, tick West only in the Region dropdown. Three rows stay visible: 1007 (Laptops, ₹1,76,000),",
        "1009 (Pens, ₹9,000) and 1012 (Laptops, ₹88,000); the row numbers on the left turn blue.",
        "West AND Laptops leaves exactly two rows — 1007 and 1012 — with visible revenue of ₹2,64,000 (₹1,76,000 + ₹88,000).",
        "!Exercise 2 — Revenue > 20000 and the live total",
        "Number Filters > Greater Than > 20000 leaves 6 rows: 1003, 1007, 1008, 1010, 1011, 1012.",
        "The live total =SUBTOTAL(9, G" + str(flt["start"]) + ":G" + str(flt["last"]) + ") then shows ₹9,32,800 (sum of visible rows only);",
        "with the filter cleared the same cell shows ₹9,97,150, the sum of all 12 rows. The lesson's generic form is =SUBTOTAL(9, G2:G200).",
        "!Exercise 3 — This month, then copy",
        "Date Filters > This Month shows the rows dated in the current month. The list spans Jan–Feb 2026, so in February only order 1012",
        "(₹88,000) remains. Select the visible rows including headers, Ctrl + C, paste into a blank sheet with Ctrl + V:",
        "the paste contains only the visible rows, never the hidden ones.",
        "!Exercise 4 — Clear and remove",
        "Data > Clear shows all rows again but keeps the dropdown arrows; Ctrl + Shift + L removes the arrows completely.",
    ])
    row += 1

    # --- Lesson: Excel Tables ---
    row = put_section(ws, row, "Lesson: Excel Tables", width=8)
    row = sol_lines(ws, row, [
        "!Exercise 1 — Create and rename the Table",
        "Ctrl + T inside the 6-row list, tick My table has headers, OK; then Table Design > Table Name, type Sales, Enter.",
        "!Exercise 2 — Revenue-per-Unit calculated column",
        "Type =Sales[@Revenue]/Sales[@Units] (or =[@Revenue]/[@Units]) in the first cell of the empty column and press Enter:",
        "   1001 → 120 · 1002 → 150 · 1003 → 16000 · 1004 → 60 · 1005 → 2400 · 1006 → 150",
        "The whole column fills at once, and any new row already has the formula waiting.",
        "!Exercise 3 — Totals row",
        "Tick Total Row on the Table Design tab; the Revenue cell with Sum shows ₹6,07,350 (a SUBTOTAL, so it follows filters).",
        "!Exercise 4 — South revenue with SUMIF",
        "=SUMIF(Sales[Region], \"South\", Sales[Revenue]) returns ₹25,950 (6750 + 19200).",
        "Add new South order 1013 (Priya Nair, South, Staplers, 30 units, ₹4,500) in the row below the Table:",
        "the Table expands and the formula updates to ₹30,400 with no range editing.",
    ])
    row += 1

    # --- Lesson: Data Validation ---
    row = put_section(ws, row, "Lesson: Data Validation", width=8)
    val = facts["validation"]
    row = sol_lines(ws, row, [
        "!Exercise 1 — Region dropdown",
        "Allow = List, Source = North,South,East,West on the Region cells D" + str(val["start"]) + ":D" + str(val["last"]) +
        " (the lesson's D2:D500 assumes a taller sheet). Improved version: source the range of the four list cells below the table",
        "with $ anchors, or name the range and use =RegionList.",
        "!Exercise 2 — Units rule",
        "Allow = Whole number, between, Minimum 1, Maximum 10000 on F" + str(val["start"]) + ":F" + str(val["last"]) +
        " (lesson: F2:F500). Stop alert text: \"Units must be a whole number between 1 and 10000\".",
        "Typing 0 is refused.",
        "!Exercise 3 — Input message",
        "Title \"Region\", message \"Choose one of the four regions from the dropdown\" — it appears whenever a Region cell is selected.",
        "!Exercise 4 — Circle Invalid Data",
        "Red ovals mark exactly two pre-existing values:",
        "   Order 1003 — Region typed as \"north\" (wrong case, not one of the four list values) → fix to North.",
        "   Order 1004 — Units 0 (outside 1–10000) → fix to a whole number between 1 and 10000.",
        "Then Clear Validation Circles.",
    ])
    row += 1

    # --- Lesson: Conditional Formatting ---
    row = put_section(ws, row, "Lesson: Conditional Formatting", width=8)
    top5_ids = ", ".join(str(r[0]) for r in TOP5_ROWS)
    row = sol_lines(ws, row, [
        "!Exercise 1 — Greater Than 20000",
        "Six revenue cells highlight: 1003 (₹3,52,000), 1007 (₹1,76,000), 1008 (₹24,000), 1010 (₹2,40,000), 1011 (₹52,800) and 1012 (₹88,000).",
        "!Exercise 2 — Top 5 items",
        "Top/Bottom Rules > Top 10 Items changed to 5 highlights 1003 (352000), 1010 (240000), 1007 (176000), 1012 (88000), 1011 (52800).",
        "The sort order of the sheet stays untouched.",
        "!Exercise 3 — Data Bars",
        "Every Revenue cell shows a bar proportional to its value; the ₹3,52,000 laptop order (1003) dwarfs the pen orders visually.",
        "!Exercise 4 — Whole-row formula rule",
        "New Rule > Use a formula, =$F2=\"Overdue\", light red fill. Exactly two rows colour across: INV-2042 (Priya Nair, ₹12,750)",
        "and INV-2044 (Sneha Iyer, ₹8,800). When a status changes to Paid, that row decolours itself — no repainting, ever.",
    ])


def build_tables_validation(path):
    wb = new_workbook("Sorting", "Filtering", "Tables", "Validation",
                      "Conditional Formatting", "Solution")
    facts = {
        "sorting": build_sorting(wb["Sorting"]),
        "filtering": build_filtering(wb["Filtering"]),
        "validation": build_validation(wb["Validation"]),
        "cf": build_cond_formatting(wb["Conditional Formatting"]),
    }
    build_tables(wb["Tables"])
    build_tv_solution(wb["Solution"], facts)
    wb.save(path)
    return path

# __APPEND_4__

# ---------------------------------------------------------------------------
# data-cleaning.xlsx — exercise sheets
# ---------------------------------------------------------------------------

DUP_HEADERS = ["Customer", "Email", "Region", "City"]


def build_duplicates(ws):
    set_widths(ws, [16, 26, 10, 12])
    row = put_title(ws, "Level 3 · Removing Duplicates — Exercise",
                    "Sunrise Traders customer list · data-cleaning workbook")
    row = note(ws, row, (
        "1) Highlight every repeated customer email: select the Email column, Home > Conditional Formatting > Highlight Cells Rules > "
        "Duplicate Values, OK. Count them. (Tip: switch the dialog to Unique to spot near-typo values that appear exactly once.)\n"
        "2) Add a helper column that shows how many times each email appears: type =COUNTIF($B$7:$B$16, B7) and fill down, "
        "then filter the helper to values greater than 1 to isolate the repeat rows.\n"
        "3) Create a backup copy of this sheet (right-click the tab > Move or Copy > Create a copy, name it CustomerList-BACKUP), "
        "then on the original run Data > Remove Duplicates with ALL columns ticked and note how many rows Excel reports removed.\n"
        "4) One customer appears twice with different addresses. Decide by hand which row is the correct one and delete the outdated row — "
        "Remove Duplicates always keeps the FIRST row, which is not always the best one."
    ))
    row += 1
    row = put_section(ws, row, "Customer list (entry order) — clean this", width=len(DUP_HEADERS))
    header_row = row
    start = put_headers(ws, row, DUP_HEADERS)
    end = put_rows(ws, start, DUP_ROWS)
    row = end + 1
    row = answer_grid(ws, row,
                      ["Task", "Your result"],
                      ["Repeated emails highlighted (count)",
                       "COUNTIF helper values (first three rows)",
                       "Remove Duplicates result message",
                       "Partial duplicate — which row did you keep?"])
    ws.freeze_panes = ws.cell(row=start, column=1)
    return {"start": start, "last": end - 1}


def build_splitting(ws):
    set_widths(ws, [18, 12, 14, 14])
    row = put_title(ws, "Level 3 · Text to Columns — Exercise",
                    "Sunrise Traders import · data-cleaning workbook")
    row = note(ws, row, (
        "1) Split the \"Last name, First name\" Customer column into two columns. Insert two empty columns to the right of the names "
        "first (select column B, right-click > Insert, twice) — Text to Columns overwrites whatever sits to the right of the destination. "
        "Then select the name column, Data > Text to Columns > Delimited > Next, tick Comma only, Next, set the destination, Finish.\n"
        "2) The Units column is stored as text (left-aligned, =SUM() returns 0). Fix the whole column in one pass: select ONLY the Units column, "
        "Text to Columns > Delimited > Next, untick every delimiter, Next, keep General, Finish. Values become right-aligned real numbers.\n"
        "3) Flash Fill: type \"Arjun\" next to \"Mehta, Arjun\" and \"Priya\" next to \"Nair, Priya\" in an empty column, then press Ctrl + E. "
        "Spot-check the results — Flash Fill guesses.\n"
        "4) The product codes hide the region in positions 1–5 (e.g. NORTH0142): select the code column, Text to Columns > Fixed width > Next, "
        "click between the 5th and 6th characters to place a break line, Next, choose a destination with a free column beside it, Finish."
    ))
    row += 1
    row = put_section(ws, row, "Imported data — all three problem types in one sheet", width=4)
    header_row = row
    start = put_headers(ws, row, ["Customer", "Units (text)", "Product code", "First name (Flash Fill)"])
    end = put_rows(ws, start, [r + [None] for r in SPLIT_ROWS])
    row = end + 1
    row = answer_grid(ws, row,
                      ["Task", "Your result"],
                      ["Split names (first row)",
                       "Units after conversion — what does =SUM() show now?",
                       "Flash Fill results (spot-checked?)",
                       "Region codes extracted"])
    ws.freeze_panes = ws.cell(row=start, column=1)
    return {"start": start, "last": end - 1}


def build_findreplace(ws):
    set_widths(ws, [10, 24, 10, 10])
    row = put_title(ws, "Level 3 · Find and Replace — Exercise",
                    "Sunrise Traders supplier log · data-cleaning workbook")
    row = note(ws, row, (
        "1) A supplier was renamed: Ctrl + H, Find \"Sunrise Traders\", Replace \"Sunrise Trading Co\", Replace All — "
        "and read the replacement count before accepting (it also matches inside longer text such as \"Sunrise Traders - North\").\n"
        "2) Replace the lowercase region values \"north\" with \"North\" without touching the correct ones: Ctrl + H, Options >>, "
        "tick Match case, Replace All.\n"
        "3) Fill every blank Region cell with the value from the cell above: select the Region range, Ctrl + G > Special > Blanks > OK, "
        "type = and press the Up arrow once, then Ctrl + Enter. Finish with Paste Special > Values so the formulas survive the next sort. "
        "(Do the Match case replacement first so the filled values are correct.)\n"
        "4) Find every formula on the sheet that returns an error value: Ctrl + G > Special > Formulas, tick Errors only, OK — "
        "Excel selects them all at once."
    ))
    row += 1
    row = put_section(ws, row, "Supplier log — fix this", width=4)
    header_row = row
    start = put_headers(ws, row, ["Order ID", "Supplier", "Region", "Units"])
    end = put_rows(ws, start, FINDREPLACE_ROWS)
    row = end + 1
    row = put_section(ws, row, "Formula audit area — two of these formulas return errors", width=4)
    fr = start
    ws.cell(row=row, column=1, value='=VLOOKUP("ACME-999", $A$' + str(fr) + ':$B$' + str(fr + 7) + ', 2, 0)')
    ws.cell(row=row, column=3, value="=1/0")
    row += 2
    row = answer_grid(ws, row,
                      ["Task", "Your result"],
                      ["Replace All count for \"Sunrise Traders\"",
                       "Match case replacements (\"north\" → \"North\")",
                       "Blanks filled (which regions appeared?)",
                       "Error formulas found"])
    ws.freeze_panes = ws.cell(row=start, column=1)
    return {"start": start, "last": end - 1}

# __APPEND_5__

def build_cleaning(ws):
    set_widths(ws, [6, 18, 26, 12, 12])
    row = put_title(ws, "Level 3 · The Complete Data Cleaning Workflow — Exercise",
                    "Messy customer list · data-cleaning workbook")
    row = note(ws, row, (
        "Step 0 — inspect BEFORE changing anything. Work on a copy of the sheet (right-click the tab > Move or Copy > Create a copy) and keep the original.\n"
        "1) List every problem type you can find: alignment (left-aligned numbers are text), case chaos, extra spaces, combined name fields, "
        "dates stored as text, units stored as text, repeated customers. Write the fix order down — TRIM before de-duplicating, "
        "structure before the duplicate definition.\n"
        "2) Clean the Customer column: insert a helper column, type =PROPER(TRIM(B7)) and fill down, then copy the helper and "
        "Paste Special > Values over the Customer column, and delete the helper. \"arjun mehta \" becomes \"Arjun Mehta\".\n"
        "3) Convert the Date column and the Units column to real types with Text to Columns: Delimited, untick all delimiters, "
        "then choose Date (day-month order for this file) for dates and General for units. Both columns should right-align and "
        "=ISNUMBER() checks return TRUE.\n"
        "4) Finish the job: Ctrl + T, rename the Table Customers; highlight duplicates on Email (Conditional Formatting > Duplicate Values); "
        "back up the sheet; then Data > Remove Duplicates on the defining columns. Verify: LEN spot-checks, ISNUMBER, zero highlighted "
        "duplicates, and a test SUMIF you can confirm by hand."
    ))
    row += 1
    row = put_section(ws, row, "The messy customer list — every row needs work", width=5)
    header_row = row
    start = put_headers(ws, row, ["No.", "Customer", "Email", "Date (text)", "Units (text)"])
    end = put_rows(ws, start, CLEAN_ROWS)
    row = end + 1
    row = answer_grid(ws, row,
                      ["Task", "Your result"],
                      ["Problem inventory (before any fix)",
                       "Customer column after PROPER(TRIM)",
                       "Dates and units after conversion (SUM of units?)",
                       "Final row count after Remove Duplicates"])
    ws.freeze_panes = ws.cell(row=start, column=1)
    return {"start": start, "last": end - 1}

# __APPEND_6__

def build_dc_solution(ws, facts):
    """Solution sheet covering all four data-cleaning lessons."""
    set_widths(ws, [16, 26, 10, 12, 10])
    row = put_title(ws, "Solutions — Data cleaning workbook",
                    "Finished states matching the lesson exercises")
    row += 1

    # --- Lesson: Removing Duplicates ---
    row = put_section(ws, row, "Lesson: Removing Duplicates", width=5)
    dup = facts["duplicates"]
    row = sol_lines(ws, row, [
        "!Exercise 1 — Highlight duplicates",
        "Duplicate Values on the Email column lights 5 cells: arjun.mehta@sunrise.in on 3 rows (the exact copy and the newer Pune row)",
        "and priya.nair@sunrise.in on 2 rows. Note that rahul.verma@sunris.in (row 9) does NOT light — it is unique, but it is a",
        "near-typo of rahul.verma@sunrise.in. Use the Unique option to catch values like this one.",
        "!Exercise 2 — COUNTIF helper",
        "=COUNTIF($B$" + str(dup["start"]) + ":$B$" + str(dup["last"]) + ", B" + str(dup["start"]) + ") filled down shows, row by row:",
        "   3, 3, 2, 1, 1, 3, 2, 1, 1, 1",
        "Filter the helper to values greater than 1 and exactly 5 repeat rows appear (the lesson's form =COUNTIF($B$2:$B$500, B2) assumes headers in row 1).",
        "!Exercise 3 — Backup, then Remove Duplicates (all columns ticked)",
        "Excel reports 2 duplicate values removed — the exact copies of Arjun Mehta (row 2) and Priya Nair (row 7).",
        "8 rows remain, stored safely on the CustomerList-BACKUP sheet too.",
        "!Exercise 4 — The partial duplicate",
        "Rows 1 and 6 share the same email but have different addresses (Delhi vs Pune). Remove Duplicates keeps the FIRST row (Delhi),",
        "so review by hand: keep the newer Pune row and delete the stale Delhi row manually. Final result: 7 unique customers.",
    ])
    row, _, _ = _dup_solution_table(ws, row, dup)
    row += 1

    # --- Lesson: Text to Columns ---
    row = put_section(ws, row, "Lesson: Text to Columns and Data Splitting", width=5)
    row = sol_lines(ws, row, [
        "!Exercise 1 — Delimited split of the names",
        "Comma split (two empty columns inserted first) produces:",
        "   Mehta | Arjun · Nair | Priya · Verma | Rahul · Iyer | Sneha · Krishnan | Meera · Shah | Vikram",
        "!Exercise 2 — Numbers stored as text",
        "Text to Columns > Delimited, NO delimiters, General format converts \"120\" \"45\" \"8\" \"200\" \"60\" \"30\" into real numbers.",
        "Alignment flips to the right and =SUM() now returns " + str(SPLIT_UNITS_SUM) + " instead of 0.",
        "!Exercise 3 — Flash Fill (Ctrl + E)",
        "From the two typed examples (\"Arjun\", \"Priya\") the column fills: Arjun, Priya, Rahul, Sneha, Meera, Vikram — always spot-check.",
        "!Exercise 4 — Fixed width split of the product codes",
        "Break line after position 5 gives:",
        "   NORTH | 0142 · SOUTH | 0207 · EAST | 0113 · WEST | 0058 · NORTH | 0321 · SOUTH | 0154",
    ])
    row += 1

    # --- Lesson: Find and Replace ---
    row = put_section(ws, row, "Lesson: Find and Replace", width=5)
    fr_facts = facts["findreplace"]
    row = sol_lines(ws, row, [
        "!Exercise 1 — Supplier renamed",
        "Ctrl + H, Find \"Sunrise Traders\", Replace \"Sunrise Trading Co\", Replace All: Excel reports 6 replacements —",
        "orders 1001, 1002, 1003, 1005, 1006 and 1007, including inside \"Sunrise Traders - North\" and \"Sunrise Traders - South\".",
        "Global Traders and Metro Suppliers are untouched.",
        "!Exercise 2 — Match case",
        "Find \"north\", Replace \"North\", Match case ticked: 2 replacements (orders 1002 and 1007). \"North\" cells are skipped.",
        "!Exercise 3 — Fill the blanks",
        "Go To Special > Blanks selects the 2 empty Region cells; = + Up arrow + Ctrl + Enter fills each with the value above:",
        "order 1003 shows North (after the Match case fix) and order 1005 shows South. Finish with Paste Special > Values.",
        "!Exercise 4 — Error formulas",
        "The workbook contains 2 error formulas, both selected by Go To Special > Formulas > Errors:",
        "   =VLOOKUP(\"ACME-999\", ...) → #N/A (no such supplier) and =1/0 → #DIV/0!",
    ])
    row += 1

    # --- Lesson: Complete Data Cleaning Workflow ---
    row = put_section(ws, row, "Lesson: The Complete Data Cleaning Workflow", width=5)
    cl = facts["cleaning"]
    row = sol_lines(ws, row, [
        "!Exercise 1 — Problem inventory (the correct inspection finds all of these)",
        "   1. Names in mixed case (\"arjun mehta\", \"PRIYA NAIR\", \"VIKRAM SHAH\").",
        "   2. Extra trailing spaces (rows 1, 5 and 9).",
        "   3. A double space inside a name (\"Rahul  Verma\", row 3).",
        "   4. Combined \"Last, First\" name fields (rows 8 and 11).",
        "   5. Dates stored as text in three formats (05-01-2026 · 05012026 · 5/1/26).",
        "   6. Units stored as text — left-aligned; =SUM() returns 0.",
        "   7. Repeated customers — all six emails appear exactly twice.",
        "!Exercise 2 — Clean the Customer column",
        "=PROPER(TRIM(B" + str(cl["start"]) + ")) filled down, then Paste Special > Values over the original column:",
        "\"arjun mehta \" becomes \"Arjun Mehta\", \"Rahul  Verma\" becomes \"Rahul Verma\", \"SHAH, Vikram\" becomes \"Shah, Vikram\".",
        "The combined names (now \"Shah, Vikram\" and \"Krishnan, Meera\") are then split with Text to Columns and rebuilt as First Last.",
        "!Exercise 3 — Real types",
        "Text to Columns > Date (day-month order) turns the text dates into 05-Jan-2026 through 14-Jan-2026; > General turns the units",
        "into real numbers. Both columns right-align, ISNUMBER returns TRUE everywhere, and =SUM() over the units shows " + str(CLEAN_UNITS_SUM) + ".",
        "!Exercise 4 — Structure, de-duplicate, verify",
        "Ctrl + T and rename the Table Customers. Duplicate Values on Email lights all 12 rows (each customer exactly twice).",
        "Back up the sheet, then Remove Duplicates on Email: 6 duplicates removed, 6 unique rows remain:",
        "   " + " · ".join(CLEANED_CUSTOMERS),
        "Final checklist passes: LEN spot-checks clean, ISNUMBER TRUE down every numeric column, zero highlighted duplicates,",
        "=COUNTA(B:B)-1 returns 6, and a test SUMIF agrees with a value verified by hand.",
    ])
    row, _, _ = _clean_solution_table(ws, row, cl)
    return row


def _dup_solution_table(ws, row, dup):
    """Finished customer list with live COUNTIF formulas against the exercise sheet."""
    rng = "Duplicates!$B$" + str(dup["start"]) + ":$B$" + str(dup["last"])
    # Keep rows 3-6 and 8-10: the two exact copies (rows 2, 7) and the stale
    # Delhi row (row 1) are gone.
    kept = [(i, r) for i, r in enumerate(DUP_ROWS, 1) if i not in (1, 2, 7)]
    row = put_headers(ws, row, ["Customer", "Email", "Region", "City", "Email count"])
    rows = []
    for i, r in kept:
        n = dup["start"] + i - 1  # the record's original row on the Duplicates sheet
        rows.append(r + ["=COUNTIF(" + rng + ", Duplicates!B" + str(n) + ")"])
    end = put_rows(ws, row, rows)
    return end, row, None


def _clean_solution_table(ws, row, cl):
    """Finished cleaned customer list (static values, as after Paste Special > Values)."""
    cleaned = [
        ["Arjun Mehta", "arjun.mehta@sunrise.in", "05-Jan-2026", 120],
        ["Priya Nair", "priya.nair@sunrise.in", "05-Jan-2026", 45],
        ["Rahul Verma", "rahul.verma@sunrise.in", "05-Jan-2026", 8],
        ["Sneha Iyer", "sneha.iyer@sunrise.in", "07-Jan-2026", 60],
        ["Meera Krishnan", "meera.k@sunrise.in", "08-Jan-2026", 30],
        ["Vikram Shah", "vikram.shah@sunrise.in", "10-Jan-2026", 90],
    ]
    row = put_section(ws, row, "Finished customer list — after cleaning and Remove Duplicates (6 rows)", width=5)
    row = put_headers(ws, row, ["Customer", "Email", "Date", "Units"])
    end = put_rows(ws, row, cleaned, num_formats={4: "#,##0"})
    return end, row, None


def build_data_cleaning(path):
    wb = new_workbook("Duplicates", "Splitting", "FindReplace", "Cleaning", "Solution")
    facts = {
        "duplicates": build_duplicates(wb["Duplicates"]),
        "findreplace": build_findreplace(wb["FindReplace"]),
        "cleaning": build_cleaning(wb["Cleaning"]),
    }
    build_splitting(wb["Splitting"])
    build_dc_solution(wb["Solution"], facts)
    wb.save(path)
    return path

# __APPEND_7__

# ---------------------------------------------------------------------------
# large-dataset.xlsx — a deterministic 1,000-row sales dataset
# ---------------------------------------------------------------------------

PEOPLE = ["Arjun Mehta", "Priya Nair", "Rahul Verma", "Sneha Iyer",
          "Meera Krishnan", "Vikram Shah"]
REGIONS = ["North", "South", "East", "West"]
PRODUCTS = [("Notebooks", 120, 20, 300), ("Staplers", 150, 10, 120),
            ("Pens", 60, 20, 250), ("Laptops", 16000, 1, 25),
            ("Printers", 2400, 1, 15)]
FIRST_DAY = date(2026, 1, 1)
SPAN_DAYS = 304  # 01-Jan-2026 .. 31-Oct-2026


def generate_large_rows(n=1000):
    random.seed(20260301)
    rows = []
    for i in range(n):
        d = FIRST_DAY + timedelta(days=i * SPAN_DAYS // n)
        prod, price, lo, hi = random.choice(PRODUCTS)
        units = random.randint(lo, hi)
        rows.append([1001 + i, d, random.choice(PEOPLE), random.choice(REGIONS),
                     prod, units, units * price])
    return rows


def build_large_exercise(ws):
    set_widths(ws, [12, 100])
    row = put_title(ws, "Level 3 · Working with Large Datasets — Exercise",
                    "A 1,000-row Sunrise Traders sales list on the 'Large dataset' sheet")
    row = note(ws, row, (
        "The 'Large dataset' sheet holds 1,000 rows of sales data (A2:G1001) with headers in row 1. Work there, then note your results:\n"
        "1) Freeze the header row: View > Freeze Panes > Freeze Top Row. Click any Order ID cell and press Ctrl + Down Arrow — "
        "which row and which Order ID do you land on? (Ctrl + Home returns you to the top.)\n"
        "2) Convert the data to a Table named Sales (Ctrl + T, tick My table has headers, rename on the Table Design tab). "
        "Count the records with =COUNTA(Sales[Order ID]) in a free cell.\n"
        "3) Click G2 and press Ctrl + Shift + Down Arrow to select the whole Revenue column in one move. Add a live filtered total "
        "above the data in cell I1: =SUBTOTAL(9, Sales[Revenue]). Compare it with =SUM(Sales[Revenue]) — filter to one region and "
        "watch the SUBTOTAL follow the filter while the SUM stays fixed.\n"
        "4) The file feels slow while you edit: Formulas > Calculation Options > Manual, make several edits (nothing recalculates), "
        "press F9 to refresh all results, then switch back to Automatic before saving."
    ))
    row += 1
    row = answer_grid(ws, row,
                      ["Task", "Your result"],
                      ["Freeze + Ctrl + Down Arrow: row and Order ID you land on",
                       "=COUNTA(Sales[Order ID])",
                       "=SUBTOTAL(9, ...) vs =SUM(Sales[Revenue]) — what changed when you filtered?",
                       "Manual calculation: what did you check after F9?"])


def build_large_solution(ws, rows):
    set_widths(ws, [14, 110])
    last = rows[-1]
    grand = sum(r[6] for r in rows)
    per_region = {reg: sum(r[6] for r in rows if r[3] == reg) for reg in REGIONS}
    row = put_title(ws, "Solutions — Large dataset",
                    "Finished results for the four exercises (the lesson describes the solutions)")
    row += 1

    row = put_section(ws, row, "Exercise 1 — Freeze the header row and jump to the last record", width=2)
    row = sol_lines(ws, row, [
        "View > Freeze Panes > Freeze Top Row pins row 1 while you scroll.",
        "Click any Order ID cell, press Ctrl + Down Arrow: you land on row 1001 instantly — the last record,",
        "   Order ID " + str(last[0]) + " · " + last[1].strftime("%d-%b-%Y") + " · " + last[2] + " · " + last[3] + " · " +
        last[4] + " · " + str(last[5]) + " units · ₹" + format(last[6], ",") + " — with the headers still visible above.",
        "Ctrl + Home returns you to A1.",
    ])
    row += 1

    row = put_section(ws, row, "Exercise 2 — Table named Sales and the record count", width=2)
    row = sol_lines(ws, row, [
        "Ctrl + T over A1:G1001 (headers ticked), renamed to Sales on the Table Design tab.",
        "=COUNTA(Sales[Order ID]) shows 1000 — and stays correct as rows come and go.",
    ])
    row += 1

    row = put_section(ws, row, "Exercise 3 — Select the Revenue column and add the live total", width=2)
    row = sol_lines(ws, row, [
        "Click G2, Ctrl + Shift + Down Arrow selects the Revenue column to the end in one keystroke.",
        "In I1: =SUBTOTAL(9, Sales[Revenue]) with no filter = ₹" + format(grand, ",") + " — the same as =SUM(Sales[Revenue]).",
        "Filter to one region and the SUBTOTAL follows the filter while the SUM stays fixed:",
        "   North ₹" + format(per_region["North"], ",") + " · South ₹" + format(per_region["South"], ",") +
        " · East ₹" + format(per_region["East"], ",") + " · West ₹" + format(per_region["West"], ","),
    ])
    row += 1

    row = put_section(ws, row, "Exercise 4 — Manual calculation, then back to Automatic", width=2)
    row = sol_lines(ws, row, [
        "Formulas tab > Calculation Options > Manual — edits stop triggering recalculation.",
        "Press F9 to recalculate everything (Shift + F9 recalculates only the active sheet) and verify a few results.",
        "Switch back to Formulas > Calculation Options > Automatic BEFORE saving or sharing — colleagues must never trust stale numbers.",
    ])


def build_large_dataset(path):
    wb = new_workbook("Exercise", "Large dataset", "Solution")
    rows = generate_large_rows(1000)
    build_large_exercise(wb["Exercise"])

    ws = wb["Large dataset"]
    set_widths(ws, ORDER_WIDTHS)
    put_headers(ws, 1, ORDER_HEADERS)  # headers in row 1, data in A2:G1001
    put_rows(ws, 2, rows, num_formats={2: "dd-mmm-yyyy", 7: "#,##0"})
    ws.freeze_panes = None  # freezing row 1 is exercise 1 — leave it to the student

    build_large_solution(wb["Solution"], rows)
    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# Main + verification
# ---------------------------------------------------------------------------

EXPECTED = {
    "tables-validation.xlsx": ["Sorting", "Filtering", "Tables", "Validation",
                               "Conditional Formatting", "Solution"],
    "data-cleaning.xlsx": ["Duplicates", "Splitting", "FindReplace", "Cleaning", "Solution"],
    "large-dataset.xlsx": ["Exercise", "Large dataset", "Solution"],
}


def verify():
    ok = True
    for name, sheets in EXPECTED.items():
        path = OUT / name
        if not path.exists():
            print(f"FAIL  {name}: file missing")
            ok = False
            continue
        try:
            wb = load_workbook(path)
        except Exception as exc:  # noqa: BLE001
            print(f"FAIL  {name}: does not open ({exc})")
            ok = False
            continue
        missing = [s for s in sheets if s not in wb.sheetnames]
        extra = [s for s in wb.sheetnames if s not in sheets]
        size = path.stat().st_size
        status = "OK  " if not missing and not extra else "FAIL"
        if status == "FAIL":
            ok = False
        print(f"{status}  {name} ({size:,} bytes) — sheets: {', '.join(wb.sheetnames)}")
        if missing:
            print(f"      missing: {missing}")
        if extra:
            print(f"      unexpected: {extra}")
    others = sorted(p.name for p in OUT.glob("*.xlsx") if p.name not in EXPECTED)
    if others:
        print(f"FAIL  unexpected files in {OUT}: {others}")
        ok = False
    return ok


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    created = [
        build_tables_validation(OUT / "tables-validation.xlsx"),
        build_data_cleaning(OUT / "data-cleaning.xlsx"),
        build_large_dataset(OUT / "large-dataset.xlsx"),
    ]
    for path in created:
        print(f"created {path.relative_to(ROOT)}")
    print()
    if verify():
        print("\nAll Level 3 workbooks verified.")
    else:
        print("\nVerification FAILED — see messages above.")
        sys.exit(1)


if __name__ == "__main__":
    main()
