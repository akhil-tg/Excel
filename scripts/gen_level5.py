"""Generate every Level 5 practice workbook referenced by the lesson data.

Creates one Exercise/Solution workbook per file listed in the `downloads`
entries of assets/js/data/level5-1.js and level5-2.js:

    employee-records.xlsx      attendance-records.xlsx
    leave-tracker.xlsx         payroll-data.xlsx
    inventory-records.xlsx     purchase-records.xlsx
    sales-records.xlsx         budget-tracker.xlsx
    expense-tracker.xlsx       billing-records.xlsx
    mis-reports.xlsx           office-data-capstone.xlsx

All data mirrors the grid blocks and exercise solutions of the Level 5
lessons (Sunrise Traders, October/November 2026). Run from the repo root:

    .venv/Scripts/python.exe scripts/gen_level5.py
"""
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from xlcommon import (
    new_workbook, put_title, put_note, put_section, put_headers,
    put_rows, set_widths,
)

OUT_DIR = os.path.join("assets", "files", "level5")

MONEY = "#,##0"
PCT = "0.0%"            # fractions (0.833 -> 83.3%)
PCTNUM = '0.0"%"'       # ROUND(x*100,1) results (105.1 -> 105.1%)
DATE_FMT = "dd-mmm-yyyy"

# ---------------------------------------------------------------------------
# Shared Sunrise Traders data (from the lesson grids)
# ---------------------------------------------------------------------------

EMPLOYEES = [
    ("EMP-001", "Arjun Mehta", "Sales", "Sales Executive", date(2021, 4, 12), 25000, "Active"),
    ("EMP-002", "Priya Nair", "Sales", "Sales Manager", date(2020, 1, 3), 40000, "Active"),
    ("EMP-003", "Rahul Verma", "Accounts", "Accountant", date(2022, 7, 21), 28000, "Active"),
    ("EMP-004", "Sneha Iyer", "Stores", "Store Keeper", date(2023, 8, 15), 22000, "Active"),
    ("EMP-005", "Vikram Singh", "Admin", "Office Administrator", date(2022, 2, 9), 24000, "Active"),
    ("EMP-006", "Divya Menon", "HR", "HR Executive", date(2022, 9, 27), 26000, "Active"),
]


def _dv(ws, cell_range, values):
    """Attach a List data-validation dropdown to a range."""
    dv = DataValidation(type="list", formula1='"' + ",".join(values) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(cell_range)


# ---------------------------------------------------------------------------
# 1. employee-records.xlsx  (l5-employee-database)
# ---------------------------------------------------------------------------

def _employee(ws, sol):
    r = put_title(ws, "Sunrise Traders — Employee Master",
                  "Level 5 · Building an Employee Database")
    r = put_note(ws, r, "This master feeds every other Level 5 sheet. Tasks: add Data "
        "Validation dropdowns to Department (Sales, Accounts, Stores, Admin, HR) and Status "
        "(Active, Inactive); complete the headcount summary with COUNTIF (dollar-lock the "
        "ranges); then filter to everyone who joined before 2023 with Ctrl+Shift+L — only "
        "Sneha Iyer joined in 2023.", width=7)
    r = put_section(ws, r, "Employee master — one row per employee, no merged cells", width=7)
    r = put_headers(ws, r, ["Emp ID", "Name", "Department", "Designation",
                            "Joining Date", "Salary", "Status"])
    first = r
    r = put_rows(ws, r, [list(e) for e in EMPLOYEES], num_formats={5: DATE_FMT, 6: MONEY})
    last = r - 1
    _dv(ws, f"C{first}:C{last}", ["Sales", "Accounts", "Stores", "Admin", "HR"])
    _dv(ws, f"G{first}:G{last}", ["Active", "Inactive"])
    r += 1
    r = put_section(ws, r, "Headcount summary — build this with COUNTIF", width=2)
    r = put_headers(ws, r, ["Department", "Headcount"])
    rows = []
    for d in ("Sales", "Accounts", "Stores", "Admin", "HR"):
        rows.append([d, f'=COUNTIF($C${first}:$C${last},"{d}")' if sol else None])
    r = put_rows(ws, r, rows)
    r = put_rows(ws, r, [["Active employees",
                          f'=COUNTIF($G${first}:$G${last},"Active")' if sol else None]])
    set_widths(ws, [10, 16, 12, 20, 14, 10, 10])
    ws.freeze_panes = f"A{first}"


# ---------------------------------------------------------------------------
# 2. attendance-records.xlsx  (l5-attendance)
# ---------------------------------------------------------------------------

SUNDAYS = {4, 11, 18, 25}
WEEK_OFF_SAT = {10, 24}
HOLIDAYS = {2}

# Per-employee A/L marks; every other working day is P (each row: P+A+L = 24)
ATT_MARKS = {
    "EMP-001": {5: "A", 13: "A", 21: "L", 29: "L"},
    "EMP-002": {7: "L", 27: "L"},
    "EMP-003": {},
    "EMP-004": {9: "A", 15: "L", 30: "L"},
    "EMP-005": {19: "A"},
    "EMP-006": {6: "A", 14: "A", 23: "A", 28: "L"},
}


def _attendance(ws, sol):
    r = put_title(ws, "Sunrise Traders — Attendance Register",
                  "October 2026 · codes: P present, A absent, L leave, H holiday, WO week off")
    r = put_note(ws, r, "Day columns C to AG hold days 1-31. Sundays (4, 11, 18, 25) and the "
        "week-off Saturdays (10, 24) stay blank; 2-Oct is the Gandhi Jayanti holiday (H). Tasks: "
        "fill the P, A and L totals with COUNTIF (C:AG per row), add Attendance % = P/(P+A+L) — "
        "never P/31 — and shade weekends with two conditional-formatting rules: "
        "=WEEKDAY(DATE(2026,10,C$2),1)=1 and =OR(C$2=10,C$2=24). Every row's P+A+L must equal "
        "the 24 working days of October 2026.", width=12)
    r = put_headers(ws, r, ["Emp ID", "Name"] + [str(d) for d in range(1, 32)]
                    + ["P", "A", "L", "Attendance %"])
    first = r
    rows = []
    for emp_id, name, *_ in EMPLOYEES:
        marks = ATT_MARKS[emp_id]
        codes = []
        for d in range(1, 32):
            if d in SUNDAYS or d in WEEK_OFF_SAT:
                codes.append(None)
            elif d in HOLIDAYS:
                codes.append("H")
            else:
                codes.append(marks.get(d, "P"))
        rr = first + len(rows)
        totals = ([f'=COUNTIF(C{rr}:AG{rr},"P")', f'=COUNTIF(C{rr}:AG{rr},"A")',
                   f'=COUNTIF(C{rr}:AG{rr},"L")', f"=AH{rr}/(AH{rr}+AI{rr}+AJ{rr})"]
                  if sol else [None] * 4)
        rows.append([emp_id, name] + codes + totals)
    r = put_rows(ws, r, rows, num_formats={38: PCT})
    r += 1
    r = put_section(ws, r, "Monthly summary — the manager's one-page view", width=6)
    r = put_headers(ws, r, ["Emp ID", "Name", "P", "A", "L", "Attendance %"])
    srows = []
    for i, (emp_id, name, *_ ) in enumerate(EMPLOYEES):
        reg = first + i
        srows.append([emp_id, name, f"=AH{reg}", f"=AI{reg}", f"=AJ{reg}", f"=AK{reg}"]
                     if sol else [emp_id, name, None, None, None, None])
    r = put_rows(ws, r, srows, num_formats={6: PCT})
    set_widths(ws, [10, 18] + [4.5] * 31 + [5, 5, 5, 12])
    ws.freeze_panes = f"C{first}"


# ---------------------------------------------------------------------------
# 3. leave-tracker.xlsx  (l5-leave-management)
# ---------------------------------------------------------------------------

# LV-009 is the earlier approved 2-day casual leave the lesson refers to, so
# Priya Nair's Casual Taken = 2 + 3 = 5 and her balance = 12 - 5 = 7.
LEAVE_LOG = [
    ("LV-009", "EMP-002", "Priya Nair", "Casual", date(2026, 1, 12), date(2026, 1, 13), "Approved"),
    ("LV-013", "EMP-001", "Arjun Mehta", "Sick", date(2026, 9, 21), date(2026, 9, 22), "Approved"),
    ("LV-014", "EMP-002", "Priya Nair", "Casual", date(2026, 11, 5), date(2026, 11, 7), "Approved"),
    ("LV-015", "EMP-004", "Sneha Iyer", "Earned", date(2026, 12, 2), date(2026, 12, 9), "Pending"),
]


def _leave(ws, sol):
    r = put_title(ws, "Sunrise Traders — Leave Tracker",
                  "Level 5 · Leave Management System · policy: Casual 12, Sick 7, Earned 15 days a year")
    r = put_note(ws, r, "Tasks: complete the Leave Log Days column with the inclusive formula "
        "(= To − From + 1); build Casual Taken with SUMIFS on the log (Emp ID + leave type, only "
        "Status = Approved rows) and Casual Balance = Allowed − Taken. LV-015 is still Pending, "
        "so it must not touch any balance until a manager approves it.", width=7)
    r = put_section(ws, r, "Leave balances — allowed, taken, balance", width=7)
    r = put_headers(ws, r, ["Emp ID", "Name", "Casual Allowed", "Sick Allowed",
                            "Earned Allowed", "Casual Taken", "Casual Balance"])
    first = r
    r = put_rows(ws, r, [[e[0], e[1], 12, 7, 15, None, None] for e in EMPLOYEES],
                 num_formats={6: MONEY, 7: MONEY})
    r += 1
    r = put_section(ws, r, "Leave Log — every request gets one row and a running number",
                    width=8)
    r = put_headers(ws, r, ["Req No", "Emp ID", "Name", "Leave Type", "From", "To",
                            "Days", "Status"])
    lf = r
    rows = []
    for i, (req, emp_id, name, typ, d_from, d_to, status) in enumerate(LEAVE_LOG):
        rr = lf + i
        rows.append([req, emp_id, name, typ, d_from, d_to,
                     f"=F{rr}-E{rr}+1" if sol else None, status])
    r = put_rows(ws, r, rows, num_formats={5: DATE_FMT, 6: DATE_FMT})
    ll = r - 1
    if sol:  # balance formulas need the log range, so they are filled after it
        end = ll + 25
        for i in range(len(EMPLOYEES)):
            rr = first + i
            ws.cell(row=rr, column=6, value=(
                f'=SUMIFS($G${lf}:$G${end},$B${lf}:$B${end},A{rr},'
                f'$D${lf}:$D${end},"Casual",$H${lf}:$H${end},"Approved")'))
            ws.cell(row=rr, column=7, value=f"=C{rr}-F{rr}")
    _dv(ws, f"D{lf}:D{ll + 25}", ["Casual", "Sick", "Earned"])
    _dv(ws, f"H{lf}:H{ll + 25}", ["Pending", "Approved", "Rejected"])
    r += 1
    put_note(ws, r, "LV-009 is Priya Nair's earlier 2-day casual leave (Approved), so her "
        "Casual Taken is 2 + 3 = 5 and her balance is 12 − 5 = 7. LV-015 (Sneha Iyer, Earned, "
        "8 days) stays Pending and does not reduce any balance.", width=7)
    set_widths(ws, [10, 16, 13, 12, 13, 13, 9, 11])
    ws.freeze_panes = f"A{first}"


# ---------------------------------------------------------------------------
# 4. payroll-data.xlsx  (l5-payroll)
# ---------------------------------------------------------------------------

def _payroll(ws, sol):
    r = put_title(ws, "Sunrise Traders — Payroll, October 2026",
                  "Level 5 · Salary and Payroll Calculations · HRA 40% of Basic, Special 20%, "
                  "Conveyance 1,600 flat, PF 12% of Basic, Prof. Tax 200, TDS 5% of Gross above 40,000")
    r = put_note(ws, r, "Tasks: complete HRA, Conveyance and Special Allowance; add PF, "
        "Professional Tax and TDS; then compute Gross, Deductions and Net — wrap money in "
        "ROUND so stored and printed values agree. Finally build the payslip block for "
        "Divya Menon (EMP-006) from her payroll row.", width=12)
    r = put_headers(ws, r, ["Emp ID", "Name", "Basic", "HRA", "Conv.", "Special", "Gross",
                            "PF", "Prof Tax", "TDS", "Deductions", "Net"])
    first = r
    rows = []
    for i, (emp_id, name, *_ ) in enumerate(EMPLOYEES):
        rr = first + i
        if sol:
            vals = [f"=ROUND(C{rr}*0.4,0)", 1600, f"=ROUND(C{rr}*0.2,0)",
                    f"=C{rr}+D{rr}+E{rr}+F{rr}", f"=ROUND(C{rr}*0.12,0)", 200,
                    f"=IF(G{rr}>40000,ROUND(G{rr}*0.05,0),0)",
                    f"=H{rr}+I{rr}+J{rr}", f"=ROUND(G{rr}-K{rr},0)"]
        else:
            vals = [None] * 9
        rows.append([emp_id, name, EMPLOYEES[i][5]] + vals)
    nf = {c: MONEY for c in range(3, 13)}
    r = put_rows(ws, r, rows, num_formats=nf)
    last = r - 1
    trow = ["Total", ""]
    for c in range(3, 13):
        col = get_column_letter(c)
        trow.append(f"=SUM({col}{first}:{col}{last})" if sol else None)
    r = put_rows(ws, r, [trow], num_formats=nf)
    r += 1
    r = put_section(ws, r, "Payslip — one page per employee (print area, portrait)", width=4)
    pr = first + 5  # EMP-006 Divya Menon's payroll row
    slip = [
        ["", "Sunrise Traders", "Payslip for October 2026", ""],
        ["Employee", "Divya Menon (EMP-006)", "Designation", "HR Executive"],
        ["Earnings", "", "Deductions", ""],
        ["Basic", f"=C{pr}", "PF", f"=H{pr}"],
        ["HRA", f"=D{pr}", "Professional Tax", f"=I{pr}"],
        ["Conveyance", f"=E{pr}", "TDS", f"=J{pr}"],
        ["Special Allowance", f"=F{pr}", "Total Deductions", f"=K{pr}"],
        ["Gross", f"=G{pr}", "", ""],
        ["", "", "NET PAY", f"=L{pr}"],
    ]
    slip = [[v if (sol or not (isinstance(v, str) and v.startswith("="))) else None
             for v in row] for row in slip]
    r = put_rows(ws, r, slip, num_formats={2: MONEY, 4: MONEY})
    set_widths(ws, [17, 14, 12, 12, 10, 10, 10, 10, 10, 10, 12, 12])
    ws.freeze_panes = f"A{first}"


# ---------------------------------------------------------------------------
# 5. inventory-records.xlsx  (l5-inventory + l5-stock-tracking)
# ---------------------------------------------------------------------------

ITEMS = [
    ("ITM-101", "Notebook", "Stationery", "Nos", 120, 45, 200),
    ("ITM-102", "Pen", "Stationery", "Nos", 300, 15, 150),
    ("ITM-103", "Stapler", "Office Tools", "Nos", 40, 130, 20),
    ("ITM-104", "Printer", "Electronics", "Nos", 8, 8500, 8),
    ("ITM-105", "Laptop", "Electronics", "Nos", 5, 42000, 2),
]
INWARD = [
    (date(2026, 10, 3), "ITM-101", 100, 42),
    (date(2026, 10, 5), "ITM-102", 200, 13),
    (date(2026, 10, 8), "ITM-103", 25, 125),
    (date(2026, 10, 12), "ITM-101", 50, 44),
    (date(2026, 10, 15), "ITM-104", 3, 8200),
]
# The lesson grid shows the first rows; the remaining October movements make the
# outward register total the lesson's expected figures (75/140/12/3/1 units,
# value 74,535).
OUTWARD = [
    (date(2026, 10, 2), "ITM-101", 45),
    (date(2026, 10, 4), "ITM-102", 80),
    (date(2026, 10, 6), "ITM-104", 2),
    (date(2026, 10, 10), "ITM-101", 30),
    (date(2026, 10, 12), "ITM-102", 60),
    (date(2026, 10, 16), "ITM-103", 12),
    (date(2026, 10, 18), "ITM-105", 1),
    (date(2026, 10, 22), "ITM-104", 1),
]
VERDICTS = {
    "ITM-101": "Fast by quantity",
    "ITM-102": "Fastest by quantity",
    "ITM-103": "Slow",
    "ITM-104": "Slow by quantity, big by value",
    "ITM-105": "Slowest by quantity, biggest by value",
}


def _inventory(ws, sol):
    r = put_title(ws, "Sunrise Traders — Inventory Records",
                  "Level 5 · Inventory Management + Stock Tracking and Reorder Alerts")
    r = put_note(ws, r, "Item master, inward register and outward register. Tasks: compute "
        "Current Stock = Opening + SUMIFS(inward qty) − SUMIFS(outward qty); fill the Outward "
        "Rate with VLOOKUP from the item master; compute Value = Qty × Rate and Stock Value = "
        "Current × Unit Cost; add Status = IF(Current<=Reorder,\"ORDER NOW\",\"OK\"); then build "
        "the movement summary with RANK and find the top seller with INDEX-MATCH-LARGE.", width=10)
    r = put_section(ws, r, "Item Master", width=10)
    r = put_headers(ws, r, ["Item Code", "Item Name", "Category", "Unit", "Opening Stock",
                            "Unit Cost", "Reorder Level", "Current Stock", "Stock Value",
                            "Status"])
    mf = r
    rows = []
    for i, (code, name, cat, unit, opening, cost, reorder) in enumerate(ITEMS):
        rows.append([code, name, cat, unit, opening, cost, reorder, None, None, None])
    r = put_rows(ws, r, rows)
    ml = r - 1
    r = put_rows(ws, r, [["Total stock value", "", "", "", "", "", "", "",
                          f"=SUM(I{mf}:I{ml})" if sol else None, ""]])
    r += 1
    r = put_section(ws, r, "Inward register (goods received)", width=5)
    r = put_headers(ws, r, ["Date", "Item Code", "Qty", "Rate", "Value"])
    inf = r
    rows = []
    for i, (d, code, qty, rate) in enumerate(INWARD):
        rr = inf + i
        rows.append([d, code, qty, rate, f"=C{rr}*D{rr}" if sol else None])
    r = put_rows(ws, r, rows, num_formats={1: DATE_FMT, 5: MONEY})
    inl = r - 1
    r += 1
    r = put_section(ws, r, "Outward register (goods issued or sold)", width=5)
    r = put_headers(ws, r, ["Date", "Item Code", "Qty", "Rate", "Value"])
    of = r
    rows = []
    for i, (d, code, qty) in enumerate(OUTWARD):
        rr = of + i
        rows.append([d, code, qty,
                     f"=VLOOKUP(B{rr},$A${mf}:$F${ml},6,FALSE)" if sol else None,
                     f"=C{rr}*D{rr}" if sol else None])
    r = put_rows(ws, r, rows, num_formats={1: DATE_FMT, 4: MONEY, 5: MONEY})
    ol = r - 1
    if sol:  # current stock needs the register ranges, so fill after both exist
        for i, (code, *_ ) in enumerate(ITEMS):
            rr = mf + i
            ws.cell(row=rr, column=8, value=(
                f"=E{rr}+SUMIFS($C${inf}:$C${inl},$B${inf}:$B${inl},A{rr})"
                f"-SUMIFS($C${of}:$C${ol},$B${of}:$B${ol},A{rr})"))
            ws.cell(row=rr, column=9, value=f"=H{rr}*F{rr}")
            ws.cell(row=rr, column=10, value=f'=IF(H{rr}<=G{rr},"ORDER NOW","OK")')
    r += 1
    r = put_section(ws, r, "October movement — fast and slow movers, ABC ranks", width=6)
    r = put_headers(ws, r, ["Item Code", "Item", "Outward Qty", "Outward Value",
                            "Rank by Value", "Verdict"])
    sf = r
    rows = []
    for i, (code, name, *_ ) in enumerate(ITEMS):
        rr = sf + i
        rows.append([code, name,
                     f"=SUMIFS($C${of}:$C${ol},$B${of}:$B${ol},A{rr})" if sol else None,
                     f"=SUMIFS($E${of}:$E${ol},$B${of}:$B${ol},A{rr})" if sol else None,
                     f"=RANK(D{rr},$D${sf}:$D${sf + len(ITEMS) - 1})" if sol else None,
                     VERDICTS[code] if sol else None])
    r = put_rows(ws, r, rows, num_formats={4: MONEY})
    sl = r - 1
    put_rows(ws, r, [["Top seller by value (INDEX-MATCH-LARGE)", "", "",
                      f"=INDEX($B${sf}:$B${sl},MATCH(LARGE($D${sf}:$D${sl},1),"
                      f"$D${sf}:$D${sl},0))" if sol else None, "", ""]])
    set_widths(ws, [12, 11, 12, 8, 14, 11, 14, 14, 13, 12])
    ws.freeze_panes = f"A{mf}"


# ---------------------------------------------------------------------------
# 6. purchase-records.xlsx  (l5-purchase-management)
# ---------------------------------------------------------------------------

POS = [
    ("PO-2026-014", date(2026, 10, 2), "Sharma Stationery", "ITM-101 Notebook", 100, 42, "Received", "Paid"),
    ("PO-2026-015", date(2026, 10, 5), "Global Office Supplies", "ITM-104 Printer", 3, 8200, "Open", "Unpaid"),
    ("PO-2026-016", date(2026, 10, 8), "Kavya Paper Mart", "ITM-102 Pen", 200, 13, "Received", "Paid"),
    ("PO-2026-017", date(2026, 10, 11), "TechDesk Systems", "ITM-105 Laptop", 1, 41000, "Open", "Unpaid"),
    ("PO-2026-018", date(2026, 10, 14), "Sharma Stationery", "ITM-103 Stapler", 25, 125, "Cancelled", "NA"),
    ("PO-2026-019", date(2026, 10, 16), "Global Office Supplies", "ITM-104 Printer", 5, 8100, "Open", "Partial"),
]
SUPPLIERS = ["Sharma Stationery", "Global Office Supplies", "Kavya Paper Mart",
             "TechDesk Systems"]


def _purchase(ws, sol):
    r = put_title(ws, "Sunrise Traders — Purchase Order Register, October 2026",
                  "Level 5 · Purchase Management")
    r = put_note(ws, r, "Tasks: complete the Amount column (= Qty × Rate — never typed); add "
        "Status (Open, Received, Cancelled) and Payment (Paid, Unpaid, Partial, NA) dropdowns; "
        "build the status summary with COUNTIF and SUMIFS; compute supplier-wise totals (all "
        "POs vs Received only — cancelled orders must stay out of the money) and the Unpaid "
        "payables total. Expected: Open 3 POs / 1,06,100 · Received 2 / 6,800 · Cancelled 1 / "
        "3,125 · Unpaid 65,600.", width=9)
    r = put_headers(ws, r, ["PO No", "Date", "Supplier", "Item", "Qty", "Rate",
                            "Amount", "Status", "Payment"])
    first = r
    rows = []
    for i, (po, d, sup, item, qty, rate, status, pay) in enumerate(POS):
        rr = first + i
        rows.append([po, d, sup, item, qty, rate, f"=E{rr}*F{rr}" if sol else None,
                     status, pay])
    r = put_rows(ws, r, rows, num_formats={2: DATE_FMT, 7: MONEY})
    last = r - 1
    _dv(ws, f"H{first}:H{last + 25}", ["Open", "Received", "Cancelled"])
    _dv(ws, f"I{first}:I{last + 25}", ["Paid", "Unpaid", "Partial", "NA"])
    r += 1
    r = put_section(ws, r, "Status summary — the daily chase list", width=3)
    r = put_headers(ws, r, ["Status", "Count", "Value"])
    rows = []
    for s in ("Open", "Received", "Cancelled"):
        rows.append([s,
                     f'=COUNTIF($H${first}:$H${last},"{s}")' if sol else None,
                     f'=SUMIFS($G${first}:$G${last},$H${first}:$H${last},"{s}")' if sol else None])
    r = put_rows(ws, r, rows, num_formats={3: MONEY})
    r += 1
    r = put_section(ws, r, "Supplier-wise totals — quote received-only figures in talks", width=3)
    r = put_headers(ws, r, ["Supplier", "All POs", "Received only"])
    sfirst = r
    rows = []
    for i, s in enumerate(SUPPLIERS):
        rr = sfirst + i
        rows.append([s,
                     f"=SUMIFS($G${first}:$G${last},$C${first}:$C${last},A{rr})" if sol else None,
                     f'=SUMIFS($G${first}:$G${last},$C${first}:$C${last},A{rr},'
                     f'$H${first}:$H${last},"Received")' if sol else None])
    r = put_rows(ws, r, rows, num_formats={2: MONEY, 3: MONEY})
    r += 1
    r = put_section(ws, r, "Payment status — the accounts team's payment run", width=2)
    r = put_headers(ws, r, ["Measure", "Amount"])
    rows = [["Fully Unpaid (pay now)",
             f'=SUMIFS($G${first}:$G${last},$I${first}:$I${last},"Unpaid")'],
            ["Partially paid (balance due)",
             f'=SUMIFS($G${first}:$G${last},$I${first}:$I${last},"Partial")']]
    r = put_rows(ws, r, [[v[0], v[1] if sol else None] for v in rows], num_formats={2: MONEY})
    set_widths(ws, [14, 12, 24, 18, 8, 10, 12, 11, 10])
    ws.freeze_panes = f"A{first}"


# ---------------------------------------------------------------------------
# 7. sales-records.xlsx  (l5-sales-management)
# ---------------------------------------------------------------------------

SALES_LOG = [
    (date(2026, 10, 2), "INV-2026-301", "Mehta Book House", "Notebook", 500, 60, "Arjun Mehta"),
    (date(2026, 10, 3), "INV-2026-302", "Northstar Retail", "Printer", 10, 10500, "Priya Nair"),
    (date(2026, 10, 6), "INV-2026-303", "Bright Office Co.", "Pen", 2000, 20, "Rahul Verma"),
    (date(2026, 10, 9), "INV-2026-304", "Shree Ganesh Traders", "Stapler", 400, 175, "Sneha Iyer"),
]
# Target vs actual, October 2026 (actuals are the lesson's figures for the
# complete October log; the rows above are its first four invoices).
TARGETS = [
    ("Arjun Mehta", 800000, 925000),
    ("Priya Nair", 1000000, 1084000),
    ("Rahul Verma", 600000, 582000),
    ("Sneha Iyer", 500000, 467500),
]
CUSTOMERS = [
    ("Mehta Book House", 960000),
    ("Northstar Retail", 771500),
    ("Bright Office Co.", 745000),
    ("Shree Ganesh Traders", 582000),
]


def _sales(ws, sol):
    r = put_title(ws, "Sunrise Traders — Sales Log and Review, October 2026",
                  "Level 5 · Sales Management and Tracking")
    r = put_note(ws, r, "The log holds the October invoices row by row with Total = Qty × Rate "
        "and a Salesperson dropdown. Month-level figures (October 30,58,500; September "
        "26,84,500; the per-person actuals and customer totals) come from the complete October "
        "log and are filled in on the Solution sheet. Tasks: compute the 2-October daily total "
        "with SUMIFS and DATE(2026,10,2); build Target vs Actual with Achievement % = "
        "Actual/Target; find the top performer with MAX-MATCH-INDEX; reconcile the customer "
        "totals to 30,58,500.", width=8)
    r = put_section(ws, r, "Sales Log — one row per invoice", width=8)
    r = put_headers(ws, r, ["Date", "Invoice No", "Customer", "Product", "Qty", "Rate",
                            "Total", "Salesperson"])
    first = r
    rows = []
    for i, (d, inv, cust, prod, qty, rate, person) in enumerate(SALES_LOG):
        rr = first + i
        rows.append([d, inv, cust, prod, qty, rate, f"=E{rr}*F{rr}" if sol else None, person])
    r = put_rows(ws, r, rows, num_formats={1: DATE_FMT, 6: MONEY, 7: MONEY})
    last = r - 1
    _dv(ws, f"H{first}:H{last + 25}",
        ["Arjun Mehta", "Priya Nair", "Rahul Verma", "Sneha Iyer"])
    r += 1
    r = put_section(ws, r, "Monthly totals and growth", width=2)
    r = put_headers(ws, r, ["Measure", "Value"])
    mfirst = r
    rows = [["September 2026 total revenue", 2684500],
            ["October 2026 total revenue", 3058500],
            ["Growth % (Oct vs Sep) = (Oct − Sep) / Sep",
             (f"=(B{mfirst + 1}-B{mfirst})/B{mfirst}" if sol else None)],
            ["Total for 2 October 2026 (from this log)",
             f"=SUMIFS($G${first}:$G${last},$A${first}:$A${last},DATE(2026,10,2))"
             if sol else None]]
    r = put_rows(ws, r, rows, num_formats={2: MONEY})
    ws.cell(row=mfirst + 2, column=2).number_format = PCT
    r += 1
    r = put_section(ws, r, "Target vs actual — the monthly review screen", width=4)
    r = put_headers(ws, r, ["Salesperson", "Target", "Actual", "Achievement %"])
    tfirst = r
    rows = []
    for i, (name, target, actual) in enumerate(TARGETS):
        rr = tfirst + i
        rows.append([name, target, actual if sol else None,
                     f"=D{rr}/C{rr}" if sol else None])
    r = put_rows(ws, r, rows, num_formats={2: MONEY, 3: MONEY, 4: PCT})
    tlast = r - 1
    r = put_rows(ws, r, [["Top performer (October)",
                          f"=INDEX($A${tfirst}:$A${tlast},MATCH(MAX($D${tfirst}:$D${tlast}),"
                          f"$D${tfirst}:$D${tlast},0))" if sol else None]])
    r += 1
    r = put_section(ws, r, "Customer-wise totals — must reconcile to the month total", width=2)
    r = put_headers(ws, r, ["Customer", "Oct revenue"])
    cfirst = r
    rows = []
    for i, (cust, rev) in enumerate(CUSTOMERS):
        rows.append([cust, rev if sol else None])
    r = put_rows(ws, r, rows, num_formats={2: MONEY})
    r = put_rows(ws, r, [["Total (must equal 30,58,500)",
                          f"=SUM(B{cfirst}:B{cfirst + len(CUSTOMERS) - 1})" if sol else None]],
                 num_formats={2: MONEY})
    set_widths(ws, [40, 16, 20, 12, 10, 10, 12, 18])
    ws.freeze_panes = f"A{first}"


# ---------------------------------------------------------------------------
# 8. budget-tracker.xlsx  (l5-budget-preparation)
# ---------------------------------------------------------------------------

BUDGET = [
    ("Office Rent", 30000, 30000),
    ("Salaries (Net)", 240760, 240760),
    ("Electricity", 8500, 9100),
    ("Internet and Phone", 2500, 2650),
    ("Transport and Delivery", 6000, 5400),
    ("Stationery", 3000, 3200),
    ("Marketing", 8000, 7600),
    ("Miscellaneous", 5000, 5900),
]
EXPLAIN = [
    ["Electricity", 600, "+7.1%", "Extra festival-week working days in the shop"],
    ["Internet and Phone", 150, "+6.0%", "One extra data top-up"],
    ["Stationery", 200, "+6.7%", "Bulk pen order for November invoicing"],
    ["Miscellaneous", 900, "+18.0%", "Emergency printer repair"],
    ["Transport and Delivery", -600, "-10.0%", "Two deliveries combined per trip"],
    ["Marketing", -400, "-5.0%", "One newspaper insert postponed to November"],
]


def _budget(ws, sol):
    r = put_title(ws, "Sunrise Traders — Budget vs Actual, October 2026",
                  "Level 5 · Monthly Budget Preparation")
    r = put_note(ws, r, "Tasks: complete the Variance column (= Actual − Budget, never typed) "
        "and the Status column (=IF(Variance>0,\"Over\",IF(Variance<0,\"Under\",\"On budget\"))) "
        "for all eight expense heads; compute the sales variance and achievement against the "
        "29,10,000 target; then apply the 5% rule — every head off budget by more than 5% of "
        "its own budget needs a written explanation line (Solution sheet). Expected totals: "
        "budget 3,03,760 · actual 3,04,610 · variance +850.", width=5)
    r = put_section(ws, r, "Expense block — red means over, green means under", width=5)
    r = put_headers(ws, r, ["Expense Head", "Budget", "Actual", "Variance", "Status"])
    first = r
    rows = []
    for i, (head, budget, actual) in enumerate(BUDGET):
        rr = first + i
        rows.append([head, budget, actual, f"=C{rr}-B{rr}" if sol else None,
                     f'=IF(D{rr}>0,"Over",IF(D{rr}<0,"Under","On budget"))' if sol else None])
    r = put_rows(ws, r, rows, num_formats={2: MONEY, 3: MONEY, 4: MONEY})
    last = r - 1
    tr = r
    r = put_rows(ws, r, [["Total",
                          f"=SUM(B{first}:B{last})" if sol else None,
                          f"=SUM(C{first}:C{last})" if sol else None,
                          f"=C{tr}-B{tr}" if sol else None,
                          f'=IF(D{tr}>0,"Over",IF(D{tr}<0,"Under","On budget"))' if sol else None]],
                 num_formats={2: MONEY, 3: MONEY, 4: MONEY})
    r += 1
    r = put_section(ws, r, "Income block — for income, over budget is good news", width=5)
    r = put_headers(ws, r, ["Income Head", "Budget", "Actual", "Variance", "Achievement %"])
    ir = r
    r = put_rows(ws, r, [["Sales", 2910000, 3058500,
                          f"=C{ir}-B{ir}" if sol else None,
                          f"=ROUND(C{ir}/B{ir}*100,1)" if sol else None]],
                 num_formats={2: MONEY, 3: MONEY, 4: MONEY})
    ws.cell(row=ir, column=5).number_format = PCTNUM
    if sol:
        r += 1
        r = put_section(ws, r, "Heads needing a written explanation "
                        "(variance beyond 5% of the head's own budget)", width=4)
        r = put_headers(ws, r, ["Head", "Variance", "% of head budget", "Explanation line"])
        r = put_rows(ws, r, EXPLAIN, num_formats={2: MONEY})
    set_widths(ws, [24, 12, 12, 12, 14])
    ws.freeze_panes = f"A{first}"


# ---------------------------------------------------------------------------
# 9. expense-tracker.xlsx  (l5-expense-tracker)
# ---------------------------------------------------------------------------

V = [
    (date(2026, 11, 2), "EXP-2026-031", "Vikram Singh", "Stationery", "Printer paper 10 reams", 3200, "Cash", "Y"),
    (date(2026, 11, 3), "EXP-2026-032", "Divya Menon", "Travel", "Bus fare - bank visit", 480, "Cash", "N"),
    (date(2026, 11, 5), "EXP-2026-033", "Sneha Iyer", "Fuel", "Delivery van diesel", 4600, "Card", "Y"),
    (date(2026, 11, 7), "EXP-2026-034", "Vikram Singh", "Meals", "Staff tea and biscuits", 850, "Cash", "Y"),
    (date(2026, 11, 9), "EXP-2026-035", "Rahul Verma", "Repairs", "Printer service", 1500, "UPI", "Y"),
    (date(2026, 11, 10), "EXP-2026-036", "Vikram Singh", "Stationery", "Pens and registers", 1750, "Cash", "Y"),
    (date(2026, 11, 12), "EXP-2026-037", "Divya Menon", "Travel", "Courier to customer", 350, "UPI", "Y"),
    (date(2026, 11, 14), "EXP-2026-038", "Sneha Iyer", "Fuel", "Delivery van diesel", 4400, "Card", "Y"),
    (date(2026, 11, 17), "EXP-2026-039", "Vikram Singh", "Repairs", "Tube light replacement", 620, "Cash", "Y"),
    (date(2026, 11, 19), "EXP-2026-040", "Divya Menon", "Postage", "Speed post invoices", 260, "Cash", "Y"),
]
CATEGORIES = ["Stationery", "Fuel", "Repairs", "Travel", "Meals", "Postage"]


def _expense(ws, sol):
    r = put_title(ws, "Sunrise Traders — Expense and Reimbursement Tracker",
                  "November 2026 (first half) · vouchers EXP-2026-031 to EXP-2026-040 · "
                  "Level 5 · Office Expense and Reimbursement Tracker")
    r = put_note(ws, r, "One row per voucher. Tasks: build the category summary with SUMIFS "
        "(rows 2-11 style ranges); build the reimbursement sheet by Paid By; add mode-wise "
        "totals; filter Bill? = N to chase missing bills. All three reports must total the "
        "same as the log — 18,010 — and that equality is the reconciliation check.", width=8)
    r = put_section(ws, r, "Expense log — one row per voucher, amounts as plain numbers",
                    width=8)
    r = put_headers(ws, r, ["Date", "Voucher No", "Paid By", "Category", "Description",
                            "Amount", "Mode", "Bill?"])
    first = r
    r = put_rows(ws, r, [list(row) for row in V],
                 num_formats={1: DATE_FMT, 6: MONEY})
    last = r - 1
    _dv(ws, f"C{first}:C{last + 25}", [e[1] for e in EMPLOYEES])
    _dv(ws, f"D{first}:D{last + 25}",
        ["Travel", "Stationery", "Fuel", "Meals", "Repairs", "Postage", "Other"])
    _dv(ws, f"G{first}:G{last + 25}", ["Cash", "UPI", "Card"])
    _dv(ws, f"H{first}:H{last + 25}", ["Y", "N"])
    r += 1
    r = put_section(ws, r, "Category summary — SUMIFS on the log", width=3)
    r = put_headers(ws, r, ["Category", "Total", "Share of spend"])
    cfirst = r
    ctot = cfirst + len(CATEGORIES)
    rows = []
    for i, cat in enumerate(CATEGORIES):
        rr = cfirst + i
        rows.append([cat,
                     f'=SUMIFS($F${first}:$F${last},$D${first}:$D${last},"{cat}")'
                     if sol else None,
                     f"=ROUND(B{rr}/$B${ctot}*100,1)" if sol else None])
    r = put_rows(ws, r, rows, num_formats={2: MONEY})
    r = put_rows(ws, r, [["Total", f"=SUM(B{cfirst}:B{ctot - 1})" if sol else None,
                          f"=ROUND(B{ctot}/$B${ctot}*100,1)" if sol else None]],
                 num_formats={2: MONEY})
    ws.cell(row=ctot, column=3).number_format = PCTNUM
    r += 1
    r = put_section(ws, r, "Reimbursement sheet — who gets paid back", width=2)
    r = put_headers(ws, r, ["Employee", "To Reimburse"])
    rfirst = r
    rows = []
    for emp in ("Sneha Iyer", "Vikram Singh", "Rahul Verma", "Divya Menon"):
        rr = rfirst + len(rows)
        rows.append([emp,
                     f'=SUMIFS($F${first}:$F${last},$C${first}:$C${last},"{emp}")'
                     if sol else None])
    r = put_rows(ws, r, rows, num_formats={2: MONEY})
    r = put_rows(ws, r, [["Total", f"=SUM(B{rfirst}:B{rfirst + 3})" if sol else None]],
                 num_formats={2: MONEY})
    r += 1
    r = put_section(ws, r, "Mode-wise totals — plan the petty-cash refill", width=3)
    r = put_headers(ws, r, ["Mode", "Total", "Vouchers"])
    rows = []
    for mode in ("Cash", "UPI", "Card"):
        rows.append([mode,
                     f'=SUMIFS($F${first}:$F${last},$G${first}:$G${last},"{mode}")'
                     if sol else None,
                     f'=COUNTIF($G${first}:$G${last},"{mode}")' if sol else None])
    r = put_rows(ws, r, rows, num_formats={2: MONEY})
    r = put_rows(ws, r, [["Vouchers with Bill? = N (chase before reimbursement day)",
                          "EXP-2026-032 — Divya Menon, bus fare 480" if sol else None]])
    set_widths(ws, [46, 16, 16, 13, 26, 10, 8, 8])
    ws.freeze_panes = f"A{first}"


# ---------------------------------------------------------------------------
# 10. billing-records.xlsx  (l5-invoice-billing)
# ---------------------------------------------------------------------------

PRODUCTS = [("Notebook", 60), ("Ball Pen (box of 100)", 20), ("Stapler", 175),
            ("HP Laser Printer", 10500), ("A4 Paper (box of 5)", 320),
            ("Ink Cartridge", 850)]
INVOICES = [
    {"no": "INV-2026-350", "d": date(2026, 11, 5), "cust": "Northstar Retail",
     "items": [("HP Laser Printer", 2), ("Ink Cartridge", 10), ("A4 Paper (box of 5)", 5)],
     "disc": 1100},
    {"no": "INV-2026-351", "d": date(2026, 11, 8), "cust": "Mehta Book House",
     "items": [("Notebook", 500), ("Ball Pen (box of 100)", 2000)],
     "disc": 2000},
]


def _billing(ws, sol):
    r = put_title(ws, "Sunrise Traders — Billing Records, November 2026",
                  "Level 5 · Billing and Invoice Creation · 18% GST on the taxable value "
                  "(after discount), ROUND(x, 2) on tax and total")
    r = put_note(ws, r, "Tasks: verify invoice INV-2026-350 line by line (subtotal 31,100; "
        "taxable 30,000; GST 5,400; total 35,400); build INV-2026-351 for Mehta Book House "
        "(500 Notebooks and 2,000 Ball Pen boxes, 2,000 discount → total 80,240) pulling every "
        "rate from the product master with exact-match VLOOKUP; then add both invoices to the "
        "sales log and prove the log reconciles to 1,15,640.", width=8)
    r = put_section(ws, r, "Products — the master (rates live here, never on the invoice)",
                    width=2)
    r = put_headers(ws, r, ["Product", "Rate"])
    pf = r
    r = put_rows(ws, r, [list(p) for p in PRODUCTS], num_formats={2: MONEY})
    pl = r - 1
    total_cells = []
    for inv in INVOICES:
        r += 1
        r = put_section(ws, r, f"Invoice {inv['no']} — {inv['cust']}", width=4)
        r = put_rows(ws, r, [["Invoice No", inv["no"]], ["Date", inv["d"]],
                             ["Bill To", inv["cust"]]], num_formats={2: DATE_FMT})
        r = put_headers(ws, r, ["Item", "Qty", "Rate", "Amount"])
        i1 = r
        rows = []
        for i, (item, qty) in enumerate(inv["items"]):
            rr = i1 + i
            rows.append([item, qty,
                         f"=VLOOKUP(A{rr},$A${pf}:$B${pl},2,FALSE)" if sol else None,
                         f"=B{rr}*C{rr}" if sol else None])
        i2 = i1 + len(inv["items"]) - 1
        r = put_rows(ws, r, rows, num_formats={3: MONEY, 4: MONEY})
        sub, disc, tax, gst, tot = r, r + 1, r + 2, r + 3, r + 4
        rows = [["Subtotal", "", "", f"=SUM(D{i1}:D{i2})" if sol else None],
                ["Discount", "", "", inv["disc"]],
                ["Taxable Value", "", "", f"=D{sub}-D{disc}" if sol else None],
                ["GST @ 18%", "", "", f"=ROUND(D{tax}*0.18,2)" if sol else None],
                ["Invoice Total", "", "", f"=ROUND(D{tax}+D{gst},2)" if sol else None]]
        r = put_rows(ws, r, rows, num_formats={4: MONEY})
        total_cells.append(tot)
    r += 1
    r = put_section(ws, r, "Sales log — November so far (one row per invoice)", width=4)
    r = put_headers(ws, r, ["Date", "Invoice No", "Customer", "Total"])
    lf = r
    rows = []
    for inv, tot_row in zip(INVOICES, total_cells):
        rr = lf + len(rows)
        rows.append([inv["d"], inv["no"], inv["cust"],
                     f"=D{tot_row}" if sol else None])
    r = put_rows(ws, r, rows, num_formats={1: DATE_FMT, 4: MONEY})
    r = put_rows(ws, r, [["Total (must equal the sum of invoice grand totals)", "", "",
                          f"=SUM(D{lf}:D{lf + len(INVOICES) - 1})" if sol else None]],
                 num_formats={4: MONEY})
    set_widths(ws, [44, 12, 12, 14])
    ws.freeze_panes = f"A{pf}"


# ---------------------------------------------------------------------------
# 11. mis-reports.xlsx  (l5-mis-reports)
# ---------------------------------------------------------------------------

HEADLINE = [
    ("Total sales", 3058500, "Target 2910000"),
    ("Gross payroll", 271600, "Budget head"),
    ("Net paid out", 240760, "Bank transfer"),
    ("Office expenses", 304610, "Budget 303760"),
]
MIS_PERSON = [("Arjun Mehta", 735000, 850000), ("Priya Nair", 1000000, 1084000),
              ("Rahul Verma", 750000, 727200), ("Sneha Iyer", 425000, 397300)]
MIS_PRODUCT = [("Notebook", 1240000), ("A4 Paper (box of 5)", 680000),
               ("HP Laser Printer", 520000), ("Ink Cartridge", 360000),
               ("Ball Pen (box of 100)", 158500), ("Stapler", 100000)]
MIS_OFFICE = [("Office Rent", 30000, 30000), ("Electricity", 8500, 9100),
              ("Internet and Phone", 2500, 2650), ("Transport and Delivery", 6000, 5400),
              ("Stationery", 3000, 3200), ("Marketing", 8000, 7600),
              ("Miscellaneous", 5000, 5900)]


def _mis(ws, sol):
    r = put_title(ws, "Sunrise Traders — MIS, October 2026",
                  "Level 5 · Management Reports and MIS · one sheet that answers the month")
    r = put_note(ws, r, "Every actual is computed or reconciled — nothing is retyped. Tasks: "
        "complete the person-wise table (achievement = ROUND(actual/target*100, 1), status = "
        "IF(achievement>=100,\"On target\",\"Below\")); build the product-wise report and prove "
        "it reconciles to the sales log total 30,58,500; complete the office-expense block and "
        "state the two headline readings (sales 105.1% of target; office expense vs the "
        "3,03,760 budget).", width=4)
    r = put_section(ws, r, "Headline block — the whole month in a few rows", width=4)
    r = put_headers(ws, r, ["Measure", "Value", "Against", "Reading"])
    hfirst = r
    rows = []
    for i, (measure, value, against) in enumerate(HEADLINE):
        rr = hfirst + i
        reading = {0: f"=ROUND(B{rr}/2910000*100,1)",
                   3: f"=ROUND(B{rr}/303760*100,1)"}.get(i, "" if i else None)
        if i in (1, 2):
            reading = "Matches payroll sheet" if i == 1 else "Matches payslip totals"
            reading = reading if sol else None
        rows.append([measure, value, against, reading if sol else None])
    r = put_rows(ws, r, rows, num_formats={2: MONEY, 4: PCTNUM})
    r += 1
    r = put_section(ws, r, "Sales by person — October 2026", width=5)
    r = put_headers(ws, r, ["Salesperson", "Target", "Actual", "Achievement %", "Status"])
    pfirst = r
    rows = []
    for i, (name, target, actual) in enumerate(MIS_PERSON):
        rr = pfirst + i
        rows.append([name, target, actual if sol else None,
                     f"=ROUND(C{rr}/B{rr}*100,1)" if sol else None,
                     f'=IF(D{rr}>=100,"On target","Below")' if sol else None])
    r = put_rows(ws, r, rows, num_formats={2: MONEY, 3: MONEY, 4: PCTNUM})
    ptot = r
    r = put_rows(ws, r, [["Total", f"=SUM(B{pfirst}:B{ptot - 1})" if sol else None,
                          f"=SUM(C{pfirst}:C{ptot - 1})" if sol else None,
                          f"=ROUND(C{ptot}/B{ptot}*100,1)" if sol else None, ""]],
                 num_formats={2: MONEY, 3: MONEY, 4: PCTNUM})
    r += 1
    r = put_section(ws, r, "Sales by product — October 2026", width=3)
    r = put_headers(ws, r, ["Product", "Total", "Share"])
    dfirst = r
    dtot = dfirst + len(MIS_PRODUCT)
    rows = []
    for i, (prod, total) in enumerate(MIS_PRODUCT):
        rr = dfirst + i
        rows.append([prod, total if sol else None,
                     f"=ROUND(B{rr}/$B${dtot}*100,1)" if sol else None])
    r = put_rows(ws, r, rows, num_formats={2: MONEY})
    r = put_rows(ws, r, [["Total", f"=SUM(B{dfirst}:B{dtot - 1})" if sol else None, ""]],
                 num_formats={2: MONEY})
    ws.cell(row=dtot, column=3).value = None
    r += 1
    r = put_section(ws, r, "Office expense by head — October 2026 (salaries stay on the "
                    "headline block: gross 2,71,600 booked, net 2,40,760 paid)", width=4)
    r = put_headers(ws, r, ["Head", "Budget", "Actual", "Variance"])
    ofirst = r
    rows = []
    for i, (head, budget, actual) in enumerate(MIS_OFFICE):
        rr = ofirst + i
        rows.append([head, budget, actual, f"=C{rr}-B{rr}" if sol else None])
    r = put_rows(ws, r, rows, num_formats={2: MONEY, 3: MONEY, 4: MONEY})
    otot = r
    r = put_rows(ws, r, [["Operating total",
                          f"=SUM(B{ofirst}:B{otot - 1})" if sol else None,
                          f"=SUM(C{ofirst}:C{otot - 1})" if sol else None,
                          f"=C{otot}-B{otot}" if sol else None]],
                 num_formats={2: MONEY, 3: MONEY, 4: MONEY})
    if sol:
        r += 1
        r = put_section(ws, r, "Reconciliation checks — every block must agree with its "
                        "source before the MIS is circulated", width=2)
        r = put_headers(ws, r, ["Check", "Result"])
        rows = [["Person report total − month sales total (30,58,500)",
                 f"=C{ptot}-3058500"],
                ["Product report total − month sales total (30,58,500)",
                 f"=B{dtot}-3058500"],
                ["Operating expense actual + Salaries net (2,40,760) − budget actual total "
                 "(3,04,610)", f"=C{otot}+240760-304610"]]
        r = put_rows(ws, r, rows, num_formats={2: MONEY})
    set_widths(ws, [56, 14, 14, 14, 12])
    ws.freeze_panes = f"A{hfirst}"


# ---------------------------------------------------------------------------
# 12. office-data-capstone.xlsx  (l5-office-capstone)
# ---------------------------------------------------------------------------

KEYS_TABLE = [
    ("Employee Master", "Emp ID (EMP-001...)", "Attendance, Leave, Payroll"),
    ("Attendance (per month)", "Emp ID + the month sheet", "Payroll working days"),
    ("Leave Log", "Req No (LV-...)", "Leave balance sheet"),
    ("Sales Log", "Invoice No (INV-2026-...)", "MIS, budget income"),
    ("Expense Log", "Voucher No (EXP-2026-...)", "Reimbursement, budget actuals"),
    ("Product Master", "Product name", "Invoices, inventory"),
    ("MIS", "Formulas only — no typed numbers", "The owner's monthly read"),
]
CHECKLIST = [
    ("Sales log total = MIS person report total", "=3058500-3058500"),
    ("Sales log total = MIS product report total", "=3058500-3058500"),
    ("Payroll gross = MIS salary expense", "=271600-271600"),
    ("Payroll net = budget salaries actual", "=240760-240760"),
    ("Budget expense total = expense actuals", "=304610-304610"),
]


def _capstone(ws, sol):
    r = put_title(ws, "Sunrise Traders — Office Data Capstone",
                  "October 2026 closed · November 2026 open · Level 5 · Office Data "
                  "Management Capstone")
    r = put_note(ws, r, "One workbook, one sheet per job, linked by keys. Tasks: run the full "
        "October reconciliation checklist and show every check as zero; post November's two "
        "invoices (INV-2026-350: 35,400 and INV-2026-351: 80,240) and reconcile the partial "
        "month (1,15,640 so far; expense log first half 18,010 with one bill still to arrive). "
        "To start November: archive as Sunrise-Traders-2026-10-CLOSED.xlsx, work in a fresh "
        "copy with cleared transaction logs, and continue the series at INV-2026-350 / "
        "EXP-2026-031.", width=6)
    r = put_section(ws, r, "One set of keys — every sheet traces back to a master or a "
                    "numbered document", width=3)
    r = put_headers(ws, r, ["Sheet", "Its key", "Feeds"])
    r = put_rows(ws, r, [list(t) for t in KEYS_TABLE])
    r += 1
    r = put_section(ws, r, "October close — reconciliation checklist (each subtraction "
                    "must show zero)", width=3)
    r = put_headers(ws, r, ["Check", "Formula", "Result"])
    rows = []
    for check, formula in CHECKLIST:
        rows.append([check, formula if sol else None, "0 — pass" if sol else None])
    r = put_rows(ws, r, rows, num_formats={2: MONEY})
    r += 1
    r = put_section(ws, r, "November 2026 — partial month, noted not closed", width=4)
    r = put_headers(ws, r, ["Date", "Invoice No", "Customer", "Total"])
    lf = r
    rows = [[date(2026, 11, 5), "INV-2026-350", "Northstar Retail", 35400],
            [date(2026, 11, 8), "INV-2026-351", "Mehta Book House", 80240]]
    r = put_rows(ws, r, rows, num_formats={1: DATE_FMT, 4: MONEY})
    r = put_rows(ws, r, [["Sales log total (November so far)", "", "",
                          f"=SUM(D{lf}:D{lf + 1})" if sol else None]],
                 num_formats={4: MONEY})
    r = put_headers(ws, r, ["Item", "Value"])
    rows = [["Expense log first-half total (EXP-2026-031 to EXP-2026-040)", 18010],
            ["Vouchers logged", 10],
            ["Bill still to arrive",
             "EXP-2026-032 — Divya Menon, bus fare 480"],
            ["October archive (closed, never edited again)",
             "Sunrise-Traders-2026-10-CLOSED.xlsx"],
            ["November working copy", "Sunrise-Traders-2026-11.xlsx"],
            ["November series starts", "Invoices at INV-2026-350 · vouchers at EXP-2026-031"]]
    r = put_rows(ws, r, [[a, b if sol else None] for a, b in rows], num_formats={2: MONEY})
    set_widths(ws, [50, 22, 20, 14])
    ws.freeze_panes = "A5"


# ---------------------------------------------------------------------------
# Driver
# ---------------------------------------------------------------------------

BUILDERS = [
    ("employee-records.xlsx", _employee),
    ("attendance-records.xlsx", _attendance),
    ("leave-tracker.xlsx", _leave),
    ("payroll-data.xlsx", _payroll),
    ("inventory-records.xlsx", _inventory),
    ("purchase-records.xlsx", _purchase),
    ("sales-records.xlsx", _sales),
    ("budget-tracker.xlsx", _budget),
    ("expense-tracker.xlsx", _expense),
    ("billing-records.xlsx", _billing),
    ("mis-reports.xlsx", _mis),
    ("office-data-capstone.xlsx", _capstone),
]


def verify():
    print("\nVerification:")
    ok = True
    for fname, _ in BUILDERS:
        path = os.path.join(OUT_DIR, fname)
        try:
            wb = load_workbook(path)
            assert wb.sheetnames == ["Exercise", "Solution"], \
                f"sheets are {wb.sheetnames}"
            print(f"  OK  {path}  ({os.path.getsize(path):,} bytes, "
                  f"sheets: {', '.join(wb.sheetnames)})")
        except Exception as exc:  # noqa: BLE001 - report every failure
            ok = False
            print(f"  FAIL  {path}: {exc}")
    if ok:
        print(f"All {len(BUILDERS)} Level 5 workbooks generated and verified.")
    return ok


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    for fname, build in BUILDERS:
        wb = new_workbook("Exercise", "Solution")
        for sheet, sol in (("Exercise", False), ("Solution", True)):
            build(wb[sheet], sol)
        path = os.path.join(OUT_DIR, fname)
        wb.save(path)
        print(f"created {path}")
    if not verify():
        sys.exit(1)


if __name__ == "__main__":
    main()
