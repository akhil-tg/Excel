"""Generate Level 7 (Practical Projects) practice workbooks.

Creates exercise + solution pairs for all ten Level 7 project lessons:

  1.  Employee Management System
  2.  Attendance Management System
  3.  Payroll Calculator
  4.  Inventory Management System
  5.  Sales Tracking System
  6.  Monthly Expense Tracker
  7.  Invoice Generator
  8.  Purchase Management System
  9.  Interactive Sales Dashboard
  10. Complete Office MIS

Run from the repo root:
    .venv\\Scripts\\python.exe scripts\\gen_level7.py
"""
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from openpyxl.styles import Border, Font, PatternFill, Alignment, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, Reference

from xlcommon import (
    TEAL, TEAL_DARK, BAND, GRAY, WHITE,
    new_workbook, put_title, put_note, put_section, put_headers, put_rows,
    set_widths, blank_solution,
)

OUT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                    "..", "assets", "files", "level7"))

MONEY = "#,##0"
PCT0 = "0%"
PCT1 = "0.0%"
DATEFMT = "DD-MMM-YY"

_GREEN = "e7f6ec"
_RED = "fde8e8"
_AMBER = "fef3c7"
_GREY = "e5e7eb"

_FIRST_THIN = Side(style="thin", color="cbd5e1")
_BOX = Border(left=_FIRST_THIN, right=_FIRST_THIN, top=_FIRST_THIN, bottom=_FIRST_THIN)


def style_cell(ws, row, col, value=None, fmt=None, bold=False, fill=None, border=True):
    c = ws.cell(row=row, column=col)
    if value is not None:
        c.value = value
    if fmt:
        c.number_format = fmt
    if bold:
        c.font = Font(bold=True)
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if border:
        c.border = _BOX
    return c


def add_validation(ws, rng, source, prompt=None):
    dv = DataValidation(type="list", formula1='"%s"' % source, allow_blank=True)
    if prompt:
        dv.prompt = prompt
    ws.add_data_validation(dv)
    dv.add(rng)
    return dv


def equal_rule(ws, rng, value, fill):
    from openpyxl.formatting.rule import FormulaRule
    first = rng.split(":")[0].replace("$", "")
    formula = 'EXACT(%s,"%s")' % (first, value)
    ws.conditional_formatting.add(
        rng,
        FormulaRule(formula=[formula],
                    fill=PatternFill("solid", fgColor=fill)))


# ----------------------------------------------------------------------------
# Shared people data (Sunrise Traders)
# ----------------------------------------------------------------------------
EMPLOYEES = [
    ("E101", "Arjun Mehta", "Sales", "Sales Executive", "15-Mar-21", "Active", 32000),
    ("E102", "Priya Nair", "Accounts", "Accountant", "02-Jun-20", "Active", 38000),
    ("E103", "Rahul Verma", "Stores", "Store Keeper", "10-Jan-22", "Active", 26000),
    ("E104", "Sneha Iyer", "HR", "HR Executive", "21-Jul-21", "On Leave", 30000),
    ("E105", "Vikram Singh", "Sales", "Sales Manager", "05-Sep-19", "Active", 55000),
    ("E106", "Divya Menon", "Sales", "Sales Executive", "18-Feb-23", "Resigned", 28000),
]
EMP_IDS = [e[0] for e in EMPLOYEES]
EMP_NAMES = [e[1] for e in EMPLOYEES]

SUNDAYS = (2, 9, 16, 23, 30)


def employee_master_rows(n_spare=0):
    rows = [list(e) for e in EMPLOYEES]
    rows += [[""] * 7 for _ in range(n_spare)]
    return rows


def fill_employee_master(wb):
    ws = wb["Employee Master"]
    put_title(ws, "Employee Master — Sunrise Traders",
              "HR types each employee once. Every other sheet looks up this sheet.")
    r = put_headers(ws, 4, ["Emp ID", "Name", "Department", "Designation",
                            "Join Date", "Status", "Basic Salary (₹)"])
    rows = employee_master_rows(n_spare=4)
    r = put_rows(ws, r, rows, num_formats={7: MONEY})
    return ws


# ----------------------------------------------------------------------------
# Project 1 — Employee Management System
# ----------------------------------------------------------------------------
def build_employee_pair():
    ex = new_workbook("Employee Master", "Attendance Summary", "Salary Summary", "Dashboard")
    sol = new_workbook("Employee Master", "Attendance Summary", "Salary Summary", "Dashboard")

    fill_employee_master(ex)
    fill_employee_master(sol)
    add_validation(sol["Employee Master"], "C5:C14", "Sales,Stores,Accounts,HR")
    add_validation(sol["Employee Master"], "F5:F14", "Active,On Leave,Resigned")

    att_hdr = ["Emp ID", "Name", "Present", "Absent", "Leave", "Total Days", "Attendance %"]
    sal_hdr = ["Emp ID", "Name", "Department", "Basic (₹)", "HRA (₹)", "Gross (₹)"]

    # ---- exercise: only header rows on both summaries, empty dashboard
    put_title(ex["Attendance Summary"], "Attendance Summary — Sunrise Traders",
              "Type the IDs, then the day counts. Names and percentages fill themselves with VLOOKUP.")
    put_headers(ex["Attendance Summary"], 4, att_hdr)
    put_title(ex["Salary Summary"], "Salary Summary — Sunrise Traders",
              "Type the IDs; basic, department and HRA compute from the master.")
    put_headers(ex["Salary Summary"], 4, sal_hdr)
    put_title(ex["Dashboard"], "Sunrise Traders — HR Snapshot",
              "Build the department snapshot with COUNTIF and SUMIF (see the lesson steps).")

    # ---- solution: attendance summary
    ws = sol["Attendance Summary"]
    put_title(ws, "Attendance Summary — Sunrise Traders",
              "Day counts typed by HR; names and % calculate themselves.")
    r = put_headers(ws, 4, att_hdr)
    counts = [(25, 0, 0), (24, 0, 1), (23, 2, 0), (20, 0, 5), (25, 0, 0), (22, 1, 2)]
    for i, (emp, cnt) in enumerate(zip(EMP_IDS, counts)):
        rr = r + i
        ws.cell(row=rr, column=1, value=emp)
        ws.cell(row=rr, column=2,
                value="=VLOOKUP($A%d,'Employee Master'!$A$5:$B$10,2,FALSE)" % rr)
        for j, v in enumerate(cnt):
            ws.cell(row=rr, column=3 + j, value=v)
        ws.cell(row=rr, column=6, value="=C%d+D%d+E%d" % (rr, rr, rr))
        ws.cell(row=rr, column=7, value="=C%d/F%d" % (rr, rr))
        ws.cell(row=rr, column=7).number_format = PCT0
    for col in range(1, 8):
        for rr in range(r, r + 6):
            ws.cell(row=rr, column=col).border = _BOX

    # ---- solution: salary summary
    ws = sol["Salary Summary"]
    put_title(ws, "Salary Summary — Sunrise Traders",
              "Basic and department come from the master; HRA is 40% of basic.")
    r = put_headers(ws, 4, sal_hdr)
    for i, emp in enumerate(EMP_IDS):
        rr = r + i
        ws.cell(row=rr, column=1, value=emp)
        ws.cell(row=rr, column=2,
                value="=VLOOKUP($A%d,'Employee Master'!$A$5:$B$10,2,FALSE)" % rr)
        ws.cell(row=rr, column=3,
                value="=VLOOKUP($A%d,'Employee Master'!$A$5:$G$10,3,FALSE)" % rr)
        ws.cell(row=rr, column=4,
                value="=VLOOKUP($A%d,'Employee Master'!$A$5:$G$10,7,FALSE)" % rr)
        ws.cell(row=rr, column=5, value="=ROUND($D%d*40%%,0)" % rr)
        ws.cell(row=rr, column=6, value="=$D%d+$E%d" % (rr, rr))
        for col in (4, 5, 6):
            ws.cell(row=rr, column=col).number_format = MONEY
        for col in range(1, 7):
            ws.cell(row=rr, column=col).border = _BOX

    # ---- solution: dashboard
    ws = sol["Dashboard"]
    put_title(ws, "Sunrise Traders — HR Snapshot")
    ws["A3"] = "Department"
    ws["B3"] = "Headcount"
    ws["C3"] = "Basic Salary (₹)"
    for col in "ABC":
        style_cell(ws, 3, "ABC".index(col) + 1, None, bold=True, fill=TEAL)
    for i, dept in enumerate(["Sales", "Stores", "Accounts", "HR"]):
        rr = 4 + i
        ws.cell(row=rr, column=1, value=dept)
        ws.cell(row=rr, column=2, value="=COUNTIF('Employee Master'!$C$5:$C$14,$A%d)" % rr)
        ws.cell(row=rr, column=3,
                value="=SUMIF('Employee Master'!$C$5:$C$14,$A%d,'Employee Master'!$G$5:$G$14)" % rr)
        ws.cell(row=rr, column=3).number_format = MONEY
        for col in range(1, 4):
            ws.cell(row=rr, column=col).border = _BOX
    ws["A9"] = "Active Employees"
    ws["B9"] = "=COUNTIF('Employee Master'!$F$5:$F$14,\"Active\")"
    ws["A10"] = "Monthly Basic Payroll (₹)"
    ws["B10"] = "=SUM('Salary Summary'!$D$4:$D$13)"
    ws["B10"].number_format = MONEY
    ws["A9"].font = Font(bold=True)
    ws["A10"].font = Font(bold=True)
    set_widths(ws, [26, 16, 18])

    return ex, sol, "employee-management-system"


# ----------------------------------------------------------------------------
# Project 2 — Attendance Management System
# ----------------------------------------------------------------------------
def fill_attendance_register(ws, solved):
    put_title(ws, "Attendance Register — November 2025",
              "Codes: P present · A absent · L leave · WO week off · H holiday. "
              "Days 1–14 are pre-coded; complete 15–30.")
    hdr = ["Emp ID", "Name"] + [d for d in range(1, 31)] + \
          ["Present", "Absent", "Leave", "Week-off", "Holiday", "Working Days", "Attendance %"]
    r = put_headers(ws, 4, hdr)
    for i, emp in enumerate(EMPLOYEES):
        rr = r + i
        ws.cell(row=rr, column=1, value=emp[0])
        ws.cell(row=rr, column=2, value=emp[1])
        for d in range(1, 31):
            col = 2 + d
            if d <= 14:
                if d in SUNDAYS:
                    code = "WO"
                elif emp[0] == "E104" and 3 <= d <= 7:
                    code = "L"
                elif emp[0] == "E103" and d == 11:
                    code = "A"
                else:
                    code = "P"
            elif solved:
                code = "WO" if d in SUNDAYS else "P"
            else:
                code = None
            if code is not None:
                ws.cell(row=rr, column=col, value=code)
        if solved:
            first, last = "C", "AF"
            ws.cell(row=rr, column=33, value='=COUNTIF($C%d:$AF%d,"P")' % (rr, rr))
            ws.cell(row=rr, column=34, value='=COUNTIF($C%d:$AF%d,"A")' % (rr, rr))
            ws.cell(row=rr, column=35, value='=COUNTIF($C%d:$AF%d,"L")' % (rr, rr))
            ws.cell(row=rr, column=36, value='=COUNTIF($C%d:$AF%d,"WO")' % (rr, rr))
            ws.cell(row=rr, column=37, value='=COUNTIF($C%d:$AF%d,"H")' % (rr, rr))
            ws.cell(row=rr, column=38, value="=30-$AJ%d-$AK%d" % (rr, rr))
            ws.cell(row=rr, column=39, value="=AG%d/AL%d" % (rr, rr))
            ws.cell(row=rr, column=39).number_format = PCT0
        for col in range(1, 40):
            ws.cell(row=rr, column=col).border = _BOX
    if solved:
        ws["B9"] = "Present per day"
        ws["B10"] = "Absent per day"
        ws["B9"].font = Font(bold=True)
        ws["B10"].font = Font(bold=True)
        for d in range(1, 31):
            col = 2 + d
            letter = ws.cell(row=4, column=col).coordinate.replace("4", "")
            ws.cell(row=9, column=col, value='=COUNTIF(%s$5:%s$10,"P")' % (letter, letter))
            ws.cell(row=10, column=col, value='=COUNTIF(%s$5:%s$10,"A")' % (letter, letter))
        add_validation(ws, "C5:AF10", "P,A,L,WO,H")
        # weekend shading + absence highlight
        from openpyxl.formatting.rule import FormulaRule
        ws.conditional_formatting.add(
            "C5:AF10",
            FormulaRule(formula=["OR(C$4=2,C$4=9,C$4=16,C$4=23,C$4=30)"],
                        fill=PatternFill("solid", fgColor=_GREY)))
        equal_rule(ws, "C5:AF10", "A", _RED)
    set_widths(ws, [10, 16] + [3.5] * 30 + [9, 9, 9, 9, 9, 12, 12])


def fill_attendance_summary(ws, solved):
    put_title(ws, "Attendance Summary — November 2025",
              "Plain references to the register — one source of truth, many views.")
    if not solved:
        return
    r = put_headers(ws, 4, ["Emp ID", "Name", "Present", "Absent", "Leave",
                            "Working Days", "Attendance %"])
    for i in range(6):
        rr = r + i
        reg = rr - r + 5  # register data row
        ws.cell(row=rr, column=1, value="='Attendance Register'!A%d" % reg)
        ws.cell(row=rr, column=2, value="='Attendance Register'!B%d" % reg)
        ws.cell(row=rr, column=3, value="='Attendance Register'!AG%d" % reg)
        ws.cell(row=rr, column=4, value="='Attendance Register'!AH%d" % reg)
        ws.cell(row=rr, column=5, value="='Attendance Register'!AI%d" % reg)
        ws.cell(row=rr, column=6, value="='Attendance Register'!AL%d" % reg)
        ws.cell(row=rr, column=7, value="='Attendance Register'!AM%d" % reg)
        ws.cell(row=rr, column=7).number_format = PCT0
        for col in range(1, 8):
            ws.cell(row=rr, column=col).border = _BOX
    set_widths(ws, [10, 16, 10, 10, 10, 13, 13])


def build_attendance_pair():
    ex = new_workbook("Attendance Register", "Summary")
    sol = new_workbook("Attendance Register", "Summary")
    fill_attendance_register(ex["Attendance Register"], solved=False)
    fill_attendance_summary(ex["Summary"], solved=False)
    fill_attendance_register(sol["Attendance Register"], solved=True)
    fill_attendance_summary(sol["Summary"], solved=True)
    return ex, sol, "attendance-management-system"


# ----------------------------------------------------------------------------
# Project 3 — Payroll Calculator
# ----------------------------------------------------------------------------
PAYSLIP_LAYOUT = [  # (row, label_left, label_right) pairs; values in B and D
    (6, "Name", "Department"),
    (7, "Designation", "Join Date"),
    (8, "Basic (₹)", "HRA (₹)"),
    (9, "PF (₹)", "Gross (₹)"),
    (10, "Tax (₹)", "Net (₹)"),
]


def fill_payslip(ws, solved):
    put_title(ws, "Sunrise Traders — Salary Slip", "Type an Emp ID below; every field fills itself.")
    ws["A4"] = "Emp ID:"
    ws["A4"].font = Font(bold=True)
    if solved:
        ws["B4"] = "E101"
        for row, lab_l, lab_r in PAYSLIP_LAYOUT:
            ws.cell(row=row, column=1, value=lab_l).font = Font(bold=True)
            ws.cell(row=row, column=3, value=lab_r).font = Font(bold=True)
        m = "MATCH($B$4,'Payroll'!$A$5:$A$10,0)"
        ws["B6"] = "=IFERROR(INDEX('Payroll'!$B$5:$B$10,%s),\"\")" % m
        ws["D6"] = "=IFERROR(VLOOKUP($B$4,'Employee Master'!$A$5:$C$10,3,FALSE),\"\")"
        ws["B7"] = "=IFERROR(VLOOKUP($B$4,'Employee Master'!$A$5:$D$10,4,FALSE),\"\")"
        ws["D7"] = "=IFERROR(VLOOKUP($B$4,'Employee Master'!$A$5:$E$10,5,FALSE),\"\")"
        for col, pcol in ((2, "C"), (4, "D")):
            ws.cell(row=8, column=col, value="=IFERROR(INDEX('Payroll'!$%s$5:$%s$10,%s),\"\")" % (pcol, pcol, m))
            ws.cell(row=8, column=col).number_format = MONEY
        for row, pcol in ((9, "E"), (10, "G")):
            ws.cell(row=row, column=2, value="=IFERROR(INDEX('Payroll'!$%s$5:$%s$10,%s),\"\")" % (pcol, pcol, m))
            ws.cell(row=row, column=2).number_format = MONEY
        ws["D9"] = "=IFERROR(INDEX('Payroll'!$F$5:$F$10,%s),\"\")" % m
        ws["D9"].number_format = MONEY
        ws["D10"] = "=IFERROR(INDEX('Payroll'!$H$5:$H$10,%s),\"\")" % m
        ws["D10"].number_format = MONEY
        ws["D10"].font = Font(bold=True)
        ws["D10"].fill = PatternFill("solid", fgColor=BAND)
        for row, _l, _r in PAYSLIP_LAYOUT:
            for col in (1, 2, 3, 4):
                ws.cell(row=row, column=col).border = _BOX
        ws.print_area = "A1:D12"
    else:
        put_note(ws, 12, "Build the payslip: labels are laid out. Put the Emp ID input in B4 and "
                 "pull each field with INDEX + MATCH wrapped in IFERROR (see lesson).", width=4)
        for row, lab_l, lab_r in PAYSLIP_LAYOUT:
            ws.cell(row=row, column=1, value=lab_l).font = Font(bold=True)
            ws.cell(row=row, column=3, value=lab_r).font = Font(bold=True)
    set_widths(ws, [14, 22, 12, 18])


def build_payroll_pair():
    ex = new_workbook("Employee Master", "Payroll", "Payslip")
    sol = new_workbook("Employee Master", "Payroll", "Payslip")
    for wb in (ex, sol):
        ws = wb["Employee Master"]
        put_title(ws, "Employee Master — Sunrise Traders",
                  "Basic Salary is column G, rows 5 to 10. Every lookup column number depends on this order.")
        r = put_headers(ws, 4, ["Emp ID", "Name", "Department", "Designation",
                                "Join Date", "Status", "Basic Salary (₹)"])
        put_rows(ws, r, employee_master_rows(), num_formats={7: MONEY})
    ws = ex["Payroll"]
    put_note(ws, 1, "Payroll — type the IDs in A, then build Basic/HRA/PF/Gross/Tax/Net with the "
             "lesson formulas. Keep every money figure rounded to whole rupees.", width=8)
    put_headers(ws, 4, ["Emp ID", "Name", "Basic (₹)", "HRA (₹)", "PF (₹)",
                        "Gross (₹)", "Tax (₹)", "Net (₹)"])
    fill_payslip(ex["Payslip"], solved=False)

    ws = sol["Payroll"]
    put_title(ws, "Payroll — Sunrise Traders",
              "HRA 40% of basic · PF 12% of basic · slab tax: 0 to ₹25,000, 5% to ₹50,000, 10% above.")
    r = put_headers(ws, 4, ["Emp ID", "Name", "Basic (₹)", "HRA (₹)", "PF (₹)",
                            "Gross (₹)", "Tax (₹)", "Net (₹)"])
    for i, emp in enumerate(EMP_IDS):
        rr = r + i
        ws.cell(row=rr, column=1, value=emp)
        ws.cell(row=rr, column=2, value="=VLOOKUP($A%d,'Employee Master'!$A$5:$B$10,2,FALSE)" % rr)
        ws.cell(row=rr, column=3, value="=VLOOKUP($A%d,'Employee Master'!$A$5:$G$10,7,FALSE)" % rr)
        ws.cell(row=rr, column=4, value="=ROUND($C%d*40%%,0)" % rr)
        ws.cell(row=rr, column=5, value="=ROUND($C%d*12%%,0)" % rr)
        ws.cell(row=rr, column=6, value="=$C%d+$D%d+$E%d" % (rr, rr, rr))
        ws.cell(row=rr, column=7,
                value=("=IF($F%d<=25000,0,IF($F%d<=50000,(F%d-25000)*5%%,1250+(F%d-50000)*10%%))"
                       % (rr, rr, rr, rr)))
        ws.cell(row=rr, column=8, value="=ROUND($F%d-$G%d,0)" % (rr, rr))
        for col in range(3, 9):
            ws.cell(row=rr, column=col).number_format = MONEY
        for col in range(1, 9):
            ws.cell(row=rr, column=col).border = _BOX
    tot = r + 6
    ws.cell(row=tot, column=1, value="Total Net Payroll (₹)").font = Font(bold=True)
    ws.cell(row=tot, column=8, value="=SUM($H$5:$H$10)").number_format = MONEY
    fill_payslip(sol["Payslip"], solved=True)
    return ex, sol, "payroll-calculator"


# ----------------------------------------------------------------------------
# Project 4 — Inventory Management System
# ----------------------------------------------------------------------------
ITEMS = [  # code, name, unit, opening, rate, reorder level
    ("ITM001", "Notebooks", "Nos", 200, 40, 100),
    ("ITM002", "Pens", "Nos", 500, 12, 150),
    ("ITM003", "Staplers", "Nos", 80, 150, 30),
    ("ITM004", "Printers", "Units", 15, 6500, 5),
    ("ITM005", "Laptops", "Units", 10, 42000, 3),
]
INWARD = [  # date, code, qty  (rate comes from the master)
    (date(2026, 11, 2), "ITM002", 200), (date(2026, 11, 3), "ITM001", 150),
    (date(2026, 11, 5), "ITM003", 20), (date(2026, 11, 7), "ITM002", 100),
    (date(2026, 11, 10), "ITM005", 2), (date(2026, 11, 12), "ITM004", 5),
]
OUTWARD = [
    (date(2026, 11, 2), "ITM001", 50), (date(2026, 11, 4), "ITM002", 150),
    (date(2026, 11, 6), "ITM003", 25), (date(2026, 11, 8), "ITM002", 100),
    (date(2026, 11, 10), "ITM005", 3), (date(2026, 11, 12), "ITM004", 5),
    (date(2026, 11, 14), "ITM001", 70), (date(2026, 11, 15), "ITM002", 210),
    (date(2026, 11, 17), "ITM003", 50), (date(2026, 11, 19), "ITM005", 1),
    (date(2026, 11, 20), "ITM004", 3),
]


def fill_register_sheet(ws, entries, solved):
    put_title(ws, ws.title + " Register — November 2026",
              "Type date, item code (dropdown) and quantity. Name, rate and value fill themselves.")
    put_headers(ws, 4, ["Date", "Item Code", "Item Name", "Quantity", "Rate (₹)", "Value (₹)"])
    for i, (d, code, qty) in enumerate(entries):
        rr = 5 + i
        c = ws.cell(row=rr, column=1, value=d)
        c.number_format = DATEFMT
        ws.cell(row=rr, column=2, value=code)
        ws.cell(row=rr, column=4, value=qty)
        if solved:
            ws.cell(row=rr, column=3,
                    value="=IF($B%d=\"\",\"\",VLOOKUP($B%d,'Item Master'!$A$5:$B$9,2,FALSE))" % (rr, rr))
            ws.cell(row=rr, column=5,
                    value="=IF($B%d=\"\",\"\",VLOOKUP($B%d,'Item Master'!$A$5:$E$9,5,FALSE))" % (rr, rr))
            ws.cell(row=rr, column=6, value="=IF($B%d=\"\",\"\",ROUND($D%d*$E%d,0))" % (rr, rr, rr))
            for col in (5, 6):
                ws.cell(row=rr, column=col).number_format = MONEY
        for col in range(1, 7):
            ws.cell(row=rr, column=col).border = _BOX
    if solved:
        add_validation(ws, "B5:B50", "ITM001,ITM002,ITM003,ITM004,ITM005")
    set_widths(ws, [12, 12, 14, 10, 10, 12])


def build_inventory_pair():
    ex = new_workbook("Item Master", "Inward", "Outward", "Current Stock")
    sol = new_workbook("Item Master", "Inward", "Outward", "Current Stock")
    for wb in (ex, sol):
        ws = wb["Item Master"]
        put_title(ws, "Item Master — Sunrise Traders",
                  "One row per item. Opening stock, rate and reorder level live only here.")
        r = put_headers(ws, 4, ["Item Code", "Item Name", "Unit", "Opening Stock",
                                "Rate (₹)", "Reorder Level"])
        put_rows(ws, r, [list(i) for i in ITEMS], num_formats={5: MONEY, 6: MONEY})
    fill_register_sheet(ex["Inward"], INWARD, solved=False)
    fill_register_sheet(ex["Outward"], OUTWARD, solved=False)
    fill_register_sheet(sol["Inward"], INWARD, solved=True)
    fill_register_sheet(sol["Outward"], OUTWARD, solved=True)

    hdr = ["Item Code", "Item Name", "Opening", "Inward", "Outward", "Closing",
           "Rate (₹)", "Value (₹)", "Reorder Level", "Status"]
    ws = ex["Current Stock"]
    put_note(ws, 1, "Current Stock — build the live balance here: opening from the master, "
             "inward/outward with SUMIFS over the registers, closing = opening + inward − outward.", width=10)
    put_headers(ws, 4, hdr)

    ws = sol["Current Stock"]
    put_title(ws, "Current Stock — Sunrise Traders",
              "Closing = Opening + Inward − Outward, live from the registers. Reorder level lives in the master.")
    r = put_headers(ws, 4, hdr)
    for i, item in enumerate(ITEMS):
        rr = r + i
        ws.cell(row=rr, column=1, value=item[0])
        ws.cell(row=rr, column=2, value="=VLOOKUP($A%d,'Item Master'!$A$5:$B$9,2,FALSE)" % rr)
        ws.cell(row=rr, column=3, value="=VLOOKUP($A%d,'Item Master'!$A$5:$D$9,4,FALSE)" % rr)
        ws.cell(row=rr, column=4, value="=SUMIFS(Inward!$D$5:$D$50,Inward!$B$5:$B$50,$A%d)" % rr)
        ws.cell(row=rr, column=5, value="=SUMIFS(Outward!$D$5:$D$50,Outward!$B$5:$B$50,$A%d)" % rr)
        ws.cell(row=rr, column=6, value="=$C%d+$D%d-$E%d" % (rr, rr, rr))
        ws.cell(row=rr, column=7, value="=VLOOKUP($A%d,'Item Master'!$A$5:$E$9,5,FALSE)" % rr)
        ws.cell(row=rr, column=8, value="=ROUND($F%d*$G%d,0)" % (rr, rr))
        ws.cell(row=rr, column=9, value="=VLOOKUP($A%d,'Item Master'!$A$5:$F$9,6,FALSE)" % rr)
        ws.cell(row=rr, column=10, value='=IF($F%d<=$I%d,"REORDER","OK")' % (rr, rr))
        for col in (7, 8):
            ws.cell(row=rr, column=col).number_format = MONEY
        for col in range(1, 11):
            ws.cell(row=rr, column=col).border = _BOX
    equal_rule(ws, "J5:J9", "REORDER", _RED)
    tot = r + 5
    ws.cell(row=tot, column=7, value="Total Stock Value (₹)").font = Font(bold=True)
    ws.cell(row=tot, column=8, value="=SUM($H$5:$H$9)").number_format = MONEY
    set_widths(ws, [11, 12, 10, 10, 10, 10, 10, 12, 13, 11])
    return ex, sol, "inventory-management-system"


# ----------------------------------------------------------------------------
# Project 5 — Sales Tracking System  (shared with Projects 9 and 10)
# ----------------------------------------------------------------------------
REGIONS = ["North", "South", "East", "West"]
PRODUCTS = ["Notebooks", "Laptops", "Printers", "Ink & Toner", "Stationery Kits"]
SALESPEOPLE = EMP_NAMES
TARGETS = [150000, 200000, 120000, 250000, 130000, 160000]

# (date, region, product, salesperson, qty, rate)
SALES_ROWS = [
    (date(2026, 8, 3), "North", "Notebooks", "Priya Nair", 1200, 45),
    (date(2026, 8, 5), "North", "Laptops", "Divya Menon", 1, 45000),
    (date(2026, 8, 8), "East", "Laptops", "Sneha Iyer", 1, 42000),
    (date(2026, 8, 12), "East", "Laptops", "Sneha Iyer", 1, 38000),
    (date(2026, 8, 15), "West", "Notebooks", "Priya Nair", 600, 45),
    (date(2026, 8, 18), "South", "Printers", "Divya Menon", 5, 8000),
    (date(2026, 8, 22), "South", "Printers", "Sneha Iyer", 5, 8000),
    (date(2026, 8, 26), "East", "Printers", "Priya Nair", 5, 6800),
    (date(2026, 9, 2), "North", "Printers", "Vikram Singh", 4, 8000),
    (date(2026, 9, 4), "East", "Laptops", "Arjun Mehta", 1, 38500),
    (date(2026, 9, 7), "East", "Notebooks", "Rahul Verma", 700, 45),
    (date(2026, 9, 9), "South", "Stationery Kits", "Rahul Verma", 17, 1000),
    (date(2026, 9, 11), "West", "Notebooks", "Vikram Singh", 460, 50),
    (date(2026, 9, 14), "West", "Notebooks", "Sneha Iyer", 220, 50),
    (date(2026, 9, 16), "West", "Printers", "Arjun Mehta", 5, 6800),
    (date(2026, 9, 18), "East", "Ink & Toner", "Vikram Singh", 40, 850),
    (date(2026, 9, 21), "West", "Ink & Toner", "Rahul Verma", 30, 850),
    (date(2026, 9, 23), "West", "Ink & Toner", "Sneha Iyer", 15, 700),
    (date(2026, 9, 25), "West", "Stationery Kits", "Vikram Singh", 35, 1000),
    (date(2026, 9, 28), "West", "Stationery Kits", "Divya Menon", 24, 1000),
    (date(2026, 9, 30), "West", "Stationery Kits", "Priya Nair", 20, 1000),
    (date(2026, 9, 30), "West", "Stationery Kits", "Sneha Iyer", 10, 900),
    (date(2026, 10, 2), "North", "Notebooks", "Sneha Iyer", 1500, 45),
    (date(2026, 10, 5), "North", "Laptops", "Divya Menon", 1, 55000),
    (date(2026, 10, 7), "North", "Ink & Toner", "Arjun Mehta", 45, 900),
    (date(2026, 10, 9), "South", "Notebooks", "Arjun Mehta", 1000, 45),
    (date(2026, 10, 12), "South", "Laptops", "Priya Nair", 1, 43500),
    (date(2026, 10, 14), "South", "Ink & Toner", "Priya Nair", 40, 850),
    (date(2026, 10, 16), "South", "Laptops", "Sneha Iyer", 1, 48000),
    (date(2026, 10, 20), "East", "Printers", "Rahul Verma", 5, 8000),
    (date(2026, 10, 24), "West", "Notebooks", "Divya Menon", 500, 50),
]
MONTHS = ["Aug-26", "Sep-26", "Oct-26"]


def fill_sales_log(ws):
    put_title(ws, "Sales Log — Sunrise Traders",
              "One row per invoice. Amount = Qty × Rate; Month = TEXT helper for the summaries.")
    r = put_headers(ws, 4, ["Date", "Invoice No", "Region", "Product", "Salesperson",
                            "Qty", "Rate (₹)", "Amount (₹)", "Month"])
    for i, (d, region, product, person, qty, rate) in enumerate(SALES_ROWS):
        rr = r + i
        c = ws.cell(row=rr, column=1, value=d)
        c.number_format = DATEFMT
        ws.cell(row=rr, column=2, value="INV-%d" % (2001 + i))
        ws.cell(row=rr, column=3, value=region)
        ws.cell(row=rr, column=4, value=product)
        ws.cell(row=rr, column=5, value=person)
        ws.cell(row=rr, column=6, value=qty)
        ws.cell(row=rr, column=7, value=rate).number_format = MONEY
        ws.cell(row=rr, column=8, value="=F%d*G%d" % (rr, rr)).number_format = MONEY
        ws.cell(row=rr, column=9, value='=TEXT(A%d,"mmm-yy")' % rr)
        for col in range(1, 10):
            ws.cell(row=rr, column=col).border = _BOX
    set_widths(ws, [11, 11, 9, 15, 14, 8, 10, 12, 9])


def fill_sales_monthly(ws, solved):
    put_title(ws, "Monthly Summary",
              "Revenue per month from the log, growth, best month — plus the region block.")
    if not solved:
        put_headers(ws, 4, ["Month", "Revenue", "Growth %"])
        put_note(ws, 5, "Type Aug-26 / Sep-26 / Oct-26 exactly as the log's Month helper spells them, "
                 "then total each month with SUMIFS and compute growth. See the lesson steps.", width=6)
        return
    r = put_headers(ws, 4, ["Month", "Revenue", "Growth %"])
    for i, m in enumerate(MONTHS):
        rr = r + i
        ws.cell(row=rr, column=1, value=m)
        ws.cell(row=rr, column=2,
                value="=SUMIFS('Sales Log'!$H$5:$H$204,'Sales Log'!$I$5:$I$204,$A%d)" % rr)
        ws.cell(row=rr, column=2).number_format = MONEY
        if i:
            ws.cell(row=rr, column=3, value="=($B%d-$B%d)/$B%d" % (rr, rr - 1, rr - 1))
            ws.cell(row=rr, column=3).number_format = PCT1
        for col in range(1, 4):
            ws.cell(row=rr, column=col).border = _BOX
    ws["A9"] = "Best Month"
    ws["A9"].font = Font(bold=True)
    ws["B9"] = "=INDEX($A$5:$A$7,MATCH(MAX($B$5:$B$7),$B$5:$B$7,0))"
    ws["A10"] = "Total 3-Month Revenue (₹)"
    ws["A10"].font = Font(bold=True)
    ws["B10"] = "=SUM($B$5:$B$7)"
    ws["B10"].number_format = MONEY
    ws["A12"] = "Region"
    ws["B12"] = "Revenue"
    style_cell(ws, 12, 1, None, bold=True, fill=TEAL)
    style_cell(ws, 12, 2, None, bold=True, fill=TEAL)
    for i, reg in enumerate(REGIONS):
        rr = 13 + i
        ws.cell(row=rr, column=1, value=reg)
        ws.cell(row=rr, column=2,
                value="=SUMIFS('Sales Log'!$H$5:$H$204,'Sales Log'!$C$5:$C$204,$A%d)" % rr)
        ws.cell(row=rr, column=2).number_format = MONEY
        for col in range(1, 3):
            ws.cell(row=rr, column=col).border = _BOX
    chart = BarChart()
    chart.type = "col"
    chart.title = "Monthly Revenue (₹)"
    chart.height, chart.width = 7.5, 13
    data = Reference(ws, min_col=2, min_row=4, max_row=7)
    cats = Reference(ws, min_col=1, min_row=5, max_row=7)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.legend = None
    ws.add_chart(chart, "D5")
    set_widths(ws, [24, 16, 11])



def fill_sales_performance(ws, solved):
    put_title(ws, "Performance",
              "Target vs actual per salesperson, achievement %, rank and status.")
    if not solved:
        put_headers(ws, 4, ["Salesperson", "Target (₹)", "Actual (₹)", "Achievement %",
                            "Rank", "Status"])
        put_note(ws, 5, "Type the six names and their 3-month targets, total each person's sales "
                 "with SUMIF, then compute achievement, RANK and the status flag.", width=6)
        return
    r = put_headers(ws, 4, ["Salesperson", "Target (₹)", "Actual (₹)", "Achievement %",
                            "Rank", "Status"])
    for i, (person, target) in enumerate(zip(SALESPEOPLE, TARGETS)):
        rr = r + i
        ws.cell(row=rr, column=1, value=person)
        ws.cell(row=rr, column=2, value=target).number_format = MONEY
        ws.cell(row=rr, column=3,
                value="=SUMIF('Sales Log'!$E$5:$E$204,$A%d,'Sales Log'!$H$5:$H$204)" % rr)
        ws.cell(row=rr, column=3).number_format = MONEY
        ws.cell(row=rr, column=4, value="=$C%d/$B%d" % (rr, rr))
        ws.cell(row=rr, column=4).number_format = PCT0
        ws.cell(row=rr, column=5, value="=RANK($D%d,$D$5:$D$10,0)" % rr)
        ws.cell(row=rr, column=6,
                value='=IF($D%d>=1,"Target Met",IF($D%d>=0.9,"Close","Below"))' % (rr, rr))
        for col in range(1, 7):
            ws.cell(row=rr, column=col).border = _BOX
    equal_rule(ws, "F5:F10", "Target Met", _GREEN)
    equal_rule(ws, "F5:F10", "Close", _AMBER)
    equal_rule(ws, "F5:F10", "Below", _RED)
    tot = r + 6
    ws.cell(row=tot, column=1, value="Total").font = Font(bold=True)
    ws.cell(row=tot, column=2, value="=SUM($B$5:$B$10)").number_format = MONEY
    ws.cell(row=tot, column=3, value="=SUM($C$5:$C$10)").number_format = MONEY
    ws.cell(row=tot, column=4, value="=$C%d/$B%d" % (tot, tot)).number_format = PCT0
    for col in range(1, 5):
        ws.cell(row=tot, column=col).border = _BOX
    set_widths(ws, [16, 12, 12, 14, 8, 13])


def build_sales_pair():
    ex = new_workbook("Sales Log", "Monthly Summary", "Performance")
    sol = new_workbook("Sales Log", "Monthly Summary", "Performance")
    fill_sales_log(ex["Sales Log"])
    fill_sales_log(sol["Sales Log"])
    fill_sales_monthly(ex["Monthly Summary"], solved=False)
    fill_sales_performance(ex["Performance"], solved=False)
    fill_sales_monthly(sol["Monthly Summary"], solved=True)
    fill_sales_performance(sol["Performance"], solved=True)
    return ex, sol, "sales-tracking-system"


# ----------------------------------------------------------------------------
# Project 6 — Monthly Expense Tracker
# ----------------------------------------------------------------------------
CATEGORIES = ["Rent", "Salaries", "Electricity", "Transport", "Marketing", "Miscellaneous"]
BUDGETS = [160000, 480000, 48000, 60000, 80000, 40000]

EXPENSE_ROWS = [  # (date, category, description, amount)
    (date(2026, 7, 1), "Rent", "Office rent — July", 40000),
    (date(2026, 7, 1), "Salaries", "Staff salaries — July", 118000),
    (date(2026, 7, 5), "Electricity", "Electricity bill — July", 12800),
    (date(2026, 7, 8), "Transport", "Delivery fuel — July", 12500),
    (date(2026, 7, 12), "Marketing", "Digital ads — July", 17000),
    (date(2026, 7, 28), "Miscellaneous", "Office supplies — July", 7350),
    (date(2026, 8, 1), "Rent", "Office rent — August", 40000),
    (date(2026, 8, 1), "Salaries", "Staff salaries — August", 112300),
    (date(2026, 8, 5), "Electricity", "Electricity bill — August", 12800),
    (date(2026, 8, 8), "Transport", "Delivery fuel — August", 14500),
    (date(2026, 8, 12), "Marketing", "Digital ads — August", 22000),
    (date(2026, 8, 28), "Miscellaneous", "Office supplies — August", 10500),
    (date(2026, 9, 1), "Rent", "Office rent — September", 40000),
    (date(2026, 9, 1), "Salaries", "Staff salaries — September", 115400),
    (date(2026, 9, 5), "Electricity", "Electricity bill — September", 12800),
    (date(2026, 9, 8), "Transport", "Delivery fuel — September", 14000),
    (date(2026, 9, 12), "Marketing", "Trade fair stall — September", 24000),
    (date(2026, 9, 28), "Miscellaneous", "Office supplies — September", 10000),
    (date(2026, 10, 1), "Rent", "Office rent — October", 40000),
    (date(2026, 10, 1), "Salaries", "Staff salaries — October", 116300),
    (date(2026, 10, 5), "Electricity", "Electricity bill — October", 12800),
    (date(2026, 10, 8), "Transport", "Delivery fuel — October", 14000),
    (date(2026, 10, 12), "Marketing", "Festive campaign — October", 23500),
    (date(2026, 10, 28), "Miscellaneous", "Office supplies — October", 8350),
]
EXP_MONTHS = ["Jul-26", "Aug-26", "Sep-26", "Oct-26"]


def fill_expense_log(ws):
    put_title(ws, "Expense Log — Sunrise Traders",
              "Every payment typed once. The Category dropdown keeps every summary honest.")
    r = put_headers(ws, 4, ["Date", "Category", "Description", "Amount (₹)", "Month"])
    for i, (d, cat, desc, amt) in enumerate(EXPENSE_ROWS):
        rr = r + i
        c = ws.cell(row=rr, column=1, value=d)
        c.number_format = DATEFMT
        ws.cell(row=rr, column=2, value=cat)
        ws.cell(row=rr, column=3, value=desc)
        ws.cell(row=rr, column=4, value=amt).number_format = MONEY
        ws.cell(row=rr, column=5, value='=TEXT(A%d,"mmm-yy")' % rr)
        for col in range(1, 6):
            ws.cell(row=rr, column=col).border = _BOX
    set_widths(ws, [11, 14, 30, 12, 9])


def fill_expense_category(ws, solved):
    put_title(ws, "Category Summary",
              "Actual vs budget per category. Variance = budget − actual (positive = money left).")
    if not solved:
        put_headers(ws, 4, ["Category", "Actual (₹)", "Budget (₹)", "Variance (₹)", "Status"])
        put_note(ws, 5, "Type the six categories and their budgets, total each with SUMIF over the "
                 "log, then compute variance and the status flag. See the lesson steps.", width=6)
        return
    r = put_headers(ws, 4, ["Category", "Actual (₹)", "Budget (₹)", "Variance (₹)", "Status"])
    for i, (cat, bud) in enumerate(zip(CATEGORIES, BUDGETS)):
        rr = r + i
        ws.cell(row=rr, column=1, value=cat)
        ws.cell(row=rr, column=2,
                value="=SUMIF('Expense Log'!$B$5:$B$104,$A%d,'Expense Log'!$D$5:$D$104)" % rr)
        ws.cell(row=rr, column=3, value=bud).number_format = MONEY
        ws.cell(row=rr, column=4, value="=$C%d-$B%d" % (rr, rr)).number_format = MONEY
        ws.cell(row=rr, column=5, value='=IF($B%d>$C%d,"Over Budget","Within Budget")' % (rr, rr))
        for col in range(1, 6):
            ws.cell(row=rr, column=col).border = _BOX
    equal_rule(ws, "E5:E10", "Over Budget", _RED)
    equal_rule(ws, "E5:E10", "Within Budget", _GREEN)
    tot = r + 6
    ws.cell(row=tot, column=1, value="Total").font = Font(bold=True)
    ws.cell(row=tot, column=2, value="=SUM($B$5:$B$10)").number_format = MONEY
    ws.cell(row=tot, column=3, value="=SUM($C$5:$C$10)").number_format = MONEY
    ws.cell(row=tot, column=4, value="=SUM($D$5:$D$10)").number_format = MONEY
    for col in range(1, 5):
        ws.cell(row=tot, column=col).border = _BOX
    set_widths(ws, [16, 12, 12, 12, 15])


def fill_expense_monthly(ws, solved):
    put_title(ws, "Monthly Summary (Expenses)",
              "Total spend per month and each month's share of the four-month whole.")
    if not solved:
        put_headers(ws, 4, ["Month", "Total (₹)", "% of Total"])
        put_note(ws, 5, "Type Jul-26 … Oct-26 exactly as the log's Month helper spells them, then "
                 "total each month with SUMIF and compute each share of the grand total.", width=6)
        return
    r = put_headers(ws, 4, ["Month", "Total (₹)", "% of Total"])
    for i, m in enumerate(EXP_MONTHS):
        rr = r + i
        ws.cell(row=rr, column=1, value=m)
        ws.cell(row=rr, column=2,
                value="=SUMIF('Expense Log'!$E$5:$E$104,$A%d,'Expense Log'!$D$5:$D$104)" % rr)
        ws.cell(row=rr, column=2).number_format = MONEY
        ws.cell(row=rr, column=3, value="=$B%d/SUM($B$5:$B$8)" % rr).number_format = PCT1
        for col in range(1, 4):
            ws.cell(row=rr, column=col).border = _BOX
    ws["A10"] = "Grand Total (₹)"
    ws["A10"].font = Font(bold=True)
    ws["B10"] = "=SUM($B$5:$B$8)"
    ws["B10"].number_format = MONEY
    ws["A12"] = "Reconciliation"
    ws["A12"].font = Font(bold=True)
    ws["B12"] = "=SUM('Expense Log'!$D$5:$D$104)-SUM('Category Summary'!$B$5:$B$10)"
    ws["B12"].number_format = MONEY
    set_widths(ws, [16, 12, 11])


def build_expense_pair():
    ex = new_workbook("Expense Log", "Category Summary", "Monthly Summary (Expenses)")
    sol = new_workbook("Expense Log", "Category Summary", "Monthly Summary (Expenses)")
    fill_expense_log(ex["Expense Log"])
    fill_expense_log(sol["Expense Log"])
    add_validation(sol["Expense Log"], "B5:B104",
                   ",".join(CATEGORIES))
    fill_expense_category(ex["Category Summary"], solved=False)
    fill_expense_monthly(ex["Monthly Summary (Expenses)"], solved=False)
    fill_expense_category(sol["Category Summary"], solved=True)
    fill_expense_monthly(sol["Monthly Summary (Expenses)"], solved=True)
    return ex, sol, "expense-tracker-project"


# ----------------------------------------------------------------------------
# Project 7 — Invoice Generator
# ----------------------------------------------------------------------------
CUSTOMERS = [  # id, name, address, city, GSTIN
    ("C001", "Sharma Book House", "22 Anna Salai", "Chennai", "33AABCS1429P1Z5"),
    ("C002", "Gupta Stationers", "45 Karol Bagh", "Delhi", "07AABCG2210K1Z2"),
    ("C003", "Mehra Office Supply", "5 MG Road", "Bengaluru", "29AABCM9012P1Z8"),
    ("C004", "Nair Traders", "12 Andheri East", "Mumbai", "27AABCN7754L1ZK"),
    ("C005", "Kaur Enterprises", "8 Sector 22", "Chandigarh", "04AABCK3390M1ZQ"),
    ("C006", "Reddy & Sons", "76 Begumpet", "Hyderabad", "36AACCR6108N1ZX"),
]
INVOICE_LINES = [  # description, qty, rate, gst rate
    ("Notebooks", 100, 45, 0.12),
    ("Ball Pens (box)", 300, 15, 0.12),
    ("Staplers", 10, 180, 0.12),
    ("Printers", 2, 6800, 0.18),
    ("Laptop", 1, 43500, 0.18),
]


def fill_customer_master(wb):
    ws = wb["Customer Master"]
    put_title(ws, "Customer Master — Sunrise Traders",
              "Each buyer typed once; the invoice pulls name, address and GSTIN from here.")
    r = put_headers(ws, 4, ["Customer ID", "Customer Name", "Address", "City", "GSTIN"])
    put_rows(ws, r, [list(c) for c in CUSTOMERS])
    return ws


def fill_invoice_sheet(ws, solved):
    put_title(ws, "Sunrise Traders", "12, MG Road, Bengaluru · GSTIN 29AABCS3456K1Z2")
    ws["A1"].font = Font(bold=True, size=16, color=TEAL_DARK)
    ws["E4"] = "Customer ID:"
    ws["E4"].font = Font(bold=True)
    ws["E5"] = "Invoice No:"
    ws["E6"] = "Date:"
    for cell in ("E4", "E5", "E6"):
        ws[cell].font = Font(bold=True)
    hdr = ["S.No", "Description", "Qty", "Rate (₹)", "Amount (₹)", "GST %", "GST (₹)", "Total (₹)"]
    put_headers(ws, 11, hdr)
    if solved:
        ws["B4"] = "C003"
        ws["F4"] = "INV-1042"
        ws["F5"] = date(2026, 11, 14)
        ws["F5"].number_format = DATEFMT
        cm = "'Customer Master'"
        match = "MATCH($B$4,%s!$A$5:$A$10,0)" % cm
        ws["B6"] = "=IFERROR(INDEX(%s!$B$5:$B$10,%s),\"\")" % (cm, match)
        ws["B7"] = "=IFERROR(INDEX(%s!$C$5:$C$10,%s),\"\")" % (cm, match)
        ws["B8"] = "=IFERROR(INDEX(%s!$D$5:$D$10,%s),\"\")" % (cm, match)
        ws["E8"] = "=IFERROR(INDEX(%s!$E$5:$E$10,%s),\"\")" % (cm, match)
        for rr in range(6, 9):
            ws.cell(row=rr, column=1).font = Font(bold=True)
        ws["A6"] = "Name:"
        ws["A7"] = "Address:"
        ws["A8"] = "City:"
        ws["D8"] = "GSTIN:"
        for i, (desc, qty, rate, gst) in enumerate(INVOICE_LINES):
            rr = 12 + i
            ws.cell(row=rr, column=1, value='=IF($B%d="","",ROW()-11)' % rr)
            ws.cell(row=rr, column=2, value=desc)
            ws.cell(row=rr, column=3, value=qty)
            ws.cell(row=rr, column=4, value=rate).number_format = MONEY
        for rr in range(12, 22):  # guarded formulas on all ten item lines
            ws.cell(row=rr, column=1, value='=IF($B%d="","",ROW()-11)' % rr)
            ws.cell(row=rr, column=5, value='=IF($B%d="","",ROUND($C%d*$D%d,0))' % (rr, rr, rr))
            ws.cell(row=rr, column=7, value='=IF($B%d="","",ROUND($E%d*$F%d,0))' % (rr, rr, rr))
            ws.cell(row=rr, column=8, value='=IF($B%d="","",$E%d+$G%d)' % (rr, rr, rr))
        for rr in range(12, 22):
            for col in range(1, 9):
                ws.cell(row=rr, column=col).border = _BOX
        ws["D23"] = "Subtotal"
        ws["E23"] = "=ROUND(SUM($E$12:$E$21),0)"
        ws["G23"] = "Total GST"
        ws["H23"] = "=ROUND(SUM($G$12:$G$21),0)"
        ws["D24"] = "Grand Total"
        ws["H24"] = "=$E$23+$H$23"
        for cell in ("D23", "G23", "D24"):
            ws[cell].font = Font(bold=True)
        for cell in ("E23", "H23"):
            ws[cell].number_format = MONEY
        ws["H24"].font = Font(bold=True)
        ws["H24"].number_format = MONEY
        ws["H24"].fill = PatternFill("solid", fgColor=BAND)
        ws.print_area = "A1:H25"
    else:
        put_note(ws, 5, "Raise an invoice: type the Customer ID in B4 and the invoice number in F4. "
                 "The customer block, item lines and totals are built with the lesson formulas — "
                 "every item-table formula is guarded with IF($B12=\"\",\"\",…).", width=8)
        for rr in range(12, 22):
            for col in range(1, 9):
                ws.cell(row=rr, column=col).border = _BOX
    set_widths(ws, [6, 22, 8, 11, 12, 8, 10, 12])


def fill_invoice_dataentry(ws, solved):
    put_title(ws, "Data Entry — Invoice Register",
              "One row per invoice raised. The customer name fills itself from the master.")
    put_headers(ws, 4, ["Invoice No", "Date", "Customer ID", "Customer Name", "Grand Total"])
    if solved:
        ws["A5"] = "INV-1042"
        c = ws["B5"]
        c.value = date(2026, 11, 14)
        c.number_format = DATEFMT
        ws["C5"] = "C003"
        ws["D5"] = ("=IFERROR(INDEX('Customer Master'!$B$5:$B$10,"
                    "MATCH($C5,'Customer Master'!$A$5:$A$10,0)),\"\")")
        ws["E5"] = 79474
        ws["E5"].number_format = MONEY
        for rr in range(5, 25):
            for col in range(1, 6):
                ws.cell(row=rr, column=col).border = _BOX
        for rr in range(6, 25):  # guarded lookup fills itself as IDs are typed
            ws.cell(row=rr, column=4,
                    value=('=IF($C%d="","",IFERROR(INDEX(\'Customer Master\'!$B$5:$B$10,'
                           'MATCH($C%d,\'Customer Master\'!$A$5:$A$10,0)),""))') % (rr, rr))
    set_widths(ws, [12, 12, 12, 22, 13])


def build_invoice_pair():
    ex = new_workbook("Customer Master", "Invoice", "Data Entry")
    sol = new_workbook("Customer Master", "Invoice", "Data Entry")
    fill_customer_master(ex)
    fill_customer_master(sol)
    fill_invoice_sheet(ex["Invoice"], solved=False)
    fill_invoice_dataentry(ex["Data Entry"], solved=False)
    fill_invoice_sheet(sol["Invoice"], solved=True)
    fill_invoice_dataentry(sol["Data Entry"], solved=True)
    return ex, sol, "invoice-generator"


# ----------------------------------------------------------------------------
# Project 8 — Purchase Management System
# ----------------------------------------------------------------------------
SUPPLIERS = [  # id, name, city, GSTIN, payment terms (days)
    ("S001", "Kaviraj Papers Ltd.", "Rajkot", "24AABCK1234M1ZP", 30),
    ("S002", "Metro Plastics Co.", "Delhi", "07AABCM5678P1Z3", 15),
    ("S003", "TechnoSource India", "Pune", "27AAACT9012R1Z8", 30),
    ("S004", "Balaji Traders", "Chennai", "33AABCB4567L1ZK", 45),
    ("S005", "WinEdge Solutions", "Noida", "09AABCW8890H1ZQ", 15),
]
PO_ROWS = [  # po_no, date, sup, item, qty, rate, status, recv_date, pay_status, paid
    ("PO-1001", date(2026, 11, 1), "S001", "A4 Paper Reams", 200, 95,
     "Received", date(2026, 11, 5), "Paid", 19000),
    ("PO-1002", date(2026, 11, 3), "S003", "Laser Toner Cartridges", 20, 2400,
     "Received", date(2026, 11, 8), "Part Paid", 20000),
    ("PO-1003", date(2026, 11, 5), "S002", "Ball Pen Refills", 1500, 20,
     "Open", None, "Unpaid", 0),
    ("PO-1004", date(2026, 11, 8), "S004", "Stapler Bodies", 200, 90,
     "Open", None, "Unpaid", 0),
    ("PO-1005", date(2026, 11, 12), "S005", "Office Chairs", 12, 4250,
     "Received", date(2026, 11, 14), "Unpaid", 0),
    ("PO-1006", date(2026, 11, 15), "S001", "Binder Clips (Bulk)", 400, 30,
     "Cancelled", None, "Unpaid", 0),
]


def build_purchase_pair():
    ex = new_workbook("Supplier Master", "PO Register", "Summary")
    sol = new_workbook("Supplier Master", "PO Register", "Summary")
    for wb in (ex, sol):
        ws = wb["Supplier Master"]
        put_title(ws, "Supplier Master — Sunrise Traders",
                  "Each supplier typed once; the register looks names up from here.")
        r = put_headers(ws, 4, ["Sup ID", "Supplier Name", "City", "GSTIN",
                                "Payment Terms (Days)"])
        put_rows(ws, r, [list(s) for s in SUPPLIERS])

    reg_hdr = ["PO No", "Date", "Sup ID", "Supplier Name", "Item", "Qty", "Rate (₹)",
               "PO Value (₹)", "Status", "Received Date", "Payment Status",
               "Amount Paid (₹)", "Balance Due (₹)"]
    ws = ex["PO Register"]
    put_note(ws, 1, "One row per purchase order. Build the name, value and balance formulas, "
             "then complete the status columns with the dropdowns (see the lesson).", width=13)
    put_headers(ws, 4, reg_hdr)

    ws = sol["PO Register"]
    put_title(ws, "PO Register — Sunrise Traders",
              "Balance Due = IF(status = \"Cancelled\", 0, PO value − amount paid).")
    put_headers(ws, 4, reg_hdr)
    for i, (po, d, sup, item, qty, rate, status, recv, pay, paid) in enumerate(PO_ROWS):
        rr = 5 + i
        ws.cell(row=rr, column=1, value=po)
        c = ws.cell(row=rr, column=2, value=d)
        c.number_format = DATEFMT
        ws.cell(row=rr, column=3, value=sup)
        ws.cell(row=rr, column=4,
                value="=IF($C%d=\"\",\"\",VLOOKUP($C%d,'Supplier Master'!$A$5:$B$9,2,FALSE))" % (rr, rr))
        ws.cell(row=rr, column=5, value=item)
        ws.cell(row=rr, column=6, value=qty)
        ws.cell(row=rr, column=7, value=rate).number_format = MONEY
        ws.cell(row=rr, column=8, value='=IF($C%d="","",ROUND($F%d*$G%d,0))' % (rr, rr, rr))
        ws.cell(row=rr, column=8).number_format = MONEY
        ws.cell(row=rr, column=9, value=status)
        if recv:
            c = ws.cell(row=rr, column=10, value=recv)
            c.number_format = DATEFMT
        ws.cell(row=rr, column=11, value=pay)
        ws.cell(row=rr, column=12, value=paid).number_format = MONEY
        ws.cell(row=rr, column=13, value='=IF($I%d="Cancelled",0,$H%d-$L%d)' % (rr, rr, rr))
        ws.cell(row=rr, column=13).number_format = MONEY
        for col in range(1, 14):
            ws.cell(row=rr, column=col).border = _BOX
    add_validation(sol["PO Register"], "C5:C30", "S001,S002,S003,S004,S005")
    add_validation(sol["PO Register"], "I5:I30", "Open,Received,Cancelled")
    add_validation(sol["PO Register"], "K5:K30", "Unpaid,Part Paid,Paid")
    equal_rule(sol["PO Register"], "I5:I30", "Open", _AMBER)
    equal_rule(sol["PO Register"], "I5:I30", "Received", _GREEN)
    equal_rule(sol["PO Register"], "I5:I30", "Cancelled", _GREY)
    equal_rule(sol["PO Register"], "K5:K30", "Unpaid", _RED)
    equal_rule(sol["PO Register"], "K5:K30", "Paid", _GREEN)
    from openpyxl.formatting.rule import CellIsRule
    sol["PO Register"].conditional_formatting.add(
        "M5:M30", CellIsRule(operator="greaterThan", formula=["0"],
                             fill=PatternFill("solid", fgColor=_RED)))
    set_widths(sol["PO Register"], [9, 11, 8, 20, 20, 8, 10, 11, 10, 12, 12, 12, 12])

    ws = sol["Summary"]
    put_title(ws, "Purchase Summary — Sunrise Traders",
              "Status block and supplier-wise ledger — everything reads the register, nothing is typed.")
    ws["A3"], ws["B3"], ws["C3"] = "Status", "Count", "Value (₹)"
    for col in (1, 2, 3):
        style_cell(ws, 3, col, None, bold=True, fill=TEAL)
    for i, st in enumerate(["Open", "Received", "Cancelled"]):
        rr = 4 + i
        ws.cell(row=rr, column=1, value=st)
        ws.cell(row=rr, column=2, value="=COUNTIF('PO Register'!$I$5:$I$30,$A%d)" % rr)
        ws.cell(row=rr, column=3,
                value="=SUMIF('PO Register'!$I$5:$I$30,$A%d,'PO Register'!$H$5:$H$30)" % rr)
        ws.cell(row=rr, column=3).number_format = MONEY
        for col in range(1, 4):
            ws.cell(row=rr, column=col).border = _BOX
    ws["A8"] = "Total Pending Value (Open POs)"
    ws["A8"].font = Font(bold=True)
    ws["C8"] = "=SUMIF('PO Register'!$I$5:$I$30,\"Open\",'PO Register'!$H$5:$H$30)"
    ws["C8"].number_format = MONEY
    ws["A10"], ws["B10"], ws["C10"], ws["D10"], ws["E10"], ws["F10"], ws["G10"] = \
        "Sup ID", "Supplier", "POs", "Total Value (₹)", "Open Value (₹)", "Paid (₹)", "Balance Due (₹)"
    for col in range(1, 8):
        style_cell(ws, 10, col, None, bold=True, fill=TEAL)
    for i, (sup, *_rest) in enumerate(SUPPLIERS):
        rr = 11 + i
        ws.cell(row=rr, column=1, value=sup)
        ws.cell(row=rr, column=2,
                value="=IF($A%d=\"\",\"\",VLOOKUP($A%d,'Supplier Master'!$A$5:$B$9,2,FALSE))" % (rr, rr))
        ws.cell(row=rr, column=3, value="=COUNTIF('PO Register'!$C$5:$C$30,$A%d)" % rr)
        ws.cell(row=rr, column=4,
                value="=SUMIF('PO Register'!$C$5:$C$30,$A%d,'PO Register'!$H$5:$H$30)" % rr)
        ws.cell(row=rr, column=4).number_format = MONEY
        ws.cell(row=rr, column=5,
                value=("=SUMIFS('PO Register'!$H$5:$H$30,'PO Register'!$C$5:$C$30,$A%d,"
                       "'PO Register'!$I$5:$I$30,\"Open\")") % rr)
        ws.cell(row=rr, column=5).number_format = MONEY
        ws.cell(row=rr, column=6,
                value="=SUMIF('PO Register'!$C$5:$C$30,$A%d,'PO Register'!$L$5:$L$30)" % rr)
        ws.cell(row=rr, column=6).number_format = MONEY
        ws.cell(row=rr, column=7,
                value="=SUMIF('PO Register'!$C$5:$C$30,$A%d,'PO Register'!$M$5:$M$30)" % rr)
        ws.cell(row=rr, column=7).number_format = MONEY
        for col in range(1, 8):
            ws.cell(row=rr, column=col).border = _BOX
    ws["A17"] = "Total Balance Due (₹)"
    ws["A17"].font = Font(bold=True)
    ws["G17"] = "=SUM($G$11:$G$15)"
    ws["G17"].number_format = MONEY
    set_widths(ws, [10, 20, 8, 14, 14, 12, 14])

    put_note(ex["Summary"], 1, "Build the status block and the supplier-wise ledger here — "
             "COUNTIF, SUMIF and SUMIFS over the register. See the lesson steps.", width=8)
    return ex, sol, "purchase-management-system"


# ----------------------------------------------------------------------------
# Project 9 — Interactive Sales Dashboard
# ----------------------------------------------------------------------------
def fill_dashboard_sheet(ws, solved):
    put_title(ws, "Sunrise Traders — Sales Dashboard")
    if not solved:
        put_note(ws, 3, "Build the dashboard here: one month selector in B3, KPI cards in row 5–6, "
                 "a Month Focus block, the mirrored leaderboard and the region block. "
                 "See the lesson steps.", width=8)
        return
    ws["A3"] = "Month:"
    ws["A3"].font = Font(bold=True)
    ws["B3"] = "Oct-26"
    ws["B3"].fill = PatternFill("solid", fgColor=_AMBER)
    ws["B3"].font = Font(bold=True)
    ws["B3"].border = _BOX
    add_validation(ws, "B3", "Aug-26,Sep-26,Oct-26")

    ws["A5"], ws["B5"], ws["C5"], ws["D5"] = "Total Revenue", "Team Achievement", "Best Month", "Top Performer"
    for col in range(1, 5):
        style_cell(ws, 5, col, None, bold=True, fill=TEAL)
    ws["A6"] = "='Monthly Summary'!$B$10"
    ws["A6"].number_format = MONEY
    ws["B6"] = "='Performance'!$C$11/'Performance'!$B$11"
    ws["B6"].number_format = PCT0
    ws["C6"] = "=INDEX('Monthly Summary'!$A$5:$A$7,MATCH(MAX('Monthly Summary'!$B$5:$B$7),'Monthly Summary'!$B$5:$B$7,0))"
    ws["D6"] = "=INDEX('Performance'!$A$5:$A$10,MATCH(1,'Performance'!$E$5:$E$10,0))"
    for col in range(1, 5):
        style_cell(ws, 6, col, None, bold=True, fill=_GREEN)

    ws["A8"], ws["B8"], ws["C8"] = "Measure", "Value", "Against Quarter"
    for col in range(1, 4):
        style_cell(ws, 8, col, None, bold=True, fill=TEAL)
    ws["A9"] = "Revenue — selected month"
    ws["B9"] = "=SUMIFS('Sales Log'!$H$5:$H$204,'Sales Log'!$I$5:$I$204,$B$3)"
    ws["B9"].number_format = MONEY
    ws["C9"] = "=$B$9/$A$6"
    ws["C9"].number_format = PCT1
    ws["A10"] = "Growth on previous month"
    ws["B10"] = ('=IF(MATCH($B$3,\'Monthly Summary\'!$A$5:$A$7,0)=1,"—",'
                 "INDEX('Monthly Summary'!$C$5:$C$7,MATCH($B$3,'Monthly Summary'!$A$5:$A$7,0)))")
    ws["B10"].number_format = PCT1
    for rr in (9, 10):
        for col in range(1, 4):
            ws.cell(row=rr, column=col).border = _BOX

    ws["A12"], ws["B12"] = "Product", "Selected Month"
    style_cell(ws, 12, 1, None, bold=True, fill=TEAL)
    style_cell(ws, 12, 2, None, bold=True, fill=TEAL)
    for i, prod in enumerate(PRODUCTS):
        rr = 13 + i
        ws.cell(row=rr, column=1, value=prod)
        ws.cell(row=rr, column=2,
                value=("=SUMIFS('Sales Log'!$H$5:$H$204,'Sales Log'!$D$5:$D$204,$A%d,"
                       "'Sales Log'!$I$5:$I$204,$B$3)") % rr)
        ws.cell(row=rr, column=2).number_format = MONEY
        for col in range(1, 3):
            ws.cell(row=rr, column=col).border = _BOX

    for i in range(7):  # leaderboard mirror of Performance rows 4..10 (hdr + 6)
        for col in range(1, 7):
            letter = "ABCDEF"[col - 1]
            ws.cell(row=20 + i, column=col, value="='Performance'!%s%d" % (letter, 4 + i))
            ws.cell(row=20 + i, column=col).border = _BOX
    style_cell(ws, 20, 1, None, bold=True, fill=TEAL)
    style_cell(ws, 20, 2, None, bold=True, fill=TEAL)
    style_cell(ws, 20, 3, None, bold=True, fill=TEAL)
    style_cell(ws, 20, 4, None, bold=True, fill=TEAL)
    style_cell(ws, 20, 5, None, bold=True, fill=TEAL)
    style_cell(ws, 20, 6, None, bold=True, fill=TEAL)
    for rr in range(21, 27):
        ws.cell(row=rr, column=2).number_format = MONEY
        ws.cell(row=rr, column=3).number_format = MONEY
        ws.cell(row=rr, column=4).number_format = PCT0

    ws["A28"] = "Region"
    ws["B28"] = "Revenue"
    style_cell(ws, 28, 1, None, bold=True, fill=TEAL)
    style_cell(ws, 28, 2, None, bold=True, fill=TEAL)
    for i, reg in enumerate(REGIONS):
        rr = 29 + i
        ws.cell(row=rr, column=1, value=reg)
        ws.cell(row=rr, column=2,
                value="=SUMIFS('Sales Log'!$H$5:$H$204,'Sales Log'!$C$5:$C$204,$A%d)" % rr)
        ws.cell(row=rr, column=2).number_format = MONEY
        for col in range(1, 3):
            ws.cell(row=rr, column=col).border = _BOX
    ws["A34"] = "90%-club"
    ws["A34"].font = Font(bold=True)
    ws["B34"] = "=COUNTIF('Performance'!$D$5:$D$10,\">=90%\")"

    src = ws.parent["Monthly Summary"]
    chart = BarChart()
    chart.type = "col"
    chart.title = "Monthly Revenue (₹)"
    chart.height, chart.width = 7.5, 13
    data = Reference(src, min_col=2, min_row=4, max_row=7)
    cats = Reference(src, min_col=1, min_row=5, max_row=7)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.legend = None
    ws.add_chart(chart, "D28")
    set_widths(ws, [24, 18, 16, 16])


def build_dashboard_pair():
    # finished Project 5 workbook + empty Dashboard (exercise)
    ex = new_workbook("Sales Log", "Monthly Summary", "Performance", "Dashboard")
    fill_sales_log(ex["Sales Log"])
    fill_sales_monthly(ex["Monthly Summary"], solved=True)
    fill_sales_performance(ex["Performance"], solved=True)
    fill_dashboard_sheet(ex["Dashboard"], solved=False)

    sol = new_workbook("Sales Log", "Monthly Summary", "Performance", "Dashboard")
    fill_sales_log(sol["Sales Log"])
    fill_sales_monthly(sol["Monthly Summary"], solved=True)
    fill_sales_performance(sol["Performance"], solved=True)
    fill_dashboard_sheet(sol["Dashboard"], solved=True)
    return ex, sol, "sales-dashboard"


# ----------------------------------------------------------------------------
# Project 10 — Complete Office MIS
# ----------------------------------------------------------------------------
def fill_mis_sheet(ws, solved):
    put_title(ws, "Sunrise Traders — Office MIS", "Aug–Oct 2026 · expenses tracked Jul–Oct 2026")
    if not solved:
        put_note(ws, 4, "Build the one-page read here: headline cards, the month table "
                 "(revenue vs expenses with net), the mirrored leaderboard and the expense "
                 "category block. Every cell is a label or a formula. See the lesson steps.", width=9)
        return
    ws["A3"] = "Revenue"
    ws["B3"] = "Expenses"
    ws["C3"] = "Net Position"
    for col in range(1, 4):
        style_cell(ws, 3, col, None, bold=True, fill=TEAL)
    ws["A4"] = "='Monthly Summary'!$B$10"
    ws["A4"].number_format = MONEY
    ws["B4"] = "='Category Summary'!$B$11"
    ws["B4"].number_format = MONEY
    ws["C4"] = "=$A$4-$B$4"
    ws["C4"].number_format = MONEY
    for col in range(1, 4):
        style_cell(ws, 4, col, None, bold=True, fill=_GREEN)
    ws["A5"] = "Vs target"
    ws["B5"] = "=$A$4/1010000"
    ws["B5"].number_format = PCT1
    ws["C5"] = "Vs budget"
    ws["D5"] = "=$B$4/808000"
    ws["D5"].number_format = PCT1

    ws["A7"], ws["B7"], ws["C7"], ws["D7"] = "Month", "Revenue (₹)", "Expenses (₹)", "Net (₹)"
    for col in range(1, 5):
        style_cell(ws, 7, col, None, bold=True, fill=TEAL)
    for i, m in enumerate(MONTHS + ["Jul-26"]):
        rr = 8 + i
        ws.cell(row=rr, column=1, value=m)
        if m != "Jul-26":
            ws.cell(row=rr, column=2,
                    value="=SUMIFS('Sales Log'!$H$5:$H$204,'Sales Log'!$I$5:$I$204,$A%d)" % rr)
            ws.cell(row=rr, column=2).number_format = MONEY
        ws.cell(row=rr, column=3,
                value="=SUMIFS('Expense Log'!$D$5:$D$104,'Expense Log'!$E$5:$E$104,$A%d)" % rr)
        ws.cell(row=rr, column=3).number_format = MONEY
        if m != "Jul-26":
            ws.cell(row=rr, column=4, value="=$B%d-$C%d" % (rr, rr))
            ws.cell(row=rr, column=4).number_format = MONEY
        for col in range(1, 5):
            ws.cell(row=rr, column=col).border = _BOX
    ws["A13"] = ("Note: the sales log starts Aug-26, so Jul-26 shows expenses (₹2,07,650) "
                 "with no revenue.")
    ws["A13"].font = Font(italic=True, size=9, color=GRAY)
    ws["A14"] = "Revenue total (Aug–Oct)"
    ws["A14"].font = Font(bold=True)
    ws["B14"] = "=SUM($B$8:$B$10)"
    ws["B14"].number_format = MONEY

    for i in range(7):  # leaderboard mirror of Performance rows 4..10
        for col in range(1, 7):
            letter = "ABCDEF"[col - 1]
            ws.cell(row=16 + i, column=col, value="='Performance'!%s%d" % (letter, 4 + i))
            ws.cell(row=16 + i, column=col).border = _BOX
    for col in range(1, 7):
        style_cell(ws, 16, col, None, bold=True, fill=TEAL)

    for i in range(8):  # expense category mirror of Category Summary rows 4..11 (hdr+6+total)
        for col in range(1, 6):
            letter = "ABCDE"[col - 1]
            ws.cell(row=25 + i, column=col, value="='Category Summary'!%s%d" % (letter, 4 + i))
            ws.cell(row=25 + i, column=col).border = _BOX
    for col in range(1, 6):
        style_cell(ws, 25, col, None, bold=True, fill=TEAL)
    for rr in range(26, 33):
        for col in (2, 3, 4):
            ws.cell(row=rr, column=col).number_format = MONEY
    set_widths(ws, [24, 18, 16, 16, 12, 12])


def fill_close_sheet(ws, solved):
    put_title(ws, "Reconciliation Checklist",
              "One subtraction per check — every result must read 0 before the period closes.")
    put_headers(ws, 4, ["Check", "Result", "Verdict"])
    if not solved:
        put_note(ws, 5, "Add the four checks: revenue, expenses, expense log vs category summary, "
                 "and sales log vs performance actuals. Each is a subtraction that must read 0, "
                 "with =IF(result=0,\"PASS\",\"CHECK\") as the verdict.", width=6)
        return
    checks = [
        ("Revenue: summary total vs MIS month table",
         "='Monthly Summary'!$B$10-SUM(MIS!$B$8:$B$10)"),
        ("Expenses: category total vs MIS month table",
         "='Category Summary'!$B$11-SUM(MIS!$C$8:$C$11)"),
        ("Expense log vs category summary",
         "=SUM('Expense Log'!$D$5:$D$104)-'Category Summary'!$B$11"),
        ("Sales log vs performance actuals",
         "='Performance'!$C$11-SUM('Sales Log'!$H$5:$H$204)"),
    ]
    for i, (label, formula) in enumerate(checks):
        rr = 5 + i
        ws.cell(row=rr, column=1, value=label)
        ws.cell(row=rr, column=2, value=formula).number_format = MONEY
        ws.cell(row=rr, column=3, value='=IF($B%d=0,"PASS","CHECK")' % rr)
        for col in range(1, 4):
            ws.cell(row=rr, column=col).border = _BOX
    equal_rule(ws, "C5:C8", "PASS", _GREEN)
    equal_rule(ws, "C5:C8", "CHECK", _RED)
    set_widths(ws, [44, 12, 10])


def build_mis_pair():
    sheets = ["MIS", "Close", "Sales Log", "Monthly Summary", "Performance", "Dashboard",
              "Expense Log", "Category Summary", "Monthly Summary (Expenses)"]
    ex = new_workbook(*sheets)
    sol = new_workbook(*sheets)
    for wb in (ex, sol):
        fill_sales_log(wb["Sales Log"])
        fill_sales_monthly(wb["Monthly Summary"], solved=True)
        fill_sales_performance(wb["Performance"], solved=True)
        fill_dashboard_sheet(wb["Dashboard"], solved=True)
        fill_expense_log(wb["Expense Log"])
        add_validation(wb["Expense Log"], "B5:B104", ",".join(CATEGORIES))
        fill_expense_category(wb["Category Summary"], solved=True)
        fill_expense_monthly(wb["Monthly Summary (Expenses)"], solved=True)
    fill_mis_sheet(ex["MIS"], solved=False)
    fill_close_sheet(ex["Close"], solved=False)
    fill_mis_sheet(sol["MIS"], solved=True)
    fill_close_sheet(sol["Close"], solved=True)
    return ex, sol, "office-mis-system"


# ----------------------------------------------------------------------------
# Aggregate verification (recomputed from the source data written above)
# ----------------------------------------------------------------------------
def verify_aggregates():
    from collections import defaultdict
    m_tot, p_tot, r_tot, pr_tot = (defaultdict(int) for _ in range(4))
    for d, region, product, person, qty, rate in SALES_ROWS:
        amt = qty * rate
        m_tot[d.strftime("%b-%y")] += amt
        p_tot[person] += amt
        r_tot[region] += amt
        pr_tot[product] += amt
    assert dict(m_tot) == {"Aug-26": 320000, "Sep-26": 345000, "Oct-26": 398500}, dict(m_tot)
    expected_person = dict(zip(SALESPEOPLE, [158000, 212500, 114000, 266000, 124000, 189000]))
    assert dict(p_tot) == expected_person, dict(p_tot)
    assert dict(r_tot) == {"North": 294000, "South": 267500, "East": 258000, "West": 244000}, dict(r_tot)
    assert dict(pr_tot) == {"Notebooks": 284000, "Laptops": 310000, "Printers": 220000,
                            "Ink & Toner": 144500, "Stationery Kits": 105000}, dict(pr_tot)
    total = sum(q * r for *_a, q, r in SALES_ROWS)
    assert total == 1063500, total
    assert sum(TARGETS) == 1010000
    ach = {p: p_tot[p] / t for p, t in zip(SALESPEOPLE, TARGETS)}
    assert round(ach["Divya Menon"], 2) == 1.18 and ach["Divya Menon"] == max(ach.values())
    assert round(ach["Rahul Verma"], 2) == 0.95 and round(ach["Vikram Singh"], 2) == 0.95
    ranks = sorted(ach, key=lambda p: -ach[p])
    assert ranks[0] == "Divya Menon" and ranks[-1] == "Rahul Verma"

    e_cat, e_month = defaultdict(int), defaultdict(int)
    for d, cat, desc, amt in EXPENSE_ROWS:
        e_cat[cat] += amt
        e_month[d.strftime("%b-%y")] += amt
    assert dict(e_cat) == {"Rent": 160000, "Salaries": 462000, "Electricity": 51200,
                           "Transport": 55000, "Marketing": 86500, "Miscellaneous": 36200}, dict(e_cat)
    assert dict(e_month) == {"Jul-26": 207650, "Aug-26": 212100, "Sep-26": 216200,
                             "Oct-26": 214950}, dict(e_month)
    assert sum(e_cat.values()) == 850900 and sum(BUDGETS) == 868000
    # MIS month nets
    assert 320000 - 212100 == 107900 and 345000 - 216200 == 128800 and 398500 - 214950 == 183550
    assert 1063500 - 850900 == 212600
    # inventory: closing stock and value
    inw, outw = defaultdict(int), defaultdict(int)
    for _d, code, qty in INWARD:
        inw[code] += qty
    for _d, code, qty in OUTWARD:
        outw[code] += qty
    closing = {c: o + inw[c] - outw[c] for c, n, u, o, r, rl in ITEMS}
    assert closing == {"ITM001": 230, "ITM002": 340, "ITM003": 25,
                       "ITM004": 12, "ITM005": 8}, closing
    value = sum(closing[c] * dict((i[0], i[4]) for i in ITEMS)[c] for c in closing)
    assert value == 431030, value
    # payroll
    assert sum(e[6] for e in EMPLOYEES) == 209000
    # purchase
    assert sum(q * r for *_x, q, r in [(p[4], p[5]) for p in PO_ROWS]) == 178000
    print("Aggregate verification: ALL CHECKS PASSED")
    print("  Sales months:", dict(m_tot))
    print("  Sales people:", dict(p_tot))
    print("  Regions:", dict(r_tot))
    print("  Products:", dict(pr_tot))
    print("  Expense categories:", dict(e_cat))
    print("  Expense months:", dict(e_month))


def main():
    os.makedirs(OUT, exist_ok=True)
    builders = [
        build_employee_pair,
        build_attendance_pair,
        build_payroll_pair,
        build_inventory_pair,
        build_sales_pair,
        build_expense_pair,
        build_invoice_pair,
        build_purchase_pair,
        build_dashboard_pair,
        build_mis_pair,
    ]
    verify_aggregates()
    for build in builders:
        ex, sol, name = build()
        for wb, suffix in ((ex, ""), (sol, "-solution")):
            path = os.path.join(OUT, "%s%s.xlsx" % (name, suffix))
            wb.save(path)
            print("saved", path, os.path.getsize(path), "bytes")


if __name__ == "__main__":
    main()



