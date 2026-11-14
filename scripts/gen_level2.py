# -*- coding: utf-8 -*-
"""Generate the Level 2 practice workbooks for Excel Lab.

Creates exactly the five .xlsx files referenced by the Level 2 lesson
downloads entries:

    assets/files/level2/formula-practice.xlsx     (lessons: SUM/AVERAGE/MIN/MAX, COUNT...ROUND)
    assets/files/level2/logical-functions.xlsx    (lessons: IF, AND/OR/IFS, IFERROR)
    assets/files/level2/lookup-practice.xlsx      (lessons: VLOOKUP, XLOOKUP, INDEX/MATCH)
    assets/files/level2/sumif-practice.xlsx       (lessons: SUMIF/COUNTIF, SUMIFS/COUNTIFS/AVERAGEIFS)
    assets/files/level2/text-date-practice.xlsx   (lesson: Text Functions)

Each workbook has an "Exercise" sheet (starter state, amber answer cells
left empty) and a "Solution" sheet (finished state: formulas as strings
starting with "=" so Excel computes them, mirroring the exact numbers in
the lesson exercise solutions).

Run from the repo root:

    .venv/Scripts/python.exe scripts/gen_level2.py
"""
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Font, PatternFill, Side
from xlcommon import (
    BAND,
    GRAY,
    new_workbook,
    put_headers,
    put_note,
    put_rows,
    put_section,
    put_title,
    set_widths,
)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "assets", "files", "level2")

_THIN = Side(style="thin", color="cbd5e1")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
ANSWER_FILL = PatternFill("solid", fgColor="fef9ec")  # pale amber: "type here"

# XLOOKUP and IFS post-date the xlsx formula grammar, so openpyxl must
# store them with the _xlfn prefix — Excel then shows/computes them fine.
XLOOKUP = "_xlfn.XLOOKUP"
IFS = "_xlfn.IFS"


def _answer(ws, r, c, value, fmt=None):
    """Styled answer/input cell (amber tint, bordered)."""
    cell = ws.cell(row=r, column=c, value=value)
    cell.border = _BORDER
    cell.fill = ANSWER_FILL
    if fmt:
        cell.number_format = fmt
    return cell


def _qa(ws, r, solution, question, formula=None, inp=None, ref=None, fmt=None):
    """Q&A row: question in A, optional input in B, answer in C, ref in D."""
    q = ws.cell(row=r, column=1, value=question)
    q.font = Font(size=10, color=GRAY)
    q.alignment = None
    if inp is not None:
        _answer(ws, r, 2, inp, fmt)
    else:
        ws.cell(row=r, column=2, value=None).border = _BORDER
    _answer(ws, r, 3, formula if solution else None, fmt)
    if ref:
        c = ws.cell(row=r, column=4, value=ref)
        c.font = Font(size=9, italic=True, color=GRAY)
    return r + 1


# ---------------------------------------------------------------------------
# Workbook 1: formula-practice.xlsx
# Lessons: l2-sum-average-min-max, l2-count-round
# ---------------------------------------------------------------------------

def build_formula_practice(ws, solution):
    put_title(ws, "Formula Practice — Exercise" if not solution else "Formula Practice — Solution")
    if solution:
        put_note(ws, 2, "Finished state. Every amber cell holds the completed formula. "
                        "Expected results: total 11200 · average 1600 · max 2100 · min 980 · "
                        "counts 7 / 9 / 2 · two-block sum 9200 · rounded amounts match the printed prices.")
    else:
        put_note(ws, 2, "Covers the lessons 'SUM, AVERAGE, MIN and MAX' and 'COUNT, COUNTA, COUNTBLANK and ROUND'. "
                        "Fill the amber answer cells, then compare with the Solution sheet. "
                        "The lesson's own cell references are shown beside each task.")

    # --- Section 1: daily sales -------------------------------------------
    put_section(ws, 4, "1 · Daily sales — SUM · AVERAGE · MIN · MAX · counting")
    put_headers(ws, 5, ["Day", "Sales (₹)"])
    sales = [1250, 980, None, 2100, 2000, None, 1650, 1320, 1900]
    rows = [[f"Day {i + 1}", s] for i, s in enumerate(sales)]
    put_rows(ws, 6, rows, num_formats={2: "#,##0"})

    summary = [
        ("Total — AutoSum / SUM", "=SUM(B6:B14)"),
        ("Average — AVERAGE", "=AVERAGE(B6:B14)"),
        ("Highest — MAX", "=MAX(B6:B14)"),
        ("Lowest — MIN", "=MIN(B6:B14)"),
    ]
    r = 16
    for label, formula in summary:
        c = ws.cell(row=r, column=1, value=label)
        c.font = Font(size=10, color=GRAY)
        _answer(ws, r, 2, formula if solution else None, "#,##0.00")
        r += 1

    # --- Section 2: counting check ----------------------------------------
    put_section(ws, 21, "2 · Counting check — COUNT · COUNTA · COUNTBLANK · two blocks in one SUM")
    put_headers(ws, 22, ["Question", "Input", "Answer", "Lesson ref"])
    counting = [
        ("How many cells in the sales column hold numbers?", "=COUNT(B6:B14)", "lesson 3 · D2"),
        ("How many cells in the day column are filled at all?", "=COUNTA(A6:A14)", "lesson 3 · D3"),
        ("How many cells in the sales column are still empty?", "=COUNTBLANK(B6:B14)", "lesson 3 · D4"),
        ("Add two blocks in one SUM, skipping the middle rows:", "=SUM(B6:B9, B12:B14)", "lesson 2 · D2"),
    ]
    r = 23
    for question, formula, ref in counting:
        _qa(ws, r, solution, question, formula=formula, ref=ref, fmt="#,##0")
        r += 1

    # --- Section 3: stock list, ROUND and ROUNDUP -------------------------
    put_section(ws, 28, "3 · Stock list — ROUND the money · ROUNDUP the boxes")
    put_headers(ws, 29, ["Item", "Units sold", "Price (₹)", "Amount (₹)", "Boxes to order"])
    stock = [
        ("Notebooks", 40, 45),
        ("Pens", 100, 12),
        ("Staplers", 15, 130),
        ("Printers", 2, 8500),
        ("Laptops", 1, 42000),
    ]
    r = 30
    for item, units, price in stock:
        ws.cell(row=r, column=1, value=item).border = _BORDER
        ws.cell(row=r, column=2, value=units).border = _BORDER
        ws.cell(row=r, column=3, value=price).border = _BORDER
        amount = f"=ROUND(B{r}*C{r}, 2)" if solution else f"=B{r}*C{r}"
        _answer(ws, r, 4, amount, "#,##0.00")
        boxes = f"=ROUNDUP(B{r}/10, 0)" if solution else None
        _answer(ws, r, 5, boxes, "#,##0")
        r += 1

    if solution:
        put_section(ws, 36, "4 · Blanks are not zeros — compare the averages")
        pairs = [
            ("Average with the two empty days left blank", "=AVERAGE(B6:B14)"),
            ("Average if the two empty days were typed as 0 (SUM ÷ 9)", "=SUM(B6:B14)/9"),
        ]
        r = 37
        for label, formula in pairs:
            c = ws.cell(row=r, column=1, value=label)
            c.font = Font(size=10, color=GRAY)
            _answer(ws, r, 2, formula, "#,##0.00")
            r += 1
        put_note(ws, 39, "AVERAGE skips empty cells, so it divides by the 7 recorded days (1600). "
                         "Typed zeros count as real days, so the average drops to 11200 ÷ 9 = 1244.44. "
                         "Which one is right depends on what the number means — a business decision, not an Excel one.")

    set_widths(ws, [40, 14, 14, 14, 14, 14, 14, 14])


# ---------------------------------------------------------------------------
# Workbook 2: sumif-practice.xlsx
# Lessons: l2-sumif-countif, l2-sumifs-countifs
# ---------------------------------------------------------------------------

def build_sumif_practice(ws, solution):
    put_title(ws, "Conditional Aggregation — Exercise" if not solution else "Conditional Aggregation — Solution")
    if solution:
        put_note(ws, 2, "Finished state. Expected results: South 3900 · Notebooks count 3 · Notebook revenue 5625 · "
                        "orders above 5000: 2 / 59000 · region box: North 20100, West 43125 · East+Pens 960 · "
                        "North >1000: count 3, average 6700 (see the note below) · March 6085 · March+North 1300.")
    else:
        put_note(ws, 2, "Covers the lessons 'SUMIF and COUNTIF' and 'SUMIFS, COUNTIFS and AVERAGEIFS'. "
                        "Fill the amber answer cells, then compare with the Solution sheet. "
                        "The lesson's own cell references are shown beside each task.")

    # --- Main sales record -------------------------------------------------
    put_section(ws, 4, "1 · Sunrise Traders sales record")
    put_headers(ws, 5, ["Region", "Product", "Units", "Price (₹)", "Amount (₹)"])
    record = [
        ("North", "Notebooks", 40, 45),
        ("South", "Pens", 100, 12),
        ("East", "Staplers", 15, 130),
        ("North", "Printers", 2, 8500),
        ("West", "Laptops", 1, 42000),
        ("South", "Notebooks", 60, 45),
        ("East", "Pens", 80, 12),
        ("North", "Staplers", 10, 130),
        ("West", "Notebooks", 25, 45),
    ]
    rows = [[reg, prod, units, price, f"=C{6 + i}*D{6 + i}"] for i, (reg, prod, units, price) in enumerate(record)]
    put_rows(ws, 6, rows, num_formats={3: "#,##0", 4: "#,##0", 5: "#,##0"})

    # --- SUMIF / COUNTIF answer box ----------------------------------------
    put_section(ws, 16, "2 · SUMIF & COUNTIF — your answers")
    put_headers(ws, 17, ["Question", "Input", "Answer", "Lesson ref"])
    qa_sumif = [
        ("Total revenue for the South region", None, '=SUMIF($A$6:$A$14, "South", $E$6:$E$14)', "H2"),
        ("How many orders were for Notebooks?", None, '=COUNTIF($B$6:$B$14, "Notebooks")', "H3"),
        ("Total revenue of those Notebook orders", None, '=SUMIF($B$6:$B$14, "Notebooks", $E$6:$E$14)', "H4"),
        ("How many orders are above 5000?", None, '=COUNTIF($E$6:$E$14, ">5000")', "H5"),
        ("Total value of the orders above 5000", None, '=SUMIF($E$6:$E$14, ">5000")', "H6"),
        ("Total for whatever region you type in B23 (try North, then West)", "North", "=SUMIF($A$6:$A$14, B23, $E$6:$E$14)", "G8 · H8"),
    ]
    r = 18
    for question, inp, formula, ref in qa_sumif:
        _qa(ws, r, solution, question, formula=formula, inp=inp, ref=ref, fmt="#,##0")
        r += 1

    # --- Date-stamped table -------------------------------------------------
    put_section(ws, 25, "3 · Date-stamped orders — SUMIFS · COUNTIFS · AVERAGEIFS")
    put_headers(ws, 26, ["Date", "Region", "Product", "Order ID", "Amount (₹)"])
    dated = [
        (date(2026, 2, 5), "North", "Notebooks", "ORD-101", 1800),
        (date(2026, 2, 7), "South", "Pens", "ORD-102", 1200),
        (date(2026, 2, 11), "East", "Staplers", "ORD-103", 1950),
        (date(2026, 2, 18), "North", "Printers", "ORD-104", 17000),
        (date(2026, 2, 25), "West", "Laptops", "ORD-105", 42000),
        (date(2026, 3, 2), "South", "Notebooks", "ORD-106", 2700),
        (date(2026, 3, 9), "East", "Pens", "ORD-107", 960),
        (date(2026, 3, 16), "North", "Staplers", "ORD-108", 1300),
        (date(2026, 3, 24), "West", "Notebooks", "ORD-109", 1125),
    ]
    rows = [[d, reg, prod, oid, amt] for d, reg, prod, oid, amt in dated]
    put_rows(ws, 27, rows, num_formats={5: "#,##0"})
    for r in range(27, 36):
        ws.cell(row=r, column=1).number_format = "dd-mmm"

    # --- SUMIFS / COUNTIFS / AVERAGEIFS answer box --------------------------
    put_section(ws, 37, "4 · SUMIFS · COUNTIFS · AVERAGEIFS — your answers")
    put_headers(ws, 38, ["Question", "Input", "Answer", "Lesson ref"])
    qa_sumifs = [
        ("Total where the Region is East AND the Product is Pens", None,
         '=SUMIFS($E$27:$E$35, $A$27:$A$35, "East", $B$27:$B$35, "Pens")', "H2", "#,##0"),
        ("North orders with a total above 1000 — count", None,
         '=COUNTIFS($B$27:$B$35, "North", $E$27:$E$35, ">1000")', "H3", "#,##0"),
        ("…and their average value", None,
         '=AVERAGEIFS($E$27:$E$35, $B$27:$B$35, "North", $E$27:$E$35, ">1000")', "H4", "#,##0.00"),
        ("Period start — type a date", date(2026, 3, 1), None, "H6", "dd-mmm-yyyy"),
        ("Period end — type a date", date(2026, 3, 31), None, "H7", "dd-mmm-yyyy"),
        ("Total for the period above", None,
         '=SUMIFS($E$27:$E$35, $A$27:$A$35, ">="&$B$42, $A$27:$A$35, "<="&$B$43)', "H8", "#,##0"),
        ("Region filter — type a region", "North", None, "H9", None),
        ("Total for the period AND that region", None,
         '=SUMIFS($E$27:$E$35, $B$27:$B$35, $B$45, $A$27:$A$35, ">="&$B$42, $A$27:$A$35, "<="&$B$43)', "H9 · ex. 4", "#,##0"),
    ]
    r = 39
    for question, inp, formula, ref, fmt in qa_sumifs:
        _qa(ws, r, solution, question, formula=formula, inp=inp, ref=ref, fmt=fmt)
        r += 1
    put_note(ws, 47, "Change the period start in B42 to 01-Feb-2026 and the period total becomes 63950 — "
                     "the formula never changes, only the two date cells.")
    if solution:
        put_note(ws, 49, "Note: the lesson text quotes 2 orders / 9150 for 'North above 1000', but the table holds three "
                         "North orders above 1000 (1800 · 17000 · 1300), so these formulas correctly return 3 and 6700. "
                         "Recount rows 27, 30 and 34 to check.")

    set_widths(ws, [52, 15, 14, 12, 14])


# ---------------------------------------------------------------------------
# Workbook 3: logical-functions.xlsx
# Lessons: l2-if, l2-and-or, l2-iferror
# ---------------------------------------------------------------------------

def build_logical_functions(ws, solution):
    put_title(ws, "Logical Functions — Exercise" if not solution else "Logical Functions — Solution")
    if solution:
        put_note(ws, 2, "Finished state. Trainees: Pass/Fail and Distinction grades applied, training flags set, "
                        "bonus and delivery decisions made, per-order averages clean of #DIV/0!. "
                        "The Debug corner at the bottom hides one deliberately broken formula.")
    else:
        put_note(ws, 2, "Covers the lessons 'IF: Making Decisions in Excel', 'AND, OR and IFS' and "
                        "'IFERROR: Handling Errors Gracefully'. Fill the amber answer cells (one formula per column, "
                        "filled down), then compare with the Solution sheet.")

    # --- Section 1: trainee results (IF, nested IF, IFS) --------------------
    put_section(ws, 4, "1 · Trainee results — IF · nested IF · IFS (Excel 2019+)")
    put_headers(ws, 5, ["Trainee", "Marks", "Result (Pass/Fail)", "Training status",
                        "Training done", "Grade (nested IF)", "Grade (IFS)"])
    trainees = [
        ("Arjun Mehta", 78, "Yes"),
        ("Priya Nair", 35, "Yes"),
        ("Rahul Verma", 52, "No"),
        ("Sneha Iyer", 88, "Yes"),
        ("Vikram Singh", 40, "No"),
        ("Meera Pillai", 75, "Yes"),
    ]
    r = 6
    for name, marks, training in trainees:
        ws.cell(row=r, column=1, value=name).border = _BORDER
        ws.cell(row=r, column=2, value=marks).border = _BORDER
        ws.cell(row=r, column=5, value=training).border = _BORDER
        _answer(ws, r, 3, f'=IF(B{r}>=40, "Pass", "Fail")' if solution else None)
        _answer(ws, r, 4, f'=IF(E{r}<>"Yes", "Pending", "Done")' if solution else None)
        _answer(ws, r, 6, f'=IF(B{r}>=75, "Distinction", IF(B{r}>=40, "Pass", "Fail"))' if solution else None)
        _answer(ws, r, 7, f'={IFS}(B{r}>=75, "A grade", B{r}>=40, "B grade", TRUE, "C grade")' if solution else None)
        r += 1

    # --- Section 2: bonus & delivery (AND, OR) ------------------------------
    put_section(ws, 13, "2 · Bonus & delivery rules — AND · OR inside IF")
    put_headers(ws, 14, ["Employee", "Sales (₹)", "Training done", "Bonus decision",
                         "Order (₹)", "Region", "Delivery"])
    staff = [
        ("Arjun Mehta", 52000, "Yes", 6200, "North"),
        ("Priya Nair", 48000, "Yes", 3200, "South"),
        ("Rahul Verma", 61000, "No", 5800, "East"),
        ("Sneha Iyer", 75000, "Yes", 4100, "North"),
        ("Vikram Singh", 50000, "No", 1200, "West"),
        ("Meera Pillai", 50000, "Yes", 20000, "South"),
    ]
    r = 15
    for name, sales, training, order, region in staff:
        ws.cell(row=r, column=1, value=name).border = _BORDER
        ws.cell(row=r, column=2, value=sales).border = _BORDER
        ws.cell(row=r, column=3, value=training).border = _BORDER
        ws.cell(row=r, column=5, value=order).border = _BORDER
        ws.cell(row=r, column=6, value=region).border = _BORDER
        _answer(ws, r, 4, f'=IF(AND(B{r}>=50000, C{r}="Yes"), "Bonus", "No bonus")' if solution else None)
        _answer(ws, r, 7, f'=IF(OR(E{r}>5000, F{r}="North"), "Free delivery", "Standard charges")' if solution else None)
        r += 1

    # --- Section 3: IFERROR -------------------------------------------------
    put_section(ws, 22, "3 · Sales per order — IFERROR")
    put_headers(ws, 23, ["Date", "Sales (₹)", "Orders", "Sales per order",
                         "IFERROR → 0", 'IFERROR → "No orders"'])
    daily = [
        (date(2026, 2, 2), 12500, 5),
        (date(2026, 2, 3), 9800, 0),
        (date(2026, 2, 4), 14300, 4),
        (date(2026, 2, 5), 11000, 2),
    ]
    r = 24
    for d, sales, orders in daily:
        ws.cell(row=r, column=1, value=d).number_format = "dd-mmm"
        ws.cell(row=r, column=2, value=sales).border = _BORDER
        ws.cell(row=r, column=3, value=orders).border = _BORDER
        _answer(ws, r, 4, f"=B{r}/C{r}" if solution else None, "#,##0.00")
        _answer(ws, r, 5, f"=IFERROR(B{r}/C{r}, 0)" if solution else None, "#,##0.00")
        _answer(ws, r, 6, f'=IFERROR(B{r}/C{r}, "No orders")' if solution else None)
        r += 1

    put_note(ws, 29, "Column D on the 03-Feb row shows #DIV/0! — the shop recorded no orders that day. "
                     "Column E hides it behind 0; column F shows the friendly text instead. "
                     "Exercise 4 sends you to the Solution sheet: one formula there is deliberately broken but wrapped "
                     "in IFERROR, so it shows a clean 0. Find the Debug corner below, delete the wrapper and press Enter.")
    if solution:
        put_section(ws, 31, "Debug corner — one formula is hiding a bug")
        c = ws.cell(row=32, column=1, value="This cell always shows 0. Select it, delete the IFERROR wrapper, "
                                            "press Enter and see the real error:")
        c.font = Font(size=10, color=GRAY)
        _answer(ws, 33, 2, "=IFERROR(SUMME(D24:D27), 0)")
        put_note(ws, 34, "Unwrapped, SUMME is a misspelled function name, so the formula returns #NAME?. "
                         "IFERROR was hiding a bug in the formula itself — fix the name (SUM), don't hide the error.")

    set_widths(ws, [16, 12, 18, 16, 14, 20, 18])


# ---------------------------------------------------------------------------
# Workbook 4: lookup-practice.xlsx
# Lessons: l2-vlookup, l2-xlookup, l2-index-match
# ---------------------------------------------------------------------------

def build_lookup_practice(ws, solution):
    put_title(ws, "Lookup Practice — Exercise" if not solution else "Lookup Practice — Solution")
    if solution:
        put_note(ws, 2, "Finished state. Invoice prices pulled three ways (all equal: 45 · 12 · 130 · 8500), "
                        "amounts 1800 · 1200 · 1950 · 17000, missing product handled, left-lookups working: "
                        "Priya Nair → Accounts · EMP003 → Rahul Verma · Sneha Iyer → EMP004 · Printers at position 4 → 8500 · Q2 → 150000.")
    else:
        put_note(ws, 2, "Covers the lessons 'VLOOKUP', 'HLOOKUP and XLOOKUP' and 'INDEX and MATCH'. "
                        "Lock the lookup table with $ (press F4) before filling down. "
                        "Fill the amber answer cells, then compare with the Solution sheet.")

    # --- Price list ---------------------------------------------------------
    put_section(ws, 4, "1 · Price list")
    put_headers(ws, 5, ["Product", "Price (₹)"])
    prices = [("Notebooks", 45), ("Pens", 12), ("Staplers", 130), ("Printers", 8500), ("Laptops", 42000)]
    put_rows(ws, 6, [[p, v] for p, v in prices], num_formats={2: "#,##0"})

    # --- Employee table -----------------------------------------------------
    put_section(ws, 12, "2 · Employee table")
    put_headers(ws, 13, ["ID", "Name", "Department"])
    staff = [
        ("EMP001", "Arjun Mehta", "Sales"),
        ("EMP002", "Priya Nair", "Accounts"),
        ("EMP003", "Rahul Verma", "Operations"),
        ("EMP004", "Sneha Iyer", "Marketing"),
    ]
    put_rows(ws, 14, [list(s) for s in staff])

    # --- Invoice -------------------------------------------------------------
    put_section(ws, 19, "3 · Invoice — pull the price for every row (three ways)")
    put_headers(ws, 20, ["Item", "Qty", "Price (VLOOKUP)", "Price (XLOOKUP)",
                         "Price (INDEX/MATCH)", "Amount"])
    invoice = [("Notebooks", 40), ("Pens", 100), ("Staplers", 15), ("Printers", 2)]
    r = 21
    for item, qty in invoice:
        ws.cell(row=r, column=1, value=item).border = _BORDER
        ws.cell(row=r, column=2, value=qty).border = _BORDER
        _answer(ws, r, 3, f"=VLOOKUP(A{r}, $A$6:$B$10, 2, FALSE)" if solution else None, "#,##0")
        _answer(ws, r, 4, f'={XLOOKUP}(A{r}, $A$6:$A$10, $B$6:$B$10, "Not in price list")' if solution else None, "#,##0")
        _answer(ws, r, 5, f'=IFERROR(INDEX($B$6:$B$10, MATCH(A{r}, $A$6:$A$10, 0)), "Not in price list")' if solution else None, "#,##0")
        _answer(ws, r, 6, f"=B{r}*C{r}" if solution else None, "#,##0")
        r += 1

    # --- Try it ---------------------------------------------------------------
    put_section(ws, 26, "4 · Try it — missing values · look left · scratch steps")
    put_headers(ws, 27, ["Question", "Input", "Answer", "Lesson ref"])
    qa = [
        ("Type a product that is NOT in the price list — plain VLOOKUP", "Whiteboard",
         "=VLOOKUP($B$28, $A$6:$B$10, 2, FALSE)", "VLOOKUP ex. 3", None),
        ("Wrap it in IFERROR", None,
         '=IFERROR(VLOOKUP($B$28, $A$6:$B$10, 2, FALSE), "Not in price list")', "VLOOKUP ex. 3", None),
        ("Same lookup with XLOOKUP's built-in fallback", None,
         f'={XLOOKUP}($B$28, $A$6:$A$10, $B$6:$B$10, "Not in price list")', "XLOOKUP ex. 3", None),
        ("Type an employee name — department via XLOOKUP", "Priya Nair",
         f'={XLOOKUP}($B$31, $B$14:$B$17, $C$14:$C$17, "Not an employee")', "XLOOKUP ex. 1", None),
        ("Reverse: type an ID — name via XLOOKUP (looks left)", "EMP003",
         f'={XLOOKUP}($B$32, $A$14:$A$17, $B$14:$B$17, "Unknown ID")', "XLOOKUP ex. 2", None),
        ("Department for that ID via VLOOKUP", None,
         "=VLOOKUP($B$32, $A$14:$C$17, 3, FALSE)", "VLOOKUP ex. 4", None),
        ("ID for 'Sneha Iyer' via INDEX/MATCH (looks left)", None,
         '=INDEX($A$14:$A$17, MATCH("Sneha Iyer", $B$14:$B$17, 0))', "INDEX/MATCH ex. 4", None),
        ("Scratch step 1: position of 'Printers' (MATCH)", None,
         '=MATCH("Printers", $A$6:$A$10, 0)', "INDEX/MATCH ex. 1", "#,##0"),
        ("Scratch step 2: price at that position (INDEX)", None,
         "=INDEX($B$6:$B$10, 4)", "INDEX/MATCH ex. 2", "#,##0"),
    ]
    r = 28
    for question, inp, formula, ref, fmt in qa:
        _qa(ws, r, solution, question, formula=formula, inp=inp, ref=ref, fmt=fmt)
        r += 1

    # --- Quarterly targets (horizontal) --------------------------------------
    put_section(ws, 38, "5 · Quarterly targets — a horizontal lookup")
    put_headers(ws, 39, ["Quarter", "Q1", "Q2", "Q3", "Q4"])
    put_rows(ws, 40, [["Target (₹)", 120000, 150000, 140000, 180000]],
             num_formats={2: "#,##0", 3: "#,##0", 4: "#,##0", 5: "#,##0"})
    put_headers(ws, 42, ["Question", "Input", "Answer"])
    _qa(ws, 43, solution, "Type a quarter — target via XLOOKUP across a row",
        formula=f'={XLOOKUP}($B$43, $B$39:$E$39, $B$40:$E$40, "Unknown quarter")',
        inp="Q2", fmt="#,##0")
    put_note(ws, 45, "HLOOKUP could do this too (=HLOOKUP(B43, B39:E40, 2, FALSE)). "
                     "XLOOKUP searches rows and columns alike — and needs Microsoft 365 or Excel 2021+.")

    set_widths(ws, [52, 15, 15, 15, 17, 12, 14])


# ---------------------------------------------------------------------------
# Workbook 5: text-date-practice.xlsx
# Lesson: l2-text-functions
# ---------------------------------------------------------------------------

def build_text_date_practice(ws, solution):
    put_title(ws, "Text & Date Practice — Exercise" if not solution else "Text & Date Practice — Solution")
    if solution:
        put_note(ws, 2, "Finished state. Row 1 of the report label reads PRIYA NAIR - STP — exactly the result "
                        "quoted in the lesson. Labels: family STP/STP/NTB… · items 013/014/201… · "
                        "colours BLACK/WHITE/GREEN… · cleaned names properly capitalised.")
    else:
        put_note(ws, 2, "Covers the lesson 'Text Functions' (LEFT, MID, RIGHT, TRIM, PROPER, & and TEXT). "
                        "Fill the amber answer cells, then compare with the Solution sheet. "
                        "First label expected: PRIYA NAIR - STP.")

    # --- Section 1: codes & names -------------------------------------------
    put_section(ws, 4, "1 · Product codes & customer names — LEFT · MID · RIGHT · TRIM · PROPER · &")
    put_headers(ws, 5, ["Order", "Product code", "Family (LEFT)", "Item (MID)", "Colour (RIGHT)",
                        "Customer (raw)", "Cleaned name", "Report label"])
    records = [
        ("ORD-001", "STP-013-BLACK", "  priya nair "),
        ("ORD-002", "STP-014-WHITE", "RAHUL  verma"),
        ("ORD-003", "NTB-201-GREEN", "sneha iyer"),
        ("ORD-004", "STP-015-BLACK", "  ARJUN mehta "),
        ("ORD-005", "NTB-202-WHITE", "MEERA  pillai"),
    ]
    r = 6
    for oid, code, raw in records:
        ws.cell(row=r, column=1, value=oid).border = _BORDER
        ws.cell(row=r, column=2, value=code).border = _BORDER
        ws.cell(row=r, column=6, value=raw).border = _BORDER
        _answer(ws, r, 3, f"=LEFT(B{r}, 3)" if solution else None)
        _answer(ws, r, 4, f"=MID(B{r}, 5, 3)" if solution else None)
        _answer(ws, r, 5, f"=RIGHT(B{r}, 5)" if solution else None)
        _answer(ws, r, 7, f"=PROPER(TRIM(F{r}))" if solution else None)
        _answer(ws, r, 8, f'=UPPER(G{r})&" - "&UPPER(C{r})' if solution else None)
        r += 1

    # --- Section 2: TEXT -------------------------------------------------------
    put_section(ws, 12, "2 · Dates & TEXT — labels you can join")
    put_headers(ws, 13, ["Date", "Sales (₹)", "Date label (TEXT)", "Money label (TEXT)"])
    days = [(date(2026, 11, 14), 1250.5), (date(2026, 11, 15), 980), (date(2026, 11, 16), 3045.75)]
    r = 14
    for d, sales in days:
        ws.cell(row=r, column=1, value=d).number_format = "dd-mmm-yyyy"
        ws.cell(row=r, column=2, value=sales).border = _BORDER
        _answer(ws, r, 3, f'=TEXT(A{r}, "dd-mmm-yyyy")' if solution else None)
        _answer(ws, r, 4, f'=TEXT(B{r}, "#,##0.00")' if solution else None)
        r += 1
    put_note(ws, 18, "=TEXT returns text: perfect for labels and titles, useless for further maths — "
                     "keep the real date and number in their own cells.")

    set_widths(ws, [10, 16, 14, 12, 14, 18, 16, 20])


# ---------------------------------------------------------------------------
# Build, save and verify
# ---------------------------------------------------------------------------

BUILDERS = {
    "formula-practice.xlsx": build_formula_practice,
    "logical-functions.xlsx": build_logical_functions,
    "lookup-practice.xlsx": build_lookup_practice,
    "sumif-practice.xlsx": build_sumif_practice,
    "text-date-practice.xlsx": build_text_date_practice,
}


def main():
    os.makedirs(OUT, exist_ok=True)
    created = []
    for name, builder in BUILDERS.items():
        wb = new_workbook("Exercise", "Solution")
        builder(wb["Exercise"], solution=False)
        builder(wb["Solution"], solution=True)
        for ws in wb.worksheets:
            ws.freeze_panes = "A3"  # keep title + instructions visible
        path = os.path.join(OUT, name)
        wb.save(path)
        created.append(path)

    # Verification: every file exists, opens, and has both sheets.
    print("\nVerification:")
    ok = True
    for path in created:
        size = os.path.getsize(path)
        wb = load_workbook(path)
        sheets = wb.sheetnames
        good = sheets == ["Exercise", "Solution"] and size > 0
        ok = ok and good
        print(f"  {os.path.relpath(path, ROOT)}  {size:>6} bytes  sheets={sheets}  {'OK' if good else 'FAIL'}")
    if not ok:
        sys.exit(1)
    print("All Level 2 workbooks generated and verified.")


if __name__ == "__main__":
    main()
