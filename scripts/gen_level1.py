"""Generate the Level 1 practice workbooks referenced by the lessons.

Run from the repo root:

    & .venv\\Scripts\\python.exe scripts\\gen_level1.py

Creates exactly the two files referenced by the Level 1 lesson downloads:

  assets/files/level1/beginner-practice.xlsx   (5 lessons)
  assets/files/level1/formula-basics.xlsx      (3 lessons)

Each workbook has an "Exercise" sheet (the starter state the student works
in) and a "Solution" sheet (the finished state). All data grids, expected
results and formulas mirror the lesson content in assets/js/data/level1-1.js
and level1-2.js:

beginner-practice.xlsx
  1. Data entry             -> l1-data-entry          (level1-1)
  2. Copy/paste/transpose   -> l1-copy-cut-paste      (level1-1)
  3. AutoFill & Flash Fill  -> l1-autofill-flashfill  (level1-1)
  4. Formatting practice    -> l1-cell-formatting     (level1-2)
  5. Number formatting      -> l1-number-formatting   (level1-2)

formula-basics.xlsx
  - Weekly budget project   -> l1-mini-project        (level1-2, A1:B17 layout)
  - Tax / conversion rates  -> l1-cell-references     (level1-2, rate cells G1/G3)
  - Price x Qty + PEMDAS    -> l1-basic-calculations  (level1-2)
"""
from datetime import date
from pathlib import Path

import xlcommon as xl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

GREEN, RED, YELLOW = "c6efce", "ffc7ce", "ffeb9c"  # Excel Good/Bad/Neutral fills
GOOD_FILL = PatternFill("solid", fgColor=GREEN)
BAD_FILL = PatternFill("solid", fgColor=RED)
NEUTRAL_FILL = PatternFill("solid", fgColor=YELLOW)
BAND_FILL = PatternFill("solid", fgColor=xl.BAND)

RUPEE2 = '"\u20b9" #,##0.00'   # Accounting-style rupee, 2 decimals
RUPEE0 = '"\u20b9" #,##0'      # rupee, 0 decimals
PCT0 = "0%"
PCT2 = "0.00%"
DATE_LONG = "dd-mmm-yyyy"      # 14-Mar-2026
DATE_SHORT = "dd-mmm-yy"       # 14-Mar-26
NUM2 = "#,##0.00"
NUM0 = "#,##0"

_THIN = Side(style="thin", color="cbd5e1")
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


# ---------------------------------------------------------------- helpers
def excel_serial(d):
    """Excel serial number for a date (1900 date system)."""
    return (d - date(1899, 12, 30)).days


def fmt(ws, col, r1, r2, code):
    for r in range(r1, r2 + 1):
        ws.cell(row=r, column=col).number_format = code


def bold(ws, row, col):
    ws.cell(row=row, column=col).font = Font(bold=True)


def headers_at(ws, row, col, labels):
    """Styled header row starting at (row, col) — column-flexible put_headers."""
    for i, h in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=h)
        c.font = Font(bold=True, color=xl.WHITE)
        c.fill = PatternFill("solid", fgColor=xl.TEAL)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = 22
    return row + 1


def table(ws, row, col, labels, rows):
    """Header row + bordered data rows starting at (row, col). Returns next row."""
    r = headers_at(ws, row, col, labels)
    for data in rows:
        for i, v in enumerate(data):
            c = ws.cell(row=r, column=col + i, value=v)
            c.border = BORDER
            if isinstance(v, str) and v.startswith("="):
                c.alignment = Alignment(horizontal="left")
            elif isinstance(v, (int, float, date)):
                c.alignment = Alignment(horizontal="right")
        r += 1
    return r


def label(ws, row, col, text):
    c = ws.cell(row=row, column=col, value=text)
    c.font = Font(italic=True, size=10, color=xl.GRAY)


def note(ws, row, text, min_len=170, height=46):
    """Instruction note merged across A:H (taller when long)."""
    xl.put_note(ws, row, text, width=8)
    if len(text) > min_len:
        ws.row_dimensions[row].height = height
    return row + 1


# ---------------------------------------------------- beginner-practice.xlsx
def build_beginner(ws, sol):
    """sol=False -> Exercise (starter state); sol=True -> Solution (finished)."""
    r = xl.put_title(
        ws,
        "Excel Lab \u2014 Beginner Practice (Level 1)",
        subtitle="Practice for: Entering and Editing Data \u00b7 Copy, Cut, Paste \u00b7 "
        "AutoFill and Flash Fill \u00b7 Cell Formatting \u00b7 Number Formatting. "
        + ("This Solution sheet shows the finished state \u2014 compare honestly "
           "(Ctrl + `), not just the numbers."
           if sol else
           "Work top to bottom; each section is labelled with the lesson it "
           "belongs to. The Solution sheet shows the finished state."),
    )
    ws.cell(row=r, column=1).font = Font(italic=True, size=10, color=xl.GRAY)
    r += 1

    # ---- 1. Data entry ----------------------------------------------------
    r = xl.put_section(ws, r, "1 \u00b7 Data entry \u2014 lesson: Entering and Editing Data")
    r = note(ws, r,
             "Complete the table: fill the missing employee names, prices and order "
             "dates. Every number and date you enter must sit on the RIGHT edge \u2014 "
             "left-aligned means text slipped in. One price was typed with \u201cRs.\u201d in "
             "front: fix that single cell with F2 or double-click (do not retype it) "
             "so it becomes a pure number.")
    first = r + 1
    rows = [
        ["Anita Menon", 1250, date(2026, 3, 2)],
        ["Rahul Verma" if sol else None, 45.5, date(2026, 3, 9)],
        ["Arjun Mehta", 2400 if sol else "Rs. 2400", date(2026, 3, 14)],
        ["Priya Nair", 450 if sol else None, date(2026, 3, 21)],
        ["Deepa Rao", 720, date(2026, 3, 28) if sol else None],
        ["Sunita Iyer" if sol else None, 12500, date(2026, 4, 4)],
    ]
    r = table(ws, r, 1, ["Employee", "Price", "Ordered on"], rows)
    if sol:
        fmt(ws, 2, first, first + 5, RUPEE2)
    fmt(ws, 3, first, first + 5, DATE_LONG)
    r += 1

    label(ws, r, 1, "Week's expenses \u2014 select the seven empty Amount cells first, then "
                    "fill them column by column with Enter only. Finish by making a "
                    "typing mistake and undoing it with Ctrl + Z (Ctrl + Y redoes).")
    r += 1
    week = [("Monday", 640), ("Tuesday", 480), ("Wednesday", 720), ("Thursday", 560),
            ("Friday", 850), ("Saturday", 1100), ("Sunday", 380)]
    first = r + 1
    r = table(ws, r, 1, ["Day", "Amount"],
              [[d, a if sol else None] for d, a in week])
    if sol:
        fmt(ws, 2, first, first + 6, NUM0)
    r += 1

    # ---- 2. Copy / paste / transpose --------------------------------------
    r = xl.put_section(ws, r, "2 \u00b7 Copy, cut, paste and transpose \u2014 lesson: Copy, Cut, Paste and Paste Special")
    r = note(ws, r,
             "a) Copy the header row of the table below and paste it into the paste "
             "area on the right, keeping the source formatting, then press Esc to "
             "clear the marching ants. b) Copy the whole calculated block and paste "
             "it as VALUES (Ctrl + Alt + V) under the pasted header \u2014 click a pasted "
             "cell: the formula bar must show a plain number. c) Transpose the "
             "months/sales row into a column in the transposed area. d) Move the "
             "small expense block with drag-and-drop, then move it back with "
             "Cut (Ctrl + X) and Paste.")
    first = r + 1
    rows = [
        [45, 10, f"=A{first}*B{first}"],
        [60, 5, f"=A{first + 1}*B{first + 1}"],
        [12500, 1, f"=A{first + 2}*B{first + 2}"],
        ["Grand total", None, f"=SUM(C{first}:C{first + 2})"],
    ]
    r = table(ws, r, 1, ["Price", "Qty", "Total"], rows)
    bold(ws, first + 3, 1)
    if sol:
        fmt(ws, 3, first, first + 3, NUM0)
        bold(ws, first + 3, 3)
        r2 = r
        table(ws, r2, 6, ["Price", "Qty", "Total"], [
            [45, 10, 450], [60, 5, 300], [12500, 1, 12500],
            ["Grand total", None, 13250],
        ])
        bold(ws, r2 + 4, 6)
        fmt(ws, 8, r2 + 1, r2 + 4, NUM0)
    else:
        label(ws, r, 6, "Paste area \u2014 paste the copied header row here, "
                        "then the Values paste below it.")
    r += 1

    r = table(ws, r, 1, ["Month", "Jan", "Feb", "Mar"], [["Sales", 4200, 5800, 6100]])
    label(ws, r, 1, "Transposed area \u2014 months down, figures beside them:")
    r += 1
    if sol:
        r = table(ws, r, 1, ["Month", "Sales"],
                  [["Jan", 4200], ["Feb", 5800], ["Mar", 6100]])
    r += 1

    r = table(ws, r, 1, ["Item", "Cost"],
              [["Tea", 350], ["Snacks", 480], ["Travel", 900]])
    if sol:
        fmt(ws, 2, r - 3, r - 1, NUM0)
    r += 1

    # ---- 3. AutoFill / Flash Fill -----------------------------------------
    r = xl.put_section(ws, r, "3 \u00b7 AutoFill and Flash Fill \u2014 lesson: AutoFill and Flash Fill")
    r = note(ws, r,
             "a) Type Monday in the first Weekday cell and fill the column through "
             "Sunday with the fill handle. b) 10 and 20 are typed in the Sequence "
             "column \u2014 select both cells, then drag the fill handle down to 100. "
             "c) Flash Fill the names (Ctrl + E): First, Last, then Email in the "
             "pattern first initial + last name (type amehta as your example). "
             "d) Type one reformatted phone (98470-12345) beside the raw number and "
             "Flash Fill the rest. Flash Fill results are static values \u2014 rerun "
             "Ctrl + E after editing the source.")
    wd = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    r = headers_at(ws, r, 1, ["Weekday"])
    headers_at(ws, r - 1, 3, ["Sequence (10 to 100)"])
    first = r
    for i in range(10):
        if i == 0:
            a = "Monday"
        elif i < 7:
            a = wd[i] if sol else None
        else:
            a = None
        s = 10 * (i + 1) if (i < 2 or sol) else None
        ca = ws.cell(row=r, column=1, value=a)
        cs = ws.cell(row=r, column=3, value=s)
        ca.border = BORDER
        cs.border = BORDER
        if isinstance(s, int):
            cs.alignment = Alignment(horizontal="right")
        r += 1
    r += 1

    people = [("Arjun Mehta", "Arjun", "Mehta", "amehta"),
              ("Priya Nair", "Priya", "Nair", "pnair"),
              ("Deepa Rao", "Deepa", "Rao", "drao"),
              ("Rahul Verma", "Rahul", "Verma", "rverma"),
              ("Sunita Iyer", "Sunita", "Iyer", "siyer")]
    first = r + 1
    r = table(ws, r, 1, ["Full name", "First", "Last", "Email"],
              [[p[0]] + ([p[1], p[2], p[3]] if sol else [None, None, None])
               for p in people])
    r += 1

    phones = ["9847012345", "9845112233", "9447098765", "9946332211"]
    first = r + 1
    r = table(ws, r, 1, ["Raw number", "Formatted (98470-12345)"],
              [[p, f"{p[:5]}-{p[5:]}" if sol else None] for p in phones])
    r += 1

    # ---- 4. Cell formatting ------------------------------------------------
    r = xl.put_section(ws, r, "4 \u00b7 Formatting practice \u2014 lesson: Cell Formatting")
    r = note(ws, r,
             "Make the header row bold with a light fill, and give the whole data "
             "block All Borders so it survives printing (screen gridlines do not "
             "print). The Note column holds long sentences that are cut off \u2014 apply "
             "Wrap Text and auto-fit the row heights. Add the title \u201cSunrise Traders "
             "- Weekly Expense Report\u201d above the table using Center Across Selection "
             "(Ctrl + 1 \u2192 Alignment \u2014 not Merge & Center). Finally mark the three "
             "status rows below with the Good, Bad and Neutral cell styles, and copy "
             "the finished header look to another area with the Format Painter "
             "(double-click it to paint several ranges).")
    if sol:
        c = ws.cell(row=r, column=1, value="Sunrise Traders - Weekly Expense Report")
        c.font = Font(bold=True, size=12, color=xl.TEAL_DARK)
        c.alignment = Alignment(horizontal="centerContinuous")
        r += 1
    first = r + 1
    rows = [
        ["Notebooks", 1800, "Restock for March: two cartons of 60, delivered to the store room."],
        ["Printer ink", 2400, "Black and colour cartridges, billed quarterly by the supplier."],
        ["Printer paper", 900, "Five reams of A4; the box arrived damaged and two reams were replaced."],
        ["Total", f"=SUM(B{first}:B{first + 2})", None],
    ]
    r = table(ws, r, 1, ["Item", "Cost", "Note"], rows)
    bold(ws, first + 3, 1)
    bold(ws, first + 3, 2)
    for i in range(3):
        if sol:
            ws.cell(row=first + i, column=3).alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[first + i].height = 28
    if sol:
        fmt(ws, 2, first, first + 3, NUM0)
        bold(ws, first + 3, 3)
    r += 1

    styles = [("Invoice 118 checked against the bill", "Good", GOOD_FILL),
              ("Printer bill is missing", "Bad", BAD_FILL),
              ("Courier charge still to confirm", "Neutral", NEUTRAL_FILL)]
    first = r + 1
    r = table(ws, r, 1, ["Sample row", "Cell style"],
              [[s[0], s[1] if sol else None] for s in styles])
    if sol:
        for i, (_, _, fill) in enumerate(styles):
            ws.cell(row=first + i, column=1).fill = fill
            ws.cell(row=first + i, column=2).fill = fill
    r += 1

    # ---- 5. Number formatting ----------------------------------------------
    r = xl.put_section(ws, r, "5 \u00b7 Number formatting \u2014 lesson: Number, Date, Currency and Percent Formatting")
    r = note(ws, r,
             "Apply Accounting format (\u20b9, two decimals) to the Price column; Percent "
             "to the Rate column \u2014 a cell showing 18% stores 0.18 (check the formula "
             "bar); Short Date to the Ordered on column, whose dates are currently "
             "showing as plain serial numbers like 46100; #,##0.00 and then Comma "
             "style to the Total column \u2014 the totals themselves never change; and "
             "Number format with 0 decimals plus the 1000 separator to the Count "
             "column after entering 12500 and 45.")
    items = [
        ("Notebook", 250, 45.5, 0.18, date(2026, 3, 14)),
        ("Printer ink", 1, 12500, 0.075, date(2026, 3, 21)),
        ("Pen pack", 12, 60, 0.15, date(2026, 3, 28)),
        ("File box", 12500, 2.5, 0.12, date(2026, 4, 4)),
        ("Stapler", 45, 220, 0.18, date(2026, 4, 11)),
    ]
    first = r + 1
    rows = []
    for i, (name, count, price, rate, dt) in enumerate(items):
        rows.append([
            name,
            count if (sol or i < 3) else None,
            price,
            f"=B{first + i}*C{first + i}",
            rate,
            dt if sol else excel_serial(dt),
        ])
    r = table(ws, r, 1, ["Item", "Count", "Price", "Total", "Rate", "Ordered on (serial)"],
              rows)
    if sol:
        fmt(ws, 2, first, first + 4, NUM0)
        fmt(ws, 3, first, first + 4, RUPEE2)
        fmt(ws, 4, first, first + 4, NUM2)
        fmt(ws, 5, first, first + 4, PCT2)
        fmt(ws, 6, first, first + 4, DATE_SHORT)
    r += 1

    label(ws, r, 1, "Display check \u2014 after formatting, the values below must display "
                    "exactly as shown in the second column (formatting is cosmetic: "
                    "the stored values never change).")
    r += 1
    first = r + 1
    rows = [[12500, "12,500.00"], [45.5, "45.50"], [0.5, "0.50"]]
    r = table(ws, r, 1, ["Raw value", "Should display as"],
              rows if not sol else [[12500, 12500], [45.5, 45.5], [0.5, 0.5]])
    if sol:
        fmt(ws, 2, first, first + 2, NUM2)
    r += 1

    ws.freeze_panes = None
    xl.set_widths(ws, [24, 14, 30, 14, 12, 18, 14, 14])

# ------------------------------------------------------ formula-basics.xlsx
def build_formula(ws, sol):
    """Budget at its canonical A1:B17 layout so the lesson formulas
    (=SUM(B5:B6), =B16*$B$2, ...) work exactly as printed in the lesson."""
    c = ws.cell(row=1, column=1, value="Sunrise Traders - Weekly Budget (Week 46)")
    c.font = Font(bold=True, size=12, color=xl.TEAL_DARK)
    ws.cell(row=2, column=1, value="Savings rate")
    b2 = ws.cell(row=2, column=2, value=0.2)
    b2.number_format = PCT0

    ws.cell(row=4, column=1, value="Income")
    ws.cell(row=5, column=1, value="Cash sales")
    ws.cell(row=5, column=2, value=18400)
    ws.cell(row=6, column=1, value="Card payments received")
    ws.cell(row=6, column=2, value=9250)
    ws.cell(row=7, column=1, value="Total income")
    ws.cell(row=7, column=2, value="=SUM(B5:B6)" if sol else None)
    ws.cell(row=9, column=1, value="Expenses")
    for rr, (lab, val) in zip(range(10, 14),
                              [("Notebooks restock", 3600), ("Printer paper", 1200),
                               ("Electricity", 2400), ("Courier", 450)]):
        ws.cell(row=rr, column=1, value=lab)
        ws.cell(row=rr, column=2, value=val)
    ws.cell(row=14, column=1, value="Total expenses")
    ws.cell(row=14, column=2, value="=SUM(B10:B13)" if sol else None)
    ws.cell(row=16, column=1, value="Balance (income - expenses)")
    ws.cell(row=16, column=2, value="=B7-B14" if sol else None)
    ws.cell(row=17, column=1, value="To savings (balance * rate)")
    ws.cell(row=17, column=2, value="=B16*$B$2" if sol else None)

    if sol:
        # Step 2 of the lesson: professional finish.
        for rr in (4, 9):
            ws.cell(row=rr, column=1).font = Font(bold=True, size=11)
            ws.cell(row=rr, column=1).fill = BAND_FILL
        for rr in (7, 14):
            for cc in (1, 2):
                ws.cell(row=rr, column=cc).font = Font(bold=True)
        for row in ws.iter_rows(min_row=1, max_row=17, min_col=1, max_col=2):
            for cell in row:
                cell.border = BORDER
        for rr in (5, 6, 7, 10, 11, 12, 13, 14, 16, 17):
            ws.cell(row=rr, column=2).number_format = RUPEE0
            ws.cell(row=rr, column=2).alignment = Alignment(horizontal="right")

    # Right-hand zone: reference-rate practice (l1-cell-references).
    label(ws, 1, 6, "Tax rate")
    g1 = ws.cell(row=1, column=7, value=0.18)
    g1.number_format = PCT0
    label(ws, 3, 6, "USD \u2192 INR rate")
    g3 = ws.cell(row=3, column=7, value=83.5)
    g3.number_format = "0.0"

    first = 4
    table(ws, 3, 3, ["Amount", "Tax"],
          [[amt, f"=C{first + i}*$G$1" if sol else None]
           for i, amt in enumerate([9500, 7200, 13400])])
    if sol:
        fmt(ws, 4, first, first + 2, NUM0)

    first = 9
    table(ws, 8, 3, ["Invoice", "USD", "INR"],
          [[inv, usd, f"=D{first + i}*$G$3" if sol else None]
           for i, (inv, usd) in enumerate([("INV-101", 250), ("INV-102", 480),
                                           ("INV-103", 320)])])
    if sol:
        fmt(ws, 5, first, first + 2, NUM0)

    note_at(ws, 13, 3, 8,
            "Tax and conversion practice \u2014 lesson: Relative and Absolute "
            "References. In the Tax column, multiply each amount by the rate cell "
            "locked with $ (click the rate cell, press F4 once): the first cell "
            "reads =C4*$G$1. Build the INR column the same way. Then change a rate "
            "cell (e.g. 0.18 to 0.21) and watch every result update \u2014 no formula "
            "edits. Practice F4 in any empty cell: it cycles G1 \u2192 $G$1 \u2192 G$1 "
            "\u2192 $G1 \u2192 G1.")

    # Price x Qty totals (l1-basic-calculations).
    xl.put_section(ws, 19, "Price \u00d7 Qty totals \u2014 lesson: Basic Calculations and Your First Formulas")
    note(ws, 20,
         "Write the missing Total formulas yourself \u2014 each row is price times qty "
         "(click the cells, do not type the addresses). One row already contains a "
         "formula with a TYPO: press Ctrl + ` to expose every formula, spot the odd "
         "one, and fix it. Then add the Grand total below with =SUM.")
    first = 22
    table(ws, 21, 2, ["Price", "Qty", "Total"],
          [[45, 10, None],
           [60, 5, "=B23*C23" if sol else "=B23*C24"],   # typo row on Exercise
           [12500, 1, None]])
    bold(ws, 25, 1)
    ws.cell(row=25, column=1, value="Grand total")
    ws.cell(row=25, column=4, value="=SUM(D22:D24)" if sol else None)
    if sol:
        fmt(ws, 4, first, first + 3, NUM0)
        bold(ws, 25, 4)

    # Order of operations (l1-basic-calculations, exercise 3).
    xl.put_section(ws, 27, "Order of operations practice \u2014 lesson: Basic Calculations (PEMDAS)")
    note(ws, 28,
         "Predict each result first, then type it as a formula (start with =) in "
         "the Answer column. PEMDAS: parentheses, exponents, multiply/divide "
         "left to right, add/subtract left to right. Expected answers are in the "
         "lesson: 3, 16 and 30.")
    table(ws, 29, 1, ["Formula", "Your answer"],
          [["10-4-3", "=10-4-3" if sol else None],
           ["2^3*2", "=2^3*2" if sol else None],
           ["(10+5)*2", "=(10+5)*2" if sol else None]])

    ws.freeze_panes = None
    xl.set_widths(ws, [26, 16, 12, 12, 12, 16, 12, 12])


def note_at(ws, row, col1, col2, text, height=46):
    ws.merge_cells(start_row=row, start_column=col1, end_row=row, end_column=col2)
    c = ws.cell(row=row, column=col1, value=text)
    c.font = Font(italic=True, size=10, color=xl.GRAY)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[row].height = height


def main():
    out = Path("assets/files/level1")
    out.mkdir(parents=True, exist_ok=True)

    wb = xl.new_workbook("Exercise", "Solution")
    build_beginner(wb["Exercise"], sol=False)
    build_beginner(wb["Solution"], sol=True)
    p1 = out / "beginner-practice.xlsx"
    wb.save(p1)

    wb = xl.new_workbook("Exercise", "Solution")
    build_formula(wb["Exercise"], sol=False)
    build_formula(wb["Solution"], sol=True)
    p2 = out / "formula-basics.xlsx"
    wb.save(p2)

    for p in (p1, p2):
        print(f"created {p} ({p.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
