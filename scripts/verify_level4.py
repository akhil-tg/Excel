"""Independent verification of the Level 4 workbooks (run after gen_level4.py)."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import load_workbook

EXPECTED = {
    "assets/files/level4/sales-data.xlsx",
    "assets/files/level4/kpi-report-data.xlsx",
    "assets/files/level4/dashboard-data.xlsx",
}

ok = True

# 1) every referenced file exists
for rel in sorted(EXPECTED):
    if not os.path.exists(rel):
        print("FAIL missing:", rel)
        ok = False
print("files present:", all(os.path.exists(p) for p in EXPECTED))

# 2) each opens cleanly, has Exercise + Solution sheets
for rel in sorted(EXPECTED):
    wb = load_workbook(rel)
    sheets = wb.sheetnames
    assert "Exercise" in sheets and "Solution" in sheets, (rel, sheets)
    for name in sheets:
        assert wb[name].max_row >= 1
    print("%-32s sheets=%s" % (os.path.basename(rel), ",".join(sheets)))

# 3) sales-data spot checks
wb = load_workbook("assets/files/level4/sales-data.xlsx")
assert "SalesTable" in wb["Exercise"].tables, wb["Exercise"].tables.keys()
assert wb["Exercise"].max_row >= 255, wb["Exercise"].max_row  # header row 5 + 250 rows
sol = wb["Solution"]
labels = {}
for r in range(1, sol.max_row + 1):
    v = sol.cell(row=r, column=1).value
    if v and v not in labels:
        labels[v] = r
region_first = labels["North"]
got = {sol.cell(row=region_first + i, column=1).value: sol.cell(row=region_first + i, column=2).value
       for i in range(4)}
expect = {"North": 645000, "South": 590000, "East": 605000, "West": 640000}
assert got == expect, got
assert sol.cell(row=region_first + 4, column=2).value == "=SUM(B%d:B%d)" % (region_first, region_first + 3)
# product table: Laptops first (sorted desc) and equals 820000
prod_row = labels["Laptops"]
assert sol.cell(row=prod_row, column=2).value == 820000, sol.cell(row=prod_row, column=2).value
# cross-tab grand total formula exists
print("sales-data spot checks OK (regions, laptops 820000, 250-row table)")

# 4) kpi-report-data spot checks
wb = load_workbook("assets/files/level4/kpi-report-data.xlsx")
ex, sol = wb["Exercise"], wb["Solution"]
assert [ex.cell(row=r, column=1).value for r in range(2, 14)] == [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
assert ex.cell(row=2, column=2).value == 500000 and ex.cell(row=2, column=3).value == 542000
assert ex.cell(row=7, column=3).value == 438000  # Jun
assert ex.cell(row=2, column=4).value is None    # answer area empty
assert sol.cell(row=2, column=4).value == "=C2/B2"
assert sol.cell(row=3, column=5).value == '=IFERROR((C3-C2)/C2,"")'
assert len(sol.conditional_formatting) >= 1
print("kpi-report-data spot checks OK (rows 2-13, formulas, conditional formatting)")

# 5) dashboard-data spot checks
wb = load_workbook("assets/files/level4/dashboard-data.xlsx")
assert "DataTable" in wb["Data"].tables, wb["Data"].tables.keys()
data = wb["Data"]
assert data.max_row >= 316, data.max_row  # header row 4 + 312 rows
months = set()
for r in range(5, 317):
    months.add(data.cell(row=r, column=1).value.month)
assert months == set(range(1, 13)), sorted(months)
ex, sol, chk = wb["Exercise"], wb["Solution"], wb["Checks"]
assert ex.cell(row=2, column=1).value == 1 and ex.cell(row=13, column=1).value == 12
assert ex.cell(row=2, column=3).value is None  # empty Revenue answer area
rev = [sol.cell(row=r, column=3).value for r in range(2, 14)]
assert all(isinstance(v, int) for v in rev), rev
assert chk.cell(row=5, column=2).value == "=SUM(DataTable[Revenue])"
assert chk.cell(row=6, column=2).value == "=SUM(Solution!C2:C13)"
print("dashboard-data spot checks OK (312 rows, all 12 months, live checks)")

print("VERIFICATION", "PASSED" if ok else "HAD FAILURES")
