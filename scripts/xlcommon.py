"""Shared helpers for generating Excel Lab practice workbooks.

Every generator script (gen_level*.py) imports this module so all workbooks
look consistent. Run generator scripts from the repo root:

    python scripts/gen_level1.py

Conventions used by the course site:
- Workbooks contain an "Exercise" sheet (student works here) and a
  "Solution" sheet with the finished state.
- Titles/notes live in the first rows; tables start a couple of rows down.
- Formulas are written as strings starting with "=" so Excel computes them.
"""
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

TEAL = "0d9488"        # header fill (site accent)
TEAL_DARK = "0f766e"
BAND = "f0fdfa"        # light teal band for section labels
GRAY = "64748b"
WHITE = "ffffff"
AMBER = "f59e0b"

_THIN = Side(style="thin", color="cbd5e1")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def new_workbook(*sheet_names):
    """Create a workbook with the given sheets (default sheet removed)."""
    wb = Workbook()
    wb.remove(wb.active)
    for name in sheet_names:
        wb.create_sheet(name)
    return wb


def put_title(ws, title, subtitle=None, row=1):
    """Big teal title in row 1 (and italic subtitle in row 2)."""
    c = ws.cell(row=row, column=1, value=title)
    c.font = Font(bold=True, size=14, color=TEAL_DARK)
    if subtitle:
        c2 = ws.cell(row=row + 1, column=1, value=subtitle)
        c2.font = Font(italic=True, size=10, color=GRAY)
        return row + 2
    return row + 1


def put_note(ws, row, text, width=8):
    """Merged instruction note (italic, gray)."""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=width)
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(italic=True, size=10, color=GRAY)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[row].height = 30
    return row + 1


def put_section(ws, row, text, width=8):
    """Light teal band with a bold section label."""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=width)
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(bold=True, size=11, color=TEAL_DARK)
    c.fill = PatternFill("solid", fgColor=BAND)
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[row].height = 20
    return row + 1


def put_headers(ws, row, headers, fill=TEAL):
    """Styled table header row. Freezes panes below it."""
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = Font(bold=True, color=WHITE)
        c.fill = PatternFill("solid", fgColor=fill)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = _BORDER
    ws.freeze_panes = ws.cell(row=row + 1, column=1)
    ws.row_dimensions[row].height = 22
    return row + 1


def put_rows(ws, row, rows, num_formats=None):
    """Write data rows with light borders. num_formats: {col_index: fmt}."""
    for r in rows:
        for i, v in enumerate(r, 1):
            c = ws.cell(row=row, column=i, value=v)
            c.border = _BORDER
            if isinstance(v, str) and v.startswith("="):
                c.alignment = Alignment(horizontal="left")
            elif isinstance(v, (int, float)):
                c.alignment = Alignment(horizontal="right")
            if num_formats and i in num_formats:
                c.number_format = num_formats[i]
        row += 1
    return row


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def blank_solution(ws, row, col, cells, num_formats=None):
    """Fill specific cells on a Solution sheet with values/formulas.

    cells: list of (row, col, value) tuples.
    """
    for r, c_i, v in cells:
        c = ws.cell(row=r, column=c_i, value=v)
        if num_formats and c_i in num_formats:
            c.number_format = num_formats[c_i]
    return row
