# Excel Lab — pilot explainer clip: Level 1, "What Is Excel?"
# Style: English slides + simple English narration (en-IN voice)
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from clipgen import build

LESSON = "Level 1 · What Is Excel?"
OUT = os.path.join("assets", "files", "clips", "l1-what-is-excel.mp4")

slides = [
    {
        "title": "What is Excel?",
        "bullets": [
            "A big grid full of small boxes",
            "Each box is called a cell",
            "Excel does the maths for you",
        ],
        "narr": "Excel is a big grid, full of small boxes. Each box is called a cell. One cell holds one piece of information. Excel does the maths for you. That is what makes Excel different from paper.",
        "caption": "Excel = a big grid of cells",
    },
    {
        "title": "Rows, Columns and Cell Addresses",
        "grid": {
            "title": "Sunrise Traders",
            "headers": ["A", "B"],
            "rows": [
                ["Employee", "Salary"],
                ["Arjun Mehta", "32000"],
                ["Priya Nair", "28500"],
            ],
            "highlight": [[1, 1]],
        },
        "narr": "Columns are named with letters — A, B, C. Rows are numbered — one, two, three. Where column B meets row three, the cell address is B3. Priya's salary sits in cell B3.",
        "caption": "B3 = column B + row 3",
    },
    {
        "title": "Your First Sales Table",
        "grid": {
            "title": "Sunrise Traders — sales",
            "headers": ["Item", "Price", "Qty", "Total"],
            "rows": [
                ["Notebook", "45", "10", "=B2*C2"],
                ["Pen pack", "60", "5", "=B3*C3"],
                ["Printer", "12500", "1", "=B4*C4"],
            ],
            "highlight": [[0, 3]],
        },
        "narr": "This is a small sales table. Price is forty five, quantity is ten. In the Total cell, we wrote a formula: equals B2 times C2. That means, price times quantity. The result, four hundred and fifty, Excel calculates automatically.",
        "caption": "Formula =B2*C2 gives 450 automatically",
    },
    {
        "title": "Open a New Workbook",
        "steps": [
            "Open Excel from the Start menu",
            "Click Blank workbook",
            "Or just press Ctrl + N",
        ],
        "narr": "To start new work, we need a new workbook. Open Excel from the Start menu, and click Blank workbook. A new empty grid opens. If Excel is already open, just press Control plus N.",
        "caption": "Ctrl + N = new workbook",
    },
    {
        "title": "Save Your Work — Ctrl + S",
        "bullets": [
            "File name: Sunrise Traders - Sales Week 1",
            "File type: .xlsx",
            "After the first save, Ctrl + S is enough",
        ],
        "narr": "One office rule: an unsaved file does not exist. Press Control plus S. The first time, the Save As screen appears. Give a clear name. Keep the file type as xlsx. After that, Control plus S updates the file automatically.",
        "caption": "Ctrl + S keeps your file safe",
    },
    {
        "title": "Remember This",
        "bullets": [
            "Excel = a grid of cells",
            "Address = column + row (like B3)",
            "Formulas start with =",
            "Never forget Ctrl + S!",
        ],
        "narr": "Remember this. Excel is a grid, and every box is a cell. Every cell has an address: column letter plus row number. A formula starts with the equals sign. Save your work with Control plus S. Now go back, and practise the exercises.",
        "caption": "Now go back and practise the exercises",
    },
]

if __name__ == "__main__":
    build(slides, LESSON, OUT, voice="en-IN-PrabhatNeural")
