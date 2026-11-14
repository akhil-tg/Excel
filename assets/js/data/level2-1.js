// Excel Lab — Level 2 (Formulas and Functions), part 1 of 2
// Metadata + lessons 1-7. Part 2 lives in level2-2.js.

window.LEVEL_META = window.LEVEL_META || {};
window.LEVEL_LESSONS = window.LEVEL_LESSONS || {};

window.LEVEL_META[2] = {
  id: 2,
  title: "Formulas and Functions",
  icon: "fx",
  difficulty: "Beginner",
  subtitle: "Master the functions that do the real work: SUM, IF, lookups, text and dates",
  description: "Learn the essential Excel functions step by step — from SUM and IF to VLOOKUP, XLOOKUP, INDEX/MATCH, text, date and dynamic array functions."
};

(window.LEVEL_LESSONS[2] = window.LEVEL_LESSONS[2] || []).push(
  {
    id: "l2-function-basics",
    title: "How Functions Work",
    difficulty: "Beginner",
    duration: 12,
    version: "",
    objectives: [
      "Break any function into its three parts: the equals sign, the name and the arguments",
      "Type functions quickly using the autocomplete suggestion list",
      "Use the Insert Function (fx) dialog when you cannot remember a function name",
      "Recognise and fix the most common syntax mistakes"
    ],
    intro: "A function is a ready-made calculation that Excel already knows how to do. Instead of typing =B2+B3+B4+B5, you type `=SUM(B2:B5)` and Excel does the adding for you. Excel has hundreds of these ready-made calculations, and in this level you will learn the ones that office work actually uses every day: totals, averages, decisions, lookups, text and dates.\n\nEvery function follows the same simple shape. Once you can read that shape, every new function in this course becomes easier, because you already know where the name goes, where the inputs go, and what Excel expects from you.",
    sections: [
      {
        heading: "The anatomy of a function",
        blocks: [
          { type: "p", text: "Look at this formula carefully — it contains every part a function can have:" },
          {
            type: "formula",
            name: "SUM",
            syntax: "=SUM(number1, [number2], ...)",
            example: "=SUM(C2:C4)",
            explanation: "SUM is the function name. C2:C4 is the argument — the data the function works on. The equals sign tells Excel that a formula is starting."
          },
          {
            type: "ul",
            items: [
              "**The equals sign ( = )** always comes first. It tells Excel: \"what follows is a formula, please calculate it.\"",
              "**The function name** (here `SUM`) says what kind of calculation to do.",
              "**Round brackets** hold the arguments — the inputs the function works with.",
              "**Arguments are separated by commas.** In `=ROUND(A2, 2)` there are two arguments: the number and the number of digits."
            ]
          },
          { type: "p", text: "Arguments can be single cells like `C2`, ranges like `C2:C4`, plain numbers like `2`, text like `\"North\"`, or even other functions tucked inside. Everything you will learn in this level fits inside this one simple shape." },
          {
            type: "grid",
            title: "A first function on the sheet",
            rows: [
              ["", "Product", "Price (₹)"],
              ["", "Notebooks", "45"],
              ["", "Pens", "12"],
              ["", "Staplers", "130"],
              ["Total", "", "=SUM(C2:C4)"]
            ]
          },
          { type: "p", text: "Cell `C5` shows 187. If the price of Pens changes from 12 to 15, the total updates to 190 on its own — no retyping. This is the whole point of formulas: the sheet recalculates itself when the data changes." }
        ]
      },
      {
        heading: "Arguments, commas and the tooltip",
        blocks: [
          { type: "p", text: "Some functions need one argument, some need two or three, and some can take many. You never have to memorise this, because Excel reminds you: as soon as you type the opening bracket, a small tooltip appears under the cell." },
          {
            type: "grid",
            title: "The tooltip while typing =ROUND(",
            rows: [
              ["You type", "Excel shows underneath"],
              ["=ROUND(", "ROUND(number, num_digits)"]
            ]
          },
          {
            type: "ul",
            items: [
              "**Bold arguments are required.** In the SUM tooltip, `number1` is bold — you must supply it.",
              "**Square brackets mean optional.** `[number2]` in SUM is optional; you can stop after one range.",
              "**Commas separate arguments.** `=ROUND(A2, 2)` has one comma and two arguments.",
              "**Nesting adds brackets.** In `=IF(A2>40, \"Pass\", \"Fail\")`, every bracket you open must be closed."
            ]
          },
          { type: "tip", text: "If the tooltip disappears, click inside the formula in the formula bar and it comes back. The argument currently shown in bold is the one Excel expects you to type next — the bold part moves along as you type each comma." }
        ]
      },
      {
        heading: "Typing functions with autocomplete",
        blocks: [
          { type: "p", text: "Typing function names by hand is slow and invites spelling mistakes. Excel's autocomplete list fixes both problems." },
          {
            type: "steps",
            items: [
              "Click the cell and type `=` to start the formula.",
              "Type the first letters, for example `=SU`. A dropdown list appears showing SUM, SUMIF, SUMIFS and every other function that starts with SU.",
              "Use the [[Arrow Down]] key to move through the list. A tooltip beside the list explains what the selected function does.",
              "Press [[Tab]] to accept the highlighted function. Excel completes the name and types the opening bracket for you.",
              "Select or type the arguments, then press [[Enter]]. Excel adds the closing bracket automatically."
            ]
          },
          { type: "p", text: "Autocomplete is not just about speed. It also protects you from misspelled names: if a name you remember is not in the list, you misremembered it. And when you type an opening bracket after the name, Excel often highlights the range you are selecting on the sheet, so you can confirm you picked the right cells." }
        ]
      },
      {
        heading: "Insert Function and editing functions",
        blocks: [
          { type: "p", text: "When you know what you want to calculate but not which function does it, use Insert Function. It is a built-in search engine for calculations." },
          {
            type: "steps",
            items: [
              "Click the **fx** button to the left of the formula bar, or press [[Shift]] + [[F3]].",
              "In the search box, describe the task in plain words, for example \"count cells\" or \"round a number\", and click Go.",
              "Pick a function from the list — the dialog shows a short description of each one — and click OK.",
              "Fill in each argument box. The Function Arguments dialog explains every argument and shows a live preview of the result at the bottom.",
              "Click OK to place the finished formula in the cell."
            ]
          },
          { type: "p", text: "To edit a function later, click the cell and press [[F2]], or simply edit the text in the formula bar. Double-clicking a cell also enters edit mode. For long formulas with many nested brackets, reopening the fx dialog is often easier than editing raw text, because each argument sits in its own labelled box." }
        ]
      },
      {
        heading: "Common syntax mistakes",
        blocks: [
          {
            type: "table",
            headers: ["Mistake", "What you see", "Fix"],
            rows: [
              ["Missing equals sign", "The cell shows the typed text itself, like SUM(C2:C4), with no result", "Start every formula with ="],
              ["Misspelled function name", "#NAME? error", "Check the spelling or pick the name from the autocomplete list"],
              ["Missing comma between arguments", "Excel refuses the formula, or the dialog shows an error", "Separate arguments with commas"],
              ["Missing closing bracket", "Excel refuses to accept the formula when you press Enter", "Count your brackets — every open bracket needs a close"]
            ]
          },
          { type: "warning", text: "When Excel rejects a formula, it usually highlights the exact problem spot in the formula bar. Do not delete the whole formula and start again — press [[Esc]], read the highlighted part, and fix only that." }
        ]
      }
    ],
    mistakes: [
      {
        wrong: "Typing a calculation without the equals sign, like SUM(C2:C4)",
        right: "=SUM(C2:C4)",
        why: "Without =, Excel stores the text as plain text instead of calculating. The cell just repeats what you typed."
      },
      {
        wrong: "=SUMME(C2:C4)",
        right: "=SUM(C2:C4)",
        why: "A misspelled function name is not a function. Excel returns the #NAME? error because it cannot find a function with that name."
      },
      {
        wrong: "=ROUND(A2 2)",
        right: "=ROUND(A2, 2)",
        why: "Arguments must be separated by commas. Without the comma Excel sees one broken argument and refuses the formula."
      },
      {
        wrong: "=IF(A2>40, \"Pass\" \"Fail\")",
        right: "=IF(A2>40, \"Pass\", \"Fail\")",
        why: "The two results are two separate arguments, so they need a comma between them."
      }
    ],
    shortcuts: [
      { keys: "Tab", action: "Accept the highlighted function from the autocomplete list" },
      { keys: "Shift + F3", action: "Open the Insert Function dialog" },
      { keys: "F2", action: "Edit the formula in the active cell" }
    ],
    exercises: [
      {
        task: "In a blank sheet, type Notebooks, Pens and Staplers in cells B2 to B4 and their prices 45, 12 and 130 in cells C2 to C4. Build the total in C5 using only the autocomplete list — type =SU and let Excel finish the job.",
        hint: "Press [[Tab]] to accept the suggestion; Excel adds the opening bracket for you.",
        solution: "Select C5, type =SU, press [[Tab]], select C2:C4, press [[Enter]]. The cell shows =SUM(C2:C4) and the result 187."
      },
      {
        task: "Use the fx button to insert the ROUND function and round the total in C5 to zero decimal places.",
        hint: "In the Function Arguments dialog, type C5 in the number box and 0 in the num_digits box.",
        solution: "The formula =ROUND(C5, 0) returns 187. The dialog shows the live result before you click OK."
      },
      {
        task: "Deliberately break the formula in C5 by deleting the equals sign, press Enter, and look at the cell. Then undo and repair it properly.",
        hint: "A formula without = is just text to Excel.",
        solution: "The cell shows the literal text SUM(C2:C4). Put the equals sign back: =SUM(C2:C4) calculates again."
      }
    ],
    quiz: [
      {
        q: "What must every Excel formula start with?",
        options: ["The equals sign ( = )", "The function name", "An opening bracket", "A cell reference"],
        answer: 0,
        explanation: "The equals sign tells Excel that the cell contains a formula to calculate. Without it, Excel stores your typing as plain text."
      },
      {
        q: "In the tooltip SUM(number1, [number2], ...), what do the square brackets around number2 tell you?",
        options: ["The argument is optional", "The argument must be a range", "The argument must be text", "The argument is an error"],
        answer: 0,
        explanation: "Square brackets mark optional arguments. SUM works fine with just one range; extra arguments are your choice."
      },
      {
        q: "You type =SU and the autocomplete list appears. What is the quickest way to accept the highlighted function?",
        options: ["Press Tab", "Delete it and retype by hand", "Click another cell", "Press Esc"],
        answer: 0,
        explanation: "[[Tab]] completes the function name and adds the opening bracket. It is faster and safer than typing the full name."
      },
      {
        q: "A cell shows #NAME? after you enter =SUMME(A2:A5). What is the most likely cause?",
        options: ["The function name is misspelled", "The range is too small", "A number is formatted as text", "The file has not been saved"],
        answer: 0,
        explanation: "#NAME? means Excel does not recognise a name in the formula. SUMME is not a function; the correct name is SUM."
      },
      {
        q: "In =ROUND(A2, 2), which part is the second argument?",
        options: ["2", "A2", "ROUND", "The comma"],
        answer: 0,
        explanation: "Arguments are read left to right between the brackets: A2 is first, 2 is second. The comma is only the separator."
      }
    ],
    summary: [
      "Every function has the shape =NAME(argument1, argument2, ...) — the equals sign first, then the name, then bracketed arguments separated by commas.",
      "The tooltip shows which argument Excel expects next; bold means required, square brackets mean optional.",
      "Autocomplete plus [[Tab]] types function names faster and without spelling mistakes.",
      "The fx (Insert Function) dialog helps you find and build functions you do not know by name.",
      "The #NAME? error usually means a misspelled function name."
    ],
    downloads: [],
    keywords: ["function anatomy", "function arguments", "function tooltip", "autocomplete", "Tab key", "Insert Function", "fx button", "syntax mistakes", "NAME error"]
  },
  {
    id: "l2-sum-average-min-max",
    title: "SUM, AVERAGE, MIN and MAX",
    difficulty: "Beginner",
    duration: 14,
    version: "",
    objectives: [
      "Total a column in two keystrokes with AutoSum",
      "Write SUM by hand for one range or several ranges",
      "Use AVERAGE, MIN and MAX to summarise sales data",
      "Decide when a whole-column reference like E:E is helpful and when it is risky",
      "Explain how blank cells behave differently from zeros"
    ],
    intro: "Every office report needs four basic questions answered: how much in total, what is typical, what is the biggest and what is the smallest. SUM, AVERAGE, MIN and MAX answer exactly these questions, and they are the four functions you will use most often in daily work.\n\nIn this lesson we work with the sales record of Sunrise Traders, a wholesale office supplies business. The same little table appears in your practice workbook, so you can follow along cell by cell.",
    sections: [
      {
        heading: "AutoSum: the fastest total in Excel",
        blocks: [
          { type: "p", text: "AutoSum is a button that writes SUM for you. Select the cell directly **below** a column of numbers (or to the **right** of a row of numbers), then press [[Alt]] + [[=]]. Excel guesses which cells you want to add, writes the formula, and waits for you to confirm." },
          {
            type: "steps",
            items: [
              "Click the cell where the total should go, usually just under the last number.",
              "Press [[Alt]] + [[=]], or click Home, then AutoSum on the far right of the ribbon.",
              "Excel draws a moving border around the range it guessed, for example E2:E10.",
              "Check the guess. If it is right, press [[Enter]]. If it is wrong, select the correct range first, then press [[Enter]]."
            ]
          },
          { type: "warning", text: "Always glance at the range AutoSum guessed before pressing Enter. If there is a gap in the column — even one empty cell — Excel stops guessing at the gap and your total misses rows." },
          {
            type: "grid",
            title: "Sunrise Traders sales record",
            rows: [
              ["Region", "Product", "Units", "Price (₹)", "Amount (₹)"],
              ["North", "Notebooks", "40", "45", "=C2*D2"],
              ["South", "Pens", "100", "12", "=C3*D3"],
              ["East", "Staplers", "15", "130", "=C4*D4"],
              ["North", "Printers", "2", "8500", "=C5*D5"],
              ["West", "Laptops", "1", "42000", "=C6*D6"],
              ["South", "Notebooks", "60", "45", "=C7*D7"],
              ["East", "Pens", "80", "12", "=C8*D8"],
              ["North", "Staplers", "10", "130", "=C9*D9"],
              ["West", "Notebooks", "25", "45", "=C10*D10"]
            ]
          },
          { type: "p", text: "The Amount column multiplies units by price. Now the totals go underneath, in row 11." }
        ]
      },
      {
        heading: "Writing SUM by hand",
        blocks: [
          {
            type: "formula",
            name: "SUM",
            syntax: "=SUM(number1, [number2], ...)",
            example: "=SUM(E2:E10)",
            explanation: "Adds every number in the range E2 to E10. A colon between two cell references means \"from this cell to that cell\"."
          },
          { type: "p", text: "The colon is the range operator: `E2:E10` means all nine cells from E2 through E10. You can also give SUM several separate ranges at once, separated by commas. This is useful when the rows you need are not next to each other." },
          {
            type: "formula",
            name: "SUM with several ranges",
            syntax: "=SUM(range1, range2, ...)",
            example: "=SUM(E2:E5, E8:E10)",
            explanation: "Adds two blocks of rows and skips the rows in between. The comma means \"and also\"."
          },
          { type: "tip", text: "You can type the range or select it with the mouse. When you start typing =SUM( and then drag across the cells, Excel writes the range for you — this is the safest way to avoid typos like E2:E01." }
        ]
      },
      {
        heading: "AVERAGE, MIN and MAX",
        blocks: [
          {
            type: "grid",
            title: "Summary row for the sales record (row 11 to 14 in column F)",
            rows: [
              ["Measure", "Formula", "Result"],
              ["Total sales", "=SUM(E2:E10)", "70035"],
              ["Average sale", "=AVERAGE(E2:E10)", "7781.67"],
              ["Biggest sale", "=MAX(E2:E10)", "42000"],
              ["Smallest sale", "=MIN(E2:E10)", "960"]
            ]
          },
          {
            type: "ul",
            items: [
              "`=SUM(E2:E10)` gives 70035 — the month's total revenue.",
              "`=AVERAGE(E2:E10)` gives 7781.67 — the average invoice value. Managers use this as a quick \"typical order\" figure.",
              "`=MAX(E2:E10)` gives 42000 — the single biggest sale (the laptop order). Useful for spotting your best deal.",
              "`=MIN(E2:E10)` gives 960 — the smallest sale. In stock work, MIN over a quantity column shows which item needs reordering first."
            ]
          },
          { type: "p", text: "All four functions accept ranges exactly the same way: one range, several ranges separated by commas, or a whole column. Once you understand SUM, the other three come free." }
        ]
      },
      {
        heading: "Whole-column references: E:E",
        blocks: [
          { type: "p", text: "Instead of `=SUM(E2:E10)` you can write `=SUM(E:E)`, which adds every cell in column E. This is tempting because it is short and new rows are included automatically." },
          {
            type: "ul",
            items: [
              "**Advantage:** if you add sales in row 11, 12 and beyond, `=SUM(E:E)` keeps counting. A fixed range like E2:E10 does not.",
              "**Advantage:** you never have to update the formula when the table grows.",
              "**Risk:** the formula includes the whole column — the header, any notes, and any stray numbers above or below the table.",
              "**Risk:** on very large sheets, whole-column formulas in many cells can slow the workbook down."
            ]
          },
          { type: "warning", text: "SUM ignores text, so a header like \"Amount (₹)\" does not break =SUM(E:E). But COUNTA does count text, and a stray number parked somewhere in the column silently changes the total. In professional reports, a clean fixed range is usually the safer habit." }
        ]
      },
      {
        heading: "Blanks are not zeros",
        blocks: [
          { type: "p", text: "This trips up almost every beginner, so read it twice. An empty cell and a cell containing 0 behave differently in every function in this lesson." },
          {
            type: "table",
            headers: ["Cell content", "SUM", "AVERAGE", "COUNT"],
            rows: [
              ["Empty cell", "Ignored", "Ignored", "Not counted"],
              ["Typed 0", "Added as 0", "Included — pulls the average down", "Counted as a number"],
              ["Text such as \"N/A\"", "Ignored", "Ignored", "Not counted"]
            ]
          },
          { type: "p", text: "Example: your delivery team records units sold, but two days have no entry at all (empty cells). AVERAGE gives the average of the days that were actually recorded. If the team had typed 0 for those days instead, the average would drop, because the zeros now count as real days with no sales. Which is right depends on what the number means — that is a business decision, not an Excel one." },
          { type: "p", text: "In plain arithmetic the rules change again: an empty cell used in a calculation acts like 0. `=E2*2` returns 0 when E2 is empty. So blanks are skipped by SUM and AVERAGE, but treated as zero inside formulas that multiply or divide them." }
        ]
      }
    ],
    mistakes: [
      {
        wrong: "=SUM(E2 E10)",
        right: "=SUM(E2:E10)",
        why: "A range needs a colon. Without it Excel refuses the formula; with a plus sign like =SUM(E2+E10) you would add only two cells instead of the whole range."
      },
      {
        wrong: "Trusting the AutoSum guess without checking it",
        right: "Look at the highlighted range, then press Enter",
        why: "A blank cell or a subtotal row in the middle of the column makes AutoSum stop early, and the total silently misses rows."
      },
      {
        wrong: "Expecting AVERAGE to count empty cells as zero",
        right: "Understand that AVERAGE skips blanks",
        why: "Blank days are excluded from the average; only typed zeros count. Enter zeros deliberately when \"no sales\" should pull the average down."
      },
      {
        wrong: "Using COUNTA over a whole column with headers, =COUNTA(E:E)",
        right: "Use a fixed range like =COUNTA(E2:E10)",
        why: "COUNTA counts text too, so the header \"Amount (₹)\" inflates the count by one. Whole-column references and COUNTA mix badly."
      }
    ],
    shortcuts: [
      { keys: "Alt + =", action: "AutoSum — insert =SUM() for the neighbouring numbers" },
      { keys: "Enter", action: "Confirm the AutoSum range after checking it" }
    ],
    exercises: [
      {
        task: "Open the formula-practice workbook, Exercise sheet. In cell B12, use AutoSum to total the daily sales in B2:B10.",
        hint: "Click B12 first, then press [[Alt]] + [[=]] and check that the marching-ants border covers B2:B10.",
        solution: "B12 contains =SUM(B2:B10)."
      },
      {
        task: "On the same sheet, put the average daily sale in B13, the highest in B14 and the lowest in B15.",
        hint: "Same range, three different functions.",
        solution: "B13: =AVERAGE(B2:B10). B14: =MAX(B2:B10). B15: =MIN(B2:B10)."
      },
      {
        task: "The Solution sheet of the same workbook shows a month where two days have blank sales cells. Compare the average with and without those days recorded as 0 and note the difference.",
        hint: "Type 0 into the blanks on a copy of the sheet and watch the AVERAGE result.",
        solution: "With blanks, AVERAGE divides only by the days that have numbers. With zeros, it divides by every day, so the average is lower."
      },
      {
        task: "In cell D2 of the Exercise sheet, add the two separate blocks A2:A5 and A8:A10 in one formula.",
        hint: "Ranges separated by a comma inside the same SUM.",
        solution: "=SUM(A2:A5, A8:A10) adds both blocks and skips the rows in between."
      }
    ],
    quiz: [
      {
        q: "Which keystroke inserts an AutoSum formula?",
        options: ["Alt + =", "Ctrl + S", "Shift + Enter", "Ctrl + A"],
        answer: 0,
        explanation: "[[Alt]] + [[=]] writes =SUM() around the numbers next to the active cell. You confirm with Enter after checking the guessed range."
      },
      {
        q: "What does =SUM(E2:E5, E8:E10) calculate?",
        options: ["The total of both ranges, skipping rows 6 and 7", "Only the first range", "An error, because SUM takes one range only", "The average of the two ranges"],
        answer: 0,
        explanation: "Commas let SUM add several ranges in one formula. Here rows 6 and 7 are left out on purpose."
      },
      {
        q: "Your range A2:A10 has numbers in A2:A6 and empty cells in A7:A10. What does =AVERAGE(A2:A10) return?",
        options: ["The average of A2:A6 only", "Zero", "The average treating blanks as zero", "An error"],
        answer: 0,
        explanation: "AVERAGE ignores empty cells completely, so only the five real numbers are averaged."
      },
      {
        q: "What is the main risk of =SUM(E:E) in a report?",
        options: ["It includes every cell in column E, including stray numbers outside the table", "It cannot handle text in the header", "It only works in the first column", "It must be retyped every month"],
        answer: 0,
        explanation: "SUM ignores text so the header is safe, but any stray number anywhere in column E silently joins the total."
      },
      {
        q: "Which function returns the smallest value in a range?",
        options: ["MIN", "SMALL", "LOW", "BOTTOM"],
        answer: 0,
        explanation: "MIN returns the smallest number in the range. SMALL also exists but needs an extra k argument; LOW and BOTTOM are not Excel functions."
      }
    ],
    summary: [
      "AutoSum ([[Alt]] + [[=]]) writes SUM for you — always check the range it guesses before pressing Enter.",
      "SUM, AVERAGE, MIN and MAX all accept ranges: =SUM(E2:E10), several ranges =SUM(E2:E5, E8:E10), or whole columns =SUM(E:E).",
      "Whole-column references are convenient but can pick up stray numbers; fixed ranges are the safer professional habit.",
      "Empty cells are ignored by SUM and AVERAGE; typed zeros are counted. Choose deliberately which one your data means.",
      "MIN and MAX answer \"best and worst\" — the biggest invoice and the smallest sale in one formula each."
    ],
    downloads: [
      {
        file: "assets/files/level2/formula-practice.xlsx",
        label: "Formula practice workbook",
        note: "Exercise and Solution sheets for SUM, AVERAGE, MIN, MAX, COUNT, COUNTA and ROUND"
      }
    ],
    keywords: ["SUM", "AVERAGE", "MIN", "MAX", "AutoSum", "Alt plus equals", "range", "whole column reference", "blank vs zero", "sales report"]
  },
  {
    id: "l2-count-round",
    title: "COUNT, COUNTA, COUNTBLANK and ROUND",
    difficulty: "Beginner",
    duration: 12,
    version: "",
    objectives: [
      "Choose correctly between COUNT, COUNTA and COUNTBLANK",
      "Use ROUND, ROUNDUP and ROUNDDOWN to change stored values",
      "Explain the difference between rounding a value and formatting a cell",
      "Round money amounts to two decimal places the professional way"
    ],
    intro: "Counting is not one job but three: how many cells hold numbers, how many cells are filled at all, and how many are empty. Excel gives each job its own function — COUNT, COUNTA and COUNTBLANK. Mixing them up is one of the most common beginner errors in reports.\n\nRounding is the second skill here. After this lesson you will know why the ROUND function and the cell format with two decimals are two completely different things — a difference that matters every time money is involved.",
    sections: [
      {
        heading: "Three ways to count",
        blocks: [
          {
            type: "formula",
            name: "COUNT",
            syntax: "=COUNT(value1, [value2], ...)",
            example: "=COUNT(C2:C6)",
            explanation: "Counts cells that contain numbers. Text, empty cells and dates-as-text are all ignored."
          },
          {
            type: "formula",
            name: "COUNTA",
            syntax: "=COUNTA(value1, [value2], ...)",
            example: "=COUNTA(B2:B6)",
            explanation: "Counts every non-empty cell — numbers, text, anything. Use it for \"how many entries do we have?\""
          },
          {
            type: "formula",
            name: "COUNTBLANK",
            syntax: "=COUNTBLANK(range)",
            example: "=COUNTBLANK(C2:C6)",
            explanation: "Counts empty cells in a range. Takes exactly one range and answers \"what is still missing?\""
          },
          {
            type: "grid",
            title: "Counting a partly filled stock list",
            rows: [
              ["", "Item", "Units sold"],
              ["", "Notebooks", "40"],
              ["", "Pens", ""],
              ["", "Staplers", "15"],
              ["", "Printers", "2"],
              ["", "Laptops", ""],
              ["Numbers only", "=COUNT(C2:C6)", "3"],
              ["Filled cells", "=COUNTA(B2:B6)", "5"],
              ["Empty cells", "=COUNTBLANK(C2:C6)", "2"]
            ]
          },
          { type: "p", text: "Read the grid like a report. `COUNT(C2:C6)` returns 3 because only three items have units sold recorded. `COUNTA(B2:B6)` returns 5 because all five item names are filled. `COUNTBLANK(C2:C6)` returns 2 — the Pens and Laptops rows still need data. In office work, COUNTBLANK is a quick data-quality check before you send a report." }
        ]
      },
      {
        heading: "When each count is used in real work",
        blocks: [
          {
            type: "ul",
            items: [
              "**COUNT** — \"how many sales did we record?\" Works on amount and quantity columns, which hold numbers.",
              "**COUNTA** — \"how many customers are in the list?\" Works on name and city columns, which hold text.",
              "**COUNTBLANK** — \"how many invoices are missing a date?\" A blank count above zero means the report is not ready.",
              "**COUNT on a text column returns 0** — that is correct behaviour, not a bug. Names are not numbers."
            ]
          },
          { type: "tip", text: "A neat identity: COUNT + COUNTBLANK on the same single-column range always equals the total number of cells in that range. If it does not, some cells hold something unexpected — worth investigating." }
        ]
      },
      {
        heading: "ROUND, ROUNDUP and ROUNDDOWN",
        blocks: [
          {
            type: "formula",
            name: "ROUND",
            syntax: "=ROUND(number, num_digits)",
            example: "=ROUND(1234.567, 2)",
            explanation: "Rounds the value itself. With num_digits 2 the example returns 1234.57 — the stored value really changes."
          },
          {
            type: "ul",
            items: [
              "`=ROUND(1234.567, 2)` returns 1234.57 — two decimal places, the standard for money.",
              "`=ROUND(1234.567, 0)` returns 1235 — no decimals.",
              "`=ROUND(12345, -2)` returns 12300 — a negative num_digits rounds to tens, hundreds, thousands.",
              "`=ROUNDUP(1234.561, 2)` returns 1234.57 — always away from zero, even for 1234.561.",
              "`=ROUNDDOWN(1234.569, 2)` returns 1234.56 — always towards zero, chopping the extra digits off."
            ]
          },
          { type: "p", text: "ROUNDUP and ROUNDDOWN are for business rules, not maths. If notebooks are packed 25 to a carton, 58 notebooks need 2.32 cartons — you cannot order a third of a carton, so `=ROUNDUP(58/25, 0)` gives the correct answer 3. Rounding down would leave seven notebooks unpacked." }
        ]
      },
      {
        heading: "ROUND versus formatting: the money mistake",
        blocks: [
          { type: "p", text: "This is the part to remember. Cell formatting changes only what you **see**. The ROUND function changes the **value** that is stored in the cell." },
          {
            type: "grid",
            title: "Same cell, two very different realities",
            rows: [
              ["", "A", "B"],
              ["1", "What you do", "What the cell really holds"],
              ["2", "Format with 2 decimals", "1234.567 (unchanged)"],
              ["3", "=ROUND(1234.567, 2)", "1234.57 (really changed)"]
            ]
          },
          { type: "p", text: "Suppose an invoice line shows ₹123.46 after formatting, and another shows ₹123.46 after rounding, but both started from different raw values. Format the display and the true value still has its extra digits — a long invoice totalled from formatted cells can end up a few paise off from the sum of the figures printed on the page. Auditors hate that. When the printed total must match the printed lines, ROUND the values first." },
          {
            type: "formula",
            name: "Rounding money the professional way",
            syntax: "=ROUND(amount, 2)",
            example: "=ROUND(C2*D2, 2)",
            explanation: "Multiply, then round the result to two decimals. The stored value matches what the invoice prints."
          },
          { type: "warning", text: "Formatting is still useful — use it for readability everywhere. But when a downstream total must equal the sum of the displayed numbers, round the values with ROUND and format afterwards." }
        ]
      }
    ],
    mistakes: [
      {
        wrong: "Using COUNT on a name column, =COUNT(B2:B50)",
        right: "Use COUNTA for text: =COUNTA(B2:B50)",
        why: "COUNT counts numbers only. On a list of customer names it returns 0, which looks like the list is empty."
      },
      {
        wrong: "Believing the 2-decimal number format has rounded the value",
        right: "Use =ROUND(value, 2) when the stored value itself must change",
        why: "Formatting changes display only. Totals built from unrounded values can disagree with the printed figures."
      },
      {
        wrong: "=ROUNDDOWN(58/25, 0) for cartons needed",
        right: "=ROUNDUP(58/25, 0)",
        why: "Rounding down gives 2 cartons for 58 notebooks — seven notebooks with nowhere to go. Business quantities usually need ROUNDUP."
      },
      {
        wrong: "Typing =COUNTBLANK(B2, C2, D2)",
        right: "=COUNTBLANK(B2:D2)",
        why: "COUNTBLANK takes one range, not a list of cells. Give it a proper range with a colon."
      }
    ],
    shortcuts: [
      { keys: "Alt + =", action: "AutoSum can also insert COUNT variants — open its dropdown to see them" }
    ],
    exercises: [
      {
        task: "Open the formula-practice workbook, Exercise sheet. In cell D2 count how many cells in the range B2:B10 contain numbers.",
        hint: "Numbers only — that is one specific counting function.",
        solution: "=COUNT(B2:B10)"
      },
      {
        task: "In cell D3 count how many cells in A2:A10 are filled with anything at all, and in D4 count how many in B2:B10 are still empty.",
        hint: "One function counts non-empty cells, the other counts blanks.",
        solution: "D3: =COUNTA(A2:A10). D4: =COUNTBLANK(B2:B10)."
      },
      {
        task: "The Exercise sheet computes amounts as units times price in column E. Change the first amount formula so the stored value is rounded to two decimals.",
        hint: "Wrap the multiplication in a function that changes the value, not the display.",
        solution: "E2 becomes =ROUND(C2*D2, 2), then fill down to E10."
      },
      {
        task: "Staplers are supplied 10 per box. In F2 of the Exercise sheet, work out how many full boxes to order for the 15 units sold, always rounding up.",
        hint: "Divide by 10, then round in the safe direction for shortages.",
        solution: "=ROUNDUP(15/10, 0) returns 2 boxes."
      }
    ],
    quiz: [
      {
        q: "A range holds three names, two numbers and one empty cell. What does =COUNT over that range return?",
        options: ["2", "3", "5", "6"],
        answer: 0,
        explanation: "COUNT counts numbers only, so just the two numeric cells count. Names and the blank are ignored."
      },
      {
        q: "Which function counts every non-empty cell, text included?",
        options: ["COUNTA", "COUNT", "COUNTBLANK", "COUNTIF"],
        answer: 0,
        explanation: "COUNTA counts anything that is not empty. COUNT is numbers-only and COUNTBLANK counts the gaps."
      },
      {
        q: "What does =ROUND(1234.567, 2) return?",
        options: ["1234.57", "1234.56", "1235", "1200"],
        answer: 0,
        explanation: "The third decimal is 7, so rounding to two decimals gives 1234.57 — and the stored value really becomes 1234.57."
      },
      {
        q: "A cell is formatted to show 2 decimals but holds 1234.567. Another cell holds =ROUND(1234.567, 2). Both display 1234.57. What differs?",
        options: ["The second cell's stored value is truly 1234.57", "Nothing — both are identical", "The first cell now holds text", "The second cell ignores the decimals"],
        answer: 0,
        explanation: "Formatting changes only the display; ROUND changes the stored value. Sums built from the two cells can disagree."
      },
      {
        q: "58 notebooks ship 25 per carton and you must order whole cartons. Which formula gives the right number of cartons?",
        options: ["=ROUNDUP(58/25, 0)", "=ROUND(58/25, 2)", "=ROUNDDOWN(58/25, 0)", "=COUNT(58/25)"],
        answer: 0,
        explanation: "58/25 is 2.32 cartons; you cannot order part of a carton, so round up to 3. Rounding down would leave notebooks unpacked."
      }
    ],
    summary: [
      "COUNT counts numbers, COUNTA counts any non-empty cell, COUNTBLANK counts empty cells.",
      "COUNTBLANK is a fast data-quality check: blanks in a report column mean missing data.",
      "ROUND(value, 2) changes the stored value; the 2-decimal format changes only the display.",
      "For money, round the value with ROUND and let formatting handle readability.",
      "ROUNDUP and ROUNDDOWN encode business rules like whole cartons and full billing days."
    ],
    downloads: [
      {
        file: "assets/files/level2/formula-practice.xlsx",
        label: "Formula practice workbook",
        note: "Exercise and Solution sheets for SUM, AVERAGE, MIN, MAX, COUNT, COUNTA and ROUND"
      }
    ],
    keywords: ["COUNT", "COUNTA", "COUNTBLANK", "ROUND", "ROUNDUP", "ROUNDDOWN", "rounding vs formatting", "two decimals", "money rounding"]
  },
  {
    id: "l2-if",
    title: "IF: Making Decisions in Excel",
    difficulty: "Beginner",
    duration: 15,
    version: "",
    objectives: [
      "Write =IF(condition, value_if_true, value_if_false) from memory",
      "Use all six comparison operators, including \"not equal\" with <>",
      "Put text criteria in quotes and know when numbers must not be quoted",
      "Build realistic Pass/Fail, bonus and stock-reorder formulas",
      "Nest one IF inside another for a simple three-way result"
    ],
    intro: "So far every function has calculated something. IF is different: it decides. \"If the marks are at least 40, write Pass, otherwise write Fail.\" \"If stock is below 10, flag it for reorder.\" Almost every checklist, payslip and stock report in an office runs on IF.\n\nIF is also your first function with true and false outcomes, so it introduces an idea the rest of this level builds on: Excel can test a condition and return different values depending on the answer.",
    sections: [
      {
        heading: "The shape of a decision",
        blocks: [
          {
            type: "formula",
            name: "IF",
            syntax: "=IF(logical_test, value_if_true, value_if_false)",
            example: "=IF(B2>=40, \"Pass\", \"Fail\")",
            explanation: "Excel tests B2 >= 40. If the test is true the cell shows Pass; if not, it shows Fail. Exactly one of the two results appears."
          },
          { type: "p", text: "Three parts, two commas: the **condition** to test, the **result when true**, and the **result when false**. All three are required — leave out the third and Excel shows FALSE whenever the test fails, which looks broken." },
          {
            type: "grid",
            title: "Pass/Fail list for a training test",
            rows: [
              ["Student", "Marks", "Result"],
              ["Arjun Mehta", "78", "=IF(B2>=40, \"Pass\", \"Fail\")"],
              ["Priya Nair", "35", "=IF(B3>=40, \"Pass\", \"Fail\")"],
              ["Rahul Verma", "52", "=IF(B4>=40, \"Pass\", \"Fail\")"],
              ["Sneha Iyer", "88", "=IF(B5>=40, \"Pass\", \"Fail\")"]
            ]
          },
          { type: "p", text: "The results column shows Pass, Fail, Pass, Pass. Notice what happened when the marks were copied down: Excel adjusted each row's reference automatically — B2 became B3, then B4. One formula, four decisions." }
        ]
      },
      {
        heading: "The six comparison operators",
        blocks: [
          {
            type: "table",
            headers: ["Operator", "Meaning", "Example test", "True when..."],
            rows: [
              ["=", "equal to", "A2=\"North\"", "the region is North"],
              ["<>", "not equal to", "A2<>\"North\"", "the region is anything except North"],
              [">", "greater than", "E2>5000", "the amount is 5000.01 or more"],
              ["<", "less than", "D2<10", "the stock is 9 or fewer"],
              [">=", "greater than or equal", "B2>=40", "the marks are 40 or more"],
              ["<=", "less than or equal", "D2<=10", "the stock is 10 or fewer"]
            ]
          },
          { type: "p", text: "Pick >= or <= when the boundary itself belongs to one side. \"Pass at 40\" means 40 passes, so the test is `B2>=40`, not `B2>40`. Choosing between them is a business decision — get it wrong and someone exactly on the boundary falls through." },
          { type: "tip", text: "The \"not equal\" operator is two characters, `<>`, and is read \"is different from\". It is perfect for flagging exceptions: =IF(A2<>\"Paid\", \"Follow up\", \"OK\")." }
        ]
      },
      {
        heading: "Text in quotes, numbers without",
        blocks: [
          { type: "p", text: "Whenever you compare against **text**, wrap the text in double quotes: `=IF(A2=\"North\", \"Home zone\", \"Other zone\")`. Forget the quotes and Excel thinks North is a name it should know — the formula returns the #NAME? error." },
          { type: "p", text: "Numbers are the opposite: **no quotes**. `=IF(B2>=40, \"Pass\", \"Fail\")` compares against the number 40. Writing `\"40\"` compares against the text \"40\", and although Excel often still matches it, mixing text and number comparisons invites subtle bugs. The results \"Pass\" and \"Fail\" are in quotes because they are text you want to display." },
          {
            type: "grid",
            title: "Stock reorder list",
            rows: [
              ["Product", "Stock", "Status"],
              ["Notebooks", "120", "=IF(B2<=10, \"Reorder\", \"OK\")"],
              ["Pens", "8", "=IF(B3<=10, \"Reorder\", \"OK\")"],
              ["Staplers", "45", "=IF(B4<=10, \"Reorder\", \"OK\")"],
              ["Printers", "3", "=IF(B5<=10, \"Reorder\", \"OK\")"]
            ]
          },
          { type: "p", text: "Pens and Printers drop to Reorder; the others are OK. The storekeeper now knows exactly which lines need a purchase order, without reading 500 rows." }
        ]
      },
      {
        heading: "More real decisions: bonus eligibility",
        blocks: [
          { type: "p", text: "Sales teams get a bonus when their monthly sales cross a target. The condition tests a number, and both results are text." },
          {
            type: "grid",
            title: "Bonus eligibility, target ₹50,000",
            rows: [
              ["Employee", "Sales (₹)", "Eligible?"],
              ["Arjun Mehta", "52000", "=IF(B2>=50000, \"Bonus\", \"No bonus\")"],
              ["Priya Nair", "48000", "=IF(B3>=50000, \"Bonus\", \"No bonus\")"],
              ["Rahul Verma", "61000", "=IF(B4>=50000, \"Bonus\", \"No bonus\")"],
              ["Sneha Iyer", "75000", "=IF(B5>=50000, \"Bonus\", \"No bonus\")"]
            ]
          },
          { type: "p", text: "Arjun and Sneha cross the line, Priya misses it by 2,000. If the rules later change — say the target becomes 45,000 — you edit one formula and refill the column, or better, test against a cell that holds the target so the sheet has a single place to change it." }
        ]
      },
      {
        heading: "One IF inside another (a first taste)",
        blocks: [
          { type: "p", text: "Sometimes two outcomes are not enough. A grading sheet needs Distinction for 75 and above, Pass for 40 to 74, and Fail below that. You can put a second IF in place of the first formula's false result:" },
          {
            type: "code",
            lang: "excel",
            text: "=IF(B2>=75, \"Distinction\", IF(B2>=40, \"Pass\", \"Fail\"))"
          },
          { type: "p", text: "Excel reads this from the outside in: is B2 at least 75? If yes, Distinction. If not, run the inner IF: is B2 at least 40? If yes, Pass, otherwise Fail. The order matters — test the top band first, because a value of 80 would also satisfy \"at least 40\"." },
          { type: "warning", text: "Two levels of nesting like this is fine. Stacking five or six IFs makes formulas unreadable and fragile. Later levels of this course teach cleaner tools for many-way decisions (IFS, and lookup tables), so do not build towers of IFs out of habit." }
        ]
      }
    ],
    mistakes: [
      {
        wrong: "=IF(B2>40, \"Pass\", \"Fail\") for a pass mark of 40",
        right: "=IF(B2>=40, \"Pass\", \"Fail\")",
        why: "With >, a student scoring exactly 40 fails. >= includes the boundary, which is what \"pass at 40\" means."
      },
      {
        wrong: "=IF(A2=North, \"Home\", \"Away\")",
        right: "=IF(A2=\"North\", \"Home\", \"Away\")",
        why: "Text criteria need double quotes. Without them Excel looks for a defined name called North and returns #NAME?."
      },
      {
        wrong: "=IF(B2>=40, \"Pass\")",
        right: "=IF(B2>=40, \"Pass\", \"Fail\")",
        why: "The false result is not optional in practice. If the test fails with no third argument, the cell shows FALSE — confusing in a report."
      },
      {
        wrong: "Writing the grade formula with the lowest band first: =IF(B2>=40, \"Pass\", IF(B2>=75, \"Distinction\", \"Fail\"))",
        right: "=IF(B2>=75, \"Distinction\", IF(B2>=40, \"Pass\", \"Fail\"))",
        why: "A score of 80 satisfies \"at least 40\" too, so it would be labelled Pass and never reach the Distinction test. Test the top band first."
      }
    ],
    shortcuts: [],
    exercises: [
      {
        task: "Open the logical-functions workbook, Exercise sheet. Column B holds test marks. In C2 write a Pass/Fail formula with a pass mark of 40 and fill it down.",
        hint: "40 must pass, so the operator needs the boundary included.",
        solution: "C2: =IF(B2>=40, \"Pass\", \"Fail\"), filled down to the last student."
      },
      {
        task: "In column D of the same sheet, flag any employee who has NOT completed training. Column E holds Yes or No for completion.",
        hint: "\"Not equal\" is two characters and works on text in quotes.",
        solution: "D2: =IF(E2<>\"Yes\", \"Pending\", \"Done\"), filled down."
      },
      {
        task: "In column F, give a Distinction for marks 75 and above, otherwise reuse the Pass/Fail logic from exercise 1.",
        hint: "Put a second IF where the false result goes, testing the top band first.",
        solution: "F2: =IF(B2>=75, \"Distinction\", IF(B2>=40, \"Pass\", \"Fail\")), filled down."
      },
      {
        task: "In column G, mark stock for reorder when quantity is 10 or fewer. Column H holds the stock counts. Compare your results with the Solution sheet.",
        hint: "10 itself must be flagged.",
        solution: "G2: =IF(H2<=10, \"Reorder\", \"OK\"), filled down."
      }
    ],
    quiz: [
      {
        q: "Which formula writes Pass when marks in B2 are 40 or more, and Fail otherwise?",
        options: ["=IF(B2>=40, \"Pass\", \"Fail\")", "=IF(B2>40, \"Pass\", \"Fail\")", "=IF(B2<=40, \"Pass\", \"Fail\")", "=IF(B2=40, \"Pass\")"],
        answer: 0,
        explanation: "The pass mark itself must pass, so the test is >=. The plain > would fail a student on exactly 40."
      },
      {
        q: "Which operator means \"not equal to\" in Excel?",
        options: ["<>", "!=", "=~", "><"],
        answer: 0,
        explanation: "Excel uses <> for \"is different from\". The != style comes from programming languages and does not work in Excel."
      },
      {
        q: "In =IF(A2=\"North\", \"Home\", \"Other\"), why are the quotes needed?",
        options: ["North is text being compared and text criteria need quotes", "Quotes make the formula faster", "Quotes convert North to a number", "Excel requires quotes around all arguments"],
        answer: 0,
        explanation: "Text values inside formulas are written in double quotes. Without them Excel looks for a named range called North and shows #NAME?."
      },
      {
        q: "What does =IF(B2>=40, \"Pass\") show when B2 holds 20?",
        options: ["FALSE", "Fail", "Pass", "An error"],
        answer: 0,
        explanation: "With no value_if_false argument, IF displays FALSE when the test fails. Always supply the third argument."
      },
      {
        q: "In =IF(B2>=75, \"Distinction\", IF(B2>=40, \"Pass\", \"Fail\")), what does a mark of 68 produce?",
        options: ["Pass", "Distinction", "Fail", "An error"],
        answer: 0,
        explanation: "68 fails the first test (not >= 75), so the inner IF runs: 68 is >= 40, giving Pass."
      }
    ],
    summary: [
      "IF has three required parts: the condition, the true result and the false result.",
      "The six operators are =, <>, >, <, >= and <= — choose >= or <= when the boundary belongs to one side.",
      "Text criteria and text results go in double quotes; numbers stay unquoted.",
      "IF drives real office work: Pass/Fail lists, bonus eligibility and stock reorder flags.",
      "One nested IF handles a third outcome, with the top band tested first; deep nesting is covered properly in Level 6."
    ],
    downloads: [
      {
        file: "assets/files/level2/logical-functions.xlsx",
        label: "Logical functions workbook",
        note: "Exercise and Solution sheets for IF, AND, OR, IFS and IFERROR"
      }
    ],
    keywords: ["IF function", "logical test", "comparison operators", "not equal", "text criteria", "Pass Fail formula", "nested IF", "stock reorder"]
  },
  {
    id: "l2-and-or",
    title: "AND, OR and IFS",
    difficulty: "Beginner",
    duration: 14,
    version: "Excel 2019+",
    objectives: [
      "Explain that AND and OR test several conditions and return TRUE or FALSE",
      "Combine AND and OR inside IF to test two or more rules at once",
      "Use the IFS function to replace several nested IFs in one readable formula",
      "Build the nested-IF alternative that works in older Excel versions"
    ],
    intro: "Real decisions rarely depend on one condition. A bonus needs the sales target met **and** attendance above 90 percent. A customer gets free delivery from the North zone **or** on orders above 5,000. AND and OR are the functions that join conditions together.\n\nAt the end of the lesson you will meet IFS, a newer function that replaces towers of nested IFs. IFS needs Excel 2019 or Microsoft 365 — if your office still runs Excel 2016, the nested-IF alternative shown here does the same job.",
    sections: [
      {
        heading: "AND and OR: two functions that answer yes or no",
        blocks: [
          {
            type: "formula",
            name: "AND",
            syntax: "=AND(logical1, [logical2], ...)",
            example: "=AND(B2>=50000, C2=\"Yes\")",
            explanation: "Returns TRUE only when every condition inside is true. One false condition makes the whole AND false."
          },
          {
            type: "formula",
            name: "OR",
            syntax: "=OR(logical1, [logical2], ...)",
            example: "=OR(A2=\"North\", E2>5000)",
            explanation: "Returns TRUE when at least one condition inside is true. It only returns FALSE when all of them are false."
          },
          { type: "p", text: "On their own, these functions show the words TRUE or FALSE in the cell. Try it: `=AND(2>1, 3>2)` shows TRUE; `=AND(2>1, 3<2)` shows FALSE; `=OR(3<2, 4<3)` shows FALSE. Handy for learning, but in real work you almost never leave them standing alone — you put them inside IF." },
          {
            type: "table",
            headers: ["Condition A", "Condition B", "AND(A, B)", "OR(A, B)"],
            rows: [
              ["TRUE", "TRUE", "TRUE", "TRUE"],
              ["TRUE", "FALSE", "FALSE", "TRUE"],
              ["FALSE", "TRUE", "FALSE", "TRUE"],
              ["FALSE", "FALSE", "FALSE", "FALSE"]
            ]
          },
          { type: "p", text: "AND is strict: every rule must pass. OR is generous: any one rule is enough. Say this sentence out loud when you choose between them — \"the customer qualifies if the region is North **and** the order is above 5,000\" means AND; \"if the region is North **or** the order is above 5,000\" means OR." }
        ]
      },
      {
        heading: "Putting AND inside IF: the bonus rule",
        blocks: [
          { type: "p", text: "Sunrise Traders pays a bonus when an employee both hits the ₹50,000 sales target **and** completes the training. Two conditions, so AND goes inside IF as the condition part:" },
          {
            type: "grid",
            title: "Bonus with two conditions",
            rows: [
              ["Employee", "Sales (₹)", "Training done", "Result"],
              ["Arjun Mehta", "52000", "Yes", "=IF(AND(B2>=50000, C2=\"Yes\"), \"Bonus\", \"No bonus\")"],
              ["Priya Nair", "48000", "Yes", "=IF(AND(B3>=50000, C3=\"Yes\"), \"Bonus\", \"No bonus\")"],
              ["Rahul Verma", "61000", "No", "=IF(AND(B4>=50000, C4=\"Yes\"), \"Bonus\", \"No bonus\")"],
              ["Sneha Iyer", "75000", "Yes", "=IF(AND(B5>=50000, C5=\"Yes\"), \"Bonus\", \"No bonus\")"]
            ]
          },
          { type: "p", text: "Read each row. Arjun passes both rules — Bonus. Priya misses the target — No bonus. Rahul beats the target but skipped training — No bonus, because AND demands both. Sneha passes both — Bonus. Compare with lesson 4, where Rahul earned a bonus on sales alone; the second rule changed the decision." },
          { type: "p", text: "The same pattern handles free delivery: `=IF(AND(B2>=5000, A2=\"North\"), \"Free delivery\", \"Standard charges\")`." }
        ]
      },
      {
        heading: "OR inside IF: multiple ways to qualify",
        blocks: [
          { type: "p", text: "A customer gets priority service if they are in the North zone **or** their monthly spend is above 20,000. Either one is enough:" },
          {
            type: "grid",
            title: "Priority service with OR",
            rows: [
              ["Customer region", "Spend (₹)", "Service level"],
              ["North", "8000", "=IF(OR(A2=\"North\", B2>20000), \"Priority\", \"Standard\")"],
              ["South", "25000", "=IF(OR(A3=\"North\", B3>20000), \"Priority\", \"Standard\")"],
              ["East", "12000", "=IF(OR(A4=\"North\", B4>20000), \"Priority\", \"Standard\")"],
              ["West", "22000", "=IF(OR(A5=\"North\", B5>20000), \"Priority\", \"Standard\")"]
            ]
          },
          { type: "p", text: "Rows 2, 3 and 5 qualify — the North region in row 2, the big spend in rows 3 and 5. Row 4 has neither, so it stays Standard. With OR, one satisfied condition carries the whole test." },
          { type: "warning", text: "AND and OR can take many conditions, and they can be mixed: =IF(OR(AND(A2=\"North\", B2>5000), C2=\"VIP\"), ...) tests two sub-conditions then the VIP flag. If a formula grows past two levels of mixing, step back and simplify — readability is part of correctness." }
        ]
      },
      {
        heading: "IFS: many conditions, one function (Excel 2019+)",
        blocks: [
          { type: "p", text: "Nested IFs grow ugly fast. IFS writes the same logic as a clean list of pairs: each condition followed by its result." },
          {
            type: "formula",
            name: "IFS",
            syntax: "=IFS(logical_test1, value_if_true1, logical_test2, value_if_true2, ...)",
            example: "=IFS(B2>=75, \"A grade\", B2>=40, \"B grade\", TRUE, \"C grade\")",
            explanation: "Excel checks the tests in order and returns the result of the first one that is true. The final TRUE acts as \"everything else\"."
          },
          {
            type: "code",
            lang: "excel",
            text: "IFS version (Excel 2019+):\n=IFS(B2>=75, \"A grade\", B2>=40, \"B grade\", TRUE, \"C grade\")\n\nNested IF version (all Excel versions):\n=IF(B2>=75, \"A grade\", IF(B2>=40, \"B grade\", \"C grade\"))"
          },
          { type: "p", text: "Both formulas grade 80 as A, 60 as B and 20 as C. The IFS version is easier to read and easier to extend — add another pair to create a fourth grade. Note the last pair uses the condition TRUE, which is always satisfied, so it catches everything the earlier tests missed." },
          { type: "warning", text: "If no test in IFS is true and you did not include a final TRUE pair, IFS returns the #N/A error. Always finish with TRUE and a default result. If you are on Excel 2016 or older, use the nested IF version instead — IFS simply does not exist there." }
        ]
      }
    ],
    mistakes: [
      {
        wrong: "=IF(B2>=50000 AND C2=\"Yes\", \"Bonus\", \"No bonus\")",
        right: "=IF(AND(B2>=50000, C2=\"Yes\"), \"Bonus\", \"No bonus\")",
        why: "AND is a function, not a word to drop between conditions. It needs its own brackets, and the whole AND(...) sits inside IF as the single condition."
      },
      {
        wrong: "Using OR when the rule needs every condition",
        right: "Use AND when all conditions must pass, OR when any one is enough",
        why: "=IF(OR(B2>=50000, C2=\"Yes\"), ...) gives Rahul a bonus for hitting the target despite skipping training. The business rule says both."
      },
      {
        wrong: "=IFS(B2>=40, \"B grade\", B2>=75, \"A grade\")",
        right: "Order tests from highest to lowest: =IFS(B2>=75, \"A grade\", B2>=40, \"B grade\", TRUE, \"C grade\")",
        why: "IFS stops at the first true test. An 80 meets \"at least 40\" first and gets the wrong grade."
      },
      {
        wrong: "Leaving no catch-all in IFS: =IFS(B2>=75, \"A\", B2>=40, \"B\")",
        right: "=IFS(B2>=75, \"A\", B2>=40, \"B\", TRUE, \"C\")",
        why: "A mark of 20 satisfies neither test, so IFS has no result to return and shows #N/A."
      }
    ],
    shortcuts: [],
    exercises: [
      {
        task: "Open the logical-functions workbook, Exercise sheet. Employees get a bonus when sales (column B) are at least 50000 AND training (column C) shows Yes. Write the formula in D2 and fill down.",
        hint: "AND(...) with its own brackets becomes the condition of IF.",
        solution: "D2: =IF(AND(B2>=50000, C2=\"Yes\"), \"Bonus\", \"No bonus\"), filled down."
      },
      {
        task: "Customers get free delivery when the order (column E) is above 5000 OR the region (column F) is North. Write the formula in G2 and fill down.",
        hint: "Either condition alone qualifies.",
        solution: "G2: =IF(OR(E2>5000, F2=\"North\"), \"Free delivery\", \"Standard charges\"), filled down."
      },
      {
        task: "On your own copy of the sheet, grade the marks in column B with IFS: A grade at 75 or more, B grade at 40 or more, C grade below that. Then rebuild the same result with a nested IF and compare the two formulas.",
        hint: "Finish IFS with TRUE and the default result; test the top band first in both versions.",
        solution: "IFS: =IFS(B2>=75, \"A grade\", B2>=40, \"B grade\", TRUE, \"C grade\"). Nested IF: =IF(B2>=75, \"A grade\", IF(B2>=40, \"B grade\", \"C grade\")). Both give identical results."
      },
      {
        task: "Check the Solution sheet of the workbook and find the row where a sales target was met but the bonus was still refused. Explain in one sentence why.",
        hint: "The refused row had a second condition that failed.",
        solution: "AND requires every condition to be true; the employee met the sales target but had not completed training, so AND returned FALSE."
      }
    ],
    quiz: [
      {
        q: "When does AND(condition1, condition2) return TRUE?",
        options: ["Only when both conditions are true", "When at least one is true", "When neither is true", "Always"],
        answer: 0,
        explanation: "AND is strict: every condition inside must be true. One false condition turns the whole result FALSE."
      },
      {
        q: "A customer qualifies for free delivery from North OR on orders above 5000. Which formula is correct?",
        options: ["=IF(OR(B2>5000, A2=\"North\"), \"Free\", \"Standard\")", "=IF(AND(B2>5000, A2=\"North\"), \"Free\", \"Standard\")", "=IF(B2>5000 AND A2=\"North\", \"Free\", \"Standard\")", "=OR(IF(B2>5000, \"Free\", \"Standard\"))"],
        answer: 0,
        explanation: "Either rule alone qualifies, so OR goes inside IF. The AND version would demand both at once."
      },
      {
        q: "What does IFS return if none of its tests is true and there is no final TRUE pair?",
        options: ["The #N/A error", "FALSE", "Zero", "The last value again"],
        answer: 0,
        explanation: "IFS needs one true test to return something. With no match it shows #N/A — always finish with TRUE and a default result."
      },
      {
        q: "Which version of Excel first included the IFS function?",
        options: ["Excel 2019", "Excel 2013", "Excel 2010", "Excel 2007"],
        answer: 0,
        explanation: "IFS arrived in Excel 2019 and Microsoft 365. In older versions you write the same logic with nested IFs."
      },
      {
        q: "In grading, why must IFS test B2>=75 before B2>=40?",
        options: ["Every score of 80 also satisfies \"at least 40\", so the first true test would win", "IFS requires descending numbers", "Tests must be alphabetical", "It makes no difference"],
        answer: 0,
        explanation: "IFS stops at the first true test. Testing the low band first would label a top scorer with the lower grade."
      }
    ],
    summary: [
      "AND returns TRUE only when every condition is true; OR returns TRUE when any one condition is true.",
      "In real work AND and OR almost always live inside IF as the condition part: =IF(AND(...), value, value).",
      "AND is for rules where every condition must pass; OR is for rules where one pass is enough.",
      "IFS lists condition-and-result pairs and stops at the first true test — always end with TRUE plus a default result.",
      "IFS needs Excel 2019 or Microsoft 365; nested IFs do the same job in older versions."
    ],
    downloads: [
      {
        file: "assets/files/level2/logical-functions.xlsx",
        label: "Logical functions workbook",
        note: "Exercise and Solution sheets for IF, AND, OR, IFS and IFERROR"
      }
    ],
    keywords: ["AND function", "OR function", "IFS function", "multiple conditions", "TRUE FALSE", "bonus eligibility", "nested IF alternative"]
  },
  {
    id: "l2-iferror",
    title: "IFERROR: Handling Errors Gracefully",
    difficulty: "Beginner",
    duration: 12,
    version: "",
    objectives: [
      "Recognise the five most common Excel error values and what each means",
      "Wrap a risky formula in IFERROR to show a clean fallback instead of an error",
      "Decide when an error should be fixed at the source instead of hidden",
      "Keep reports free of #DIV/0! without hiding real mistakes"
    ],
    intro: "Nothing makes a report look unfinished like a column of #DIV/0! and #N/A. The calculations may be right — but the person reading the report sees red error text and loses trust in the numbers. IFERROR is the polite face of a formula: it catches errors and shows something sensible instead.\n\nThere is a discipline to it, though. An error is often information — it tells you the data behind the formula is broken. This lesson teaches you both halves: fix the error when you can, and use IFERROR when a clean fallback is genuinely the right answer.",
    sections: [
      {
        heading: "The errors you will actually meet",
        blocks: [
          {
            type: "table",
            headers: ["Error", "Cause", "Typical office situation"],
            rows: [
              ["#DIV/0!", "Dividing by zero or by an empty cell", "A day with no sales divided into an amount"],
              ["#N/A", "The value you searched for was not found", "A product code that does not exist in the price list"],
              ["#VALUE!", "The wrong kind of input — usually text where a number is expected", "A quantity column that contains \"two\" instead of 2"],
              ["#REF!", "A formula points to cells that no longer exist", "Someone deleted a column your formula was using"],
              ["#NAME?", "A name Excel does not recognise — usually a misspelled function", "=SUMME(C2:C4) instead of =SUM(C2:C4)"]
            ]
          },
          { type: "p", text: "Two of these are worth extra attention. #DIV/0! appears whenever the bottom of a division is 0 or empty — extremely common in daily-rate and percentage calculations. #N/A appears in lookup formulas, which you will meet later this level; it means \"the thing I was asked to find is not in the list\"." },
          { type: "tip", text: "Click a cell showing an error and a small diamond icon appears beside it. Clicking the icon offers **Show Calculation Steps** and **Help on this error** — Excel explains exactly which part failed. Use it before reaching for IFERROR." }
        ]
      },
      {
        heading: "The IFERROR pattern",
        blocks: [
          {
            type: "formula",
            name: "IFERROR",
            syntax: "=IFERROR(value, value_if_error)",
            example: "=IFERROR(B2/C2, 0)",
            explanation: "Excel calculates B2/C2 first. If it works, the cell shows the result; if it produces any error, the cell shows the fallback — here 0 — instead."
          },
          { type: "p", text: "The formula takes two arguments: the calculation you actually want, and what to show if that calculation errors. The fallback can be a number, text in quotes, an empty-looking pair of quotes `\"\"`, or even another formula." },
          {
            type: "grid",
            title: "Daily average per order — before and after IFERROR",
            rows: [
              ["Date", "Sales (₹)", "Orders", "Before IFERROR", "With IFERROR"],
              ["02-Feb", "12500", "5", "=B2/C2", "=IFERROR(B2/C2, 0)"],
              ["03-Feb", "9800", "0", "=B3/C3", "=IFERROR(B3/C3, 0)"],
              ["04-Feb", "14300", "4", "=B4/C4", "=IFERROR(B4/C4, 0)"]
            ]
          },
          { type: "p", text: "On 03-Feb the shop recorded no orders, so `=B3/C3` shows #DIV/0! while `=IFERROR(B3/C3, 0)` shows 0. The report stays readable, and the zero communicates \"no orders that day\"." },
          { type: "p", text: "A text fallback reads even better in a summary sent to a manager: `=IFERROR(B2/C2, \"No orders\")`. Choose the fallback that makes sense to a human reader, not just something that hides the red." }
        ]
      },
      {
        heading: "Fix first, wrap second",
        blocks: [
          { type: "p", text: "IFERROR does not know why the formula failed. It treats a genuine problem and a data-entry mistake exactly the same, and quietly replaces both with your fallback. That is its strength and its danger." },
          {
            type: "ul",
            items: [
              "**If the error means missing data** (no orders on Sunday), a fallback like 0 or \"—\" is the right answer. Wrap it.",
              "**If the error means your formula is wrong** — a misspelled function name, a deleted column, a broken range — IFERROR hides the bug. Fix the formula first.",
              "**Build and test unwrapped.** Write the formula, confirm it returns correct numbers, and only then wrap it in IFERROR for delivery.",
              "**Prefer fixing the source.** A #VALUE! from a quantity typed as \"two\" is cured by correcting the data, not by hiding the error."
            ]
          },
          { type: "warning", text: "A sheet where every formula is wrapped in IFERROR can look clean while producing nonsense. If you inherit a report full of IFERROR formulas, temporarily remove a few wrappers and look at what they are hiding before you trust the numbers." }
        ]
      }
    ],
    mistakes: [
      {
        wrong: "Wrapping a broken formula straight into IFERROR: =IFERROR(SUMME(B2:B10), 0)",
        right: "Fix the formula first, then wrap: =IFERROR(SUM(B2:B10), 0)",
        why: "The misspelled name is a bug in your work, not a data problem. IFERROR would hide it and the cell would quietly show 0."
      },
      {
        wrong: "=IFERROR(B2/C2, No orders)",
        right: "=IFERROR(B2/C2, \"No orders\")",
        why: "Text fallbacks need double quotes. Without them Excel looks for a defined name and shows #NAME? — an error inside your error handler."
      },
      {
        wrong: "Using IFERROR on formulas that never produce errors, like =IFERROR(B2+C2, 0)",
        right: "Keep IFERROR for genuinely risky calculations: division and lookups",
        why: "Adding two numbers cannot fail unless the inputs are already errors. Extra wrappers hide future problems and make every formula slower to read."
      }
    ],
    shortcuts: [],
    exercises: [
      {
        task: "Open the logical-functions workbook, Exercise sheet. Column B holds sales, column C holds the number of orders, and some days have no orders. In D2 build sales per order and fill down. You will see #DIV/0! on some rows.",
        hint: "Simple division first: =B2/C2.",
        solution: "D2: =B2/C2, filled down. Days with 0 orders show #DIV/0!."
      },
      {
        task: "Now wrap the same division so error rows show 0 instead, in column E.",
        hint: "The fallback goes in as the second argument.",
        solution: "E2: =IFERROR(B2/C2, 0), filled down. The report is now error-free."
      },
      {
        task: "Change the fallback so days without orders display the text \"No orders\" instead of 0, in column F.",
        hint: "Text fallbacks need double quotes.",
        solution: "F2: =IFERROR(B2/C2, \"No orders\"), filled down."
      },
      {
        task: "The Solution sheet of the workbook contains one deliberately broken formula hidden inside an IFERROR. Find it and explain what the unwrapped formula does wrong.",
        hint: "Temporarily delete the IFERROR wrapper and press Enter to see the true error.",
        solution: "Removing the wrapper reveals the underlying error — a wrong reference or misspelled name — which IFERROR had been silently covering up."
      }
    ],
    quiz: [
      {
        q: "What does the #DIV/0! error mean?",
        options: ["A formula tried to divide by zero or by an empty cell", "A function name is misspelled", "A value was not found in a list", "A referenced column was deleted"],
        answer: 0,
        explanation: "Division by zero or by a blank cell produces #DIV/0!. It is the most common error in daily-rate and per-order calculations."
      },
      {
        q: "Which error appears when a lookup cannot find the value it was asked to find?",
        options: ["#N/A", "#DIV/0!", "#VALUE!", "#REF!"],
        answer: 0,
        explanation: "#N/A means \"not available\" — the search value is not in the list. You will meet it properly in the VLOOKUP lesson."
      },
      {
        q: "What does =IFERROR(B2/C2, 0) show when C2 is empty?",
        options: ["0", "#DIV/0!", "Empty", "C2"],
        answer: 0,
        explanation: "B2/C2 would give #DIV/0! because an empty divisor counts as zero. IFERROR catches it and returns the fallback 0."
      },
      {
        q: "Why is it wise to test a formula before wrapping it in IFERROR?",
        options: ["IFERROR hides bugs in the formula itself, so you should confirm it works first", "IFERROR slows the workbook down permanently", "IFERROR deletes the original formula", "Wrapped formulas cannot be copied"],
        answer: 0,
        explanation: "IFERROR cannot tell a broken formula from missing data. Verify the unwrapped formula first, then add the wrapper for delivery."
      },
      {
        q: "Which IFERROR fallback correctly shows the text \"Not found\"?",
        options: ["=IFERROR(A2/B2, \"Not found\")", "=IFERROR(A2/B2, Not found)", "=IFERROR(A2/B2, 'Not found')", "=IFERROR(\"Not found\", A2/B2)"],
        answer: 0,
        explanation: "Text in formulas needs double quotes, and the calculation comes first. The last option swaps the two arguments."
      }
    ],
    summary: [
      "The five everyday errors: #DIV/0! (divide by zero), #N/A (not found), #VALUE! (wrong input type), #REF! (deleted cells), #NAME? (unknown name).",
      "IFERROR(value, fallback) shows the fallback instead of any error — perfect for division and lookups.",
      "Fallbacks can be 0, \"\" for blank-looking cells, or friendly text like \"No orders\".",
      "Fix the root cause when the error reflects a real problem; use IFERROR when a clean fallback is the business answer.",
      "Test formulas unwrapped, then add IFERROR last — it hides your own bugs as easily as data gaps."
    ],
    downloads: [
      {
        file: "assets/files/level2/logical-functions.xlsx",
        label: "Logical functions workbook",
        note: "Exercise and Solution sheets for IF, AND, OR, IFS and IFERROR"
      }
    ],
    keywords: ["IFERROR", "error handling", "DIV/0", "N/A error", "VALUE error", "REF error", "NAME error", "fallback value", "clean reports"]
  },
  {
    id: "l2-sumif-countif",
    title: "SUMIF and COUNTIF",
    difficulty: "Intermediate",
    duration: 16,
    version: "",
    objectives: [
      "Write =SUMIF(range, criteria, sum_range) and =COUNTIF(range, criteria)",
      "Build criteria as numbers, text, wildcards and cell references",
      "Line up sum_range with the criteria range correctly",
      "Answer real sales questions: totals per region, counts per product, orders above a limit"
    ],
    intro: "SUM tells you the total of everything. But office questions are rarely that broad: how much did the **North** region sell? How many **Notebook** orders came in? What is the total of orders **above 5,000**? SUMIF and COUNTIF answer questions with a condition attached — add up only the rows that match.\n\nThese two functions are the everyday workhorses of sales reports, stock lists and invoice summaries. Master them well here, because SUMIFS, COUNTIFS and AVERAGEIFS in the next lesson are their direct big siblings.",
    sections: [
      {
        heading: "COUNTIF: how many rows match?",
        blocks: [
          {
            type: "formula",
            name: "COUNTIF",
            syntax: "=COUNTIF(range, criteria)",
            example: "=COUNTIF(B2:B10, \"Notebooks\")",
            explanation: "Counts the cells in B2:B10 that are equal to Notebooks. Two arguments only: where to look, and what to look for."
          },
          {
            type: "grid",
            title: "Sunrise Traders sales record",
            rows: [
              ["Region", "Product", "Units", "Price (₹)", "Amount (₹)"],
              ["North", "Notebooks", "40", "45", "=C2*D2"],
              ["South", "Pens", "100", "12", "=C3*D3"],
              ["East", "Staplers", "15", "130", "=C4*D4"],
              ["North", "Printers", "2", "8500", "=C5*D5"],
              ["West", "Laptops", "1", "42000", "=C6*D6"],
              ["South", "Notebooks", "60", "45", "=C7*D7"],
              ["East", "Pens", "80", "12", "=C8*D8"],
              ["North", "Staplers", "10", "130", "=C9*D9"],
              ["West", "Notebooks", "25", "45", "=C10*D10"]
            ]
          },
          { type: "p", text: "`=COUNTIF(B2:B10, \"Notebooks\")` returns 3 — Notebooks appear in rows 2, 7 and 10. `=COUNTIF(A2:A10, \"North\")` returns 3 as well. Text criteria go in double quotes, exactly as they did inside IF in lesson 4." }
        ]
      },
      {
        heading: "The four shapes of criteria",
        blocks: [
          { type: "p", text: "The criteria argument is more flexible than it first looks. It can be plain text, a number, a comparison, or a cell reference. These four shapes cover almost every report question." },
          {
            type: "table",
            headers: ["Criteria shape", "Example", "Question it answers"],
            rows: [
              ["Text", "\"North\"", "How many rows are the North region?"],
              ["Number", "42", "How many rows show exactly 42?"],
              ["Comparison", "\">5000\"", "How many rows are greater than 5000?"],
              ["Cell reference", "G2 (which holds \"South\")", "How many rows match whatever is typed in G2?"],
              ["Wildcard", "\"Pen*\"", "How many rows start with Pen?"]
            ]
          },
          { type: "p", text: "Comparisons go **inside the quotes**: `\">5000\"` is one piece of text that Excel interprets as a rule. And comparisons can join a cell reference with an ampersand, so the threshold lives in a cell you can change without editing the formula: `\">\"&G3` builds \">5000\" from whatever G3 holds." },
          {
            type: "grid",
            title: "COUNTIF with different criteria",
            rows: [
              ["Question", "Formula", "Result"],
              ["North rows", "=COUNTIF(A2:A10, \"North\")", "3"],
              ["Sales above 5000", "=COUNTIF(E2:E10, \">5000\")", "2"],
              ["Notebook orders", "=COUNTIF(B2:B10, \"Notebooks\")", "3"],
              ["Products starting Pen", "=COUNTIF(B2:B10, \"Pen*\")", "2"],
              ["South rows (from G2)", "=COUNTIF(A2:A10, G2)", "2"]
            ]
          },
          { type: "p", text: "The cell-reference form is the professional habit: put South in `G2` and the formula counts South. Change G2 to East and the same formula counts East. The report becomes interactive without touching the formula." }
        ]
      },
      {
        heading: "SUMIF: add up the matching rows",
        blocks: [
          {
            type: "formula",
            name: "SUMIF",
            syntax: "=SUMIF(range, criteria, [sum_range])",
            example: "=SUMIF(A2:A10, \"North\", E2:E10)",
            explanation: "Looks in A2:A10 for North, and adds the Amount values from the matching rows in E2:E10. The ranges run parallel, row for row."
          },
          { type: "p", text: "Read the three arguments as a sentence: \"look **here** for **this**, and add up **that**.\" The first range is where the condition lives, the second is the condition, the third is what gets added. When you leave out the third argument, SUMIF adds the criteria range itself — correct when you are summing the same numbers you are testing, like `=SUMIF(E2:E10, \">5000\")` for the total of big orders." },
          {
            type: "grid",
            title: "SUMIF on the sales record",
            rows: [
              ["Question", "Formula", "Result"],
              ["North total", "=SUMIF(A2:A10, \"North\", E2:E10)", "20100"],
              ["Notebook revenue", "=SUMIF(B2:B10, \"Notebooks\", E2:E10)", "5625"],
              ["Big orders only", "=SUMIF(E2:E10, \">5000\")", "59000"],
              ["Units for Pens", "=SUMIF(B2:B10, \"Pens\", C2:C10)", "180"]
            ]
          },
          { type: "p", text: "Check the North total by hand: rows 2, 5 and 9 give 1800 + 17000 + 1300 = 20100. Notebook revenue is 1800 + 2700 + 1125 = 5625. Big orders are the Printer and Laptop rows: 17000 + 42000 = 59000. Units for Pens: 100 + 80 = 180. Every result traces back to visible rows — that is how you should verify your own SUMIF formulas." }
        ]
      },
      {
        heading: "Wildcards: matching text patterns",
        blocks: [
          {
            type: "ul",
            items: [
              "**Asterisk ( * )** stands for any number of characters. `\"Pen*\"` matches Pens, Pencils, Pen holders.",
              "**Question mark ( ? )** stands for exactly one character. `\"North ?\"`... more practically, `\"N???h\"` matches North.",
              "**Wildcards work in both COUNTIF and SUMIF** — anywhere text criteria are allowed.",
              "**Tilde ( ~ ) escapes a wildcard** if you truly need to search for a literal asterisk: `\"~*\"`."
            ]
          },
          { type: "p", text: "In the product list, `\"Pen*\"` matches only Pens (2 orders), while `\"Pen\"` without the wildcard matches only the exact word Pen — which appears zero times here. Notebooks need no wildcard because the name is complete, but in real product lists names often come with sizes or codes attached (\"Notebook A4 Ruled\"), and `\"Notebook*\"` catches them all." },
          { type: "tip", text: "Wildcards also power the AVERAGEIF sibling: =AVERAGEIF(A2:A10, \"North\", E2:E10) gives the average North sale, 20100 divided by 3, which is 6700. Same argument order as SUMIF." }
        ]
      },
      {
        heading: "Keeping the ranges aligned",
        blocks: [
          { type: "p", text: "The criteria range and the sum range must start at the same row and run the same length. Excel reads them row by row: row 2 of the criteria range decides whether row 2 of the sum range is added." },
          {
            type: "ul",
            items: [
              "Correct: `=SUMIF(A2:A10, \"North\", E2:E10)` — both start at row 2, both cover 9 rows.",
              "Also correct: `=SUMIF(A2:A10, \"North\", E5:E13)` — Excel matches by position, adding the amounts 3 rows down. Legal but confusing; avoid it.",
              "Wrong: ranges of different lengths invite mistakes — Excel tolerates them, but the results stop meaning anything."
            ]
          },
          { type: "warning", text: "If a SUMIF returns a number that looks impossible, check the two ranges first — an off-by-one start row is the number-one cause of wrong conditional totals. The safest method: select both ranges from the same row and drag." }
        ]
      }
    ],
    mistakes: [
      {
        wrong: "=COUNTIF(B2:B10, Notebooks)",
        right: "=COUNTIF(B2:B10, \"Notebooks\")",
        why: "Text criteria need double quotes. Without them Excel looks for a defined name called Notebooks and shows #NAME?."
      },
      {
        wrong: "=SUMIF(A2:A10, \">5000\", E2:E10)",
        right: "=SUMIF(E2:E10, \">5000\", E2:E10) or simply =SUMIF(E2:E10, \">5000\")",
        why: "The comparison must run on the column being tested. Testing regions against >5000 matches nothing and returns 0."
      },
      {
        wrong: "Putting the comparison outside the quotes: =SUMIF(E2:E10, >\"5000\")",
        right: "=SUMIF(E2:E10, \">5000\")",
        why: "The whole rule — operator and number together — is one piece of text: \">5000\". Splitting them breaks the criteria."
      },
      {
        wrong: "Criteria range and sum range starting on different rows",
        right: "Start both ranges on the same row: =SUMIF(A2:A10, G2, E2:E10)",
        why: "Excel pairs the ranges position by position. An off-by-one start makes every row test its neighbour, and the total is silently wrong."
      }
    ],
    shortcuts: [],
    exercises: [
      {
        task: "Open the sumif-practice workbook, Exercise sheet. It holds a sales list with Region in column A, Product in B and Amount in E. In cell H2, total the revenue for the South region.",
        hint: "Look in A, match text in quotes, add from E.",
        solution: "H2: =SUMIF(A2:A10, \"South\", E2:E10), which returns 3900 (1200 + 2700)."
      },
      {
        task: "In cell H3, count how many orders were for Notebooks, and in H4 total their revenue.",
        hint: "One formula counts, the other sums the matching amounts.",
        solution: "H3: =COUNTIF(B2:B10, \"Notebooks\") returns 3. H4: =SUMIF(B2:B10, \"Notebooks\", E2:E10) returns 5625."
      },
      {
        task: "In cell H5, count the orders above 5000 and in H6 total their value, so the manager sees how much revenue comes from big tickets.",
        hint: "The comparison lives inside the quotes.",
        solution: "H5: =COUNTIF(E2:E10, \">5000\") returns 2. H6: =SUMIF(E2:E10, \">5000\") returns 59000."
      },
      {
        task: "Put a region name in cell G8 (try North, then change it to West). In H8 build a formula that totals revenue for whatever region is in G8. Compare with the Solution sheet.",
        hint: "The criteria argument can simply be the cell reference.",
        solution: "H8: =SUMIF(A2:A10, G8, E2:E10). With North it returns 20100; change G8 to West and it returns 43125."
      }
    ],
    quiz: [
      {
        q: "Which formula totals the Amounts (column E) for the North region listed in column A?",
        options: ["=SUMIF(A2:A10, \"North\", E2:E10)", "=SUMIF(E2:E10, \"North\", A2:A10)", "=SUMIF(A2:A10, E2:E10, \"North\")", "=SUM(A2:A10, \"North\")"],
        answer: 0,
        explanation: "The order is: criteria range, criteria, sum range. Look in A for North, add from E."
      },
      {
        q: "How do you write a \"greater than 5000\" criterion in COUNTIF?",
        options: ["\">5000\"", ">5000 without quotes", "\"5000>\"", "=\">5000\""],
        answer: 0,
        explanation: "The operator and number form one piece of text inside the quotes. Excel reads \">5000\" as a rule."
      },
      {
        q: "=COUNTIF(B2:B10, \"Pen*\") counts how many rows in the sales record?",
        options: ["2 — Pens matches, exact Pen does not exist", "0 — wildcards do not work in COUNTIF", "1 — only exact matches count", "10 — it matches everything"],
        answer: 0,
        explanation: "The asterisk means \"anything after Pen\", so both Pens rows match. Plain \"Pen\" would match no row at all."
      },
      {
        q: "What happens if you leave out the third argument in =SUMIF(E2:E10, \">5000\")?",
        options: ["SUMIF adds the criteria range itself", "Excel returns an error", "SUMIF returns 0", "SUMIF counts instead of sums"],
        answer: 0,
        explanation: "Without sum_range, SUMIF sums the same range it tests — perfect when you are adding the very numbers you are filtering."
      },
      {
        q: "Why is a cell-reference criterion like =COUNTIF(A2:A10, G2) useful in reports?",
        options: ["You can change the region in G2 and the formula follows, without editing the formula", "It is faster than typing text", "It works in older Excel versions only", "It ignores blank cells"],
        answer: 0,
        explanation: "The report becomes interactive: type a new value in G2 and the count updates. One formula serves every region."
      }
    ],
    summary: [
      "COUNTIF counts matching rows; SUMIF adds the matching rows' values: look here, for this, add that.",
      "Criteria come in four shapes: text in quotes, plain numbers, comparisons inside quotes like \">5000\", and cell references.",
      "Wildcards extend text criteria: * for any run of characters, ? for one character.",
      "The sum range must align with the criteria range row for row — off-by-one starts silently corrupt totals.",
      "AVERAGEIF works with the same argument pattern as SUMIF and averages the matches instead of totalling them."
    ],
    downloads: [
      {
        file: "assets/files/level2/sumif-practice.xlsx",
        label: "Conditional aggregation workbook",
        note: "Exercise and Solution sheets for SUMIF, COUNTIF, AVERAGEIF, SUMIFS, COUNTIFS and AVERAGEIFS"
      }
    ],
    keywords: ["SUMIF", "COUNTIF", "AVERAGEIF", "criteria", "wildcards", "cell reference criteria", "conditional sum", "conditional count", "region totals"]
  },
);
