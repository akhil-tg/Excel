// Excel Lab — Level 3: Data Management (part 2: lessons 6-10)
window.LEVEL_LESSONS = window.LEVEL_LESSONS || {};

(window.LEVEL_LESSONS[3] = window.LEVEL_LESSONS[3] || []).push(

  {
    id: "l3-remove-duplicates",
    title: "Removing Duplicates",
    difficulty: "Beginner",
    duration: 10,
    version: "",
    objectives: [
      "Find duplicates visually with conditional formatting before deleting anything",
      "Use Remove Duplicates and choose which columns define a duplicate",
      "Count duplicates with COUNTIF",
      "Build a safe backup habit before any deletion"
    ],
    intro: "Duplicate rows are silent poison for a business file. The customer list shows Arjun Mehta three times, so the mail merge sends him three emails, the COUNTA-based headcount is wrong, and a SUMIF over the data double-counts his orders.\n\nThis lesson gives you a safe routine: first find and understand the duplicates, then decide what \"duplicate\" really means for your table, and only then delete — with a backup taken first.",
    sections: [
      {
        heading: "First, look before you delete",
        blocks: [
          { type: "p", text: "Never delete duplicates straight away. First make them visible, because they teach you how the data went wrong. Conditional formatting is the fastest way (you learned it earlier in this level):" },
          { type: "steps", items: [
            "Select the column or columns where repeats matter, for example the Email column `B2:B500`.",
            "Home > Conditional Formatting > Highlight Cells Rules > Duplicate Values.",
            "Click OK — every repeated value lights up in red.",
            "Read the pattern: are these copy-paste accidents, or two customers who genuinely share a value?"
          ] },
          { type: "grid", title: "A customer list with Duplicate Values highlighting on the Email column",
            rows: [
              ["Customer", "Email", "Region"],
              ["Arjun Mehta", "arjun.mehta@sunrise.in", "North"],
              ["Arjun Mehta", "arjun.mehta@sunrise.in", "North"],
              ["Priya Nair", "priya.nair@sunrise.in", "South"],
              ["Rahul Verma", "rahul.verma@sunrise.in", "East"]
            ] },
          { type: "p", text: "In this mini-example, rows 2 and 3 highlight — the same customer entered twice. Rows 4 and 5 stay clean. Seeing the duplicates also tells you something important: are they **full-row duplicates** (every column identical, safe to delete outright) or **partial duplicates** (same name but one row has a newer address, so the \"newer\" row should win)?" },
          { type: "tip", text: "Switch the Duplicate Values dialog from Duplicate to Unique to highlight the opposite — the values that appear exactly once. That is a quick way to spot typos that almost match: a customer who appears once with a misspelled email is more dangerous than an exact repeat." }
        ]
      },
      {
        heading: "Counting duplicates with COUNTIF",
        blocks: [
          { type: "p", text: "Colouring shows you duplicates; a formula counts them. A helper column with COUNTIF tells you how many times each value appears in the whole list:" },
          { type: "formula", name: "COUNTIF",
            syntax: "=COUNTIF(range, criteria)",
            example: "=COUNTIF($B$2:$B$500, B2)",
            explanation: "Counts how many cells in B2:B500 are equal to B2. Anchored with $ so the range stays fixed while the formula fills down: a value appearing three times shows 3 on all three of its rows." },
          { type: "p", text: "Two uses in practice. To get the **total number of repeat rows** for a status report: filter the helper column to values greater than 1 and read `=SUBTOTAL(2, H2:H500)` (counts visible numbers), or simply subtract: `=COUNTA(B2:B500)-SUMPRODUCT((COUNTIF(B2:B500,B2:B500)=1)*1)` is overkill — a count of rows where the helper is 1 tells you how many unique entries exist. For day-to-day work, the helper column plus a filter is enough: filter to values above 1, inspect those rows, and decide row by row." },
          { type: "p", text: "The helper column also answers the partial-duplicate question: sort by the email column and the helper column side by side, and identical values with **different** other columns are easy to compare and merge by hand." }
        ]
      },
      {
        heading: "Remove Duplicates — and which columns define a duplicate",
        blocks: [
          { type: "p", text: "The deletion tool is Data > Remove Duplicates. The dialog lists every column with a tick box, and this is where the thinking happens: a row is a duplicate only in the columns you tick." },
          { type: "ul", items: [
            "**All columns ticked** — removes rows that are completely identical. This is the safe default for a log where every field matters.",
            "**Only Email ticked** — keeps the **first** occurrence of each email and deletes later rows, even if those rows contained newer information. Dangerous when the later row is the better one.",
            "**Name + Region ticked** — treats two rows as duplicates when those two fields match. Useful when one customer genuinely has multiple orders and you want one row per customer-region pair."
          ] },
          { type: "steps", items: [
            "Click any cell in the data (ideally inside a Table).",
            "Data > Remove Duplicates.",
            "Tick exactly the columns that define \"same customer/record\" for this task.",
            "Click OK and read the result message: Excel reports how many duplicate values were found and how many unique values remain."
          ] },
          { type: "warning", text: "Remove Duplicates keeps the first row and deletes later ones — the first row in the current order, not the best row. If some duplicates contain newer or more complete data, sort so the rows you want to keep come first, or clean up by hand after inspecting." }
        ]
      },
      {
        heading: "Backup first: the habit that saves jobs",
        blocks: [
          { type: "p", text: "Remove Duplicates is instant and permanent — there is no \"are you sure?\" preview of which rows will die. The professional routine costs thirty seconds:" },
          { type: "steps", items: [
            "Right-click the sheet tab > Move or Copy > tick Create a copy. Name the copy CustomerList-BACKUP.",
            "Work on the original sheet only.",
            "Remove duplicates, verify the result (compare row counts, spot-check highlighted values).",
            "Only when everything checks out, delete the backup sheet — or keep it until the next full save cycle if your data is critical."
          ] },
          { type: "p", text: "Ctrl + Z also undoes a Remove Duplicates immediately, but not after you have saved, closed and reopened the file. The backup sheet is the version that survives anything, including the moment your manager says \"actually, send me the version from this morning\"." },
          { type: "quote", text: "Deleting data is the one spreadsheet action you cannot reason your way out of later. A backup sheet is cheaper than any recovery." }
        ]
      }
    ],
    mistakes: [
      {
        wrong: "Running Remove Duplicates with only the Name column ticked on a customer list.",
        right: "Tick the columns that truly define a duplicate — usually email, phone or a customer ID, plus the fields that must match.",
        why: "Two different customers can share a name (or one can have two addresses). Over-ticking keeps too much; under-ticking deletes real records. Decide the definition first, then tick."
      },
      {
        wrong: "Removing duplicates on a live sheet that other formulas, pivots and charts point at, with no backup.",
        right: "Copy the sheet first (Move or Copy > Create a copy), then run Remove Duplicates on the original.",
        why: "The operation is instant and unpreviewable. When it goes wrong, Ctrl + Z only works until the next save — the backup sheet works forever."
      },
      {
        wrong: "Assuming the row that survives is the most recent or most complete one.",
        right: "Sort the data so the rows you want to keep come first, or handle conflicting duplicates manually before running the tool.",
        why: "Remove Duplicates keeps the first occurrence in the current order. A stale address entered on Monday can silently replace the correct one from Friday."
      }
    ],
    shortcuts: [
      { keys: "Alt + A + M", action: "Open the Remove Duplicates dialog" },
      { keys: "Alt + H + L + H + D", action: "Highlight Duplicates (Conditional Formatting)" },
      { keys: "Ctrl + Z", action: "Undo a Remove Duplicates immediately" }
    ],
    exercises: [
      {
        task: "On the Duplicates sheet of the data cleaning workbook, highlight every repeated customer email with conditional formatting.",
        hint: "Conditional Formatting > Highlight Cells Rules > Duplicate Values.",
        solution: "Select the Email column, Home > Conditional Formatting > Highlight Cells Rules > Duplicate Values, OK. Repeated emails turn red. Count them, then compare with the Solution sheet."
      },
      {
        task: "Add a helper column that shows how many times each email appears in the list, then filter it to find the repeat rows.",
        hint: "COUNTIF with an anchored range, filled down.",
        solution: "In a free column type =COUNTIF($B$2:$B$500, B2) and fill down. Every repeated email shows a number above 1 on each of its rows. Filter the helper to greater than 1 to isolate them."
      },
      {
        task: "Create a backup copy of the sheet, then remove rows where every column is identical.",
        hint: "Right-click the sheet tab > Move or Copy > Create a copy. Then Remove Duplicates with all columns ticked.",
        solution: "Copy the sheet as CustomerList-BACKUP. On the original, Data > Remove Duplicates, keep all columns ticked, OK. Excel reports how many duplicates were removed. The Solution sheet shows the expected final row count."
      },
      {
        task: "One customer appears twice with different addresses. Keep the correct one without losing data.",
        hint: "Remove Duplicates always keeps the first row — sort or inspect first.",
        solution: "Filter or sort to bring the two rows together, read both, delete the outdated row manually (or sort so the correct row is first and then run Remove Duplicates on the defining columns). Manual review is the right tool when duplicates are not identical."
      }
    ],
    quiz: [
      {
        q: "Which conditional formatting rule makes repeated values visible before you delete anything?",
        options: ["Data Bars", "Duplicate Values", "Color Scales", "Icon Sets"],
        answer: 1,
        explanation: "Highlight Cells Rules > Duplicate Values colours every repeated value, so you can see the pattern — and spot the near-duplicate typos that exact-matching tools would miss."
      },
      {
        q: "In the Remove Duplicates dialog, what do the column tick boxes control?",
        options: [
          "Which columns get deleted",
          "Which columns must match for two rows to count as duplicates",
          "Which columns are sorted",
          "Nothing — Excel decides automatically"
        ],
        answer: 1,
        explanation: "A row is treated as a duplicate only in the ticked columns. All columns ticked = fully identical rows; a single ID column ticked = one row per ID."
      },
      {
        q: "Which formula counts how many times the email in B2 appears in B2:B500?",
        options: [
          "=COUNTA(B2)",
          "=COUNTIF($B$2:$B$500, B2)",
          "=SUM(B2:B500)",
          "=COUNTIF(B2, $B$2:$B$500)"
        ],
        answer: 1,
        explanation: "COUNTIF's first argument is the anchored range, the second is the value to count. Filling it down shows each row how many twins it has."
      },
      {
        q: "When two duplicate rows differ slightly, which row does Remove Duplicates keep?",
        options: [
          "The row with the newest date",
          "A random row",
          "The first row in the current order",
          "It refuses to run and asks you"
        ],
        answer: 2,
        explanation: "It keeps the first occurrence and deletes the later ones. If a later row holds better data, sort it to the top first — or edit by hand."
      },
      {
        q: "Why take a backup sheet before Remove Duplicates?",
        options: [
          "Excel requires it",
          "Undo stops working after the file is saved or closed",
          "Remove Duplicates destroys the whole sheet",
          "The backup speeds up the operation"
        ],
        answer: 1,
        explanation: "Ctrl + Z covers the immediate next step only. A copied sheet (Move or Copy > Create a copy) is the durable safety net for an instant, unpreviewable deletion."
      }
    ],
    summary: [
      "Highlight duplicates with Conditional Formatting first, and look for near-duplicate typos with the Unique option.",
      "`=COUNTIF($B$2:$B$500, B2)` filled down counts how many times each value appears.",
      "Remove Duplicates deletes rows that match in the ticked columns and keeps the first occurrence.",
      "Always copy the sheet as a backup before any removal — undo does not survive a save.",
      "When duplicates differ in content, decide manually which row is the correct one."
    ],
    downloads: [
      { file: "assets/files/level3/data-cleaning.xlsx",
        label: "Data cleaning workbook",
        note: "Exercise and Solution sheets with messy customer records for cleaning practice" }
    ],
    keywords: ["remove duplicates", "duplicate values", "countif duplicates", "backup sheet", "unique records", "clean customer list"]
  },

  {
    id: "l3-text-to-columns",
    title: "Text to Columns and Data Splitting",
    difficulty: "Beginner",
    duration: 12,
    version: "",
    objectives: [
      "Split text by a delimiter such as a comma or space",
      "Split by fixed width positions",
      "Convert numbers stored as text into real numbers",
      "Split full names into first and last name columns",
      "Know when Flash Fill is the faster alternative"
    ],
    intro: "Data arrives joined together more often than you would expect: full names in one cell, addresses crammed into a single field, and — most famously — a whole column of \"numbers\" that Excel refuses to sum because they are actually text.\n\nText to Columns is the built-in splitting tool, and its third page hides one of the most valuable fixes in office Excel: converting text numbers into real ones. This lesson covers both, plus Flash Fill as the modern shortcut.",
    sections: [
      {
        heading: "Delimited split: commas, spaces and tabs",
        blocks: [
          { type: "p", text: "A delimiter is the character that separates the pieces — comma in CSV exports, space between first and last name, semicolon in some European files, tab in pasted database output. The wizard walks you through it:" },
          { type: "steps", items: [
            "Select the column with the combined text, for example `B2:B500` containing \"Mehta, Arjun\".",
            "Data > Text to Columns. Choose **Delimited**, click Next.",
            "Tick the delimiter your data uses — Comma. A live preview shows the split. Untick the ones you do not want (Tab is on by default).",
            "Click Next, choose the destination (the cell where the first result should start, like `=B2`'s neighbour `C2`) and click Finish."
          ] },
          { type: "grid", title: "Before: one column with joined text. After a comma split into two columns",
            rows: [
              ["Customer (before)", "Last name", "First name"],
              ["Mehta, Arjun", "Mehta", "Arjun"],
              ["Nair, Priya", "Nair", "Priya"],
              ["Verma, Rahul", "Verma", "Rahul"]
            ] },
          { type: "warning", text: "Text to Columns writes over whatever sits to the right of the destination. Always insert enough empty columns between your data and anything important before you click Finish — a split into three parts overwrites three columns." }
        ]
      },
      {
        heading: "Fixed width split",
        blocks: [
          { type: "p", text: "Choose **Fixed width** when the pieces sit at exact character positions rather than being separated by a character. Classic case: system exports like `NORTH0142` where positions 1-5 are the region code and positions 6-10 are the branch number." },
          { type: "p", text: "In the wizard's preview pane, click at a character position to place a break line (drag to move it, double-click to remove it). Excel cuts at those positions regardless of delimiters. It is a niche tool, but when the data is positional it is the only tool that works without formulas." }
        ]
      },
      {
        heading: "The classic office fix: numbers stored as text",
        blocks: [
          { type: "p", text: "Symptoms you will meet within your first month in any data job: a column of numbers that left-aligns (real numbers align right), `=SUM()` returns 0, a green triangle appears on every cell, and lookups fail because `\"120\"` does not match `120`. The values are text that looks like numbers — usually from a CSV export or a system paste." },
          { type: "p", text: "The fastest reliable repair for a whole column is Text to Columns with no delimiter at all:" },
          { type: "steps", items: [
            "Select the affected column, for example the Units column `F2:F500`.",
            "Data > Text to Columns. Choose Delimited, click Next.",
            "Untick **every** delimiter so nothing splits, and click Next.",
            "On the last page choose **General** as the column data format, click Finish."
          ] },
          { type: "p", text: "General format makes Excel re-read each value and decide its type: \"120\" becomes the number 120, alignment flips to the right, SUM starts working. The column must be a single column — select only it. The same wizard page also has a Date option, which repairs dates imported as text such as `20260105` when you tell Excel the year-month-day order." },
          { type: "tip", text: "Quicker single-cell checks: a small green triangle in the cell corner with an exclamation mark offers \"Convert to Number\" when you select the cells. It works well for a few cells; for thousands, Text to Columns is faster and more consistent." }
        ]
      },
      {
        heading: "Flash Fill: the fast alternative",
        blocks: [
          { type: "p", text: "Flash Fill (Excel 2013 and later — covered in Level 1) often beats the wizard for simple splits. Type the desired result of the first two rows in a new column — \"Arjun\" next to \"Mehta, Arjun\", \"Priya\" next to \"Nair, Priya\" — and Excel detects the pattern. Press [[Ctrl]] + [[E]] or Data > Flash Fill and the rest of the column fills itself." },
          { type: "table",
            headers: ["Situation", "Better tool", "Why"],
            rows: [
              ["One obvious pattern (first name, area code, initials)", "Flash Fill", "Type two examples, Ctrl + E, done"],
              ["A structured delimiter split of thousands of rows", "Text to Columns", "Deterministic, no pattern guessing, no typing"],
              ["Numbers stored as text", "Text to Columns (General)", "Flash Fill produces text results — it cannot fix data types"],
              ["Complex or variable patterns", "Formulas or Power Query (Level 6)", "Full control and repeatability"]
            ] },
          { type: "warning", text: "Flash Fill is a guess. It can be confidently wrong when some rows deviate from the pattern it detected — a name with a middle initial, a double space. Always spot-check a few rows after Ctrl + E, and remember its output is static text: new rows do not fill themselves the way formulas do." }
        ]
      }
    ],
    mistakes: [
      {
        wrong: "Running a comma split with the destination next to important columns and overwriting them.",
        right: "Insert empty columns first so the split has room to land.",
        why: "Text to Columns overwrites cells to the right of the destination without asking. The undo works — if you notice immediately and have not saved."
      },
      {
        wrong: "Trying to fix text numbers with Flash Fill or by retyping values one by one.",
        right: "Use Text to Columns > General on the single column, or the Convert to Number option from the green-triangle menu.",
        why: "Flash Fill copies the text characters; it does not change the data type. The wizard's General format genuinely converts thousands of cells in one pass."
      },
      {
        wrong: "Selecting three columns at once before the text-numbers fix.",
        right: "Select exactly one column for the Text to Columns repair.",
        why: "The wizard works per selection, and with multiple columns selected the delimiter page behaves differently — the reliable route is one column at a time."
      }
    ],
    shortcuts: [
      { keys: "Alt + A + E", action: "Open Text to Columns wizard" },
      { keys: "Ctrl + E", action: "Flash Fill based on the examples you typed" },
      { keys: "Ctrl + Z", action: "Undo a split that landed badly" }
    ],
    exercises: [
      {
        task: "On the Splitting sheet of the data cleaning workbook, split the full-name column \"Last name, First name\" into two columns.",
        hint: "Delimited split with Comma. Insert empty columns to the right first.",
        solution: "Insert two empty columns next to the names. Select the name column, Data > Text to Columns > Delimited > Next, tick Comma only, Next, set destination to the first empty column, Finish. Compare with the Solution sheet."
      },
      {
        task: "A Units column shows numbers that are left-aligned and SUM returns 0. Fix the whole column in one pass.",
        hint: "Text to Columns, Delimited, untick all delimiters, General format.",
        solution: "Select the Units column only, Data > Text to Columns > Delimited > Next, untick Tab (and any other ticked delimiter), Next, keep General, Finish. Values turn right-aligned and =SUM() now adds them correctly."
      },
      {
        task: "Use Flash Fill to create a First name column from the combined names.",
        hint: "Type the first two expected results by hand, then Ctrl + E.",
        solution: "In an empty column type \"Arjun\" next to \"Mehta, Arjun\" and \"Priya\" next to \"Nair, Priya\". Press Ctrl + E — the rest fills automatically. Spot-check a few rows, since Flash Fill guesses the pattern."
      },
      {
        task: "A product-code column contains entries like NORTH0142 where positions 1-5 are the region code. Extract the region code.",
        hint: "Fixed width split with a break line after position 5.",
        solution: "Select the column, Data > Text to Columns > Fixed width > Next. Click between the 5th and 6th characters to place a break line, Next, choose a destination with an empty column beside it, Finish. The region code lands in its own column."
      }
    ],
    quiz: [
      {
        q: "You want to split \"Mehta, Arjun\" into two cells. Which wizard type fits?",
        options: ["Fixed width", "Delimited with Comma", "Delimited with Space", "General only"],
        answer: 1,
        explanation: "The comma is the separator character, so Delimited with Comma is the match. Delimited with Space would split on the space after the comma instead."
      },
      {
        q: "A column of numbers is actually text: SUM returns 0. What is the one-pass repair?",
        options: [
          "Retype each value",
          "Text to Columns > Delimited with no delimiters > General format",
          "Apply a currency format",
          "Flash Fill the column"
        ],
        answer: 1,
        explanation: "Running the wizard with General makes Excel re-evaluate every value and convert \"120\" into the number 120. Flash Fill and formats change appearance, not the underlying type."
      },
      {
        q: "What does Flash Fill (Ctrl + E) do?",
        options: [
          "Fills empty cells with zeros",
          "Fills a column by detecting the pattern from examples you typed",
          "Converts text to numbers",
          "Sorts the selected column"
        ],
        answer: 1,
        explanation: "Flash Fill learns from the examples in the first cells and completes the column. Its results are static text and pattern-based, so always spot-check."
      },
      {
        q: "Why should you insert empty columns before a Text to Columns split?",
        options: [
          "The wizard requires blank columns",
          "The split results overwrite the cells to the right of the destination",
          "It makes the split faster",
          "Empty columns keep the file small"
        ],
        answer: 1,
        explanation: "Results are written starting at the destination and continue rightward, replacing any content there. Room to land prevents accidental overwrites."
      },
      {
        q: "Which wizard page also fixes dates that were imported as text, such as 20260105?",
        options: [
          "Step 1: Delimited or Fixed width",
          "Step 3: the Column data format option Date",
          "The Finish button",
          "Text to Columns cannot fix dates"
        ],
        answer: 1,
        explanation: "Step 3 lets you mark a column as Date and specify the order of year, month and day, converting text dates into real dates Excel can filter and calculate with."
      }
    ],
    summary: [
      "Data > Text to Columns splits a column by delimiter (comma, space, tab) or by fixed positions.",
      "The wizard overwrites cells to the right of the destination — insert empty columns first.",
      "Text to Columns > General is the classic one-pass fix for numbers stored as text.",
      "Flash Fill (Ctrl + E) fills patterns from examples — fast, but static and guess-based.",
      "Step 3 also converts text dates to real dates when you specify the date order."
    ],
    downloads: [
      { file: "assets/files/level3/data-cleaning.xlsx",
        label: "Data cleaning workbook",
        note: "Exercise and Solution sheets with messy customer records for cleaning practice" }
    ],
    keywords: ["text to columns", "delimited", "fixed width", "split names", "numbers as text", "convert to number", "flash fill", "ctrl e", "csv import"]
  },

  {
    id: "l3-find-replace",
    title: "Find and Replace",
    difficulty: "Beginner",
    duration: 12,
    version: "",
    objectives: [
      "Use Ctrl + F and Ctrl + H across a sheet or the whole workbook",
      "Control searches with Match Case, Entire Cell and Within options",
      "Search with the wildcards * and ?",
      "Use Go To Special to select blanks, constants and formulas, and fill blanks quickly",
      "Replace safely inside formulas"
    ],
    intro: "Every dataset you inherit has something wrong in bulk: a supplier renamed with two spellings, \"N.A\" where numbers should be, region codes that changed last quarter. Editing hundreds of cells by hand is not work — it is punishment.\n\nFind and Replace fixes hundreds of cells in seconds, and Go To Special — its often-overlooked partner — selects exactly the cells you need to fix: all blanks, all formulas, all constants. Together they are the fastest bulk-editing tools in Excel.",
    sections: [
      {
        heading: "Find and Replace basics",
        blocks: [
          { type: "steps", items: [
            "Press [[Ctrl]] + [[F]] to open Find, or [[Ctrl]] + [[H]] to open Replace directly.",
            "Type what to find — for example \"Srini\" — and, in Replace, what it should become — \"Srinivasan\".",
            "Optional: click Options to reach the power settings (next section).",
            "Click Find Next / Replace to walk through matches one by one, or Replace All to fix every match at once."
          ] },
          { type: "p", text: "Two habits worth building immediately. First, **Replace All without checking the count is a gamble** — read the confirmation message (\"All done. We made 42 replacements.\") and judge whether 42 sounds right for your data. Second, **always have a backup or an undo plan**: Ctrl + Z undoes a Replace All, but not after you have saved and closed." },
          { type: "warning", text: "Replace All also matches inside longer words. Replacing \"Pen\" with \"Pens\" would turn \"Pen\" into \"Pens\" but also \"Pencil\" into \"Pencilss\" if you search values (and \"Pencil\" cells contain the letters P-e-n). Use Entire cell match or case sensitivity to control this." }
        ]
      },
      {
        heading: "The options that make it safe",
        blocks: [
          { type: "p", text: "Click Options >> in the dialog to reveal the settings that separate careful users from victims:" },
          { type: "table",
            headers: ["Option", "What it changes", "When you need it"],
            rows: [
              ["Within: Sheet / Workbook", "Search only this sheet or every sheet", "A supplier name changed across 12 sheets"],
              ["Match case", "\"north\" and \"North\" become different matches", "Fixing lowercase regions without touching correct ones"],
              ["Entire cell contents", "Only exact whole-cell matches", "Replacing the value \"Pen\" but not \"Pens\" or \"Pencil\""],
              ["Look in: Formulas / Values / Comments", "Searches the formula text or the displayed result", "Editing formulas in bulk vs cleaning displayed values"]
            ] },
          { type: "p", text: "The **Look in: Formulas** setting is the one people discover too late. By default, Ctrl + F searches formulas — so searching for \"120\" finds the number 120 in a cell and also finds `=SUM(B2:B120)` inside a far-away formula, because the formula text contains 120. When your goal is data cleaning, switch Look in to Values; when your goal is bulk-editing formulas, keep Formulas and stay very precise." }
        ]
      },
      {
        heading: "Wildcards: * and ?",
        blocks: [
          { type: "p", text: "Find supports two wildcards in both Find and Replace fields (and in COUNTIF-type criteria):" },
          { type: "ul", items: [
            "**`*` matches any number of characters.** Searching `Verma*` finds \"Verma\", \"Verma Industries\" and \"Verma & Sons\". Searching `*@sunrise.in` finds every email ending in that domain.",
            "**`?` matches exactly one character.** Searching `Pen?` finds \"Pens\" and \"Peng\" but not \"Pen\" and not \"Pencil\".",
            "**To search for a literal `*` or `?`**, put a tilde before it: `~*` finds an actual asterisk."
          ] },
          { type: "p", text: "Wildcards make Replace All safe and surgical. To change every \"Verma-something\" supplier to a single name, search `Verma*` and replace with the new name — one click instead of a hundred edits. To standardise old product codes like `NB-12` through `NB-99`, search `NB-??` (two characters after the dash) and replace with the new prefix pattern." }
        ]
      },
      {
        heading: "Go To Special: select all blanks, constants or formulas",
        blocks: [
          { type: "p", text: "Press [[Ctrl]] + [[G]] (or F5) and click Special — or Home > Find & Select > Go To Special. This dialog does not find text; it selects whole categories of cells at once:" },
          { type: "table",
            headers: ["Option", "Selects", "Typical use"],
            rows: [
              ["Blanks", "Every empty cell in the selection", "Fill gaps in a report quickly"],
              ["Constants", "Typed values only, not formulas", "Find cells someone overtyped with a hard number"],
              ["Formulas", "Cells containing formulas (by result type: numbers, text, errors)", "Audit where formulas live before restructuring"],
              ["Errors", "All error values like #N/A, #REF!", "Locate every broken formula in one selection"],
              ["Visible cells only", "Just the unhidden rows", "Copy from a filtered sheet without hidden rows"]
            ] },
          { type: "p", text: "Selection is the point: everything Excel can do to one cell, it can do to a whole selection. Select all errors and press Delete; select all constants and apply a fill; select all blanks and fill them in one command — which is the famous next trick." }
        ]
      },
      {
        heading: "Filling blanks quickly",
        blocks: [
          { type: "p", text: "A report has Region filled only on the first row of each group, leaving gaps below — fine for humans, terrible for filters, SUMIFs and pivots. Filling every blank with the value above takes four moves:" },
          { type: "steps", items: [
            "Select the Region column range `D2:D500`.",
            "Go To Special > Blanks > OK. Every empty cell is selected (the active cell is the first one).",
            "Type `=` and press the Up arrow once — the formula becomes `=D2`, pointing at the cell above the first blank.",
            "Press [[Ctrl]] + [[Enter]] to fill the same formula into **all selected cells**, each one referencing its own cell above. Then copy the column and Paste Special > Values to turn the formulas into plain data."
          ] },
          { type: "grid", title: "Before: gaps after the group header. After the fill: every row carries its region",
            rows: [
              ["Order ID", "Region"],
              ["1001", "North"],
              ["1002", "=D2"],
              ["1003", "=D3"],
              ["1004", "South"]
            ] },
          { type: "warning", text: "The formulas point at the cell above, so they are only stable until the next sort. That is why the final Paste Special > Values step matters: convert the filled formulas into plain values before you sort, filter or move anything." }
        ]
      }
    ],
    mistakes: [
      {
        wrong: "Replace All with \"Pen\" to \"Pens\" across the sheet, then finding \"Pencilss\" in the data.",
        right: "Tick Entire cell contents (for values) or Match case, and read the replacement count before accepting.",
        why: "Replace matches inside longer text by default. A blanket Replace All without options can silently corrupt hundreds of unrelated cells."
      },
      {
        wrong: "Searching for a number like 120 and getting confused by formula matches such as `=SUM(B2:B120)`.",
        right: "Set Look in to Values when cleaning data, or Formulas when editing formulas — and know which one you want.",
        why: "Find searches formula text by default. Every bulk edit needs the right Look in target or it finds the wrong layer of the sheet."
      },
      {
        wrong: "Filling blanks with the `=cell above` trick and then sorting the table immediately.",
        right: "Convert the filled formulas to values with Paste Special > Values before sorting or filtering.",
        why: "The fill formulas reference positions, not values. After a sort they point at different rows and the regions scramble — exactly the sorting risk from earlier in this level."
      }
    ],
    shortcuts: [
      { keys: "Ctrl + F", action: "Open Find" },
      { keys: "Ctrl + H", action: "Open Replace" },
      { keys: "Ctrl + G", action: "Open Go To (click Special for Go To Special)" },
      { keys: "Shift + F4", action: "Repeat the last Find without reopening the dialog" },
      { keys: "Ctrl + Enter", action: "Fill the entry into all selected cells" }
    ],
    exercises: [
      {
        task: "On the FindReplace sheet of the data cleaning workbook, a supplier was renamed: replace every \"Sunrise Traders\" variant with the new name \"Sunrise Trading Co\".",
        hint: "Ctrl + H, consider whether Entire cell or a wildcard fits, read the replacement count.",
        solution: "Ctrl + H: Find \"Sunrise Traders\", Replace \"Sunrise Trading Co\", Replace All. Read the count message. If some cells contain extra text like \"Sunrise Traders - North\", use Find \"Sunrise Traders\" which still matches inside longer text. Check the Solution sheet."
      },
      {
        task: "Replace the lowercase region values \"north\" with \"North\" without touching the correctly typed ones.",
        hint: "Options >> Match case.",
        solution: "Ctrl + H, Find \"north\", Replace \"North\", Options >> tick Match case, Replace All. Only lowercase occurrences change; \"North\" cells are skipped because case differs."
      },
      {
        task: "Fill every blank cell in the Region column with the value from the cell above.",
        hint: "Go To Special > Blanks, then = and Up arrow, then Ctrl + Enter, then paste as values.",
        solution: "Select the Region range, Ctrl + G > Special > Blanks > OK. Type = and press Up arrow, then Ctrl + Enter. Every blank now shows the region above it. Copy the column, Paste Special > Values to fix them in place."
      },
      {
        task: "Find every formula on the sheet that returns an error value.",
        hint: "Go To Special has a direct option for this.",
        solution: "Ctrl + G > Special > Formulas > tick Errors only > OK. Excel selects all error cells at once so you can inspect or clear them together. The Solution sheet lists what the workbook contains."
      }
    ],
    quiz: [
      {
        q: "Which shortcut opens the Replace dialog directly?",
        options: ["Ctrl + F", "Ctrl + H", "Ctrl + G", "Ctrl + R"],
        answer: 1,
        explanation: "Ctrl + H is Replace; Ctrl + F is Find; Ctrl + G is Go To (with the Special button); Ctrl + R fills right."
      },
      {
        q: "What does the wildcard pattern Verma* match?",
        options: [
          "Only the exact text Verma",
          "Any text starting with Verma, such as Verma Industries",
          "Any text containing Verma anywhere",
          "Only six-letter words"
        ],
        answer: 1,
        explanation: "The asterisk stands for any number of characters after the anchor. Verma* matches anything that begins with Verma; *Verma would match it anywhere."
      },
      {
        q: "You want to replace the whole-cell value \"Pen\" but not \"Pens\" or \"Pencil\". Which option prevents the damage?",
        options: ["Match case", "Entire cell contents", "Look in: Formulas", "Within: Workbook"],
        answer: 1,
        explanation: "Entire cell contents matches only cells whose full value equals the search text. Shorter or longer values like Pens and Pencil are ignored."
      },
      {
        q: "Go To Special > Blanks, then typing = and pressing the Up arrow, then Ctrl + Enter achieves what?",
        options: [
          "Deletes all blank cells",
          "Fills every selected blank with a copy of the cell above it",
          "Highlights blanks in yellow",
          "Counts the blank cells"
        ],
        answer: 1,
        explanation: "Go To Special selects all blanks at once; the relative formula =cell-above filled with Ctrl + Enter copies the value above into every gap. Finish with Paste Special > Values before sorting."
      },
      {
        q: "Searching the value 120 also matches a formula =SUM(B2:B120). Which setting stops that?",
        options: [
          "Match entire cell contents",
          "Look in: Values instead of Formulas",
          "Within: Sheet",
          "It is not possible"
        ],
        answer: 1,
        explanation: "Find searches the formula text by default. Look in: Values searches what cells display, so only actual values like 120 match."
      }
    ],
    summary: [
      "Ctrl + F finds, Ctrl + H replaces; Replace All reports a count — always read it.",
      "Options >> gives Within (sheet/workbook), Match case, Entire cell and Look in Formulas vs Values.",
      "Wildcards: `*` matches any run of characters, `?` exactly one, and `~*` a literal asterisk.",
      "Go To Special (Ctrl + G > Special) selects all blanks, constants, formulas or errors at once.",
      "Fill blanks via Go To Special > Blanks + `=up` + Ctrl + Enter, then paste as values before sorting."
    ],
    downloads: [
      { file: "assets/files/level3/data-cleaning.xlsx",
        label: "Data cleaning workbook",
        note: "Exercise and Solution sheets with messy customer records for cleaning practice" }
    ],
    keywords: ["find and replace", "ctrl f", "ctrl h", "wildcards", "match case", "entire cell", "go to special", "fill blanks", "replace in formulas"]
  },

  {
    id: "l3-data-cleaning",
    title: "The Complete Data Cleaning Workflow",
    difficulty: "Intermediate",
    duration: 18,
    version: "",
    objectives: [
      "Inspect messy data and list its problems before fixing anything",
      "Apply TRIM, PROPER and date standardisation to a real customer list",
      "Split, de-duplicate and convert data types in a working order",
      "Run a final checklist before trusting a cleaned dataset",
      "Know when a job has outgrown formulas and needs Power Query"
    ],
    intro: "This lesson puts everything from this level together into one professional routine. You receive a customer list like the ones that actually arrive by email: names typed five different ways, stray spaces, dates as text, numbers as text, duplicates everywhere.\n\nThe discipline is the same whether the file has 50 rows or 50,000: inspect first, fix in a planned order, verify at the end. Follow this workflow and the data cleaning workbook becomes a template you can reuse on every messy file you meet.",
    sections: [
      {
        heading: "Step 0: inspect before you touch anything",
        blocks: [
          { type: "p", text: "Cleaning without looking first is how good columns get destroyed. Open the messy sheet in the workbook and spend five minutes noticing:" },
          { type: "ul", items: [
            "**Alignment** — left-aligned numbers are text; right-aligned are real numbers.",
            "**Green triangles** — Excel's own hint of text-stored-numbers.",
            "**Extra spaces** — visible as ragged alignment, or prove them with `=LEN(A2)` compared against a clean twin value.",
            "**Case chaos** — \"ARJUN MEHTA\", \"arjun mehta\" and \"Arjun  Mehta\" (two spaces) in the same column.",
            "**Date formats** — 05-01-2026, 5/1/26 and 05-Jan all in one column.",
            "**Duplicates** — highlight with the Duplicate Values rule from the previous lesson."
          ] },
          { type: "p", text: "Then write down the fix order before starting. The order matters: for example, TRIM before de-duplicating, or the extra space in \"Priya Nair \" keeps two \"different\" names alive. A wrong order means doing the same job twice." },
          { type: "tip", text: "Work on a copy of the sheet and keep the original untouched (right-click the tab > Move or Copy > Create a copy). Every cleaning step below is safe to undo only until the next save — the untouched original is your reference." },
          { type: "grid", title: "A taste of what you will find in the workbook: every row of this sample needs work",
            rows: [
              ["Customer", "Date", "Units"],
              ["arjun mehta ", "05-01-2026", "120"],
              ["PRIYA NAIR", "05012026", "45"],
              ["Rahul  Verma", "5/1/26", "8"],
              ["arjun mehta", "06-01-2026", "200"]
            ] },
          { type: "p", text: "Row by row: mixed case and a trailing space on Arjun, a date in system format on Priya, a double space on Rahul and month-day ambiguity everywhere, units stored as text (left-aligned), and Arjun appearing twice. Five rows, five problem types — the real sheet repeats this pattern across hundreds of records." }
        ]
      },
      {
        heading: "Text fixes: TRIM and PROPER",
        blocks: [
          { type: "p", text: "TRIM removes leading, trailing and doubled spaces — the number one invisible cause of failed matches. `\"Priya Nair \"` does not equal `\"Priya Nair\"`, so VLOOKUP misses it and Remove Duplicates keeps both rows. Fix in helper columns, not in place:" },
          { type: "formula", name: "TRIM",
            syntax: "=TRIM(text)",
            example: "=TRIM(B2)",
            explanation: "Returns the text with all extra spaces removed: leading, trailing and double spaces collapse to single separators." },
          { type: "formula", name: "PROPER",
            syntax: "=PROPER(text)",
            example: "=PROPER(TRIM(B2))",
            explanation: "Capitalises the first letter of each word: \"arjun mehta\" becomes \"Arjun Mehta\". Combining it with TRIM fixes spacing and case in one pass." },
          { type: "steps", items: [
            "Insert a helper column next to Customer. In the first row type `=PROPER(TRIM(B2))` and fill down.",
            "Check the results for false hits: PROPER also capitalises \"ltd\" into \"Ltd\" — usually desirable for names, watch it for values like \"3m\".",
            "Select the helper column, copy, then Paste Special > Values onto the original column.",
            "Delete the helper column. The name column is now clean, single-spaced and consistently capitalised."
          ] },
          { type: "warning", text: "TRIM removes the ordinary space character only. Data pasted from web pages often contains non-breaking spaces that TRIM ignores; if a value still has a phantom space after TRIM, replace it via Ctrl + H using [[Alt]] + [[0]] + [[1]] + [[6]] + [[0]] typed on the numeric keypad in the Find box (leave Replace empty)." }
        ]
      },
      {
        heading: "Dates and numbers: make them real types",
        blocks: [
          { type: "p", text: "Dates stored as text cannot be filtered by month, sorted chronologically or used in calculations — and text numbers break SUM. The fixes come straight from the Text to Columns lesson:" },
          { type: "ul", items: [
            "**Text dates like 20260105:** Text to Columns > Delimited, no delimiters, Step 3 choose **Date** with YMD order, Finish.",
            "**Mixed-looking dates like 05-01-2026:** decide what they mean in this file (day-month or month-day), then convert once and check a known row against the source system.",
            "**Text numbers:** select the column, Text to Columns > Delimited, no delimiters, **General**, Finish.",
            "**Verify:** the column right-aligns, `=SUM()` returns a sensible number, and `=ISNUMBER(C2)` filled down returns TRUE everywhere."
          ] },
          { type: "p", text: "Standardise the display too: select the date column and apply one explicit date format (for example 14-Mar-2026) via Ctrl + 1. A cleaned file should show one format everywhere — mixed formats are how \"05-01\" gets misread as May first." }
        ]
      },
      {
        heading: "Split, de-duplicate, finalise",
        blocks: [
          { type: "p", text: "With types fixed, handle structure. If the file has \"Last, First\" combined names or address blobs, split them with Text to Columns (empty columns first). If city and region are jammed together, split there too. Structure before de-duplication: you want the duplicate definition to use the final, clean columns." },
          { type: "steps", items: [
            "Convert the finished range to an Excel Table with [[Ctrl]] + [[T]] and name it Customers.",
            "Highlight duplicates with Conditional Formatting > Duplicate Values on the defining columns (Email or Customer ID).",
            "Review: same customer twice with one newer address? Fix by hand. Fully identical rows? Proceed.",
            "Back up the sheet (Move or Copy > Create a copy), then Data > Remove Duplicates with the defining columns ticked.",
            "Read the result message — the removed count should match what the highlighting suggested."
          ] },
          { type: "p", text: "The final checklist before you trust the file — run it every time, it takes two minutes:" },
          { type: "ol", items: [
            "No leading/trailing spaces: spot-check with `=LEN()` on a few rows.",
            "All names in one consistent case, dates in one format.",
            "Numbers right-aligned and SUM-able; `=ISNUMBER()` TRUE down each numeric column.",
            "Duplicate count is zero: run the Duplicate Values highlight once more — it should light up nothing.",
            "Row count matches expectations: `=COUNTA(A:A)-1` for the data rows.",
            "A test SUMIF or VLOOKUP from the earlier levels returns the value you can verify by hand on one row."
          ] },
          { type: "quote", text: "Cleaning is not finished when the data looks right. It is finished when a test formula agrees with a value you verified by hand." }
        ]
      },
      {
        heading: "When to graduate to Power Query",
        blocks: [
          { type: "p", text: "Everything above lives in formulas and ribbon tools, which is perfect for one-off fixes and files up to a few thousand rows. But notice what you did: a fixed sequence of steps applied to a file. Next month the same supplier sends the same messy export again — and you repeat all of it by hand." },
          { type: "p", text: "Power Query (Data > Get Data, covered in Level 6) records those steps once — trim, change case, change type, split, remove duplicates — and replays them on any new file with one Refresh. It also handles files too big for comfortable formula work. The rule of thumb: formulas and this lesson's workflow for today's file; Power Query when the same cleaning repeats monthly or the data exceeds a few hundred thousand rows." }
        ]
      }
    ],
    mistakes: [
      {
        wrong: "Cleaning in a random order — removing duplicates first and trimming spaces afterwards.",
        right: "Fix spaces and case first (TRIM, PROPER), then split columns, then de-duplicate.",
        why: "\"Priya Nair \" and \"Priya Nair\" are not duplicates until the space is gone. Wrong order means duplicates survive the first pass and the job must be repeated."
      },
      {
        wrong: "Typing fixes directly over the original data with no copy and no helper columns.",
        right: "Work on a copied sheet; use helper columns for TRIM/PROPER and paste as values over the originals.",
        why: "In-place fixing destroys your ability to compare before and after. When a formula misfires, the original column is gone — helpers make every step reversible."
      },
      {
        wrong: "Declaring the file clean because it \"looks right\".",
        right: "Run the verification checklist: LEN checks, ISNUMBER checks, a final duplicate highlight and a hand-verified test formula.",
        why: "Invisible problems (non-breaking spaces, one text number among 499 real ones) do not show on screen. Verification formulas catch what eyes cannot."
      }
    ],
    shortcuts: [
      { keys: "Ctrl + T", action: "Convert the cleaned range into a Table" },
      { keys: "Alt + A + E", action: "Open Text to Columns for date and number conversions" },
      { keys: "Ctrl + Alt + V", action: "Paste Special (Values) to replace helpers with plain data" },
      { keys: "Ctrl + 1", action: "Open Format Cells to standardise date display" }
    ],
    exercises: [
      {
        task: "Open the Cleaning sheet of the data cleaning workbook. Inspect it and list every problem type you can find before changing anything.",
        hint: "Look for alignment, green triangles, case chaos, extra spaces, text dates and duplicates.",
        solution: "A correct inspection finds: names in mixed case, extra trailing spaces, dates stored as text, units stored as text, repeated customer rows and combined name fields. The Solution sheet lists the full problem inventory."
      },
      {
        task: "Clean the Customer column: remove extra spaces and standardise capitalisation.",
        hint: "Helper column with =PROPER(TRIM(B2)), fill down, paste as values back.",
        solution: "Insert a helper column, type =PROPER(TRIM(B2)) and fill down. Copy the helper, Paste Special > Values over the Customer column, delete the helper. \"arjun mehta \" becomes \"Arjun Mehta\". Compare with the Solution sheet."
      },
      {
        task: "Convert the date column and the units column to real types.",
        hint: "Text to Columns with no delimiters: Date format for dates, General for numbers.",
        solution: "For the date column: Data > Text to Columns > Delimited, untick all, Step 3 choose Date and the correct order, Finish. For units: same route but General. Both columns should right-align and ISNUMBER checks return TRUE."
      },
      {
        task: "Finish the job: convert to a Table, highlight remaining duplicates, back up the sheet and remove them.",
        hint: "Ctrl + T, Duplicate Values highlight, Move or Copy backup, Remove Duplicates with defining columns ticked.",
        solution: "Ctrl + T and rename the Table Customers. Highlight duplicates on Email, review each pair, back up the sheet, then Data > Remove Duplicates. Verify with the checklist: LEN spot-checks, ISNUMBER, zero highlighted duplicates, and a SUMIF you can confirm by hand. The Solution sheet shows the expected final state."
      }
    ],
    quiz: [
      {
        q: "Why should TRIM run before Remove Duplicates?",
        options: [
          "TRIM is faster than Remove Duplicates",
          "Rows differing only by a stray space are not duplicates until the space is removed",
          "Remove Duplicates deletes TRIM formulas",
          "Order does not matter"
        ],
        answer: 1,
        explanation: "\"Priya Nair \" and \"Priya Nair\" compare as different text, so Remove Duplicates keeps both. Clean the text first, then the duplicate definition is meaningful."
      },
      {
        q: "What does =PROPER(TRIM(B2)) return for the input \"arjun  mehta \" (two spaces, trailing space)?",
        options: [
          "ARJUN MEHTA",
          "arjun mehta",
          "Arjun Mehta",
          "\"arjun  mehta \" unchanged"
        ],
        answer: 2,
        explanation: "TRIM collapses the double space and strips the trailing one; PROPER capitalises each word. The combined formula fixes spacing and case in one pass."
      },
      {
        q: "A units column of text numbers needs conversion for 5,000 rows. What is the fastest correct method?",
        options: [
          "Flash Fill",
          "Retype the values",
          "Text to Columns > Delimited, no delimiters, General format",
          "Apply a number format from the ribbon"
        ],
        answer: 2,
        explanation: "The wizard's General format re-evaluates every value and converts text to real numbers in one pass. Flash Fill and formats cannot change the underlying data type."
      },
      {
        q: "Which check proves a numeric column is genuinely numeric after cleaning?",
        options: [
          "The column looks right-aligned",
          "=ISNUMBER(C2) filled down returns TRUE for every row",
          "Ctrl + S saves without errors",
          "The cells have no fill colour"
        ],
        answer: 1,
        explanation: "ISNUMBER tests the actual data type cell by cell. Alignment is a hint, but one text value among hundreds can hide; ISNUMBER exposes it."
      },
      {
        q: "When does a cleaning job justify moving from this formula workflow to Power Query?",
        options: [
          "Never — formulas are always better",
          "When the same cleaning sequence repeats regularly or the file grows very large",
          "Only when the file has more than 10 columns",
          "As soon as any duplicate exists"
        ],
        answer: 1,
        explanation: "Power Query records your cleaning steps once and replays them on every new file with one Refresh — the right tool when the work repeats monthly or the data is too large for formula-level edits."
      }
    ],
    summary: [
      "Inspect first: alignment, green triangles, case chaos, spaces, text dates, duplicates — and write the fix order down.",
      "Fix text with `=PROPER(TRIM(...))` in helper columns, then paste as values over the originals.",
      "Convert text dates and text numbers with Text to Columns (Date / General), and verify with ISNUMBER.",
      "Structure, de-duplicate with a backup, then run the two-minute verification checklist before trusting the file.",
      "Same cleanup every month, or huge files? That is Power Query's job — Level 6."
    ],
    downloads: [
      { file: "assets/files/level3/data-cleaning.xlsx",
        label: "Data cleaning workbook",
        note: "Exercise and Solution sheets with messy customer records for cleaning practice" }
    ],
    keywords: ["data cleaning", "trim", "proper", "text to columns", "remove duplicates", "isnumber", "text numbers", "standardize dates", "power query", "cleaning workflow"]
  },

  {
    id: "l3-large-datasets",
    title: "Working with Large Datasets",
    difficulty: "Intermediate",
    duration: 14,
    version: "",
    objectives: [
      "Freeze panes so headers stay visible in thousand-row files",
      "Navigate huge sheets with Ctrl + Arrow and the Name Box",
      "Switch calculation to manual when a big file becomes slow",
      "Count records and slice data sensibly with filters"
    ],
    intro: "From here on, the files get serious: the workbook for this lesson has 1,000 rows of sales data, and real jobs start at 50,000. Everything you learned still works — but at this size, small habits decide whether Excel feels fast or frozen.\n\nThis lesson is a survival kit for big files: keep your bearings, move without scrolling, stop Excel from recalculating what you are not looking at, and slice data down to the rows that matter.",
    sections: [
      {
        heading: "Freeze panes: never lose your headers",
        blocks: [
          { type: "p", text: "On row 600 of a sales list, the columns mean nothing without the headers. Freezing locks rows (and columns) on screen while everything else scrolls." },
          { type: "steps", items: [
            "Click the cell **below** the header row and **right** of any label columns you want pinned — for `A2:G1001` data with headers in row 1, click `A2`.",
            "View > Freeze Panes > Freeze Panes.",
            "A thin line appears under row 1; scroll to row 600 and the headers are still there.",
            "Undo anytime: View > Freeze Panes > Unfreeze Panes."
          ] },
          { type: "ul", items: [
            "Freeze Top Row — one click for the standard case (header row only).",
            "Freeze First Column — pins column A when your data is very wide.",
            "Freeze Panes — the general option: everything above and left of the selected cell is pinned."
          ] },
          { type: "tip", text: "Frozen panes only affect your view, not the file's data or what colleagues see on their machines. There is no downside to using them everywhere." },
          { type: "grid", title: "Scrolled to row 600 with Freeze Top Row active — the headers from row 1 stay pinned above",
            rows: [
              ["Order ID", "Date", "Salesperson", "Region", "Revenue"],
              ["1600", "14-Sep-2026", "Sneha Iyer", "West", "26400"],
              ["1601", "14-Sep-2026", "Arjun Mehta", "North", "6750"]
            ] },
          { type: "p", text: "Whatever row you scroll to, the header names remain in view above the data, so \"is that column Revenue or Units?\" never becomes a guess. In the workbook, scroll deep into the 1,000-row list with and without freezing to feel the difference." }
        ]
      },
      {
        heading: "Tables and navigation at scale",
        blocks: [
          { type: "p", text: "Convert the 1,000 rows to a Table ([[Ctrl]] + [[T]], name it Sales) before anything else. At this size the benefits compound: the Table's structured references keep every formula correct as rows are added, filtering is built in, and the header row stays aware even without freezing (Tables keep headers in view when you scroll inside them — the column letters area shows the header names). Formulas written as `=SUMIF(Sales[Region], \"North\", Sales[Revenue])` never need range maintenance again." },
          { type: "p", text: "For movement, forget the mouse wheel. The keyboard travels by blocks of data:" },
          { type: "table",
            headers: ["Keys", "Where you land", "Use it for"],
            rows: [
              ["[[Ctrl]] + Down Arrow", "Last filled cell of the column", "Jump to the bottom of 1,000 rows instantly"],
              ["[[Ctrl]] + Up Arrow", "First filled cell of the column", "Back to the top"],
              ["[[Ctrl]] + Shift + Down Arrow", "Selects from here to the last filled cell", "Select a whole column of data for formatting or formulas"],
              ["[[Ctrl]] + Home", "Cell A1", "Reset your position"],
              ["[[Ctrl]] + End", "Last used cell of the sheet", "Discover leftover data far to the right or below"]
            ] },
          { type: "p", text: "The Name Box (the small box left of the formula bar) is the teleporter: type `G600` and press [[Enter]] to jump straight there, or type a defined name like Sales to select the whole Table. Combine with Go To Special ([[Ctrl]] + [[G]] > Special) — for example, selecting all blanks across a 1,000-row range in one step, then filling them as you practised in the Find and Replace lesson." }
        ]
      },
      {
        heading: "When the file crawls: manual calculation",
        blocks: [
          { type: "p", text: "Symptom: every edit takes seconds, the status bar shows \"Calculating (4 processors): 25%\", typing feels sticky. Cause: with automatic calculation (the default), Excel recalculates affected formulas after **every** change — and in a file with thousands of formulas, that tax hits each keystroke." },
          { type: "steps", items: [
            "Formulas tab > Calculation Options > **Manual**.",
            "Work, paste and edit freely — formulas stop updating on every change.",
            "When you want fresh results, press [[F9]] (recalculate everything) or [[Shift]] + [[F9]] (only the active sheet).",
            "**Switch back to Automatic** before saving or sharing the file."
          ] },
          { type: "warning", text: "Manual calculation changes nothing about stored values until you press F9 — a half-updated file can show old numbers next to new ones, which looks exactly like a bug. The warning sign is a cell whose stored result ignores the data it points at. Never leave a workbook on Manual without saying so; colleagues will trust stale numbers." },
          { type: "p", text: "Two more speed habits for big files: delete genuinely unused ranges (Ctrl + End shows leftovers far outside the data — select and delete those rows/columns, then save), and keep conditional formatting rules on data-sized ranges rather than entire columns, as you learned in the conditional formatting lesson." }
        ]
      },
      {
        heading: "Counting and slicing 1,000+ rows",
        blocks: [
          { type: "p", text: "First questions with any big file: how many records do I actually have, and what is in here? Count records with COUNTA on the key column:" },
          { type: "formula", name: "COUNTA",
            syntax: "=COUNTA(range)",
            example: "=COUNTA(Sales[Order ID])",
            explanation: "Counts every non-empty cell in the Order ID column — the number of records in the Table. Put it in a status cell next to the Table as a permanent record counter." },
          { type: "p", text: "Then slice with the filter skills from this level. On 1,000 rows the discipline matters more than on 50:" },
          { type: "ul", items: [
            "**Filter to one question at a time.** Region = North AND month = March gives a focused, checkable subset; six filters at once gives a subset nobody can verify.",
            "**Use a SUBTOTAL cell as a control total.** `=SUBTOTAL(9, Sales[Revenue])` above the data shows the visible total — compare it against the grand total `=SUM(Sales[Revenue])` to see how much the filter removed.",
            "**Copy filtered results to a working sheet** when a slice needs its own sorting or cleaning, as practised in the filtering lesson.",
            "**Sort inside the Table with levels** rather than eyeballing: largest revenue within each region is two clicks in the Sort dialog."
          ] },
          { type: "p", text: "Notice the pattern across this whole level: Table + filters + SUBTOTAL + structured references turn a 1,000-row wall of data into an interactive summary that updates itself. When even this is not enough — hundreds of thousands of rows, or data arriving from files and databases — that is the boundary of spreadsheet work and the doorway to Power Query and Power Pivot in Level 6." },
          { type: "quote", text: "At 1,000 rows you stop reading data and start querying it. The tools are the same ones you already know — applied with discipline." }
        ]
      }
    ],
    mistakes: [
      {
        wrong: "Scrolling with the mouse wheel to find row 850, then losing the header row and misreading columns.",
        right: "Freeze the header row and jump with Ctrl + Down Arrow or the Name Box.",
        why: "Without frozen headers, a 1,000-row file is a sea of identical-looking rows; wrong-column mistakes (Revenue vs Units) become easy and silent."
      },
      {
        wrong: "Leaving a shared workbook on Manual calculation after a speed fix.",
        right: "Switch back to Formulas > Calculation Options > Automatic before saving or sharing.",
        why: "Colleagues will see stale formula results and trust them. Manual mode is a working session tool, not a file setting to keep."
      },
      {
        wrong: "Applying conditional formatting or validation to entire columns like `A:A` on a 100,000-row file.",
        right: "Set Applies To to the real data range, or the Table columns.",
        why: "Whole-column rules make Excel evaluate millions of cells repeatedly; data-sized ranges keep large files responsive."
      }
    ],
    shortcuts: [
      { keys: "Alt + W + F + F", action: "Freeze Panes (View tab)" },
      { keys: "Ctrl + Down Arrow", action: "Jump to the last filled cell of the column" },
      { keys: "Ctrl + Shift + Down Arrow", action: "Select down to the last filled cell" },
      { keys: "Ctrl + End", action: "Jump to the last used cell of the sheet" },
      { keys: "F9", action: "Recalculate all formulas (Manual mode)" },
      { keys: "Ctrl + G", action: "Go To — type an address or use Special" }
    ],
    exercises: [
      {
        task: "Open the Large dataset sheet of the exercise workbook (1,000 rows). Freeze the header row and jump to the last record.",
        hint: "View > Freeze Top Row; Ctrl + Down Arrow from any cell in a data column.",
        solution: "View > Freeze Panes > Freeze Top Row. Click any Order ID cell, press Ctrl + Down Arrow — you land on row 1001 instantly with the headers still visible above. Ctrl + Home returns you to the top."
      },
      {
        task: "Convert the data to a Table named Sales and count the records with a formula.",
        hint: "Ctrl + T, rename in Table Design, then COUNTA on the Order ID column.",
        solution: "Ctrl + T, confirm headers, OK; name it Sales in Table Design. In a free cell type =COUNTA(Sales[Order ID]) — it shows 1000 and stays correct as rows come and go."
      },
      {
        task: "Select the whole Revenue column of data in one move and add a live filtered total above the Table.",
        hint: "Ctrl + Shift + Down Arrow for the selection; SUBTOTAL 9 respects filters.",
        solution: "Click G2 (first revenue cell), Ctrl + Shift + Down Arrow selects the column to the end. Above the data, type =SUBTOTAL(9, Sales[Revenue]) — filter to any region and watch the visible total update while =SUM(Sales[Revenue]) stays fixed."
      },
      {
        task: "The file feels slow while you edit. Switch calculation off, make several edits, then recalculate once and restore the default.",
        hint: "Formulas > Calculation Options > Manual; F9 to recalculate; switch back to Automatic.",
        solution: "Formulas tab > Calculation Options > Manual. Edit, paste and sort freely — nothing recalculates. Press F9 to refresh all results, verify a few, then set Calculation Options back to Automatic before saving."
      }
    ],
    quiz: [
      {
        q: "Which feature keeps row 1 visible while you scroll to row 900?",
        options: ["Split screen", "Freeze Panes", "Hide rows", "Zoom out"],
        answer: 1,
        explanation: "Freeze Top Row (View tab) pins the header row on screen. Split and zoom change the view differently and do not pin headers reliably."
      },
      {
        q: "What does Ctrl + Shift + Down Arrow do in a data column?",
        options: [
          "Deletes the cells below",
          "Selects from the current cell to the last filled cell of the column",
          "Sorts the column descending",
          "Jumps to the next sheet"
        ],
        answer: 1,
        explanation: "It extends the selection through the block of data — one keystroke selects a thousand cells for formatting, copying or filling."
      },
      {
        q: "A file recalculates slowly after every edit. Which setting gives relief, and what must you remember afterwards?",
        options: [
          "Delete all formulas; and nothing else",
          "Formulas > Calculation Options > Manual — and switch back to Automatic before sharing",
          "Turn off the computer's other programs; nothing to restore",
          "Save more often; nothing to restore"
        ],
        answer: 1,
        explanation: "Manual mode stops per-keystroke recalculation (F9 refreshes on demand). Leaving it on in a shared file means colleagues see stale results."
      },
      {
        q: "Which formula counts the records in a Table named Sales using its Order ID column?",
        options: [
          "=COUNT(Sales)",
          "=COUNTA(Sales[Order ID])",
          "=SUM(Sales[Order ID])",
          "=SUBTOTAL(2, Sales)"
        ],
        answer: 1,
        explanation: "COUNTA counts non-empty cells — one per record. SUM would add IDs together, and plain COUNT ignores text values like IDs."
      },
      {
        q: "Why keep a =SUBTOTAL(9, ...) cell above a large filtered table?",
        options: [
          "It replaces the Table's totals row",
          "It shows the total of visible rows, so you always know how much the filter removed",
          "It sorts the data automatically",
          "It speeds up the file"
        ],
        answer: 1,
        explanation: "SUBTOTAL respects the filter. Comparing it with the full SUM shows the size of the visible slice at a glance — a control total for disciplined slicing."
      }
    ],
    summary: [
      "Freeze the header row so 1,000-row files stay readable while you scroll.",
      "Navigate by block: Ctrl + Arrow to jump, Ctrl + Shift + Arrow to select, Name Box to teleport.",
      "Excel Tables keep large data structured: growing ranges, built-in filters, structured references.",
      "Manual calculation (Formulas > Calculation Options) unblocks slow files — press F9 for fresh numbers and restore Automatic before sharing.",
      "Count records with `=COUNTA(Sales[Order ID])` and slice 1,000+ rows with focused filters plus a SUBTOTAL control total."
    ],
    downloads: [
      { file: "assets/files/level3/large-dataset.xlsx",
        label: "Large sales dataset (1,000 rows)",
        note: "A realistic 1,000-row sales dataset for large-data practice; solutions are described in the lesson" }
    ],
    keywords: ["large datasets", "freeze panes", "ctrl arrow", "name box", "manual calculation", "f9", "counta", "large files slow", "filter large data"]
  }
);
