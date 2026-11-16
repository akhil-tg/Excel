// Excel Lab — Level 1 (Excel Fundamentals) — metadata + lessons 1-6
window.LEVEL_META = window.LEVEL_META || {};
window.LEVEL_LESSONS = window.LEVEL_LESSONS || {};

window.LEVEL_META[1] = {
  id: 1,
  title: "Excel Fundamentals",
  subtitle: "Start from zero: the interface, cells, data entry, formatting and your first formulas",
  description: "Build a rock-solid foundation: understand the Excel interface, enter and format data confidently, and write your first formulas with correct references.",
  difficulty: "Beginner",
  icon: "grid"
};

(window.LEVEL_LESSONS[1] = window.LEVEL_LESSONS[1] || []).push(
  {
    id: "l1-what-is-excel",
    title: "What Is Excel and Why Learn It?",
    difficulty: "Beginner",
    duration: 12,
    version: "",
    objectives: [
      "Explain what a spreadsheet is and what Excel is used for",
      "Describe how rows, columns and cells work together",
      "Open Excel and create a blank workbook",
      "Save a workbook with Ctrl + S in the xlsx format"
    ],
    intro: "Excel is the most used office program in the world, after word processors. Almost every office job needs it. Accounts, sales, HR and stores all use it daily.\n\nThis lesson starts from zero. You will learn what a spreadsheet actually is. You will see who uses Excel in a normal working day. Then you will open your first blank workbook and save it safely. No prior experience is needed.",
    keyIdea: "Excel is a big grid of boxes called cells, and each cell holds one piece of information.",
    sections: [
      {
        heading: "What is a spreadsheet?",
        blocks: [
          { type: "p", text: "A **spreadsheet** is a very large grid of boxes. Each box is called a **cell**. You can type a number, some text, or a date into any cell." },
          { type: "p", text: "You can also ask Excel to calculate things for you. This is the big difference from a piece of paper: the sheet does the maths." },
          { type: "p", text: "The grid is organised in **columns** and **rows**. Columns run down the sheet and are named with letters: A, B, C, and so on. Rows run across the sheet and are numbered: 1, 2, 3, and so on." },
          { type: "p", text: "Where a column and a row meet, there is a cell. The cell where column B meets row 3 is called `B3`." },
          { type: "grid", title: "A small sales list for Sunrise Traders", rows: [
            ["Item", "Price", "Qty", "Total"],
            ["Notebook", "45", "10", "450"],
            ["Pen pack", "60", "5", "300"],
            ["Printer", "12500", "1", "12500"]
          ] },
          { type: "p", text: "Look at the grid above. Each cell holds exactly one piece of information. `A2` holds the text \"Notebook\". `B2` holds the price 45. `D2` holds the total 450, which Excel calculated from the price and the quantity. One grid like this can replace pages of hand-written calculations." }
        ]
      },
      {
        heading: "Who uses Excel in real offices?",
        blocks: [
          { type: "p", text: "Excel is not just for accountants. Walk into any office and you will find someone using it. Here are real examples of how different people use it every day:" },
          { type: "table", headers: ["Role", "Typical Excel work"], rows: [
            ["Accounts clerk", "Entering daily bills and payments, adding up totals, checking if amounts match the bank statement"],
            ["HR staff", "Keeping an employee list with names, joining dates, salaries and leave taken"],
            ["Sales team", "Recording each day's orders, adding up monthly sales, comparing this month against last month"],
            ["Store and inventory", "Tracking stock of items like notebooks and printers, flagging items that are running low"],
            ["Managers", "Reading reports, checking budgets, making simple charts for meetings"]
          ] },
          { type: "p", text: "Notice something: in every case the person is doing the same three things. They are **entering data**, **calculating totals**, and **making the result look clean** for someone else to read. Those three skills are exactly what Level 1 of this course teaches." }
        ]
      },
      {
        heading: "What you will be able to do after this course",
        blocks: [
          { type: "p", text: "By the end of Level 1 you will be able to do real office work, not just theory. You will be able to:" },
          { type: "ul", items: [
            "Find your way around the Excel screen confidently and quickly",
            "Enter text, numbers and dates correctly, and fix mistakes without panic",
            "Copy, move and fill data so you never type the same thing twice",
            "Format a sheet with currency, borders and clean headings so it looks professional",
            "Write your first formulas like `=B2*C2` and `=SUM(B2:B10)`, and understand cell references"
          ] },
          { type: "p", text: "Later levels build on this: functions like `VLOOKUP`, charts, and analysing bigger data sets. But every one of those starts with the solid base you build here." }
        ]
      },
      {
        heading: "Opening Excel and creating a blank workbook",
        blocks: [
          { type: "p", text: "A file made in Excel is called a **workbook**. A workbook can hold one or more **worksheets**. A worksheet is one individual grid page. When you start a new piece of work, you create a new blank workbook." },
          { type: "steps", items: [
            "Click the Start button (or press the Windows key) and type \"Excel\". Click **Excel** in the results.",
            "Excel opens and shows the Start screen with recent files and templates.",
            "Click **Blank workbook** — the large plain option with a green plus sign.",
            "A fresh empty grid opens, waiting for your data."
          ] },
          { type: "p", text: "If Excel is already open, you can start a new workbook any time. Press [[Ctrl]] + [[N]]. A blank workbook appears immediately, on top of the one you were working in." }
        ]
      },
      {
        heading: "Saving your work — the habit that saves jobs",
        blocks: [
          { type: "p", text: "Here is a true office rule: **an unsaved file does not exist**. If the power fails before you save, everything you typed is gone. Saving takes one second, so make it a habit." },
          { type: "steps", items: [
            "Press [[Ctrl]] + [[S]]. The first time, Excel shows the Save As screen.",
            "Choose a folder you will remember, for example **Documents** or a folder called **Excel Practice**.",
            "Type a clear name, such as \"Sunrise Traders - Sales Week 1\".",
            "Keep the file type as **Excel Workbook (*.xlsx)** — that is the standard modern format.",
            "Click **Save**."
          ] },
          { type: "p", text: "After the first save, pressing [[Ctrl]] + [[S]] updates the same file silently. No dialog appears. The modern `.xlsx` format keeps formulas, formatting and multiple sheets intact. It opens in Excel 2007 and every later version." },
          { type: "tip", text: "Press [[Ctrl]] + [[S]] every few minutes. Also save just after finishing something important, like a long list or a formula. The file name stays the same; only the contents are updated." },
          { type: "warning", text: "Want to keep an original copy and work on a new version? Use **File > Save As** and give the file a new name, such as \"Sales Week 2\". If you just press [[Ctrl]] + [[S]], you overwrite the original." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Typing prices as text like \"Rs. 45\" or \"45/-\" so you can add currency symbols.", right: "Type the pure number 45, and add the currency look later with formatting (covered in the number formatting lesson).", why: "Excel cannot add text. \"Rs. 45\" looks fine but `=SUM()` will return 0 or ignore the cell, and you will not notice until the totals are wrong." },
      { wrong: "Working for an hour without saving, then the laptop battery dies.", right: "Press [[Ctrl]] + [[S]] after finishing any meaningful chunk of work.", why: "Unsaved work is lost permanently if Excel or Windows closes unexpectedly." },
      { wrong: "Giving files names like \"New sheet final 2\".", right: "Use clear names with company and period, like \"Sunrise Traders - Sales Week 1\".", why: "In a real office, colleagues must find your file later. A clear name saves everyone time, including you in three weeks." }
    ],
    shortcuts: [
      { keys: "Ctrl + N", action: "Create a new blank workbook" },
      { keys: "Ctrl + S", action: "Save the current workbook" }
    ],
    exercises: [
      { task: "In a blank workbook, rebuild the Sunrise Traders grid from this lesson: headers Item, Price, Qty, Total in row 1, then Notebook, Pen pack and Printer in rows 2 to 4. Save the file as \"Sunrise Traders - Sales Week 1.xlsx\" in your Documents folder.", hint: "Click each cell once, type, and press [[Enter]] to move down. Press [[Ctrl]] + [[S]] when finished.", solution: "A1: Item, B1: Price, C1: Qty, D1: Total; A2: Notebook, B2: 45, C2: 10; A3: Pen pack, B3: 60, C3: 5; A4: Printer, B4: 12500, C4: 1. File saved via Ctrl + S with the xlsx type." },
      { task: "Make a second workbook for your own monthly expenses: columns Month, Rent, Food, Travel. Fill in three rows of realistic amounts, then save it as \"My Expenses.xlsx\" and close Excel. Reopen the file from Documents to confirm it saved correctly.", hint: "Use [[Ctrl]] + [[N]] for the new workbook. After closing, use the Start menu search or File > Open > Recent to find it again.", solution: "Any realistic rows are fine, for example Jan / 12000 / 6500 / 2200. The key habits practised are Ctrl + N, entering a small table, Ctrl + S, closing, and reopening." }
    ],
    quiz: [
      { q: "What is a cell in Excel?", options: ["The box where a column and a row meet", "The whole Excel file", "Only the top-left box", "A saved copy of your work"], answer: 0, explanation: "A cell is a single box in the grid, named by its column letter and row number, such as B3. The whole file is called a workbook." },
      { q: "Which address points to the cell in column C, row 7?", options: ["C7", "7C", "C-7", "Row C7"], answer: 0, explanation: "Cell addresses always start with the column letter followed by the row number, so column C row 7 is C7." },
      { q: "What happens the second time you press Ctrl + S after the file is already saved?", options: ["Excel asks for a new file name every time", "A new copy of the file is created", "Nothing happens at all", "The file updates silently with no dialog"], answer: 3, explanation: "After the first Save As, Ctrl + S simply updates the existing file. No dialog appears." },
      { q: "Why should you type 45 instead of \"Rs. 45\" when entering a price you plan to add up?", options: ["Excel can only calculate pure numbers, not text", "Text uses more disk space", "Rs. 45 is not allowed in a cell", "Numbers are always shown in red"], answer: 0, explanation: "Excel treats \"Rs. 45\" as text and ignores it in calculations. Type the pure number and apply currency formatting later." },
      { q: "Which file format is the standard for saving a modern Excel workbook?", options: [".xlsx", ".doc", ".txt", ".jpg"], answer: 0, explanation: "The .xlsx format is the standard Excel Workbook format since 2007. It preserves formulas, formatting and multiple sheets." }
    ],
    summary: [
      "A spreadsheet is a grid of cells; each cell holds one value and cells are named like B3 (column letter, row number).",
      "Excel is used everywhere: accounts, HR, sales and inventory all rely on entering data, calculating and formatting.",
      "A workbook is the file; it contains one or more worksheets.",
      "Create a new workbook with Ctrl + N or Blank workbook on the Start screen.",
      "Save early and often with Ctrl + S, keep the .xlsx format, and use Save As for new versions."
    ],
    downloads: [],
    keywords: ["what is excel", "spreadsheet basics", "rows columns cells", "create blank workbook", "save excel file", "xlsx format", "excel for beginners", "why learn excel"]
  },
  {
    id: "l1-excel-interface",
    title: "The Excel Interface and Navigation",
    difficulty: "Beginner",
    duration: 15,
    version: "",
    objectives: [
      "Name the main parts of the Excel screen and what each one does",
      "Use the Name Box and formula bar to inspect any cell",
      "Navigate a large sheet with keyboard shortcuts",
      "Show, hide and use the Quick Access Toolbar and the Ribbon"
    ],
    intro: "The first time Excel opens, it shows a lot of things at once. There is a giant grid, a strip of buttons on top, small boxes above the grid, and coloured tabs at the bottom. It can feel overwhelming.\n\nThis lesson gives you a guided tour of every part of that screen. Once you can name each part and know what it is for, Excel stops being scary. It becomes a tool you control.",
    keyIdea: "The Excel screen has named parts, and each part does one clear job.",
    sections: [
      {
        heading: "The Ribbon: tabs, groups and dialog launchers",
        blocks: [
          { type: "p", text: "The **Ribbon** is the wide strip of buttons across the top. It is divided into **tabs**: Home, Insert, Page Layout, Formulas, Data, Review and View. Each tab holds buttons for one kind of task. Beginners spend most of their time on the **Home** tab. Font, colour, number format, alignment and borders are all there." },
          { type: "p", text: "Within a tab, related buttons sit together in **groups** with a small label underneath, such as Font, Alignment and Number on the Home tab." },
          { type: "table", headers: ["Tab", "What you find there"], rows: [
            ["Home", "Formatting: bold, font size, fill colour, borders, number formats, cell styles"],
            ["Insert", "Charts, pictures, tables, and other objects you add to the sheet"],
            ["Page Layout", "Print settings: margins, orientation, print area"],
            ["Formulas", "The function library and tools for checking calculations"],
            ["Data", "Sorting, filtering, text to columns and data tools"],
            ["Review", "Spell check, comments and sheet protection"],
            ["View", "Zoom, freezing rows, and switching window layout"]
          ] },
          { type: "p", text: "Some groups have a tiny arrow in the bottom-right corner. That is a **dialog launcher**. Clicking it opens the full dialog box with every option for that group. Example: on the Home tab, click the launcher in the Number group. The full **Format Cells** dialog opens. It is the same dialog you get with [[Ctrl]] + [[1]], which you will meet properly in the number formatting lesson." }
        ]
      },
      {
        heading: "The Name Box and the formula bar",
        blocks: [
          { type: "p", text: "Directly above the grid on the left is a small white box called the **Name Box**. It always shows the address of the **active cell** — the one cell with the thick green border. If it shows `B4`, then B4 is active. You can also type an address like `D25` into the Name Box. Press [[Enter]] to jump straight there." },
          { type: "p", text: "Next to it is the **formula bar**, a long white strip. When a cell is active, this bar shows what is really inside the cell. This matters because cells can display one thing and contain another. A cell may show \"January\" but actually contain a date value. Or it may show 450 because it contains the formula `=B2*C2`. Click any cell and read the formula bar to see the truth." },
          { type: "grid", title: "What you see vs what the formula bar shows", rows: [
            ["", "B", "C"],
            ["2", "Notebook", "45"],
            ["3", "Pen pack", "60"],
            ["4", "Total", "105"]
          ] },
          { type: "p", text: "In the grid above, cell `C4` displays **105**. If you click on it, the formula bar shows `=C2+C3`. The grid shows the result; the formula bar shows the instruction. Getting comfortable with this difference is one of the most important habits in Excel." },
          { type: "p", text: "The many small rectangles filling the rest of the screen are **cells**. Only one is the active cell at a time (though you can select many, as the next lesson shows). Everything you type goes into the active cell." }
        ]
      },
      {
        heading: "Status bar, zoom and the Quick Access Toolbar",
        blocks: [
          { type: "p", text: "At the very bottom of the window is the **status bar**. It quietly gives useful information. It shows the current zoom level. When you select a group of numbers, it also shows instant totals like **Sum**, **Count** and **Average** — with no formula. Try it: select a few cells with numbers and look at the bottom-right corner." },
          { type: "p", text: "The zoom controls sit on the status bar too. Drag the slider, or click the plus and minus buttons. You can zoom between roughly 10% and 400%. Zoom changes only how big things look on screen. It never changes your data or the printed size." },
          { type: "p", text: "Above the Ribbon, at the top-left, is the **Quick Access Toolbar** (QAT). It is a tiny row of one-click buttons. By default it holds **Save**, **Undo** and **Redo**. You can add any command by right-clicking a Ribbon button and choosing \"Add to Quick Access Toolbar\". Many experienced users put their favourite tools there. Then they never have to hunt through tabs." },
          { type: "tip", text: "Press [[Alt]] and Excel shows small letter labels over each tab (H for Home, N for Insert, and so on). Press the letter to open that tab using only the keyboard. This is called KeyTips. Fast office workers use the Ribbon this way." }
        ]
      },
      {
        heading: "Sheet tabs at the bottom",
        blocks: [
          { type: "p", text: "Below the grid, at the bottom-left, are the **sheet tabs**. Each tab is one worksheet in the workbook. A new workbook starts with a sheet usually named \"Sheet1\" (sometimes with extras, depending on your version). Click a tab to switch to that sheet. The next lesson shows you how to add, rename, colour and rearrange these tabs. It is a real tidy-up skill that office files use all the time." }
        ]
      },
      {
        heading: "Navigating with mouse and keyboard",
        blocks: [
          { type: "p", text: "You can always click a cell with the mouse. But in a sheet with thousands of rows, clicking and scrolling is slow. Keyboard navigation is faster, and office employers notice the difference." },
          { type: "table", headers: ["Keys", "What it does"], rows: [
            ["Arrow keys", "Move the active cell one cell up, down, left or right"],
            ["Ctrl + Arrow key", "Jump to the edge of the data block — end of a list, or bottom of a column of numbers"],
            ["Ctrl + Home", "Jump back to cell A1"],
            ["Page Down / Page Up", "Move one screenful down or up"],
            ["Name Box", "Type any address (like F500) and press Enter to jump there"]
          ] },
          { type: "p", text: "**Ctrl + Arrow** is worth practising until it is automatic. Say your data has 3,000 rows of sales. Pressing [[Ctrl]] + [[Down]] once lands you on the last row instantly. Press [[Ctrl]] + [[Home]] to return to the top. Fast workers fly through big sheets this way. They never touch the scrollbar." },
          { type: "p", text: "One more useful control: the small **^** arrow button at the top-right of the Ribbon. (You can also right-click the Ribbon and choose \"Collapse the Ribbon\".) This hides the button area and leaves only the tab names. A wide sheet then has more room. Click any tab name to show the buttons temporarily. Nothing is lost — it is only a display change." },
          { type: "warning", text: "When arrow keys seem to move the whole sheet instead of the active cell, **Scroll Lock** is turned on. Look for \"Scroll Lock\" on the status bar and turn the key off on your keyboard." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Typing a formula into a cell, seeing a strange result, and assuming Excel is broken.", right: "Click the cell and read the formula bar to see what is actually stored there.", why: "The grid shows results; the formula bar shows contents. Almost every \"Excel is wrong\" mystery is solved by reading the formula bar." },
      { wrong: "Hunting through every tab for one command each time.", right: "Learn which tab holds your common commands, or add them to the Quick Access Toolbar.", why: "Experienced users reach Home commands and Save in one or two actions. Small savings repeated all day become big speed gains." },
      { wrong: "Scrolling manually to the bottom of a 5,000-row list.", right: "Press [[Ctrl]] + [[Down]] to jump to the last filled row instantly.", why: "Keyboard jumps are immediate; scrolling takes many seconds and loses your place." }
    ],
    shortcuts: [
      { keys: "Ctrl + Home", action: "Jump to cell A1" },
      { keys: "Ctrl + Down / Up / Left / Right", action: "Jump to the edge of the current data block" },
      { keys: "Page Down / Page Up", action: "Move one screen up or down" },
      { keys: "Alt", action: "Show KeyTips to use the Ribbon from the keyboard" },
      { keys: "Ctrl + F1", action: "Collapse or expand the Ribbon" }
    ],
    exercises: [
      { task: "Open a blank workbook and take the tour: point at and name the Ribbon, Name Box, formula bar, active cell, status bar and sheet tabs without looking at this lesson. Then press [[Ctrl]] + [[Home]] and confirm the Name Box shows A1.", hint: "The Name Box sits above the grid on the left; the formula bar is the long strip next to it.", solution: "You should be able to identify: Ribbon tabs (top), Quick Access Toolbar (above the Ribbon), Name Box and formula bar (above the grid), sheet tabs (bottom-left), status bar (bottom). Ctrl + Home makes A1 active." },
      { task: "In the Beginner practice workbook, go to the Exercise sheet. Select the column of prices with numbers, then look at the status bar and write down the Sum and Count shown there. Next jump to the very last row of the data using only the keyboard.", hint: "Click the first price cell, then use [[Ctrl]] + [[Down]] to reach the end. The status bar totals appear whenever more than one cell is selected.", solution: "Selecting the price cells shows live Sum, Count and Average on the status bar with no formula needed. Ctrl + Down jumps from the first cell to the last filled cell in the column." },
      { task: "Collapse the Ribbon, then bring it back. Then press [[Alt]] and use the KeyTips to open the Home tab without touching the mouse.", hint: "Use the small arrow at the Ribbon's top-right corner, or [[Ctrl]] + [[F1]]. For KeyTips, press [[Alt]] then the letter shown on the tab.", solution: "Ctrl + F1 toggles the Ribbon. Pressing Alt shows KeyTips; pressing H opens the Home tab. Press Esc to cancel." }
    ],
    quiz: [
      { q: "What does the formula bar show?", options: ["What is actually stored in the active cell", "The sum of the selected cells", "The name of the current sheet", "Only errors in the workbook"], answer: 0, explanation: "The formula bar shows the true contents of the active cell — a formula, a number, or text — even when the cell displays something different." },
      { q: "Which shortcut jumps straight to cell A1?", options: ["Ctrl + A", "Alt + Home", "Ctrl + Home", "F1"], answer: 2, explanation: "Ctrl + Home always returns the active cell to A1. Ctrl + A selects the whole sheet instead." },
      { q: "You select several cells containing numbers and look at the bottom of the window. What do you see?", options: ["Live Sum, Count and Average on the status bar", "A printed report", "The formula bar totals", "Nothing unless you type a formula first"], answer: 0, explanation: "The status bar shows instant statistics for the selected numbers, with no formula required." },
      { q: "What is a dialog launcher?", options: ["The small arrow in a group corner that opens the full dialog with all options", "A button that starts Excel", "The green plus on a new sheet", "A key on the keyboard"], answer: 0, explanation: "Dialog launchers are the tiny bottom-right arrows in Ribbon groups. They open the complete dialog, like Format Cells from the Number group." },
      { q: "Pressing Ctrl + Down while a column of data is selected does what?", options: ["Jumps to the last filled cell of that column", "Deletes the column", "Copies the data down", "Sorts the column"], answer: 0, explanation: "Ctrl + Arrow keys jump to the edge of the current data block — a fast way to reach the end of long lists." }
    ],
    summary: [
      "The Ribbon holds tabs, tabs hold groups, and small corner arrows open full dialogs.",
      "The Name Box shows and jumps to cell addresses; the formula bar shows a cell's true contents.",
      "The status bar gives instant Sum, Count and Average for any selected numbers.",
      "Ctrl + Arrow keys and Ctrl + Home navigate large sheets far faster than scrolling.",
      "The Quick Access Toolbar keeps your favourite commands one click away."
    ],
    downloads: [],
    keywords: ["excel interface", "ribbon tabs", "formula bar", "name box", "status bar", "quick access toolbar", "sheet tabs", "excel navigation shortcuts", "collapse ribbon", "keytips"]
  },
  {
    id: "l1-workbooks-cells",
    title: "Workbooks, Worksheets, Rows, Columns and Cells",
    difficulty: "Beginner",
    duration: 15,
    version: "",
    objectives: [
      "Explain the difference between a workbook and a worksheet",
      "Add, rename, move, copy, colour and delete sheets",
      "Select cells, ranges, whole rows, columns and the entire sheet",
      "Adjust column width and row height, and hide or unhide them"
    ],
    intro: "In the last lesson you learned to find your way around the screen. Now we go one level deeper. Let us see how an Excel file is organised inside.\n\nBy the end you will handle worksheets like tabs in a paper file. You will also select and resize any part of the grid with confidence. These small skills make every later lesson easier.",
    keyIdea: "A workbook is the file you save; worksheets are the tabbed pages inside it.",
    sections: [
      {
        heading: "Workbook vs worksheet",
        blocks: [
          { type: "p", text: "Keep these two words separate in your mind:" },
          { type: "ul", items: [
            "A **workbook** is the whole file — the thing you save with [[Ctrl]] + [[S]]. Its extension is `.xlsx`.",
            "A **worksheet** is one grid page inside that file. The tabs at the bottom each represent one worksheet."
          ] },
          { type: "p", text: "Here is a useful real-life picture. The workbook is a ring binder; each worksheet is a sheet of paper inside it. An office file often has several worksheets, for example \"Sales\", \"Expenses\" and \"Summary\". All of them are saved inside one workbook. New workbooks start with one worksheet called Sheet1." }
        ]
      },
      {
        heading: "Managing sheets: add, rename, move, copy, colour, delete",
        blocks: [
          { type: "p", text: "All sheet management starts from the sheet tabs at the bottom-left, or from right-clicking a tab. Here is the full toolkit:" },
          { type: "table", headers: ["Action", "How"], rows: [
            ["Add a new sheet", "Click the **+** button to the right of the tabs, or press Shift + F11"],
            ["Rename a sheet", "Double-click the tab, type the new name, press Enter"],
            ["Delete a sheet", "Right-click the tab and choose Delete (Excel warns you — it cannot be undone)"],
            ["Move a sheet", "Drag the tab left or right"],
            ["Copy a sheet", "Right-click the tab, choose Move or Copy, tick \"Create a copy\" (or drag the tab while holding Ctrl)"],
            ["Colour a tab", "Right-click the tab, choose Tab Color, pick a colour"]
          ] },
          { type: "p", text: "Renaming sheets to meaningful names is more than decoration. A sheet called \"Expenses\" tells any colleague exactly what is inside. Formulas that read data across sheets also become easier to understand, like `=Expenses!B10` (the `!` separates the sheet name from the cell)." },
          { type: "tip", text: "Colour the tabs of a shared workbook: green for data that people enter, blue for calculated reports, grey for old archived data. Colleagues learn the colour code within a day." }
        ]
      },
      {
        heading: "Cells and their addresses",
        blocks: [
          { type: "p", text: "The grid is made of **rows** and **columns**. Rows are numbered 1, 2, 3, ... down to more than a million. Columns are lettered A, B, C, ... to XFD — over 16,000 of them. Every cell has an **address**: column letter first, then row number. `C8` means column C, row 8. Row 8 is the eighth row down; column C is the third column from the left." },
          { type: "grid", title: "Addresses in a small sheet", rows: [
            ["", "A", "B", "C"],
            ["1", "Employee", "Department", "Salary"],
            ["2", "Arjun Mehta", "Accounts", "32000"],
            ["3", "Priya Nair", "Sales", "28500"]
          ] },
          { type: "p", text: "In this grid, \"Arjun Mehta\" lives at `A2`. His salary 32000 lives at `C2`. You will use these addresses constantly. Formulas refer to cells by address — you will write your first ones in a few lessons." }
        ]
      },
      {
        heading: "Selecting cells, ranges, rows and columns",
        blocks: [
          { type: "p", text: "Before formatting, copying or deleting, you **select** what you want to act on. Selection is a skill worth practising until it feels automatic:" },
          { type: "table", headers: ["To select", "Do this"], rows: [
            ["One cell", "Click it"],
            ["A range (a block of cells)", "Click the first cell, hold the mouse button, drag to the last cell — e.g. B2 to D5 is written `B2:D5`"],
            ["Several separate cells", "Hold Ctrl and click each one"],
            ["A whole row", "Click the row number on the left edge"],
            ["A whole column", "Click the column letter at the top"],
            ["The entire sheet", "Click the small triangle above the row numbers, left of column A — or press Ctrl + A"]
          ] },
          { type: "p", text: "Notice the colon: `B2:D5` means \"from B2 to D5\". This notation appears everywhere in Excel, especially inside formulas like `=SUM(B2:B10)`." },
          { type: "tip", text: "Keyboard selection: click a cell, hold **Shift**, then press the arrow keys. The selection grows cell by cell. Shift + Ctrl + Arrow extends the selection to the end of the data. It is ideal for selecting a 2,000-row column in one action." }
        ]
      },
      {
        heading: "Column width and row height",
        blocks: [
          { type: "p", text: "Text often refuses to fit. A cell shows **#####** when a number is too wide for its column. That is Excel saying \"widen me\". Your data is not lost. To fix widths and heights:" },
          { type: "ul", items: [
            "**Auto-fit**: double-click the border line between two column letters (for example between B and C). The column snaps to the widest content. Same trick works for rows.",
            "**Exact size**: right-click the column letter, choose Column Width, and type a number. Right-click a row number for Row Height.",
            "**Drag**: drag the border between column letters to size it by eye."
          ] },
          { type: "p", text: "To change several columns at once, select them all first. Click the first letter, then drag or Ctrl+click the rest. Then apply any of the methods above. A tidy report has columns wide enough for their content. Nothing is cut off, and there are no giant empty gaps." }
        ]
      },
      {
        heading: "Hiding and unhiding",
        blocks: [
          { type: "p", text: "Sometimes a sheet has columns colleagues do not need to see. Maybe they hold old prices or private salary details. But you do not want to delete them. **Hiding** removes them from view and keeps all data and formulas safe." },
          { type: "steps", items: [
            "Select the column letter(s) to hide.",
            "Right-click the selection and choose **Hide**. The column disappears; the letter line skips straight from B to D if C is hidden.",
            "To bring it back, select the columns on both sides (B and D), right-click, and choose **Unhide**."
          ] },
          { type: "warning", text: "Hidden is not secret. Anyone who knows how to unhide can see hidden columns and sheets. Never rely on hiding to protect confidential data. A company should control access to the file itself." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Seeing ##### in a cell and retyping the number.", right: "Widen the column by double-clicking its right border.", why: "##### only means \"the number does not fit on screen\". The value is safe; the column is simply too narrow." },
      { wrong: "Deleting a sheet you think you do not need, then realising formulas on other sheets pointed at it.", right: "Hide the sheet instead of deleting, until you are certain nothing depends on it.", why: "Deleting a sheet breaks every formula that referenced it, turning them into #REF! errors, and the deletion cannot be undone." },
      { wrong: "Writing a range as B2 to D5 or B2-D5 in formulas.", right: "Use the colon: `B2:D5`.", why: "The colon is Excel's official \"from-to\" notation. Every range-based function, like `=SUM(B2:D5)`, expects it." }
    ],
    shortcuts: [
      { keys: "Ctrl + A", action: "Select the entire sheet (press again to select the region)" },
      { keys: "Shift + F11", action: "Insert a new worksheet" },
      { keys: "Shift + Arrow keys", action: "Extend the selection one cell at a time" },
      { keys: "Ctrl + Shift + Arrow keys", action: "Extend the selection to the end of the data" }
    ],
    exercises: [
      { task: "In the Beginner practice workbook, add a new sheet, rename it \"My Notes\", and colour its tab green. Then move it so it sits first, before the Exercise sheet.", hint: "Use the + button for the new sheet, double-click to rename, right-click for Tab Color, and drag the tab to reorder.", solution: "Click +, double-click the new tab and type My Notes, right-click > Tab Color > green, then drag the tab left of Exercise." },
      { task: "In the Exercise sheet of the Beginner practice workbook, select the whole data range with the keyboard only (click the first cell, then Ctrl + Shift + Arrow down and right). Auto-fit every column you used.", hint: "Ctrl + Shift + Down selects to the bottom of the column; Ctrl + Shift + Right extends across. Double-click borders between column letters to auto-fit.", solution: "The selection covers the full data block, and each column snaps to fit its widest entry, so nothing shows ##### or is cut off." },
      { task: "In the Exercise sheet, hide one column, confirm the letters skip it, then unhide it again.", hint: "Select both neighbouring columns before right-clicking to Unhide.", solution: "Right-click the column letter > Hide, then select the columns either side > right-click > Unhide. The data is untouched throughout." }
    ],
    quiz: [
      { q: "What is the difference between a workbook and a worksheet?", options: ["The workbook is the file; worksheets are the grid pages inside it", "They are the same thing", "A worksheet is the saved file; a workbook is one page", "Worksheets hold formulas only"], answer: 0, explanation: "The workbook is the whole .xlsx file. Each tab at the bottom is one worksheet within it." },
      { q: "What does the range B2:D5 mean?", options: ["Every cell from B2 across to D5", "Only the four corner cells", "Four separate single cells", "A formula error"], answer: 0, explanation: "The colon means \"through\". B2:D5 is the whole rectangle of cells with B2 at the top-left and D5 at the bottom-right." },
      { q: "A cell shows #####. What should you do?", options: ["Delete the number and retype it", "Change the number to text", "Press Ctrl + Z", "Widen the column"], answer: 3, explanation: "##### means the number is too wide to display. Auto-fit or widen the column and the number reappears." },
      { q: "How do you copy a worksheet inside the same workbook?", options: ["Right-click the tab, choose Move or Copy, tick \"Create a copy\"", "Ctrl + S the workbook twice", "Retype the data on a new sheet", "Hide the sheet"], answer: 0, explanation: "Move or Copy with \"Create a copy\" duplicates the whole sheet, data and formatting included. Dragging the tab while holding Ctrl also works." },
      { q: "Why is hiding a column not a safe way to protect salary data?", options: ["Anyone can unhide it in seconds", "Hidden columns are deleted at midnight", "Excel forbids hiding numeric columns", "Hidden data prints automatically"], answer: 0, explanation: "Hiding is display only. Unhide takes two clicks, so confidentiality needs real file protection, not hidden columns." }
    ],
    summary: [
      "A workbook is the saved file; worksheets are its pages, managed from the tabs at the bottom.",
      "Sheets can be added, renamed, moved, copied, coloured and deleted — rename them meaningfully.",
      "Every cell has an address like C8: column letter then row number.",
      "Select with clicks, Shift, and Ctrl; ranges are written like B2:D5.",
      "##### means the column is too narrow; double-click borders to auto-fit. Hide keeps data out of view but not out of reach."
    ],
    downloads: [],
    keywords: ["workbook vs worksheet", "add rename delete sheet", "move copy worksheet", "tab color", "cell address", "select range", "column width row height", "hide unhide column", "a1 notation"]
  },
  {
    id: "l1-data-entry",
    title: "Entering and Editing Data",
    difficulty: "Beginner",
    duration: 15,
    version: "",
    objectives: [
      "Enter text, numbers and dates and know how Excel treats each type",
      "Edit a cell with F2 or double-click without retyping",
      "Undo and redo mistakes confidently",
      "Enter data quickly across a range using Tab and Enter"
    ],
    intro: "Data entry is the bread and butter of office Excel work. Somebody has to type the bills, the employee list and the stock counts. Good typing makes every later calculation work.\n\nThis lesson teaches the rules Excel applies to what you type. You will learn to fix mistakes without retyping whole cells. You will also pick up small tricks that make long entry sessions fast and accurate.",
    keyIdea: "Type into a cell and press Enter; Excel decides if it is text, number or date.",
    sections: [
      {
        heading: "Typing text, numbers and dates",
        blocks: [
          { type: "p", text: "Type into the active cell. Press [[Enter]] to confirm and move down. Press [[Tab]] to confirm and move right. Press [[Esc]] if you change your mind mid-typing. As you type, Excel quietly decides what kind of data it is:" },
          { type: "table", headers: ["You type", "Excel stores it as", "Alignment in the cell"], rows: [
            ["Sunrise Traders", "Text", "Left edge"],
            ["1250", "Number", "Right edge"],
            ["14-Mar-2026", "Date (a number)", "Right edge"]
          ] },
          { type: "p", text: "The alignment is the giveaway. **Text hugs the left; numbers and dates hug the right.** This is not decoration — it is Excel telling you what it stored. A value you expected as a number but see left-aligned has been stored as text. Maths on it will fail." },
          { type: "grid", title: "Alignment reveals the data type", rows: [
            ["", "A", "B"],
            ["1", "Item", "Notebook"],
            ["2", "Price", "45"],
            ["3", "Ordered on", "14-Mar-2026"]
          ] },
          { type: "p", text: "In this grid, \"Notebook\" sits left, while 45 and 14-Mar-2026 sit right. Both right-aligned values are numbers under the hood. Dates are just numbers wearing a date format. You will see why that matters in the number formatting lesson." },
          { type: "p", text: "Dates Excel understands include `14-Mar-2026`, `14/3/2026` (depending on your regional settings), `Mar 14, 2026` and `14-Mar`. That last one is handy: typing just `14-Mar` stores this year's 14 March. Always type dates in a form Excel recognises. Then later lessons can sort and calculate with them." },
          { type: "warning", text: "Never enter a date as text like \"March 14th\" by hand with a trailing \"th\" — Excel stores it as text, sorts it wrongly, and refuses to subtract it from other dates." }
        ]
      },
      {
        heading: "Autocomplete for repeated text",
        blocks: [
          { type: "p", text: "Type the first letters of an entry you have already typed in the same column. Excel offers to finish it. Say \"Sunrise Traders\" exists above. Type \"Sun\" and Excel suggests the rest, selected after your letters. Press [[Enter]] to accept, or keep typing to override." },
          { type: "p", text: "In a real office list — employee names, supplier names, item names — autocomplete saves enormous typing time. It also keeps spellings consistent. \"Sunrise Traders\" never becomes \"Sunrise Trader\" halfway down the column." }
        ]
      },
      {
        heading: "Editing a cell: F2 and double-click",
        blocks: [
          { type: "p", text: "Beginners often retype an entire cell to fix one wrong digit. Do not. Two easy ways exist to **edit** what is already there:" },
          { type: "ul", items: [
            "Select the cell and press **F2** — the cursor goes into the cell and you edit in place.",
            "Or **double-click** the cell — the same thing, using the mouse."
          ] },
          { type: "p", text: "In edit mode you can move the cursor with arrows. Delete single characters with [[Backspace]] or [[Delete]], and insert new text. Press [[Enter]] to confirm the edit, or [[Esc]] to abandon it and keep the original. There is also a third way: click the cell and edit directly in the **formula bar**. That is comfortable for long text or formulas." },
          { type: "tip", text: "A single click on a cell selects it; a double-click (or F2) edits it. Knowing which one you want prevents the most common beginner fumbles, like accidentally typing over a good cell." }
        ]
      },
      {
        heading: "Deleting contents: Delete key vs Clear",
        blocks: [
          { type: "p", text: "Select cell(s) and press the **Delete** key. The contents vanish, but the cell itself remains, with its formatting. This is what you want 95% of the time: clearing a value while keeping the sheet structure." },
          { type: "p", text: "The difference matters when formatting is involved. Suppose a cell was coloured yellow and bold. Pressing [[Delete]] empties the value, but the yellow bold look remains. That confuses the next person. To remove value **and** formatting together, select the cells. Then choose **Home > Clear** (the little eraser icon) and pick **Clear All**." },
          { type: "table", headers: ["Command", "Removes"], rows: [
            ["Delete key", "Contents only — formatting stays"],
            ["Clear All", "Contents and formatting"],
            ["Clear Formats", "Formatting only — the value stays"]
          ] }
        ]
      },
      {
        heading: "Undo and redo: your safety net",
        blocks: [
          { type: "p", text: "Press **[[Ctrl]] + [[Z]]** to undo your last action. Press it again to undo the one before, and so on — Excel remembers a long history. Went too far back? **[[Ctrl]] + [[Y]]** redoes what you undone." },
          { type: "p", text: "Undo covers typing, deleting, formatting, moving and even sheet actions. But it has limits: it does **not** reverse a saved-and-closed session, and it cannot bring back a deleted sheet. So undo is your first-aid kit for small accidents, while [[Ctrl]] + [[S]] (saving) remains your real insurance." },
          { type: "tip", text: "The Quick Access Toolbar's curved arrow button is the same as Ctrl + Z, and the small dropdown next to it lets you jump back several steps at once." }
        ]
      },
      {
        heading: "Entering data in a range: Tab and Enter teamwork",
        blocks: [
          { type: "p", text: "Here is a trick that makes filling a table much smoother. Select the whole range you want to fill first — say `B2:D5`. Now:" },
          { type: "ul", items: [
            "Press [[Tab]] after each entry — the active cell moves right across the row.",
            "At the end of a row, [[Tab]] wraps the cursor to the start of the next row.",
            "Press [[Enter]] after an entry and the cursor moves **down** the column — and at the column's end, wraps to the top of the next column."
          ] },
          { type: "p", text: "Because you selected the range first, the cursor never wanders into empty cells outside your table. You can fill an entire grid without touching the mouse — type, Tab, type, Tab, Enter at the row end, and so on." },
          { type: "grid", title: "Range B2:D5 selected before typing", rows: [
            ["", "B", "C", "D"],
            ["2", "Item 1", "Item 2", "Item 3"],
            ["3", "Item 4", "Item 5", "Item 6"],
            ["4", "Item 7", "Item 8", "Item 9"]
          ] },
          { type: "p", text: "Real use: entering a week of bills for Sunrise Traders. Select the seven cells of the week, type each amount with [[Enter]], and Excel walks down the column for you, stopping exactly at the end." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Retyping a whole cell to fix one character.", right: "Press F2 or double-click and edit only the wrong part.", why: "Editing is faster and avoids introducing new typos into text that was otherwise correct." },
      { wrong: "Pressing Delete on a formatted cell and wondering why it is still yellow and bold.", right: "Use Home > Clear > Clear All when you want the cell visually empty too.", why: "The Delete key removes only contents. Formatting survives and can mislead whoever uses the sheet next." },
      { wrong: "Typing amounts with a currency symbol like \"$1250\" and then trying to add them.", right: "Type 1250 and apply currency formatting afterwards.", why: "A leading symbol makes Excel store the entry as text, and SUM quietly ignores text." },
      { wrong: "Entering dates as \"14.3.26\" or \"March 14th\".", right: "Use a format Excel recognises, like 14-Mar-2026 or 14/3/2026.", why: "Unrecognised dates become text: they sort wrong and cannot be used in date calculations." }
    ],
    shortcuts: [
      { keys: "F2", action: "Edit the active cell in place" },
      { keys: "Delete", action: "Clear the contents of the selected cells" },
      { keys: "Ctrl + Z", action: "Undo the last action" },
      { keys: "Ctrl + Y", action: "Redo the last undone action" },
      { keys: "Esc", action: "Cancel the current typing or edit" },
      { keys: "Enter / Tab", action: "Confirm the entry and move down / right" }
    ],
    exercises: [
      { task: "In the Beginner practice workbook's Exercise sheet, complete the data entry table: fill the missing employee names, prices and order dates. Watch the alignment — confirm every number and date you enter sits on the right edge.", hint: "Press [[Enter]] to move down a column. If an entry sits left-aligned, Excel stored it as text — check what you typed.", solution: "All numbers and dates right-align once entered correctly. Any left-aligned number usually means a stray symbol or space was typed with it." },
      { task: "One price in the Exercise sheet was typed with the text \"Rs.\" in front of it. Fix that cell without retyping the other cells, so the value becomes a pure number.", hint: "Double-click the cell (or press F2), delete the \"Rs. \" part with Backspace, and press Enter.", solution: "After removing the text prefix, the value right-aligns, showing it is now a proper number that formulas can use." },
      { task: "Select the empty range for the week's expenses in the Exercise sheet and fill it column by column using Enter, without touching the mouse. Then make a typing mistake on purpose and undo it with Ctrl + Z.", hint: "Select the range first so the cursor stays inside it. Ctrl + Z removes your last entry; Ctrl + Y brings it back.", solution: "The cursor steps down the selected column after each Enter and stops at the range edge. Ctrl + Z undoes the last entry and Ctrl + Y restores it." }
    ],
    quiz: [
      { q: "You type a product code into a cell and it hugs the left edge. What does that tell you?", options: ["Excel stored it as text", "The cell is broken", "The font is wrong", "The cell is merged"], answer: 0, explanation: "Text aligns left by default; numbers and dates align right. Left alignment is Excel's signal that it stored the entry as text." },
      { q: "Which key edits the active cell without retyping everything?", options: ["F4", "Ctrl + E", "F2", "Alt + E"], answer: 2, explanation: "F2 (or a double-click) puts the cursor inside the cell so you can edit the existing content in place." },
      { q: "What is the difference between pressing Delete and Home > Clear > Clear All?", options: ["Delete removes contents only; Clear All removes contents and formatting", "There is no difference", "Delete removes formatting only", "Clear All deletes the whole sheet"], answer: 0, explanation: "The Delete key empties values but leaves colours, borders and bold. Clear All strips both contents and formatting." },
      { q: "You selected the range B2:D5 before typing. After filling D2 and pressing Tab, where does the cursor go?", options: ["To B3, the start of the next row inside the range", "To E2, outside the range", "To D5", "To A1"], answer: 0, explanation: "Inside a selected range, Tab wraps to the first cell of the next row of the range, keeping you inside the table." },
      { q: "Which shortcut undoes your last action?", options: ["Ctrl + Z", "Ctrl + Y", "Ctrl + U", "Ctrl + Backspace"], answer: 0, explanation: "Ctrl + Z steps back through your recent actions; Ctrl + Y redoes what you undone." }
    ],
    summary: [
      "Text aligns left; numbers and dates align right — alignment reveals the data type.",
      "Autocomplete finishes repeated column entries and keeps spelling consistent.",
      "F2 or double-click edits a cell in place; Delete empties contents, Clear All also removes formatting.",
      "Ctrl + Z undoes and Ctrl + Y redoes — but saving is still the real insurance.",
      "Select a range before typing and Tab/Enter walk you through it perfectly."
    ],
    downloads: [
      { file: "assets/files/level1/beginner-practice.xlsx", label: "Beginner practice workbook", note: "Exercise and Solution sheets covering data entry, formatting, autofill and copy/paste tasks — the Exercise sheet contains the data entry table used in this lesson" }
    ],
    keywords: ["entering data in excel", "edit cell f2", "delete vs clear", "undo redo excel", "autocomplete", "tab enter range entry", "dates in excel", "text vs numbers alignment"]
  },
  {
    id: "l1-copy-cut-paste",
    title: "Copy, Cut, Paste and Paste Special",
    difficulty: "Beginner",
    duration: 18,
    version: "",
    objectives: [
      "Copy, cut and paste cells with Ctrl + C, Ctrl + X and Ctrl + V",
      "Move data with drag-and-drop and understand the marching ants border",
      "Choose the right paste option: keep formatting, values only, or transpose",
      "Copy formulas down a column safely"
    ],
    intro: "Nobody in a real office types the same data twice. Copying and moving cells is where Excel starts to feel genuinely powerful — and where beginners occasionally create chaos by pasting over the wrong thing.\n\nThis lesson covers the full copy-and-paste toolkit, including Paste Special, one of Excel's most useful hidden gems. Everything here works on the practice workbook, so follow along as you read.",
    keyIdea: "Paste Special lets you choose exactly what arrives: values only, formulas, formatting, or a transposed layout.",
    sections: [
      {
        heading: "Copy and paste: the basics",
        blocks: [
          { type: "p", text: "The trio works like this: select cells, press **[[Ctrl]] + [[C]]** to copy (or **[[Ctrl]] + [[X]]** to cut), click the destination, then **[[Ctrl]] + [[V]]** to paste." },
          { type: "p", text: "The moment you copy or cut, a moving dashed border — called **marching ants** — appears around the selection. That border means \"these cells are on the clipboard\". It stays there until you press [[Esc]], start typing something new, or perform an action that clears the clipboard. While the ants are marching, you can paste the same cells to several destinations one after another." },
          { type: "grid", title: "Before and after copying C2:C3 to E2:E3", rows: [
            ["", "C", "D", "E"],
            ["1", "Price", "", "Price"],
            ["2", "45", "", "45"],
            ["3", "60", "", "60"]
          ] },
          { type: "p", text: "Copying brings **everything**: the value, the formatting (bold, colours, number formats) and any formula — with its references adjusted, as you will see in the references lesson. Paste does not just duplicate numbers; it duplicates looks and behaviour too." }
        ]
      },
      {
        heading: "Cut, move and drag-and-drop",
        blocks: [
          { type: "p", text: "**Cut** ([[Ctrl]] + [[X]]) is copy with a difference: when you paste, the source cells are emptied. Use cut when you want to **move** data — for example, moving a block of expenses from one area of the sheet to another. Use copy when you want the data to exist in both places." },
          { type: "p", text: "There is a mouse way to move, too: select the cells, point at the **edge** of the selection until the cursor becomes a four-headed arrow, then drag to the new location and drop. Hold **Ctrl** while dragging to copy instead of move. Drag-and-drop is quick for short distances on a visible sheet; cut-and-paste is more reliable for long moves or between sheets." },
          { type: "warning", text: "Pasting onto cells that already contain data **silently replaces** what was there. There is no merge and no confirmation. If you paste over the wrong spot, press [[Ctrl]] + [[Z]] immediately — undo works on paste." }
        ]
      },
      {
        heading: "Paste Options: the small icons that matter",
        blocks: [
          { type: "p", text: "After pasting, a small clipboard icon appears beside the pasted cells. Click it (or press **Ctrl** to open it while pasting) and you get **Paste Options**. These change what the paste carries:" },
          { type: "table", headers: ["Option", "What it pastes", "When to use it"], rows: [
            ["Keep Source Formatting", "Values plus the original look", "Default — fine for most copies"],
            ["Values (123)", "Only the results, no formulas, no formatting", "Pasting calculated totals into a report that must not change"],
            ["Formulas", "Formulas only, without the source's colours and fonts", "Copying calculations into a differently styled sheet"],
            ["Formatting", "Only colours, borders, number formats — no data", "Applying one block's look to another block"],
            ["Transpose", "Rotates the data: rows become columns", "Converting a horizontal list to vertical"]
          ] },
          { type: "p", text: "The **Values** paste is the star of real office work. Imagine you calculated a summary with formulas, and you want to send the numbers to your manager in a clean sheet where nothing can break. Copy the summary, then paste as **Values**: the receiving sheet contains plain frozen numbers, safe from deleted source cells." },
          { type: "p", text: "**Transpose** is the other gem. If Sunrise Traders recorded months across a row (Jan, Feb, Mar) and you need them down a column, copy the row, choose the destination cell, then select Paste Special > Transpose. The layout flips 90 degrees. There is no simple way to do this by retyping — transpose does it in one action." }
        ]
      },
      {
        heading: "The Paste Special dialog",
        blocks: [
          { type: "p", text: "For more control, press **[[Ctrl]] + [[Alt]] + [[V]]** after copying. The **Paste Special** dialog opens with every option as a radio button: Values, Formulas, Formats, and Transpose among them — the same choices as the Paste Options icons, plus a few more like **All except borders**." },
          { type: "steps", items: [
            "Select the cells to copy and press [[Ctrl]] + [[C]].",
            "Click the destination cell.",
            "Press [[Ctrl]] + [[Alt]] + [[V]].",
            "Choose **Values** and click OK — the destination receives results only."
          ] },
          { type: "tip", text: "The shortcut to remember: Ctrl + Alt + V opens Paste Special from any copied selection. Fast users paste with Ctrl + V when the default is right, and Ctrl + Alt + V when they need Values or Transpose." }
        ]
      },
      {
        heading: "Copying formulas down a column",
        blocks: [
          { type: "p", text: "The most common copy in office life: a column of totals. Suppose `D2` contains `=B2*C2` (price times quantity). Select `D2`, copy it, select `D3:D20`, and paste. Excel does not copy the letters \"B2*C2\" literally — it adjusts each copy to its own row, so `D5` becomes `=B5*C5`." },
          { type: "grid", title: "One formula, copied down four rows", rows: [
            ["", "B", "C", "D"],
            ["1", "Price", "Qty", "Total"],
            ["2", "45", "10", "=B2*C2"],
            ["3", "60", "5", "=B3*C3"],
            ["4", "12500", "1", "=B4*C4"]
          ] },
          { type: "p", text: "This automatic adjustment is called a **relative reference**, and it is the reason copying formulas works at all. The next lessons examine exactly how it behaves — and how to control it with the $ sign when you need part of a formula to stay fixed." },
          { type: "p", text: "You can also drag the fill handle (the small square at the bottom-right of a selected cell) down the column instead of copy-paste. That is the subject of the next lesson." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Pasting a formula into a report sent to others, where the source cells might be deleted later.", right: "Use Paste Special > Values when handing over calculated results.", why: "A pasted formula still points back at your working sheet. If those cells move or vanish, the report shows #REF! errors in front of your manager." },
      { wrong: "Using cut (Ctrl + X) when you meant copy, and finding the original data gone.", right: "Pause and choose: copy keeps the source, cut empties it.", why: "The two shortcuts sit side by side and do very different things. Undo (Ctrl + Z) rescues you if you cut by mistake." },
      { wrong: "Pasting over existing data without looking, and losing the old numbers.", right: "Paste into empty cells, or check the destination first — undo works, but only if you notice.", why: "Paste replaces the destination silently. There is no merge prompt." },
      { wrong: "Rebuilding a transposed table by retyping it.", right: "Copy, then Paste Special > Transpose.", why: "Transpose flips the entire block in one step, with formatting and formulas handled correctly." }
    ],
    shortcuts: [
      { keys: "Ctrl + C", action: "Copy the selection to the clipboard" },
      { keys: "Ctrl + X", action: "Cut the selection (paste will move it)" },
      { keys: "Ctrl + V", action: "Paste from the clipboard" },
      { keys: "Ctrl + Alt + V", action: "Open the Paste Special dialog" },
      { keys: "Esc", action: "Cancel the marching ants and clear the clipboard" }
    ],
    exercises: [
      { task: "In the Beginner practice workbook's Exercise sheet, copy the header row and paste it to the empty area to the right, keeping the source formatting. Then press Esc to clear the marching ants.", hint: "Select the header cells, [[Ctrl]] + [[C]], click the destination, [[Ctrl]] + [[V]], then [[Esc]].", solution: "The pasted header looks identical to the original because Keep Source Formatting is the default paste. Esc removes the marching ants." },
      { task: "The Exercise sheet has a calculated totals block. Copy it and paste it as **Values** into the Solution-style area beside it. Click a pasted cell and check the formula bar — it should show a plain number, not a formula.", hint: "Copy with [[Ctrl]] + [[C]], then press [[Ctrl]] + [[Alt]] + [[V]] and choose Values.", solution: "After a Values paste the formula bar shows only the number. The pasted cells can never break, because no formula remains." },
      { task: "A row of months (Jan to Mar) with sales figures sits horizontally in the Exercise sheet. Use Transpose to flip it so the months run down a column instead.", hint: "Copy the row, click an empty cell, open Paste Special with [[Ctrl]] + [[Alt]] + [[V]], tick Transpose.", solution: "After transposing, the three months appear vertically with their figures beside them, ready to be used as a vertical list." },
      { task: "Move a small block of expense rows from one part of the Exercise sheet to another using drag-and-drop, then move them back using cut and paste.", hint: "Drag from the selection's border, not its middle. For the return trip: [[Ctrl]] + [[X]] then [[Ctrl]] + [[V]].", solution: "Drag-and-drop and cut-paste both relocate the rows; the original positions are emptied either way." }
    ],
    quiz: [
      { q: "What do the marching ants around a selection mean?", options: ["The cells are on the clipboard, ready to paste", "The cells contain errors", "The cells are locked", "The sheet is being saved"], answer: 0, explanation: "The moving dashed border appears after a copy or cut and stays until you press Esc, type, or clear the clipboard." },
      { q: "Which paste option gives you calculated results with no formulas attached?", options: ["Values", "Formulas", "Formatting", "Transpose"], answer: 0, explanation: "Values paste stores plain numbers. Nothing can break later because no formula references remain." },
      { q: "What does Transpose do?", options: ["Switches rows to columns and columns to rows", "Moves data to another sheet", "Converts text to numbers", "Duplicates the selection"], answer: 0, explanation: "Transpose rotates the pasted block 90 degrees — a horizontal list becomes vertical, and vice versa." },
      { q: "Cell D2 contains =B2*C2. You copy D2 and paste into D5. What does D5 now contain?", options: ["=B5*C5", "=B2*C2 exactly as before", "=D2*D5", "The number 450 as text"], answer: 0, explanation: "Copied formulas adjust to their new position — relative references shift with the paste, so D5 points at row 5." },
      { q: "Which shortcut opens the Paste Special dialog?", options: ["Ctrl + Shift + V", "F5", "Ctrl + Alt + V", "Ctrl + P"], answer: 2, explanation: "Ctrl + Alt + V opens Paste Special with all options including Values, Formats and Transpose." }
    ],
    summary: [
      "Ctrl + C copies, Ctrl + X cuts (moves on paste), Ctrl + V pastes; marching ants show the clipboard selection.",
      "Paste Options let you paste values, formulas, formatting or transpose the layout.",
      "Paste Special (Ctrl + Alt + V) gives full control — Values is the safest way to hand over results.",
      "Drag-and-drop moves data by mouse; hold Ctrl while dragging to copy.",
      "Copied formulas adjust their references automatically — relative references in action."
    ],
    downloads: [
      { file: "assets/files/level1/beginner-practice.xlsx", label: "Beginner practice workbook", note: "Exercise and Solution sheets covering data entry, formatting, autofill and copy/paste tasks — the Exercise sheet includes the copy and transpose blocks used here" }
    ],
    keywords: ["copy paste excel", "cut and move cells", "paste special values", "transpose excel", "paste options", "marching ants", "copy formulas down", "ctrl alt v", "keep source formatting"]
  },
  {
    id: "l1-autofill-flashfill",
    title: "AutoFill and Flash Fill",
    difficulty: "Beginner",
    duration: 18,
    version: "Excel 2016+",
    objectives: [
      "Use the fill handle to copy cells and create series",
      "Autofill days, months and number sequences",
      "Copy instead of series with the Ctrl key",
      "Split and reshape text with Flash Fill (Ctrl + E)"
    ],
    intro: "Typing \"Monday, Tuesday, Wednesday...\" by hand, or the same formula forty times, is exactly the work Excel was built to abolish. This lesson covers two tools that finish your typing for you: AutoFill, which extends patterns, and Flash Fill, which learns from one example and completes the column.\n\nNote: Flash Fill needs Excel 2013 or later, so every version from 2016 onward has it. AutoFill works everywhere. Both appear constantly in office work, from staff rosters to tidying imported customer lists.",
    keyIdea: "Drag the fill handle to extend any pattern, and Flash Fill (Ctrl + E) finishes typing from one example.",
    sections: [
      {
        heading: "The fill handle: Excel's fastest tool",
        blocks: [
          { type: "p", text: "Select any cell and look at its bottom-right corner. There is a tiny green square: the **fill handle**. Hover over it and the cursor becomes a thin black plus. Now drag down or across, release, and Excel fills the cells you dragged over." },
          { type: "p", text: "What gets filled depends on what you started with. A plain number like 45 gets **copied**. But text Excel recognises as a series gets **continued** — and this is where the magic is:" },
          { type: "table", headers: ["You type", "Drag the fill handle down", "Result"], rows: [
            ["Monday", "5 more cells", "Tuesday, Wednesday, Thursday, Friday, Saturday"],
            ["Jan", "11 more cells", "Feb through Dec"],
            ["1", "5 more cells", "1, 1, 1, 1, 1 (copies)"],
            ["Week 1", "3 more cells", "Week 2, Week 3, Week 4"]
          ] },
          { type: "p", text: "Excel knows the calendar (days, months) and simple patterns (\"Week 1\" becomes \"Week 2\"). This alone kills hours of typing on rosters, schedules and report headers." },
          { type: "grid", title: "AutoFill of months and a copied number", rows: [
            ["", "A", "B"],
            ["1", "Jan", "45"],
            ["2", "Feb", "45"],
            ["3", "Mar", "45"],
            ["4", "Apr", "45"]
          ] },
          { type: "p", text: "Column A grew from one \"Jan\" into the full series; column B copied the 45 unchanged. Same handle, two different behaviours — and you control which one with the Ctrl key, coming up in a moment." }
        ]
      },
      {
        heading: "Number series: two cells teach the step",
        blocks: [
          { type: "p", text: "A single number copies rather than counts — Excel cannot guess whether you wanted 1, 2, 3 or 1, 5, 9. So **teach it the step with two cells**: type `1` in the first cell and `3` below it, select **both** cells, then drag the fill handle. Excel reads the difference (step of 2) and continues: 5, 7, 9..." },
          { type: "p", text: "The same trick builds any arithmetic series: 10, 20, 30 for step-10 increases; 100, 95, 90 for countdowns (the step can be negative). Dates work the same way — type two dates a week apart, select both, and the fill handle continues weekly." },
          { type: "p", text: "There is also a menu version: after any drag, a small **AutoFill Options** button appears beside the filled cells. Click it to switch between Copy Cells, Fill Series, Fill Days, and more — useful when Excel guessed differently from what you wanted." }
        ]
      },
      {
        heading: "Ctrl to copy instead of continue",
        blocks: [
          { type: "p", text: "What if you drag \"Monday\" but want every cell to say Monday — say, a Monday-only roster column? Hold **Ctrl** while dragging: the series becomes a copy. The icon beside the cursor gains a small plus sign to confirm you are copying." },
          { type: "p", text: "The rule is easy to remember: **without Ctrl**, recognised series continue; **with Ctrl**, everything copies. For plain numbers the defaults are already copy, so Ctrl makes little difference there — but for days, months and patterned text, Ctrl flips the behaviour, and it flips it back. Release Ctrl mid-drag if you change your mind." },
          { type: "tip", text: "Double-clicking the fill handle (instead of dragging) fills down automatically, stopping at the last filled row of the neighbouring column. For a 500-row list where the formula must go all the way down, the double-click fills it instantly." }
        ]
      },
      {
        heading: "Custom lists: your own series",
        blocks: [
          { type: "p", text: "Excel's built-in series cover days and months, but offices have their own repeating sequences: department rotations, priority levels, branch names. You can teach Excel any list via **File > Options > Advanced > Edit Custom Lists**. Type your entries one per line (or import a range from a sheet), click Add, and from then on, typing the first item and dragging the fill handle continues with **your** list." },
          { type: "p", text: "Example: a company with branches in Kochi, Chennai and Hyderabad enters a roster. Add those three names as a custom list once, and every future roster fills the rotation automatically. Sorting by that custom list works too — data can be ordered in your business sequence rather than alphabetically." }
        ]
      },
      {
        heading: "Flash Fill: one example, whole column done",
        blocks: [
          { type: "p", text: "**Flash Fill** (shortcut **[[Ctrl]] + [[E]]**) solves a different problem: reshaping text. You show Excel one or two examples of what you want, and it fills the rest of the column by copying your pattern. It needs Excel 2013 or later." },
          { type: "p", text: "Realistic case 1 — splitting names. A column holds full names like \"Arjun Mehta\". You need first names in one column, last names in another:" },
          { type: "grid", title: "Flash Fill splitting full names", rows: [
            ["", "A", "B", "C"],
            ["1", "Full name", "First", "Last"],
            ["2", "Arjun Mehta", "Arjun", "Mehta"],
            ["3", "Priya Nair", "Priya", "Nair"],
            ["4", "Deepa Rao", "Deepa", "Rao"]
          ] },
          { type: "steps", items: [
            "In `B2`, type the first name from `A2`: type **Arjun** and press Enter.",
            "In `B3`, start typing **Priya** — Flash Fill may suggest the rest of the column in grey. If it does, press Enter to accept.",
            "If nothing appears, press **[[Ctrl]] + [[E]]** with `B3` active. The whole column fills.",
            "Repeat in column C with the surname \"Mehta\" and [[Ctrl]] + [[E]]."
          ] },
          { type: "p", text: "Realistic case 2 — reformatting phone numbers. Imported data stores phones as `9847012345`. You need them displayed as `98470-12345`. Type `98470-12345` in the first row beside the raw number, press [[Ctrl]] + [[E]], and the whole column converts. Numbers, dashes, brackets, spacing — Flash Fill copies whatever structure your example shows." },
          { type: "p", text: "Other uses that come up weekly in offices: building email addresses from names (`armehta@sunrisetraders.in` from Arjun Mehta), extracting item codes from long descriptions, or joining first and last name into one column. Anywhere you can say \"take this, transform it like my example\", Flash Fill can try." },
          { type: "warning", text: "Flash Fill results are **typed values, not live formulas**. If you later correct a name in column A, the Flash Filled columns will not update — rerun Flash Fill after fixing the source. For data that must stay linked to its source, formulas are the right tool; Flash Fill is for one-off clean-up jobs." }
        ]
      },
      {
        heading: "When to use which tool",
        blocks: [
          { type: "ul", items: [
            "**AutoFill** for series and copies: day or month headers, row numbers, a formula repeated down a column, copying a constant to many cells.",
            "**Flash Fill** for text reshaping: splitting, joining or reformatting text based on examples.",
            "**Copy-paste** when you need exact control of what transfers (values, formatting, transpose) — covered in the previous lesson."
          ] },
          { type: "p", text: "In practice they chain together. A typical morning: AutoFill the weekday headers across a sales sheet, Flash Fill the customer names into first and last columns, then AutoFill the total formula down the rows. Three minutes of work that once took an hour of typing." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Typing the full days-of-the-week sequence manually for every new roster.", right: "Type Monday once and drag the fill handle.", why: "AutoFill continues recognised series instantly and without typing errors." },
      { wrong: "Dragging a single cell containing 1 and expecting 2, 3, 4.", right: "Type 1 and 2 in two cells, select both, then drag — or use the AutoFill Options menu and choose Fill Series.", why: "One number gives Excel no step to continue, so it copies instead of counting." },
      { wrong: "Expecting Flash Filled text to update when the source column changes.", right: "Rerun Flash Fill (Ctrl + E) after editing the source, or keep the data linked with formulas if it must stay live.", why: "Flash Fill writes static values. It is a one-time transformation, not a living formula." },
      { wrong: "Using Flash Fill on data that keeps changing, like a growing price list.", right: "Use Flash Fill for one-off cleanups; use formulas for data that grows or updates.", why: "Flash Fill will not cover new rows automatically — you would have to rerun it every time." }
    ],
    shortcuts: [
      { keys: "Ctrl + E", action: "Flash Fill from your typed example (Excel 2013+)" },
      { keys: "Ctrl + drag fill handle", action: "Copy instead of continuing a series" },
      { keys: "Double-click fill handle", action: "Fill down to the last row of the neighbouring column" }
    ],
    exercises: [
      { task: "In the Beginner practice workbook's Exercise sheet, find the empty weekday column. Type Monday in the first cell and use the fill handle to fill the rest of the column through Sunday.", hint: "Type Monday, then drag the small green square at the cell's bottom-right corner down.", solution: "The column fills Tuesday through Sunday automatically, since days of the week are a built-in Excel series." },
      { task: "Create a number sequence 10, 20, 30 ... down to 100 in the sequence area of the Exercise sheet, using the two-cell technique.", hint: "Type 10 and 20 in two cells, select both, then drag the fill handle down.", solution: "Selecting 10 and 20 teaches Excel a step of 10, so the fill continues 30, 40 ... 100." },
      { task: "The Exercise sheet has a Full Name column. Use Flash Fill to create First Name and Last Name columns beside it, then create an email column in the pattern first initial + last name, like amehta.", hint: "Type the first two examples manually, press [[Ctrl]] + [[E]] with the next empty cell in that column active.", solution: "Ctrl + E fills First Name and Last Name correctly. For emails, type amehta beside Arjun Mehta as the pattern; two examples usually make the fill accurate for all rows." },
      { task: "The sheet's phone numbers are stored as ten digits with no separators. Use Flash Fill to reformat them as 98470-12345 style (first five digits, dash, last five).", hint: "Type the reformatted version of the first number in the cell beside it, then [[Ctrl]] + [[E]].", solution: "One typed example is enough for a consistent pattern like 98470-12345; the whole column converts instantly." }
    ],
    quiz: [
      { q: "You type \"Wednesday\" in a cell and drag its fill handle down. What happens?", options: ["The following days fill in: Thursday, Friday, and so on", "Wednesday is copied into every cell", "Excel shows an error", "The cells stay empty"], answer: 0, explanation: "Recognised series like weekdays continue automatically when you drag the fill handle. Hold Ctrl while dragging to copy instead." },
      { q: "How do you make AutoFill produce 5, 10, 15, 20?", options: ["Type 5 and 10 in two cells, select both, then drag", "Type 5 and drag with Ctrl held", "Type 5, 10, 15 manually and drag", "Flash Fill is required"], answer: 0, explanation: "Two cells teach Excel the step (difference of 5), and the fill handle continues the series. One cell alone would just copy." },
      { q: "What does holding Ctrl while dragging the fill handle do?", options: ["Speeds up the fill", "Copies the cell instead of continuing a series", "Deletes the source cell", "Fills to the right only"], answer: 1, explanation: "Ctrl flips the behaviour: series become copies and copies become series while Ctrl is held." },
      { q: "Which shortcut triggers Flash Fill?", options: ["Ctrl + E", "Ctrl + F", "Ctrl + Shift + F", "F5"], answer: 0, explanation: "Ctrl + E runs Flash Fill after you type one or two examples of the pattern you want." },
      { q: "You fix a misspelled name in the source column after using Flash Fill. What happens to the Flash Filled result?", options: ["It stays as it was — you must rerun Flash Fill", "It updates automatically", "Excel shows a broken reference error", "It deletes itself"], answer: 0, explanation: "Flash Fill writes static values, not formulas. Editing the source does not flow through; rerun Ctrl + E to refresh." }
    ],
    summary: [
      "The fill handle copies plain cells and continues recognised series like days and months.",
      "Two cells define a step for number and date series: 10 then 20 fills 30, 40 and beyond.",
      "Hold Ctrl while dragging to copy instead of continue; double-click the handle to fill down instantly.",
      "Custom Lists (File > Options > Advanced) teach Excel your own repeating sequences.",
      "Flash Fill (Ctrl + E) transforms text columns from one or two examples — but results are static values."
    ],
    downloads: [
      { file: "assets/files/level1/beginner-practice.xlsx", label: "Beginner practice workbook", note: "Exercise and Solution sheets covering data entry, formatting, autofill and copy/paste tasks — includes the weekday, sequence, names and phone number areas used here" }
    ],
    keywords: ["autofill excel", "fill handle", "excel series days months", "custom lists", "flash fill", "ctrl e", "split names excel", "format phone numbers", "fill series numbers"]
  }
);
