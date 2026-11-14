// Excel Lab — Level 6: Advanced Excel (part 2: lessons 6-9)
window.LEVEL_LESSONS = window.LEVEL_LESSONS || {};

(window.LEVEL_LESSONS[6] = window.LEVEL_LESSONS[6] || []).push(
  {
    id: "l6-intro-dax",
    title: "Introduction to DAX",
    difficulty: "Advanced",
    duration: 22,
    version: "Excel 2016+",
    objectives: [
      "Explain what DAX is and where it runs (inside the Data Model, not on worksheets)",
      "Choose correctly between a calculated column and a measure",
      "Write your first measures: SUM, COUNTROWS, DIVIDE and CALCULATE",
      "Explain why measures recalculate automatically with filters and slicers",
      "Create a measure with the right name, format and description"
    ],
    intro: "DAX (Data Analysis Expressions) is the formula language that lives inside Power Pivot and the Data Model. If the previous lesson wired your tables together, DAX is how you ask the model smart questions: total sales per filter, laptop revenue in the North region, the average order value for whichever salesperson the slicer selects.\n\nThe good news: DAX looks almost exactly like Excel formulas. `=SUM(Sales[Amount])` is recognisable at a glance. The one big idea to accept early: DAX formulas do not point at cells like A2 or B7. They point at columns of whole tables — `Sales[Amount]` means \"the Amount column of the Sales table\". That single difference changes how you think about every calculation, and this lesson walks it in small steps.",
    sections: [
      {
        heading: "What DAX is and where it runs",
        blocks: [
          { type: "p", text: "DAX is a formula language, like the Excel formula language, but it runs inside the Data Model — the compressed store of tables you built with Power Pivot. It is available in Excel 2016 and later (through Power Pivot) and in Power BI, so learning it now also prepares you for a BI tool many companies use. It is not something you type into a normal worksheet cell: `=SUM(Sales[Amount])` in A1 of a plain sheet returns #NAME?, because the sheet has no idea what the Sales table is." },
          { type: "p", text: "The syntax difference is worth writing on a sticky note. In worksheet formulas you reference single cells and ranges: `B2`, `C2:C50`. In DAX you reference tables and columns in the form TableName[ColumnName]:" },
          { type: "table", headers: ["Excel worksheet formula", "DAX equivalent idea"], rows: [
            ["=SUM(D2:D50000)", "=SUM(Sales[Amount]) — sums the whole Amount column of the Sales table"],
            ["=COUNTA(A2:A50000)", "=COUNTROWS(Sales) — counts the rows of the Sales table"],
            ["=D2*0.18 per row (calculated column)", "Sales[GST] = Sales[Amount] * 0.18 — one formula, applied to every row"],
            ["=SUMIFS with many conditions", "=CALCULATE(SUM(Sales[Amount]), filter conditions)"]
          ] },
          { type: "p", text: "Notice there are no dollar signs and no row numbers anywhere. A DAX expression describes a whole column, and the Data Model decides which rows of that column are currently visible based on the filters in your pivot — more on that in the last section, because it is the key to everything." }
        ]
      },
      {
        heading: "Calculated columns vs measures: the first big decision",
        blocks: [
          { type: "p", text: "DAX can create two kinds of things, and choosing wrongly is the most common beginner mistake. A calculated column is computed row by row, once, and stored in the table like any other column — think of it as a VLOOKUP-style helper column that lives inside the model. A measure is a single value that recalculates for whatever filter situation the pivot is showing right now." },
          { type: "table", headers: ["", "Calculated column", "Measure"], rows: [
            ["What it is", "A new column, one value per row", "One named formula, one result at a time"],
            ["When it computes", "Once, at refresh time, and the result is stored", "At the moment of display, for the current filters"],
            ["Typical use", "GST = Amount * 18%; Year = YEAR(Sales[Date])", "Total Sales, Average Order, Laptop Sales"],
            ["Cost", "Uses storage in the model", "Almost none — calculated on demand"],
            ["Wrong choice causes", "Bloated model, slow refresh", "Wrong numbers when you needed row-level values"]
          ] },
          { type: "p", text: "The rule of thumb: if the answer should change when you slice and dice, it is a measure. If it is a fact about one row (this sale's GST, this row's year, this line's margin), it is a calculated column. A \"Total Sales\" number must follow the slicers, so it is a measure. A row's GST belongs to that row, so it is a calculated column." },
          { type: "warning", text: "Beginners often build \"Total Sales\" as a calculated column, writing =SUM(Sales[Amount]) in a column context. The result surprises them: the sum appears on every row — 33 identical totals down the table. Row context does not filter the SUM. Anything that aggregates belongs in a measure." }
        ]
      },
      {
        heading: "Implicit vs explicit measures",
        blocks: [
          { type: "p", text: "When you drag Sales[Amount] into a pivot's Values area, Excel quietly creates a sum for you. That automatic aggregation is called an implicit measure — implicit because you never wrote it. It works, and for quick ad-hoc questions it is fine. Professional reports, however, are built on explicit measures: aggregations you have written yourself in DAX and given a proper name." },
          { type: "ul", items: [
            "An explicit measure has a business name — \"Total Sales\", not \"Sum of Amount\" — so the pivot header reads like a report, not a spreadsheet.",
            "You control the number format once (currency, no decimals) at the measure level, instead of reformatting every pivot you build.",
            "Other measures can reference it: Average Order Value = DIVIDE([Total Sales], [Total Orders]) builds directly on the two you already wrote.",
            "CALCULATE-based logic — the real power of DAX — can only live in an explicit measure. Implicit sums cannot filter themselves."
          ] },
          { type: "p", text: "Create one via Power Pivot > Measures > New Measure (or right-click a table in the pivot field list > Add Measure). A dialog appears with three fields worth filling in every time: Formula, the measure name, and the number format. A fourth field — Description — is optional but kind to your future colleagues." },
          { type: "tip", text: "Build the habit now: never drag Amount into Values in a model pivot. Write the measure once, then drag the measure. Your field list starts to look like a small report catalogue, and every pivot you build reuses the same definitions." }
        ]
      },
      {
        heading: "Your first four measures, step by step",
        blocks: [
          { type: "p", text: "Open the Power Pivot practice workbook, confirm the three tables (Sales, Products, Regions) are in the model with their relationships, then create these measures in order. Each one teaches one DAX building block." },
          { type: "code", lang: "excel", text: "Measure 1 — name it \"Total Sales\", format as currency:\n\n=SUM(Sales[Amount])" },
          { type: "p", text: "This is the SUM you know, pointed at a table column instead of a range. Drop it into a pivot with Category on Rows and you get revenue per category. Because it is a measure, it will also respond to every slicer and filter you add later — automatically." },
          { type: "code", lang: "excel", text: "Measure 2 — name it \"Total Orders\":\n\n=COUNTROWS(Sales)" },
          { type: "p", text: "COUNTROWS counts the rows of a table. In a pivot filtered to one salesperson, it answers \"how many sale lines did this person make?\". With no filters, it counts the whole table. Same formula, different scope — that is the magic you are about to rely on." },
          { type: "code", lang: "excel", text: "Measure 3 — name it \"Average Order\":\n\n=DIVIDE([Total Sales], [Total Orders])" },
          { type: "p", text: "Two things to notice. First, measures reference other measures by name in square brackets — this is how DAX code stays short. Second, DIVIDE is the professional division: if Total Orders is ever 0 (say, a region with no sales yet), DIVIDE returns a blank instead of a divide-by-zero error, and it accepts an optional third argument for the alternative result." },
          { type: "code", lang: "excel", text: "Measure 4 — name it \"Laptop Sales\":\n\n=CALCULATE(SUM(Sales[Amount]), Products[Category] = \"Laptops\")" },
          { type: "p", text: "CALCULATE is the most important function in DAX, and its idea is simple: calculate this expression, but under these filter conditions. Here it sums Amount, but only the rows where the related Products category is Laptops. In a pivot filtered to Priya Nair, Laptop Sales shows Priya's laptop revenue only. You just wrote the equivalent of a SUMIFS that adapts itself to every filter on the page." },
          { type: "grid", title: "One pivot, four measures, every filter combination", rows: [
            ["Row label", "Total Sales", "Total Orders", "Average Order", "Laptop Sales"],
            ["Arjun Mehta", "486000", "22", "22090", "172000"],
            ["Priya Nair", "521000", "25", "20840", "195000"],
            ["Rahul Verma", "398000", "19", "20947", "96000"],
            ["Sneha Iyer", "455000", "21", "21666", "128000"]
          ] },
          { type: "p", text: "Four measures, no copy-paste, no per-row formulas — and every figure in that grid is correct for its own row without you writing anything per person." }
        ]
      },
      {
        heading: "Why measures recalculate with filters and slicers",
        blocks: [
          { type: "p", text: "Here is the concept that makes DAX click, described without jargon. A measure is written once, but it is evaluated many times — once for every cell of the pivot where it appears. Each evaluation happens under a different set of active filters, called the filter context." },
          { type: "p", text: "Look at the grid above. The cell where \"Arjun Mehta\" row meets \"Total Sales\" column is evaluated with the filter \"Salesperson = Arjun Mehta\" active, so SUM(Sales[Amount]) sees only Arjun's rows. The Total row is evaluated with no salesperson filter, so the same formula sees all rows. Nothing about the formula changed; the visible data behind it changed." },
          { type: "ul", items: [
            "A slicer click changes the filter context for the whole pivot — every measure cell re-evaluates instantly.",
            "A row or column label adds its own filter on top: within Arjun's row, a CALCULATE for Laptops narrows the context further.",
            "This is why the same `=SUM(Sales[Amount])` shows a different number in every cell of the pivot — and why that is correct, not a bug.",
            "It is also why worksheet-style thinking (\"this cell references that cell\") does not transfer: a measure does not reference anything on the sheet. It references the model and asks \"what is filtered right now?\"."
          ] },
          { type: "p", text: "If you remember one sentence from this lesson, make it this: measures are formulas that answer the same question for every filter combination the user can produce. Everything else in DAX — CALCULATE, filter context, time intelligence — is detail hanging off that sentence. Practice the four measures here until creating one feels routine; the next step in real projects is exactly these four, plus CALCULATE conditions that grow more interesting." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Typing DAX formulas into worksheet cells: =SUM(Sales[Amount]) in A1 of a normal sheet.", right: "Create measures via Power Pivot > Measures > New Measure, or calculated columns inside the Power Pivot window.", why: "DAX runs inside the Data Model. On a worksheet the table and column names do not exist, so every such formula returns #NAME? no matter how correct the syntax is." },
      { wrong: "Building \"Total Sales\" as a calculated column so it can be \"in the table\".", right: "Aggregations belong in measures; calculated columns are for row-level facts.", why: "In a calculated column, SUM(Sales[Amount]) sums the entire column and stamps that one total on every row — it does not vary per row, and it wastes storage storing thousands of identical copies." },
      { wrong: "Writing =SUM(Sales[Amount])/COUNTROWS(Sales) inline everywhere an average is needed.", right: "Create \"Average Order\" as its own measure: =DIVIDE([Total Sales], [Total Orders]).", why: "Inline division crashes on zero-order slices and must be re-edited in every pivot. A named measure fixes the definition once, handles divide-by-zero via DIVIDE, and is reusable everywhere." },
      { wrong: "Expecting a measure to show a fixed number, and \"fixing\" it by copying the pivot values as static text.", right: "Trust the filter context: the number changes because the filters changed, which is the desired behaviour.", why: "Freezing results kills the entire benefit of measures. If a report needs a snapshot at a point in time, take the snapshot deliberately (copy-paste-values into an archive sheet) rather than replacing the live report." }
    ],
    shortcuts: [
      { keys: "Right-click", action: "A table in the pivot field list > Add Measure — the quickest route to a new explicit measure" },
      { keys: "Ctrl + Shift + U", action: "Expand the formula bar in the measure dialog for long DAX lines" },
      { keys: "Alt + F5", action: "Refresh the pivot (and the model) after source data changes" }
    ],
    exercises: [
      { task: "Open assets/files/level6/power-pivot-data.xlsx and confirm the Sales, Products and Regions tables are in the Data Model with relationships from the previous lesson. Create the measure \"Total Sales\" = SUM(Sales[Amount]) with currency format.", hint: "Power Pivot > Measures > New Measure. TableName[ColumnName] — the table names are exactly Sales, Products, Regions.", solution: "Measure formula: =SUM(Sales[Amount]) — drag it into a pivot with Category on Rows to see revenue per category." },
      { task: "Create \"Total Orders\" = COUNTROWS(Sales) and \"Average Order\" = DIVIDE([Total Sales], [Total Orders]). Put Salesperson on Rows with all three measures in Values.", hint: "Measures reference each other in brackets: [Total Sales], [Total Orders]. DIVIDE takes the division first, then what to do about zeros.", solution: "Each salesperson row shows their own totals — the same three formulas produce a different (correct) number per row because the filter context changes per row." },
      { task: "Create \"Laptop Sales\" = CALCULATE(SUM(Sales[Amount]), Products[Category] = \"Laptops\") and compare it with \"Total Sales\" per category in a pivot.", hint: "CALCULATE(expression, filter). The filter references the Products table, not Sales — the relationship carries the condition across.", solution: "=CALCULATE(SUM(Sales[Amount]), Products[Category]=\"Laptops\") — in the Laptops row it equals Total Sales; in every other category row it shows blank, because no rows match the filter there." },
      { task: "Add a slicer for Region[City] to your pivot and click through two or three cities. Watch all four measures change. Then write down in one sentence why the formulas did not change but the numbers did.", hint: "Think about which rows of the Sales table were visible to SUM each time you clicked.", solution: "The slicer changes the filter context: each click makes a different set of Sales rows visible, so the same SUM, COUNTROWS, DIVIDE and CALCULATE formulas evaluate to different, correct results." }
    ],
    quiz: [
      { q: "Where do DAX formulas live?", options: ["In normal worksheet cells", "Inside the Data Model (measures and calculated columns)", "Only in .xlsm files", "In the VBA editor"], answer: 1, explanation: "DAX is the language of Power Pivot and the Data Model. It creates measures and calculated columns there; typing it on a worksheet returns #NAME?." },
      { q: "Sales[Amount] means:", options: ["The cell at the intersection of row Sales and column Amount", "The Amount column of the Sales table", "A range named Amount on sheet Sales", "A worksheet function called Sales"], answer: 1, explanation: "DAX references are TableName[ColumnName]. There are no cell addresses or dollar signs — expressions describe whole columns and let the current filters pick the rows." },
      { q: "You need each sales row to show its own 18% GST. What do you build?", options: ["A measure", "A calculated column", "A VLOOKUP on the sheet", "A slicer"], answer: 1, explanation: "A per-row value is a row-level fact — that is a calculated column: Sales[GST] = Sales[Amount] * 0.18. Measures produce one value per filter context, not one per row." },
      { q: "What does =CALCULATE(SUM(Sales[Amount]), Products[Category]=\"Laptops\") do?", options: ["Adds two columns together", "Sums Amount only over rows where the related category is Laptops, under the current filters", "Counts the laptops in stock", "Calculates a running total"], answer: 1, explanation: "CALCULATE evaluates the expression under additional filter conditions. Combined with the relationship Sales→Products, it sums only laptop sale amounts — a self-adjusting SUMIFS." },
      { q: "Why does the same measure show a different number on each row of a pivot?", options: ["The measure is copied and edited per row", "Each cell is evaluated under that row's filter context", "DAX averages the filters randomly", "It only happens when the model has errors"], answer: 1, explanation: "A measure is evaluated once per pivot cell, each time against the filters active for that cell (row labels, columns, slicers). Same formula, different visible rows, different result — by design." }
    ],
    summary: [
      "DAX is the formula language of the Data Model (Excel 2016+ via Power Pivot, also used in Power BI); it references TableName[ColumnName], never cells.",
      "Calculated columns compute one value per row and are stored; measures compute one value per filter context on demand.",
      "Rule of thumb: if it should change with slicers, it is a measure; if it describes a single row, it is a calculated column.",
      "Implicit measures appear when you drag a column into Values; explicit measures you write yourself are named, formatted, reusable and can use CALCULATE.",
      "Start with SUM, COUNTROWS, DIVIDE and CALCULATE — and remember that a measure answers its question anew for every filter combination."
    ],
    downloads: [
      { file: "assets/files/level6/power-pivot-data.xlsx", label: "Power Pivot data model tables", note: "Three linked tables — Sales, Products and Regions — ready for the data model and DAX practice" }
    ],
    keywords: ["DAX", "measures", "calculated columns", "CALCULATE", "SUM Sales[Amount]", "COUNTROWS", "DIVIDE", "implicit vs explicit measures", "filter context", "Power Pivot formulas"]
  },
  {
    id: "l6-record-macros",
    title: "Recording Macros",
    difficulty: "Advanced",
    duration: 18,
    version: "",
    objectives: [
      "Explain what a macro is and when recording one is worth it",
      "Enable the Developer tab",
      "Record a formatting macro and run it from the Macros dialog",
      "Switch between absolute and relative recording",
      "Assign a macro to a button or shape",
      "Save correctly as .xlsm and understand macro security prompts"
    ],
    intro: "A macro is a saved list of actions that Excel can repeat for you at the click of a button. Suppose every morning you format the Sunrise Traders daily report the same way: bold header, coloured fill, column widths, freeze the top row. Doing it by hand is two minutes of clicking; recording it once turns it into one click, forever.\n\nThis lesson needs no programming. Excel watches what you do, writes the steps down in VBA (a language you will read properly two lessons from now), and can play them back exactly. You will learn where the record button lives, what Excel actually writes down, the one setting most beginners get wrong (relative vs absolute), and the saving rule that protects your work.",
    sections: [
      {
        heading: "What a macro is, and when to record one",
        blocks: [
          { type: "p", text: "Technically, a macro is a small VBA program in your workbook. Practically, it is a robot: you show it the moves once, and it repeats them instantly, identically, every time. Recording is how most people create their first macros, and it remains a professional technique even for experienced VBA writers — record the boring part, then edit the code." },
          { type: "p", text: "Good macro candidates share a pattern: the same steps, on data that keeps arriving in the same shape. Formatting the daily report, clearing last month's input sheet, applying a standard print area and header, building the same three pivots on the new export. Poor candidates: one-off fixes, anything that needs a decision each time (\"which rows to delete today?\"), and steps that depend on values you have to read and judge — recording cannot think, as this lesson shows, but it never gets bored and never mistypes." },
          { type: "steps", items: [
            "File > Options > Customize Ribbon.",
            "On the right, tick the Developer checkbox and click OK.",
            "The Developer tab appears with the Code group: Visual Basic, Macros, Record Macro, Use Relative References, Macro Security."
          ] },
          { type: "p", text: "You need the Developer tab once per computer. Note that on many company machines, macro recording is available but running macros may be restricted by policy — check with IT if buttons appear greyed out." }
        ]
      },
      {
        heading: "Record your first formatting macro",
        blocks: [
          { type: "p", text: "Open the practice workbook and go to its Exercise sheet, which holds a small daily sales report with an unformatted header row. We will record a macro that formats the header every time raw data is pasted in." },
          { type: "steps", items: [
            "Select cell A1 (always start from a known place — more on this below).",
            "Developer > Record Macro. Give it the name FormatReport, a shortcut key if you like (leave it blank for now), choose \"This Workbook\" for storage, and type a description: \"Formats the daily sales report header\". Click OK. The button changes to Stop Recording — everything you do now is being written down.",
            "Select A1:E1. Apply bold, a fill colour, and center alignment — the exact clicks you would do by hand.",
            "Widen the columns to fit, then View > Freeze Panes > Freeze Top Row.",
            "Click cell A1 again, so the macro leaves the sheet in a tidy starting state.",
            "Developer > Stop Recording. The macro is saved."
          ] },
          { type: "grid", title: "The report header before and after running FormatReport", rows: [
            ["Date", "Salesperson", "Product", "Qty", "Amount (₹)"],
            ["02-Nov-2026", "Arjun Mehta", "Staplers", "12", "2520"],
            ["02-Nov-2026", "Priya Nair", "Notebooks", "48", "5760"],
            ["03-Nov-2026", "Rahul Verma", "Pens", "100", "3500"]
          ] },
          { type: "p", text: "After running the macro, row 1 is bold, filled and centered, the columns fit their contents and the top row is frozen. To see what Excel wrote down, Developer > Visual Basic (or Alt+F11), and in the left panel double-click Modules > Module1. You will find something like this — I have tidied the colour number and spacing for readability:" },
          { type: "code", lang: "vba", text: "Sub FormatReport()\n' Formats the daily sales report header\n\n  Range(\"A1:E1\").Select\n  Selection.Font.Bold = True\n  With Selection.Interior\n    .Color = RGB(255, 217, 102)\n  End With\n  Selection.HorizontalAlignment = xlCenter\n  Range(\"A1\").Select\n\nEnd Sub" },
          { type: "p", text: "You can already read this: select the range, make the selection's font bold, give its interior a colour, center it, put the cursor back on A1. Recorded code is verbose (real recordings include many steps you did not intend, like extra Select calls), but it is a faithful transcript. Two lessons from now you will write cleaner code by hand — and even then, recording remains the fastest way to discover the right VBA words for a formatting action you cannot remember." },
          { type: "tip", text: "Before recording any macro, plan the exact steps and the starting selection. A macro remembers everything, including the accidental scroll, the wrong cell click and the spelling fix you typed over the data." }
        ]
      },
      {
        heading: "Absolute vs relative recording: the one setting to understand",
        blocks: [
          { type: "p", text: "By default, recording is absolute: if you click cell B5 while recording, the macro will always act on B5, no matter where you are when you run it. That is perfect for the report header, which always lives in A1:E1. But many jobs are \"do the same thing relative to where I am\" — \"go to the next empty row and paste\", \"total the column to my left\". For those, use relative references." },
          { type: "steps", items: [
            "Developer > Use Relative References — the toggle button highlights when it is on.",
            "Record a small macro: from any cell, press the Down arrow twice, then type your initials and press Enter. Stop recording.",
            "Run it from different starting cells. The initials land two rows below wherever the cursor was — the recording stored \"two rows down\", not a fixed address.",
            "Toggle relative mode off again before recording fixed-position work like the header macro."
          ] },
          { type: "p", text: "In the code, the difference is visible: absolute recording writes Range(\"B5\").Select; relative recording writes ActiveCell.Offset(2, 0).Range(\"A1\").Select — \"two rows down from the active cell\". The recorded wording is clunky, but the behaviour is what matters. Choose the mode that matches the job: fixed addresses for fixed layouts, relative for \"same again, one row further down\"." },
          { type: "warning", text: "The most common recorded-macro failure: recording in absolute mode a macro meant to be reused on tomorrow's data that arrives lower down the sheet, or vice versa. If a macro \"worked when I recorded it but hits the wrong cells now\", the mode was wrong for the job." }
        ]
      },
      {
        heading: "Running, buttons and keyboard shortcuts",
        blocks: [
          { type: "p", text: "A macro is only useful when it is easy to run. You have four everyday options, in rising order of convenience:" },
          { type: "ul", items: [
            "Developer > Macros (or Alt+F8): the Macros dialog lists every macro with its description; select one and click Run. Also the place to click Edit, which jumps straight into the code.",
            "Keyboard shortcut: either set one while recording (the Record Macro dialog's shortcut box, which always adds Ctrl), or add it later in the Macros dialog via Options. Choose combinations like Ctrl+Shift+R that Excel does not already use.",
            "A shape or button on the sheet: Insert > Shapes, draw a rectangle, type a label such as \"Format report\", then right-click it > Assign Macro and pick FormatReport. Now anyone — including a colleague who has never seen the Developer tab — can run the report by clicking the shape.",
            "The Forms Control button (Developer > Insert > Button) works the same way and looks more old-school official."
          ] },
          { type: "p", text: "For shared workbooks, the shape-button approach is the professional favourite: the button travels with the sheet, sits exactly where the work happens, and has a readable label. Keep the Buttons group on a single \"control\" area of the sheet — say a shaded strip above the data — so the sheet has one obvious place to interact." }
        ]
      },
      {
        heading: "Macro security, saving as .xlsm, and the Personal Macro Workbook",
        blocks: [
          { type: "p", text: "Two rules protect your macros from being silently destroyed, and one protects your computer." },
          { type: "ul", items: [
            "Save as .xlsm. A workbook containing macros must be saved as Excel Macro-Enabled Workbook (.xlsm). If you press Ctrl+S on a new macro workbook, Excel shows a warning that macros cannot be saved in .xlsx. If you choose \"Yes\" and continue saving as .xlsx, Excel saves without the macros — your code is gone, and often you only discover when the button does nothing next week.",
            "Macro security: when you open a workbook with macros, Excel shows a yellow security bar: \"Macros have been disabled. Enable Content\". Excel blocks macros from untrusted sources by default, which is correct behaviour — only click Enable Content for files you and your company trust. The rules live in File > Options > Trust Center > Trust Center Settings > Macro Settings.",
            "Personal Macro Workbook: in the Record Macro dialog you can choose to store macros in the Personal Macro Workbook (a hidden file called PERSONAL.XLSB that opens with Excel itself). Macros stored there are available in every workbook you open — ideal for your own everyday utilities. It is stored on your machine only, so do not put shared report macros there; keep those in the workbook the team uses."
          ] },
          { type: "warning", text: "Trusted locations are the clean answer for recurring files: in Trust Center settings you can mark a folder (say, your reports folder) as trusted, and workbooks opened from there skip the security bar entirely — no clicking, no risk of forgetting to enable and running half a workflow without macros." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Saving a macro workbook with Ctrl+S and clicking \"Yes\" to keep it as .xlsx.", right: "Use File > Save As and choose \"Excel Macro-Enabled Workbook (*.xlsm)\".", why: ".xlsx files cannot contain macros. Saving as .xlsx silently strips all VBA — the buttons remain on the sheet pointing at nothing." },
      { wrong: "Recording a macro in absolute mode and expecting it to work on next month's data, which sits ten rows lower.", right: "Decide per job: Use Relative References for position-independent work, absolute for fixed layouts like a header in row 1.", why: "Absolute recording hardcodes every address. The macro replays on exactly the cells from recording day, wherever the new data happens to be." },
      { wrong: "Recording while browsing: scrolling around, misclicking a cell, retyping a header you misspelled.", right: "Plan the steps, select the correct starting cell, then record in one clean pass.", why: "The recording is a transcript of everything, including mistakes. Sloppy recording produces a macro that performs your errors at machine speed." },
      { wrong: "Clicking \"Enable Content\" on any workbook that asks.", right: "Enable macros only for files from trusted sources — your own, or your company's approved locations.", why: "Macros are a common malware delivery route. Excel's warning exists because the file can now run code; treat the prompt as a real decision, not an obstacle." }
    ],
    shortcuts: [
      { keys: "Alt + F8", action: "Open the Macros dialog from anywhere to run, edit or delete a macro" },
      { keys: "Alt + F11", action: "Open the VBA editor to inspect what a recording wrote" },
      { keys: "Alt + F8 > Options", action: "Attach or change a Ctrl keyboard shortcut for a macro" }
    ],
    exercises: [
      { task: "Enable the Developer tab, then open assets/files/level6/macro-practice.xlsx and record the FormatReport macro exactly as in the lesson: bold, fill, center on A1:E1, fit the column widths, freeze the top row, return to A1. Stop recording.", hint: "Always select A1 before starting the recording and again before stopping it. The Record dialog asks for a name and where to store it — choose This Workbook.", solution: "Running the macro again (after messing up the header by hand) restores the formatting in one click — proof the steps were captured." },
      { task: "Record a second macro in RELATIVE mode: from the active cell, move two rows down and type your initials. Run it from five different starting cells and observe where the initials land. Then toggle relative mode off.", hint: "Developer > Use Relative References must be highlighted before you start this recording.", solution: "The initials always appear two rows below the active cell, whatever cell you started from — the recording stored the movement, not the address." },
      { task: "Insert a rounded rectangle shape on the sheet, label it \"Format report\", and assign the FormatReport macro to it. Delete the formatting by hand, then click the button to restore it.", hint: "Insert > Shapes; type the label; right-click the shape > Assign Macro. Note: left-click runs the macro, so right-click to select or move the shape later.", solution: "One click on the shape replays the whole format routine — and anyone using the sheet can run it without knowing the Developer tab exists." },
      { task: "Save the workbook as Sunrise-Traders-Macros.xlsm in a folder you control. Then open it again and dismiss or enable the macro security bar as appropriate.", hint: "File > Save As > file type \"Excel Macro-Enabled Workbook\". On reopening, watch for the yellow bar and choose Enable Content since this is your own file.", solution: "The workbook now carries its macros on disk. A copy saved as .xlsx would reopen with the shapes present but every macro gone." }
    ],
    quiz: [
      { q: "What does recording a macro actually do?", options: ["Films the screen so you can watch it later", "Writes your actions down as VBA code that can be replayed", "Copies the worksheet to a backup", "Speeds up Excel's recalculation"], answer: 1, explanation: "The recorder translates each click into VBA statements, saved in a module. Running the macro executes those statements — it is code, not a video." },
      { q: "You recorded clicking cell B5 in absolute mode. You run the macro while the cursor is on F2. What happens?", options: ["The macro acts on the cell two rows below F2", "The macro selects and acts on B5", "The macro acts on F2", "The macro asks which cell to use"], answer: 1, explanation: "Absolute recording hardcodes addresses: B5 is B5 forever. For \"act near wherever I am\", record with Use Relative References turned on." },
      { q: "What happens if you save a macro workbook as .xlsx?", options: ["Excel refuses outright and closes", "The file saves, but all macros are stripped out", "The macros convert to formulas", "Nothing — .xlsx holds macros fine"], answer: 1, explanation: ".xlsx cannot contain VBA. Excel warns during save; if you proceed, the code is silently dropped and the buttons on the sheet point at nothing." },
      { q: "The best way to let any colleague run a macro without training them on the Developer tab is:", options: ["Email them the VBA code", "Assign the macro to a labelled shape or button on the sheet", "Tell them to press Alt+F8 and guess the name", "Print instructions and tape them to the monitor"], answer: 1, explanation: "A labelled shape with the macro assigned travels with the sheet and runs on a single click — self-service for anyone who opens the file." },
      { q: "What is the Personal Macro Workbook (PERSONAL.XLSB)?", options: ["A hidden workbook that opens with Excel, holding macros available in every file you open", "A password-protected workbook for confidential data", "The folder where .xlsm files must be saved", "A special sheet inside each macro workbook"], answer: 0, explanation: "Store your own everyday macros there and they are available everywhere Excel runs on your machine. It is per-machine, so team macros belong in the shared workbook instead." }
    ],
    summary: [
      "A macro is VBA code that repeats your steps exactly; recording is how you create one without programming.",
      "Enable the Developer tab once: File > Options > Customize Ribbon > tick Developer.",
      "Record clean passes only: plan the steps, start from a known cell, stop before you browse. Choose absolute vs relative mode to match the job.",
      "Run from Alt+F8, a Ctrl shortcut, or — best for teams — a labelled shape with the macro assigned.",
      "Save macro workbooks as .xlsm (saving as .xlsx deletes the code), click Enable Content only for trusted files, and keep personal utilities in PERSONAL.XLSB."
    ],
    downloads: [
      { file: "assets/files/level6/macro-practice.xlsx", label: "Macro practice workbook", note: "A plain workbook with sample data to record and run your first macros (re-save as .xlsm when adding macros)" }
    ],
    keywords: ["record macro", "Developer tab", "macro-enabled workbook", "xlsm", "relative references", "absolute recording", "assign macro to button", "macro security", "Enable Content", "Personal Macro Workbook", "PERSONAL.XLSB"]
  },
  {
    id: "l6-vba-fundamentals",
    title: "VBA Fundamentals",
    difficulty: "Advanced",
    duration: 22,
    version: "",
    objectives: [
      "Open the VBA editor and create a module",
      "Write a Sub and declare variables with Dim (String, Long, Double)",
      "Refer to cells with Range and Cells, and group formatting with With blocks",
      "Communicate with the user through MsgBox and InputBox",
      "Comment code and step through it line by line with F8"
    ],
    intro: "Recording can repeat what you do, but it cannot think. It cannot loop over 4,000 rows, decide that this one is overdue, format only those rows, and write a summary sheet. For that you write VBA directly — and it is more approachable than most people expect, because VBA is English-like and you already know the objects: workbooks, worksheets, ranges.\n\nThis lesson builds the vocabulary you will use for the rest of your career in Excel automation: the editor, a Sub, variables, cell references, With blocks, messages to the user, comments, and stepping through code one line at a time. Everything here runs in any modern Excel, with no add-ons.",
    sections: [
      {
        heading: "The VBA editor: your workbench",
        blocks: [
          { type: "steps", items: [
            "Press Alt+F11 (or Developer > Visual Basic). The VBA editor opens as a separate window over Excel.",
            "Learn the three panels: the Project Explorer (top left) lists every open workbook and its parts; the Properties panel (below it) shows details of the selected item; the large code window is where you type.",
            "In the Project Explorer, find your workbook (it appears as VBAProject (Sunrise-Traders-Macros.xlsm)). Right-click it > Insert > Module.",
            "A module is a blank code page. Type your code there. Modules you insert sit under the Modules folder.",
            "Switch between Excel and the editor with Alt+F11 — it toggles."
          ] },
          { type: "p", text: "One caution before writing anything: macros saved in a workbook need the file to be .xlsm (the previous lesson's rule). Also notice worksheets have their own code pages in the Project Explorer — ignore those for now; standard modules are where everyday macro code lives." },
          { type: "tip", text: "Turn on two helpers while you learn: in the editor, Tools > Options > Editor — tick \"Require Variable Declaration\" (it adds Option Explicit at the top of new modules, catching spelling mistakes in variable names before they become silent bugs) and make sure \"Auto List Members\" is ticked, which gives you the dropdown of valid properties as you type." }
        ]
      },
      {
        heading: "Subs and variables",
        blocks: [
          { type: "p", text: "A macro is a Sub — a subroutine, a named block of instructions that runs from Sub to End Sub. Here is the smallest useful one, explained line by line:" },
          { type: "code", lang: "vba", text: "Sub FillReportTitle()\n  ' Puts the title in A1 and today's date in A2\n\n  Range(\"A1\").Value = \"Sunrise Traders - Daily Sales\"\n  Range(\"A2\").Value = \"Report date: \" & Format(Date, \"dd-mmm-yyyy\")\n\nEnd Sub" },
          { type: "ul", items: [
            "Sub FillReportTitle() — starts the routine and names it. The name must have no spaces; use capitals inside words (FillReportTitle) so it stays readable.",
            "Lines starting with an apostrophe are comments — notes for humans, ignored by Excel. Every macro you write should carry a one-line comment saying what it does.",
            "Range(\"A1\").Value = ... — sets the Value property of cell A1. Reading it right to left: the text on the right is stored into the cell on the left.",
            "& joins text pieces together. Format(Date, \"dd-mmm-yyyy\") turns today's date into \"14-Nov-2026\", so A2 reads \"Report date: 14-Nov-2026\".",
            "End Sub — the routine is finished."
          ] },
          { type: "p", text: "Variables are named boxes for values, created with the Dim statement and a type. Three types cover most office work:" },
          { type: "table", headers: ["Type", "Holds", "Examples"], rows: [
            ["String", "Text", "\"Arjun Mehta\", \"S002\", \"Overdue\""],
            ["Long", "Whole numbers", "Row numbers, quantities, counts: 12, 4500, -3"],
            ["Double", "Numbers with decimals", "Money amounts, rates: 2520.5, 0.18, 123456.78"]
          ] },
          { type: "code", lang: "vba", text: "Sub OrderSummary()\n  ' Shows a one-line summary of a sample order\n\n  Dim product As String\n  Dim qty As Long\n  Dim price As Double\n  Dim total As Double\n\n  product = \"Staplers\"\n  qty = 12\n  price = 210\n  total = qty * price\n\n  MsgBox \"Sold \" & qty & \" \" & product & \" for \" & total, vbInformation\n\nEnd Sub" },
          { type: "p", text: "Read the Dim block as a shopping list: \"I will need a text box, two number boxes\". Then values go in, calculations happen in memory (no worksheet involved), and MsgBox shows the result: \"Sold 12 Staplers for 2520\". Notice that qty and price could have come from cells instead of typed values — that is the bridge to automation: variables read from the sheet, calculations happen in code, results go back to the sheet." },
          { type: "warning", text: "Without Option Explicit (Tools > Options > Editor > Require Variable Declaration), a typo like qyt = 12 silently creates a brand-new empty variable and your totals come out wrong with no error. With it, VBA stops and complains about the undeclared name — a one-second fix instead of a half-hour hunt." }
        ]
      },
      {
        heading: "Referring to cells: Range and Cells",
        blocks: [
          { type: "p", text: "Two ways to address cells, both essential. Range uses the address you see on the sheet — best when the address is known and fixed. Cells uses row and column numbers — best when the position is calculated while the code runs, which is exactly what loops will need in the next lesson." },
          { type: "code", lang: "vba", text: "Range(\"B2\").Value = \"Chennai\"        ' the cell B2\nRange(\"A1:E1\").Font.Bold = True      ' a whole range at once\n\nCells(2, 3).Value = \"Chennai\"        ' row 2, column 3 = C2\nCells(1, 1).Value = \"Date\"           ' A1 — the top-left corner" },
          { type: "p", text: "The two notations can mix: Cells(5, 2) is B5, and for a whole row of the table you can write Range(Cells(r, 1), Cells(r, 5)). In practice, professionals use Range for constants and Cells inside loops, because \"row r, column 4\" is easy to compute but hard to spell as a letter address." },
          { type: "p", text: "When several properties of the same object are being set, a With block removes the repetition and reads cleanly:" },
          { type: "code", lang: "vba", text: "With Range(\"A1:E1\")\n  .Font.Bold = True\n  .Interior.Color = RGB(255, 217, 102)\n  .HorizontalAlignment = xlCenter\nEnd With" },
          { type: "p", text: "Everything from With to End With applies to Range(\"A1:E1\"); each inner line starts with a dot, meaning \"...of the object named above\". This is the tidy version of what the macro recorder produced in the previous lesson — same actions, half the words, no selecting at all." },
          { type: "tip", text: "You rarely need .Select. Recorded code selects everything because you clicked everything; handwritten code can set properties directly on ranges. Selecting is slower and breaks easily when the user has scrolled elsewhere. The habit \"write to the range, don't select it\" separates clean code from recorded code." }
        ]
      },
      {
        heading: "Talking to the user: MsgBox and InputBox",
        blocks: [
          { type: "p", text: "A macro that finishes in silence leaves users unsure whether anything happened. Two built-in dialogs cover everyday communication. MsgBox shows a message (with an optional icon); InputBox asks a question and returns what the user typed." },
          { type: "code", lang: "vba", text: "Sub WhoIsIt()\n  ' Asks for a name and greets the user\n\n  Dim who As String\n\n  who = InputBox(\"Which salesperson is this report for?\")\n  If who = \"\" Then Exit Sub          ' user pressed Cancel — stop politely\n\n  MsgBox \"Preparing the report for \" & who, vbInformation\n\nEnd Sub" },
          { type: "p", text: "Three details worth copying into your habits. InputBox returns a String, so it lands in a String variable. If the user presses Cancel, the result is an empty string — checking for \"\" and exiting politely prevents a report \"for nobody\". And MsgBox takes an icon argument: vbInformation for neutral news, vbExclamation for warnings, vbCritical for errors." },
          { type: "warning", text: "Do not use MsgBox as a logging system: a message box after every row processed in a 4,000-row loop means 4,000 clicks. Report once at the end (\"Formatted 62 overdue invoices\"), or write progress into a status cell on the sheet, which the user can read without dismissing anything." }
        ]
      },
      {
        heading: "Stepping through code: your best debugging tool",
        blocks: [
          { type: "p", text: "Code rarely works perfectly the first time, and VBA gives you a microscope. With the cursor inside a Sub, press F8: the editor runs one line and highlights the next in yellow. Press F8 again and again to watch your macro execute line by line, while you hover the mouse over any variable to see what it currently holds." },
          { type: "steps", items: [
            "Click inside the Sub you want to examine (so the cursor is in its code).",
            "Press F8. The first line highlights yellow — nothing has run yet; the highlight marks the line that will run when you press F8 again.",
            "Keep pressing F8. Hover over variables (like total) to see their current values as tooltips.",
            "Press Ctrl+F8 to jump a whole line without stepping into it, or click the reset square button in the toolbar to stop mid-run.",
            "When the code works, run it normally with F5 (the Play button) — no stepping."
          ] },
          { type: "p", text: "Stepping is how you answer \"why did my macro write the wrong thing?\" in minutes instead of guesswork: run to the suspicious line, look at what the variables actually contain, and the bug usually announces itself. Between F8 stepping and comments, you have the complete junior-debugging toolkit — and it is enough for 90% of everyday VBA problems." },
          { type: "p", text: "Type the OrderSummary macro from this lesson into a module, step through it once with F8, and watch product, qty, price and total fill in as you hover. That one exercise makes variables, assignment and MsgBox concrete in a way reading never will." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Writing Sub Format Report() with a space in the name.", right: "Use FormatReport or Format_Report — no spaces in macro names.", why: "VBA names cannot contain spaces. The editor rejects the line, or worse, beginners rename it five times trying variants before asking why." },
      { wrong: "Declaring no variables and assigning values straight away (qty = 12 without Dim).", right: "Declare every variable: Dim qty As Long before first use — or enable Require Variable Declaration.", why: "Undeclared variables work but make every typo a silent new variable. Declared, typed variables let VBA's own checks catch the mistakes for you." },
      { wrong: "Writing Range(\"B\" & r) when you mean row r of column C.", right: "Use Cells(r, 3).Value for \"row r, column 3\".", why: "Cells takes numbers, which loops can compute. Building letter addresses by string arithmetic is slow to write, easy to get wrong, and unnecessary." },
      { wrong: "Copying recorded code with Select everywhere and editing the values around it.", right: "Write directly against ranges: Range(\"A1:E1\").Font.Bold = True instead of Range(\"A1:E1\").Select + Selection.Font.Bold = True.", why: "Selecting depends on the right sheet being active and the scroll position being sane. Direct property setting is faster, shorter, and immune to what the user was doing at the time. The same rule applies to messages: report once after the work, never inside a loop." }
    ],
    shortcuts: [
      { keys: "Alt + F11", action: "Toggle between Excel and the VBA editor" },
      { keys: "F8", action: "Step through code one line at a time from the cursor" },
      { keys: "F5", action: "Run the Sub the cursor is in (without stepping)" },
      { keys: "Ctrl + G", action: "Open the Immediate window to print variable values while debugging" }
    ],
    exercises: [
      { task: "Open assets/files/level6/macro-practice.xlsx (saved as .xlsm from the previous lesson if you added macros there). Open the VBA editor, insert a module, and type the FillReportTitle macro. Run it with F5 and check the sheet.", hint: "Alt+F11, right-click your workbook's VBAProject > Insert > Module. The comment line and both Range lines go between Sub and End Sub.", solution: "A1 shows \"Sunrise Traders - Daily Sales\" and A2 shows \"Report date: \" followed by today's date, which updates every time the macro runs." },
      { task: "Type the OrderSummary macro, then run it once with F8 and watch the four variables fill in as you hover over them. Change the values (product \"Notebooks\", qty 48, price 120) and predict the message before running.", hint: "Click inside the Sub first so F8 knows where to start. Hover over each variable name to see its current contents.", solution: "The message reads \"Sold 48 Notebooks for 5760\" — the multiplication happens in memory, then MsgBox displays the joined text." },
      { task: "Write a macro CellTest that puts \"Staplers\" in C2 two ways: once with Range(\"C2\") and once with Cells(2, 3). Run it and confirm both lines do the same thing.", hint: "Cells takes (row, column) numbers: row 2, column 3 is C2.", solution: "Range(\"C2\").Value = \"Staplers\" and Cells(2, 3).Value = \"Staplers\" — the second form is the one loops will use, because row numbers can be computed." },
      { task: "Replace the FormatReport recording from the previous lesson with a hand-written With block version (bold, RGB(255,217,102) fill, centered, no .Select anywhere). Run it and compare with the recorded one.", hint: "Three properties inside one With Range(\"A1:E1\"): Font.Bold, Interior.Color, HorizontalAlignment = xlCenter.", solution: "With Range(\"A1:E1\") / .Font.Bold = True / .Interior.Color = RGB(255, 217, 102) / .HorizontalAlignment = xlCenter / End With — identical result, no selecting, and it works whichever cell the cursor is in." }
    ],
    quiz: [
      { q: "What does Sub ... End Sub define?", options: ["A worksheet formula", "A named block of VBA instructions that can be run as a macro", "A keyboard shortcut", "A cell reference"], answer: 1, explanation: "A Sub is a routine with a name; Excel's Macros dialog and your buttons run it. Everything between Sub and End Sub executes in order." },
      { q: "Which Dim line is appropriate for storing the amount ₹1,250.75?", options: ["Dim amount As String", "Dim amount As Long", "Dim amount As Double", "Dim amount As Cell"], answer: 2, explanation: "The value has decimals, so Double. Long only stores whole numbers (it would round the .75 away), and String would store text that cannot be calculated with." },
      { q: "Cells(4, 2) refers to which cell?", options: ["B4", "A4", "D2", "B2"], answer: 0, explanation: "Cells takes (row, column): row 4, column 2 is B4. This numeric form is what makes loops over rows possible." },
      { q: "What is the purpose of a line starting with an apostrophe in VBA?", options: ["It prints text to the worksheet", "It is a comment for humans, ignored when the code runs", "It declares a variable", "It ends the Sub"], answer: 1, explanation: "The apostrophe marks a comment. Well-commented code explains itself to whoever opens it next — including you in six months." },
      { q: "Your macro writes the wrong number somewhere. What is the fastest way to find out why?", options: ["Reread the whole macro and hope", "Step through with F8 and hover over the variables to see their live values", "Delete the macro and re-record it", "Restart Excel"], answer: 1, explanation: "F8 runs one line at a time and tooltips show what each variable holds at that moment. Watching the values change pinpoints the line where reality diverges from intention." }
    ],
    summary: [
      "Alt+F11 opens the VBA editor; Insert > Module gives you a code page; every macro lives as Sub ... End Sub.",
      "Dim creates typed variables — String for text, Long for whole numbers, Double for decimals — and Option Explicit catches typos for you.",
      "Range(\"B2\") addresses fixed cells; Cells(r, c) addresses by numbers, which is what loops need; With blocks group repeated property settings without Select.",
      "MsgBox reports results (with vbInformation / vbExclamation icons) and InputBox asks for input; always handle Cancel and report at the end, not per row.",
      "Comments (apostrophe lines) explain the code, and F8 stepping is the everyday debugger — watch variables live instead of guessing."
    ],
    downloads: [
      { file: "assets/files/level6/macro-practice.xlsx", label: "Macro practice workbook", note: "A plain workbook with sample data to record and run your first macros (re-save as .xlsm when adding macros)" },
      { file: "assets/files/level6/macro-code.bas", label: "VBA macro code file (.bas)", note: "Ready-to-import VBA module with the macros taught in the VBA lessons" }
    ],
    keywords: ["VBA", "VBA editor", "Alt F11", "Sub End Sub", "Dim variables", "String Long Double", "Range Cells", "With block", "MsgBox", "InputBox", "F8 step through", "comments"]
  },
  {
    id: "l6-vba-automation",
    title: "Automating Reports with VBA",
    difficulty: "Advanced",
    duration: 25,
    version: "",
    objectives: [
      "Loop over rows with For...Next and find the last used row",
      "Make decisions in code with If...Then...ElseIf...End If",
      "Format ranges and write values from code",
      "Create a new sheet and write a summary onto it",
      "Save the workbook with a date-stamped name",
      "Add basic error handling with On Error and assemble a full Monthly Report macro"
    ],
    intro: "This lesson assembles everything so far into the macro every office eventually needs: a Monthly Report routine. Point it at the sales data, click once, and it scans every row, highlights what matters, writes a summary onto a fresh sheet, and saves a dated copy of the workbook — the whole month-end chore in a few seconds.\n\nThe building blocks are small: one loop, one If, some formatting, one new sheet, one save. Each is introduced on its own with a runnable snippet, then the final section combines them into a single macro you can copy from the provided .bas file. A closing note explains how professionals really work: record the parts you can, edit what the recorder wrote, and hand-write only the logic.",
    sections: [
      {
        heading: "The monthly report job, planned",
        blocks: [
          { type: "p", text: "At Sunrise Traders, the month-end routine looks like this: open the Sales sheet, scan every row, highlight the laptop sales (they need a follow-up call about extended warranty), total the month, create a summary sheet with the key figures, and file the workbook as \"Sunrise Traders 2026-11-14.xlsm\". By hand: 20 minutes of careful scrolling, one missed row away from a wrong total. By macro: one click, identical every month." },
          { type: "p", text: "Planning before coding is the professional habit. We need five capabilities, and the lesson covers them in this order:" },
          { type: "ol", items: [
            "Visit every data row automatically — a For loop.",
            "Decide something about each row (is this a laptop sale?) — an If statement.",
            "Act on matches and totals — formatting and arithmetic.",
            "Create a fresh sheet and write the summary — Worksheets.Add.",
            "Save the result with today's date in the filename — SaveAs. Plus: survive surprises — On Error."
          ] },
          { type: "p", text: "The snippets below assume the Sales sheet layout used throughout: column A the date, B the salesperson, C the product category, D the amount in ₹. That is also the layout of the practice workbook, so everything here runs on it unchanged." }
        ]
      },
      {
        heading: "For loops over rows",
        blocks: [
          { type: "p", text: "A For loop repeats a block of code while a counter steps through values. For \"do something to every data row\", the counter is the row number. The only trick is finding where the data ends — hardcoding row 400 works until month 9 brings 470 rows. The standard ending-finder is one line worth memorising:" },
          { type: "code", lang: "vba", text: "Dim lastRow As Long\nDim r As Long\n\nlastRow = Cells(Rows.Count, \"A\").End(xlUp).Row\n\nFor r = 2 To lastRow\n  ' this block runs once for each data row\n  Debug.Print \"Processing row \" & r\nNext r" },
          { type: "ul", items: [
            "Cells(Rows.Count, \"A\") is the very last cell of column A (A1048576 on a normal sheet).",
            ".End(xlUp) is the keyboard equivalent of pressing Ctrl+Up from there — it lands on the last cell in column A that contains anything.",
            ".Row returns that cell's row number. Together: lastRow is exactly where your data ends, however long it is this month.",
            "For r = 2 To lastRow starts at 2 because row 1 is the header. Next r adds 1 to r and loops until the end."
          ] },
          { type: "p", text: "Debug.Print writes to the Immediate window (Ctrl+G in the editor) instead of the sheet — a quiet way to watch a loop work while testing. Inside the loop body you can read any cell of the current row with Cells(r, column): Cells(r, \"C\").Value is this row's product category." },
          { type: "tip", text: "Always compute lastRow from a column you know is filled on every data row (the date column here). Computing it from a sparsely filled column stops the loop early and silently skips data — a classic, hard-to-notice bug." }
        ]
      },
      {
        heading: "If conditions and the actions inside the loop",
        blocks: [
          { type: "p", text: "Inside the loop, each row needs a decision. VBA's If reads almost like English, and it must be closed with End If when the actions span several lines:" },
          { type: "code", lang: "vba", text: "For r = 2 To lastRow\n  If Cells(r, \"C\").Value = \"Laptops\" Then\n    ' one-line action goes here\n  ElseIf Cells(r, \"C\").Value = \"Printers\" Then\n    ' another case\n  Else\n    ' everything else\n  End If\nNext r" },
          { type: "p", text: "Comparisons work as you would hope: = for equality, <> for not-equal, >, <, >=, <= for numbers. Text comparisons are case-insensitive by default, so \"laptops\" matches \"Laptops\" — convenient for data typed by humans. Combine conditions with And and Or, exactly as in worksheet formulas: If Cells(r, \"C\").Value = \"Laptops\" And Cells(r, \"D\").Value > 100000 Then." },
          { type: "warning", text: "A cell containing a formula error (#DIV/0!, #N/A) makes reading .Value raise a type-mismatch error inside the loop. If your source data can contain error cells, test with If Not IsError(Cells(r, \"D\").Value) before doing arithmetic on the value." },
          { type: "p", text: "When a row matches, the macro should act. Highlighting a whole row is one statement thanks to the Rows collection, and totals are just arithmetic on variables — no worksheet round-trips. The money amounts accumulate in Double variables and are formatted only at the end:" },
          { type: "code", lang: "vba", text: "Dim total As Double\nDim laptopTotal As Double\n\ntotal = 0\nlaptopTotal = 0\n\nFor r = 2 To lastRow\n  total = total + Cells(r, \"D\").Value\n\n  If Cells(r, \"C\").Value = \"Laptops\" Then\n    Rows(r).Interior.Color = RGB(255, 235, 205)\n    laptopTotal = laptopTotal + Cells(r, \"D\").Value\n  End If\nNext r\n\nMsgBox \"Total: \" & Format(total, \"#,##0\") & vbNewLine & _\n       \"Laptops: \" & Format(laptopTotal, \"#,##0\"), vbInformation" },
          { type: "p", text: "Two details to notice. Format(total, \"#,##0\") turns 1689200 into \"1,689,200\" for the message — the same formatting idea as the worksheet TEXT function. And the underscore at the end of a line continues the statement on the next line, used here to keep the long MsgBox readable. The message shows once, after the loop — reporting inside the loop would mean clicking a dialog for every row." }
        ]
      },
      {
        heading: "Writing the summary to a new sheet",
        blocks: [
          { type: "p", text: "A summary buried in a message box is gone after the click; a summary on a sheet becomes part of the report. Creating a sheet and writing to it is four statements:" },
          { type: "code", lang: "vba", text: "Dim rep As Worksheet\n\nSet rep = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count))\nrep.Name = \"Report \" & Format(Date, \"yyyy-mm\")\n\nrep.Range(\"A1\").Value = \"Sunrise Traders - Monthly Report\"\nrep.Range(\"A2\").Value = \"Total sales\"\nrep.Range(\"B2\").Value = total\nrep.Range(\"A3\").Value = \"Laptop sales\"\nrep.Range(\"B3\").Value = laptopTotal\nrep.Range(\"B2:B3\").NumberFormat = \"#,##0\"\nrep.Range(\"A1\").Font.Bold = True" },
          { type: "ul", items: [
            "Set rep = ... Worksheets.Add adds a sheet and stores a reference to it in the variable rep — after this, every rep.Range(\"...\") points at the new sheet, wherever it is.",
            "After:=Sheets(Sheets.Count) places the new sheet at the very end of the workbook instead of jumping to the front.",
            "The name uses yyyy-mm (\"Report 2026-11\"), so each month's summary sheet sorts naturally. Writing the raw numbers into B2:B3 and then formatting the range keeps values and display separate.",
            "Because rep is a real worksheet reference, everything you know about ranges works: NumberFormat, Font.Bold, Interior.Color, all of it."
          ] },
          { type: "warning", text: "If a sheet named \"Report 2026-11\" already exists (you ran the macro twice this month), the .Name line fails with an error. The full macro below wraps that one line in a local On Error Resume Next so a second run simply reuses the existing sheet's twin — or you can delete the old sheet first." }
        ]
      },
      {
        heading: "Saving a dated copy, and On Error basics",
        blocks: [
          { type: "p", text: "Filing copies by date is how finance teams keep an audit trail. One SaveAs line does it — with the file format that keeps macros alive:" },
          { type: "code", lang: "vba", text: "ThisWorkbook.SaveAs Filename:= _\n  ThisWorkbook.Path & Application.PathSeparator & _\n  \"Sunrise Traders \" & Format(Date, \"yyyy-mm-dd\") & \".xlsm\", _\n  FileFormat:=xlOpenXMLWorkbookMacroEnabled" },
          { type: "ul", items: [
            "ThisWorkbook is the workbook containing the code — safer than ActiveWorkbook, which is whichever workbook happens to be in front.",
            "The name becomes \"Sunrise Traders 2026-11-14.xlsm\"; the date stamp means each day's file keeps its own copy and never overwrites the previous one.",
            "FileFormat:=xlOpenXMLWorkbookMacroEnabled is .xlsm — the constant matters, because a plain SaveAs would try to save as .xlsx and the macros would vanish.",
            "If the file already exists, Excel asks before overwriting; if the user cancels, an error is raised — which brings us to error handling."
          ] },
          { type: "p", text: "Error handling is a safety net for the surprises you cannot foresee: a missing sheet, a cancelled save, a renamed file. The basic pattern is On Error GoTo plus a labelled exit at the end of the Sub:" },
          { type: "code", lang: "vba", text: "Sub MyMacro()\n  On Error GoTo ErrHandler\n\n  ' ...all the normal code...\n  Exit Sub                 ' normal end - skips the handler\n\nErrHandler:\n  MsgBox \"The macro stopped: \" & Err.Description, vbExclamation\nEnd Sub" },
          { type: "p", text: "If any line fails, execution jumps to ErrHandler, shows the reason (Err.Description), and ends cleanly instead of leaving the user with a cryptic VBA error dialog and a half-formatted sheet. Related tools: On Error Resume Next ignores the next error (use sparingly, and locally — around one risky line such as sheet naming), and On Error GoTo 0 switches handling back off. Never wrap a whole loop in Resume Next — silent failure is worse than an error message." },
          { type: "tip", text: "Test your error handler on purpose: temporarily rename the Sales sheet to \"Sales old\" and run the macro. It should stop with a clear message instead of crashing. Handlers you have never seen fire are handlers that do not work." }
        ]
      },
      {
        heading: "Putting it together: the MonthlyReport macro",
        blocks: [
          { type: "p", text: "Here is the complete routine — every piece from this lesson in one Sub. It is included in the downloadable .bas module, so you can import it via File > Import File in the editor instead of typing it." },
          { type: "code", lang: "vba", text: "Sub MonthlyReport()\n  ' Scans the Sales sheet, highlights laptop sales,\n  ' writes a summary sheet and saves a dated copy.\n\n  Dim ws As Worksheet, rep As Worksheet\n  Dim lastRow As Long, r As Long\n  Dim total As Double, laptopTotal As Double\n\n  On Error GoTo ErrHandler\n\n  Set ws = ThisWorkbook.Worksheets(\"Sales\")\n  lastRow = ws.Cells(ws.Rows.Count, \"A\").End(xlUp).Row\n\n  total = 0\n  laptopTotal = 0\n  For r = 2 To lastRow\n    total = total + ws.Cells(r, \"D\").Value\n    If ws.Cells(r, \"C\").Value = \"Laptops\" Then\n      ws.Rows(r).Interior.Color = RGB(255, 235, 205)\n      laptopTotal = laptopTotal + ws.Cells(r, \"D\").Value\n    End If\n  Next r\n\n  Set rep = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count))\n  On Error Resume Next           ' name may already exist this month\n  rep.Name = \"Report \" & Format(Date, \"yyyy-mm\")\n  On Error GoTo ErrHandler\n\n  rep.Range(\"A1\").Value = \"Sunrise Traders - Monthly Report\"\n  rep.Range(\"A2\").Value = \"Total sales\"\n  rep.Range(\"B2\").Value = total\n  rep.Range(\"A3\").Value = \"Laptop sales\"\n  rep.Range(\"B3\").Value = laptopTotal\n  rep.Range(\"B2:B3\").NumberFormat = \"#,##0\"\n  rep.Range(\"A1\").Font.Bold = True\n\n  ThisWorkbook.SaveAs Filename:= _\n    ThisWorkbook.Path & Application.PathSeparator & _\n    \"Sunrise Traders \" & Format(Date, \"yyyy-mm-dd\") & \".xlsm\", _\n    FileFormat:=xlOpenXMLWorkbookMacroEnabled\n\n  MsgBox \"Report ready. Total \" & Format(total, \"#,##0\"), vbInformation\n  Exit Sub\n\nErrHandler:\n  MsgBox \"The macro stopped: \" & Err.Description, vbExclamation\nEnd Sub" },
          { type: "steps", items: [
            "Save a copy of the practice workbook as .xlsm (a macro-free .xlsx cannot store this code), or work on your own copy.",
            "Import the .bas file: in the editor, File > Import File, pick macro-code.bas — or type the macro into a module yourself.",
            "Run MonthlyReport from Alt+F8 or a button, and step through it once with F8 to watch lastRow, total and laptopTotal fill in.",
            "Run it a second time in the same month: the summary sheet naming is protected, and the SaveAs will ask about overwriting today's file — Cancel is safe and handled."
          ] },
          { type: "p", text: "One closing piece of professional advice. Nobody writes macros like this from a blank page. The realistic workflow is: record the mechanical parts (a format here, a paste there), read the recorded code, replace its Select-and-click style with the direct patterns from the VBA fundamentals lesson, and hand-write only the logic the recorder cannot do — the loop, the If, the summary. Record, edit, extend. That is how the MonthlyReport macro in any real company was actually built." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Hardcoding the loop end: For r = 2 To 400.", right: "Compute the real end: lastRow = Cells(Rows.Count, \"A\").End(xlUp).Row, from a column that is filled on every data row (here the date in column A).", why: "Fixed ends stop working the moment the data grows or shrinks — either crashing on empty rows or silently skipping the newest sales. And End(xlUp) only adapts if the column you measure is filled everywhere: a sparse column makes the loop stop mid-table with no warning." },
      { wrong: "Saving the automated workbook with a plain SaveAs and no FileFormat argument.", right: "Specify FileFormat:=xlOpenXMLWorkbookMacroEnabled (.xlsm).", why: "Without the format argument the save follows Excel's default (.xlsx) and strips all macros — the automation survives exactly once, then disappears." },
      { wrong: "Wrapping the entire macro in On Error Resume Next \"so it never crashes\".", right: "Use On Error GoTo ErrHandler for the routine, and Resume Next only around single known-risky lines (like naming a possibly-existing sheet).", why: "Resume Next swallows every failure silently: a wrong total, a missing sheet, a skipped section — the macro appears to succeed while producing wrong results. GoTo handler keeps failures visible and explained." },
      { wrong: "Referring to ActiveSheet and ActiveWorkbook in shared macros.", right: "Reference objects explicitly: ThisWorkbook.Worksheets(\"Sales\"), and store sheets in variables (Set ws = ...).", why: "Active objects depend on what the user has open and clicked. Explicit references run correctly no matter who clicks the button or what else they had open." }
    ],
    shortcuts: [
      { keys: "F8", action: "Step through the macro line by line to watch lastRow and the totals fill in" },
      { keys: "Ctrl + G", action: "Open the Immediate window to see Debug.Print output from loops" },
      { keys: "Alt + F8", action: "Run MonthlyReport without touching the editor" },
      { keys: "Ctrl + Up", action: "The worksheet equivalent of End(xlUp) — the pattern the lastRow line automates" }
    ],
    exercises: [
      { task: "In the practice workbook (saved as .xlsm), write a macro LoopTest that computes lastRow from column A and shows it with MsgBox. Add data rows to the sheet and run again to confirm the number follows the data.", hint: "lastRow = Cells(Rows.Count, \"A\").End(xlUp).Row, then MsgBox lastRow. The header row counts too if A1 is filled.", solution: "The number equals the last row containing a date in column A; adding or removing rows changes it on the next run — no hardcoded 400 anywhere." },
      { task: "Write a macro that loops the Sales rows and shows, in one MsgBox at the end, how many Pen sales occurred (count rows where column C is \"Pens\").", hint: "Dim penCount As Long, start at 0, add 1 inside the If inside the For, report after Next r.", solution: "For r = 2 To lastRow / If Cells(r, \"C\").Value = \"Pens\" Then penCount = penCount + 1 / End If / Next r, then MsgBox penCount." },
      { task: "Create the summary sheet code from the lesson as its own macro WriteSummary using two sample totals (total = 1859000, laptopTotal = 525000), then run it twice and observe what happens to the sheet name line.", hint: "Worksheets.Add(After:=Sheets(Sheets.Count)), name it \"Report \" & Format(Date, \"yyyy-mm\"), write labels in A and values in B, format B as \"#,##0\".", solution: "The first run creates \"Report 2026-11\"; the second fails on .Name unless you wrap it in On Error Resume Next — exactly the protection the full MonthlyReport macro includes." },
      { task: "Import assets/files/level6/macro-code.bas into the practice workbook (File > Import File in the editor), run MonthlyReport once, then test the error handling by renaming the Sales sheet and running again.", hint: "After renaming, Worksheets(\"Sales\") cannot be found — the handler should catch it. Rename the sheet back afterwards.", solution: "Instead of a cryptic VBA error, a message appears: \"The macro stopped: ...\" with the reason (subscript out of range for the missing sheet). The workbook is left in a clean state." }
    ],
    quiz: [
      { q: "What does lastRow = Cells(Rows.Count, \"A\").End(xlUp).Row compute?", options: ["The number 1048576", "The row number of the last filled cell in column A", "The count of columns used", "The last row of the used range on all sheets"], answer: 1, explanation: "Start at the bottom of column A, jump up (End(xlUp), like Ctrl+Up) to the last filled cell, and take its row number. It adapts to however much data arrived this month." },
      { q: "Inside For r = 2 To lastRow, how do you read this row's amount?", options: ["Range(\"D\" & \"r\").Value", "Cells(r, \"D\").Value", "Rows.D.Value", "Range(\"D\").Rows(r).Value"], answer: 1, explanation: "Cells(row, column) is the loop-friendly reference: Cells(r, \"D\") is row r of column D. The variable r changes each pass, so the reference follows the loop." },
      { q: "What does On Error GoTo ErrHandler do?", options: ["Deletes the code that caused the error", "Jumps to a labelled section that reports the problem cleanly instead of crashing", "Pauses the macro until the user fixes the error", "Retries the failed line forever"], answer: 1, explanation: "When any line fails, execution jumps to the ErrHandler label — typically showing Err.Description in a message box. The user gets a clear explanation instead of a cryptic VBA dialog." },
      { q: "Why specify FileFormat:=xlOpenXMLWorkbookMacroEnabled in SaveAs?", options: ["It makes the file open faster", "It saves as .xlsm so the macros stay in the file", "It encrypts the workbook", "It is the required format for pivot tables"], answer: 1, explanation: "A plain SaveAs defaults to .xlsx, which cannot contain VBA. The macro-enabled format keeps the automation alive in the saved copy." },
      { q: "The most realistic way professionals build a report macro like MonthlyReport is:", options: ["Write every line from memory with no testing", "Record the mechanical parts, then edit and extend the recorded code with hand-written logic", "Copy a macro from the internet and hope it fits", "Hire a programmer for every report"], answer: 1, explanation: "Record, edit, extend: the recorder handles formats and clicks you would otherwise look up; you hand-write the loop, If and summary logic it cannot think through." }
    ],
    summary: [
      "Find the data end with Cells(Rows.Count, \"A\").End(xlUp).Row — never hardcode the last row.",
      "For r = 2 To lastRow ... Next r visits every data row; inside it, Cells(r, \"C\").Value reads that row, and If ... End If makes decisions.",
      "Act in code, not by selecting: Rows(r).Interior.Color highlights, variables accumulate totals, Format(x, \"#,##0\") displays them readably.",
      "Write summaries to a new sheet via Set rep = Worksheets.Add(...): labels, values, NumberFormat — and guard the sheet-name line for repeat runs.",
      "Save dated copies with SaveAs and FileFormat:=xlOpenXMLWorkbookMacroEnabled, and wrap the routine in On Error GoTo so failures explain themselves. Build macros by recording, editing, then hand-writing the logic."
    ],
    downloads: [
      { file: "assets/files/level6/macro-practice.xlsx", label: "Macro practice workbook", note: "A plain workbook with sample data to record and run your first macros (re-save as .xlsm when adding macros)" },
      { file: "assets/files/level6/macro-code.bas", label: "VBA macro code file (.bas)", note: "Ready-to-import VBA module with the macros taught in the VBA lessons" }
    ],
    keywords: ["VBA automation", "For Next loop", "last row End xlUp", "If Then VBA", "Worksheets.Add", "SaveAs xlsm", "On Error GoTo", "Monthly Report macro", "date-stamped filename", "record and edit macros"]
  }
);
