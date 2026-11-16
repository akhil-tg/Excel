// Excel Lab — Level 3: Data Management (part 1: level metadata + lessons 1-5)
window.LEVEL_META = window.LEVEL_META || {};
window.LEVEL_LESSONS = window.LEVEL_LESSONS || {};

window.LEVEL_META[3] = {
  id: 3,
  title: "Data Management",
  subtitle: "Sort, filter, clean and organize data like a professional",
  description: "Take control of your data: Excel Tables, sorting, filtering, validation, conditional formatting and complete data-cleaning workflows.",
  difficulty: "Beginner",
  icon: "database"
};

(window.LEVEL_LESSONS[3] = window.LEVEL_LESSONS[3] || []).push(

  {
    id: "l3-sorting",
    title: "Sorting Data",
    difficulty: "Beginner",
    duration: 12,
    version: "",
    objectives: [
      "Sort a list A to Z or Z to A in one click, safely",
      "Sort by two or more columns, such as region first and then revenue",
      "Sort by cell colour and by your own custom lists",
      "Explain when sorting can break formulas and how Tables protect you"
    ],
    intro: "Data rarely arrives in the order you need. Your manager may want the biggest orders first, or names A to Z. Sorting does this in seconds.\n\nBut sorting moves every row. That can affect your formulas. Here you will learn to sort fast, sort safely, and spot when sorting can cause damage.",
    keyIdea: "Sorting puts rows in order, but Excel sorts whole rows together so your data never gets mixed up.",
    sections: [
      {
        heading: "Why sorting matters in office work",
        blocks: [
          { type: "p", text: "Every report starts as a list: sales records, invoices, employees, inventory. The order of that list almost never matches the question being asked. Sorting answers many questions with zero formulas:" },
          { type: "ul", items: [
            "**Which orders were the biggest?** Sort Revenue from largest to smallest and read the top rows.",
            "**Which customers do we contact today?** Sort the Next Contact date from oldest to newest.",
            "**Show me all North region rows together.** Sort by Region and the regions group themselves.",
            "**Who is next in the alphabet?** Sort the customer column A to Z for a clean printed list."
          ] },
          { type: "p", text: "Here is the key rule: when you sort, whole rows move together. Row 5 holds one full order today. After sorting, it still holds that same order — just at a new row number. Excel keeps rows together automatically when it can see where your table starts and ends." }
        ]
      },
      {
        heading: "One-click sorting and the sort warning",
        blocks: [
          { type: "p", text: "The fastest way to sort is to click any single cell inside the column you want to sort by — not the whole column — and then use the Data tab:" },
          { type: "steps", items: [
            "Click one cell in the column you want to sort by, for example the Region column.",
            "Go to the Data tab and click Sort A to Z (alphabetical, smallest to largest for numbers) or Sort Z to A (reverse, largest to smallest for numbers).",
            "Excel automatically expands the sort to cover the whole surrounding table, including the header row."
          ] },
          { type: "warning", text: "If Excel shows a Sort Warning dialog, always choose **Expand the selection**. This happens when you selected a whole column yourself. \"Continue with the current selection\" sorts only that column. The rest of the row stays behind. Order IDs no longer match their products — the data is now scrambled." },
          { type: "p", text: "Excel normally spots your header row and leaves it at the top. But sometimes it sorts the header into the data. If that happens, open Data > Sort and tick **My data has headers**." }
        ]
      },
      {
        heading: "Sorting by two or more levels",
        blocks: [
          { type: "p", text: "Real questions usually need more than one sort key. Suppose you want every region grouped together, and within each region, the biggest order on top. That is a two-level sort." },
          { type: "steps", items: [
            "Click any cell inside your data and go to Data > Sort.",
            "In the first row of the dialog, set Sort by = **Region**, Sort On = **Values**, Order = **A to Z**.",
            "Click **Add Level**. Excel adds a second row: Then by. Set it to **Revenue**, Values, **Largest to Smallest**.",
            "Click OK. Excel sorts by Region first; ties within the same region are then broken by Revenue."
          ] },
          { type: "grid", title: "Before: orders in entry order, and after sorting by Region A to Z, then Revenue largest to smallest",
            rows: [
              ["Order ID", "Salesperson", "Region", "Product", "Revenue"],
              ["1004", "Sneha Iyer", "East", "Pens", "6000"],
              ["1003", "Rahul Verma", "North", "Laptops", "352000"],
              ["1001", "Arjun Mehta", "North", "Notebooks", "14400"],
              ["1002", "Priya Nair", "South", "Staplers", "6750"]
            ] },
          { type: "p", text: "Read the result. East comes first, then North, then South. Inside North, the laptop order (₹3,52,000) sits above the notebook order (₹14,400). The order of levels matters. Swap the two levels and you get the biggest orders first, with regions only breaking ties." },
          { type: "tip", text: "Sort levels are applied top to bottom. The first level defines the main grouping; every extra level only breaks ties inside the group above it." }
        ]
      },
      {
        heading: "Sort by colour and custom lists",
        blocks: [
          { type: "p", text: "Two more options inside the Sort dialog are useful in daily work." },
          { type: "p", text: "**Sort by colour.** Maybe you marked some rows with a fill colour. Open Data > Sort and set Sort On to **Cell Colour** (or Font Colour). Then pick which colour comes first. For example, put red overdue invoices at the top. One sort level handles one colour only. Add more levels for a second colour." },
          { type: "p", text: "**Custom lists.** Alphabetical order is wrong for some data. Months should read Jan, Feb, Mar. Priorities should read High, Medium, Low. In the Sort dialog, set Order to **Custom List**. Pick a built-in list like months or days, or make your own." },
          { type: "p", text: "To make your own, type High, Medium, Low on separate lines and click Add. Sunrise Traders could define its four regions this way. Reports would then always read North, South, East, West in the same order." }
        ]
      },
      {
        heading: "When sorting breaks formulas (and how Tables help)",
        blocks: [
          { type: "p", text: "Sorting moves rows. Formulas that assume a row stays in one place can silently start returning wrong numbers. Three patterns to watch:" },
          { type: "ul", items: [
            "**Formulas pointing at a specific cell.** If cell `J2` says `=F5` because \"row 5 is our top order today\", that link follows the position, not the order. After a sort, `F5` may hold a different order entirely.",
            "**Lookup ranges without anchors.** `=VLOOKUP(A2, D2:G50, 3, 0)` uses a range that starts at row 2. Sorting the sheet does not change the range address, but the data inside it moves, so results can shift in confusing ways. Anchor important lookup tables with `=VLOOKUP(A2, $D$2:$G$50, 3, 0)` or better, keep them on a separate sheet.",
            "**Sorting part of a table.** A one-column sort with \"Continue with the current selection\" tears rows apart. This is data damage, not a formula problem, and it is the most common sorting accident in real offices."
          ] },
          { type: "p", text: "Excel Tables (next lesson) reduce this risk. A Table always sorts as one unit, so rows can never be torn apart. Its formulas use structured references that move correctly with their rows. If a formula must survive sorting, build it on a Table, not on loose cell addresses." },
          { type: "quote", text: "Sort answers \"in what order should I read this?\". Filters answer \"which rows belong to this question?\". Formulas should not depend on either." }
        ]
      }
    ],
    mistakes: [
      {
        wrong: "Selecting an entire column and clicking Sort A to Z, then clicking \"Continue with the current selection\" in the warning dialog.",
        right: "Click a single cell inside the column, or if the warning appears, choose Expand the selection.",
        why: "Sorting only the selected column leaves the other columns in place, so every row becomes a mix of two different records. Undo and fix immediately if this happens."
      },
      {
        wrong: "Sorting a large dataset and then trusting a formula like `=F5` that pointed at \"the row with the biggest order\".",
        right: "Use formulas that find the answer themselves, such as `=MAX(F2:F50)` or `=INDEX(A2:A50, MATCH(MAX(F2:F50), F2:F50, 0))`, or sort a copy of the data.",
        why: "Cell addresses do not move with their data during a sort. A position-based reference points at whatever lands in that position afterwards."
      },
      {
        wrong: "Sorting without checking whether Excel recognised the header row.",
        right: "Open Data > Sort and confirm \"My data has headers\" is ticked.",
        why: "If the header row is sorted into the data, \"Region\" sits between East and North and your first data row becomes a header. Always glance at row 1 after sorting."
      }
    ],
    shortcuts: [
      { keys: "Alt + A + S + S", action: "Open the full Sort dialog" },
      { keys: "Alt + A + S + O", action: "Sort A to Z (ascending)" },
      { keys: "Alt + A + S + D", action: "Sort Z to A (descending)" },
      { keys: "Ctrl + Z", action: "Undo a sort (the fastest fix for a bad sort)" }
    ],
    exercises: [
      {
        task: "Open the Sorting sheet of the exercise workbook. Sort the order list so the highest Revenue appears first, keeping every row intact.",
        hint: "Click one cell in the Revenue column, then Data > Sort Z to A. Watch for the sort warning and always expand the selection.",
        solution: "Select any cell in the Revenue column, Data tab > Sort Z to A. The laptop order (₹3,52,000) should now be at the top with its own Order ID, salesperson and region still on the same row. Compare with the Solution sheet."
      },
      {
        task: "On the same sheet, sort so that all regions are grouped A to Z, and within each region the smallest Revenue is on top.",
        hint: "This needs two levels: Region A to Z, then Revenue Smallest to Largest.",
        solution: "Data > Sort: Level 1 = Region, Values, A to Z. Add Level: Level 2 = Revenue, Values, Smallest to Largest. Regions appear as East, North, South, West with ascending revenue inside each group."
      },
      {
        task: "The workbook also contains a small priority column with the values High, Medium and Low. Sort it so High rows come first, then Medium, then Low.",
        hint: "Alphabetical order would give High, Low, Medium. You need a custom list.",
        solution: "Data > Sort > Sort by Priority. In Order, choose Custom List, then type High, Medium, Low on three lines in the dialog and click Add, then OK. The Solution sheet shows the expected order."
      }
    ],
    quiz: [
      {
        q: "You want to sort by Region and then by Revenue within each region. How should you set this up?",
        options: [
          "Sort by Region first, then add a second level sorted by Revenue",
          "Sort by Revenue first, then by Region — the order does not matter",
          "Sort Region and Revenue in one level separated by a comma",
          "You cannot sort by two columns at once"
        ],
        answer: 0,
        explanation: "Levels apply top to bottom. The first level groups by Region; the second level only breaks ties inside each region. Reversing the order produces a different result."
      },
      {
        q: "Excel shows the Sort Warning dialog after you selected a column and clicked Sort A to Z. What should you do?",
        options: [
          "Click Continue with the current selection — it is faster",
          "Press Escape and never sort",
          "Choose Expand the selection so the whole rows sort together",
          "Close Excel and reopen the file"
        ],
        answer: 2,
        explanation: "Expand the selection keeps each row complete. Continue with the current selection sorts only the selected column and permanently scrambles the relationship between columns."
      },
      {
        q: "Why can sorting break a formula like `=F5` that was placed to show the top order?",
        options: [
          "Sorting deletes formulas it finds in the sorted range",
          "The formula references a position, and after sorting a different record sits in F5",
          "Sorting converts formulas into static values",
          "It cannot break — Excel updates cell references during a sort"
        ],
        answer: 1,
        explanation: "Cell references point at addresses, not at records. When rows move, the address stays and the content changes, so position-based formulas return the wrong record."
      },
      {
        q: "Which sort order would you choose so that months appear as Jan, Feb, Mar instead of alphabetical order?",
        options: [
          "Sort A to Z, because that is month order",
          "Sort Z to A",
          "Sort by cell colour",
          "Order > Custom List and pick the months list"
        ],
        answer: 3,
        explanation: "Alphabetically, April would come before February. Excel ships with built-in custom lists for months and days, and you can define your own such as High, Medium, Low."
      },
      {
        q: "What is the safest first move if a sort produces a scrambled table?",
        options: [
          "Sort again in the opposite direction",
          "Press Ctrl + Z immediately to undo",
          "Delete the data and retype it",
          "Save the file to lock in the result"
        ],
        answer: 1,
        explanation: "Sorting is fully reversible with Ctrl + Z as long as you have not saved or edited anything else since. Saving a scrambled table makes the damage much harder to undo."
      }
    ],
    summary: [
      "Click a single cell in the sort column and use Data > Sort A to Z or Z to A; whole rows move together.",
      "If the Sort Warning appears, always choose Expand the selection.",
      "Use Data > Sort with multiple levels for questions like \"biggest order within each region\".",
      "Custom lists fix orders that alphabetical sorting gets wrong, such as months or High/Medium/Low.",
      "Sorting moves rows, so position-based formulas like `=F5` can silently return the wrong record."
    ],
    downloads: [
      { file: "assets/files/level3/tables-validation.xlsx",
        label: "Tables, filtering and validation workbook",
        note: "Exercise and Solution sheets for sorting, filtering, Excel Tables, validation and conditional formatting" }
    ],
    keywords: ["sort", "sort a to z", "multi level sort", "sort by color", "custom list", "sort warning", "expand selection", "reorder data"]
  },

  {
    id: "l3-filtering",
    title: "Filtering Data",
    difficulty: "Beginner",
    duration: 12,
    version: "",
    objectives: [
      "Turn filtering on and off with Ctrl + Shift + L",
      "Use text, number and date filters plus the search box",
      "Filter several columns at once and clear filters safely",
      "Use SUBTOTAL so totals follow the filter",
      "Copy only the visible, filtered rows"
    ],
    intro: "A 500-row report may hold only 20 rows you need right now. Filtering hides the rest. You can then focus, print or copy just the rows that matter.\n\nFiltering never deletes anything. The hidden rows stay in the file. One click brings them all back. That safety makes filtering a daily tool in every data job.",
    keyIdea: "Filtering hides the rows you do not need; nothing is deleted and one click brings them back.",
    sections: [
      {
        heading: "Turning the filter on and off",
        blocks: [
          { type: "steps", items: [
            "Click any cell inside your data (one cell is enough).",
            "Press [[Ctrl]] + [[Shift]] + [[L]], or click Data > Filter.",
            "Small dropdown arrows appear on every header cell: Order ID, Date, Salesperson, Region, Product, Revenue.",
            "Click an arrow, tick only the values you want to see, and click OK.",
            "Press [[Ctrl]] + [[Shift]] + [[L]] again to remove the filter completely."
          ] },
          { type: "p", text: "When a filter is on, the row numbers on the left turn blue. Hidden row numbers are skipped, like 1, 2, 9, 10. That is your sign that rows are hidden, not gone." },
          { type: "warning", text: "A filter hides rows; it does not delete them. But some Excel operations touch only visible cells, and others touch everything. Check the blue row numbers before you copy, print or delete. When you finish, remove the filter so nobody thinks data is missing." }
        ]
      },
      {
        heading: "Text, number and date filters",
        blocks: [
          { type: "p", text: "Each filter dropdown offers three things: a value list with tick boxes, a search box, and a context-aware Filters submenu. The submenu changes with the column type:" },
          { type: "table",
            headers: ["Column type", "Useful built-in filters", "Office example"],
            rows: [
              ["Text (Region, Product)", "Equals, Begins With, Contains, Does Not Contain", "Product Contains \"Stapler\" catches Staplers"],
              ["Numbers (Revenue, Units)", "Greater Than, Between, Top 10, Above Average", "Revenue Greater Than 20000 shows big orders"],
              ["Dates (Order Date, Due Date)", "Tomorrow, This Week, This Month, Last Year, Before", "Due Date Before today shows overdue invoices"]
            ] },
          { type: "p", text: "The **Top 10** option is flexible. It can show the top or bottom 5, 10 or 25 items, or the top 10 percent. Date filters also update live. \"This Month\" shows different rows next month, with no change to the filter." },
          { type: "tip", text: "Use the search box in the dropdown when a column has hundreds of values, such as customer names. Type \"priya\" and only matching values stay ticked — far faster than scrolling." }
        ]
      },
      {
        heading: "Filtering several columns at once",
        blocks: [
          { type: "p", text: "Filters on different columns work together. Each filter hides more rows. Two filters show only rows that meet **both** conditions. Filter Region = West, then Product = Laptops. You now see only laptop orders from the West." },
          { type: "grid", title: "A 12-row order list filtered to Region = West and Product = Laptops — only matching rows remain visible",
            rows: [
              ["Order ID", "Salesperson", "Region", "Product", "Revenue"],
              ["1007", "Sneha Iyer", "West", "Laptops", "176000"],
              ["1012", "Priya Nair", "West", "Laptops", "88000"]
            ] },
          { type: "p", text: "A filtered column shows a funnel icon instead of a plain arrow. So you can always see which columns are filtered. To clear one column, open its dropdown and click **Clear Filter**. To clear every filter, use Data > Clear. To remove filtering fully, press [[Ctrl]] + [[Shift]] + [[L]] again." },
          { type: "p", text: "Know the difference. **Clear** keeps the arrows but shows all rows again. **Removing the filter** ([[Ctrl]] + [[Shift]] + [[L]]) takes the arrows away too. \"The filter is still on\" causes many confused reports. Someone cleared one column but forgot the others." }
        ]
      },
      {
        heading: "SUBTOTAL: totals that follow the filter",
        blocks: [
          { type: "p", text: "Every analyst meets this in week one. The total cell shows ₹3,52,000. But the filtered rows add up to only ₹2,64,000. The reason: `=SUM()` ignores filters. It always adds everything, visible or hidden." },
          { type: "formula", name: "SUBTOTAL",
            syntax: "=SUBTOTAL(function_number, range)",
            example: "=SUBTOTAL(9, G2:G200)",
            explanation: "Adds only the visible cells in the range. The first number picks the operation: 9 for SUM, 1 for AVERAGE, 2 for COUNT, 3 for COUNTA. When the filter changes, the result changes with it." },
          { type: "table",
            headers: ["Function number", "Operation", "Example"],
            rows: [
              ["9", "SUM of visible cells", "=SUBTOTAL(9, G2:G200)"],
              ["1", "AVERAGE of visible cells", "=SUBTOTAL(1, G2:G200)"],
              ["2", "COUNT of visible numbers", "=SUBTOTAL(2, F2:F200)"],
              ["3", "COUNTA of visible non-blank cells", "=SUBTOTAL(3, A2:A200)"]
            ] },
          { type: "p", text: "Put a `=SUBTOTAL(9, ...)` above your data. It becomes a live dashboard. Filter to North and it shows North revenue. Filter to March and it shows March revenue. Excel Tables do this automatically in their totals row — one more reason to use Tables." }
        ]
      },
      {
        heading: "Copying only the filtered rows",
        blocks: [
          { type: "p", text: "Here is a very common task: filter to the rows you need, then copy them to a new sheet or an email. Excel behaves well here." },
          { type: "p", text: "When a filter is active, copying takes **only the visible cells**. Hidden rows are skipped. Paste into a blank sheet and you get a clean extract of the filtered rows. Headers come too, if you selected them." },
          { type: "steps", items: [
            "Apply the filter so only the rows you want are visible.",
            "Select the data including the header row (click the first header cell, then [[Ctrl]] + [[Shift]] + [[End]] works well for wide tables).",
            "Press [[Ctrl]] + [[C]] to copy — the marching-ants border skips hidden rows.",
            "Click a cell in the destination sheet and press [[Ctrl]] + [[V]]. The paste contains only visible rows."
          ] },
          { type: "warning", text: "This rule applies to rows hidden by a filter only. If you hid rows manually with right-click > Hide, a normal copy includes them. First press [[Alt]] + [[;]] to select visible cells only." }
        ]
      }
    ],
    mistakes: [
      {
        wrong: "Filtering a list, then deleting the \"extra\" rows to clean up the data.",
        right: "Delete filtered rows only when you truly want them gone — and remember the rows above and below the filter are hidden, not deleted. Use Data > Clear to bring everything back.",
        why: "Deleting while filtered removes only visible rows, which is sometimes exactly the intention, but many beginners delete, remove the filter later and discover they lost records they thought were already gone — or keep records they believed were deleted."
      },
      {
        wrong: "Putting `=SUM(G2:G500)` below a filtered table and expecting it to show the filtered total.",
        right: "Use `=SUBTOTAL(9, G2:G500)` or a Table's totals row.",
        why: "SUM always counts hidden rows. SUBTOTAL number 9 sums visible cells only and updates live as the filter changes."
      },
      {
        wrong: "Building a report from a sheet where three columns still have funnel icons on them.",
        right: "After finishing, press Data > Clear (or Ctrl + Shift + L if the arrows should go completely) and check the row numbers run 1, 2, 3 with no gaps.",
        why: "Leftover filters silently hide data from everyone who opens the file later — printing, pivots and formulas that read the sheet will behave in ways nobody can explain."
      }
    ],
    shortcuts: [
      { keys: "Ctrl + Shift + L", action: "Turn the filter on or off" },
      { keys: "Alt + Down Arrow", action: "Open the filter dropdown of the selected column" },
      { keys: "Alt + ;", action: "Select only visible cells in the selection" },
      { keys: "Ctrl + Shift + End", action: "Extend selection to the last used cell of the table" }
    ],
    exercises: [
      {
        task: "On the Filtering sheet of the exercise workbook, show only the orders from the West region.",
        hint: "Click a cell in the data, press Ctrl + Shift + L, open the Region dropdown and tick only West.",
        solution: "After Ctrl + Shift + L, click the Region dropdown arrow, clear Select All, tick West, click OK. Only West rows stay visible and the row numbers on the left turn blue."
      },
      {
        task: "From the same list, show only rows with Revenue greater than ₹20,000, then use a live total that follows the filter.",
        hint: "Number Filters > Greater Than. For the total, SUBTOTAL with function number 9.",
        solution: "Filter Revenue > Greater Than > 20000. In a free cell type =SUBTOTAL(9, G2:G200) — it shows the sum of visible revenue only. Change the filter to North and watch the total update."
      },
      {
        task: "Show orders due this month from the order date column, then copy just those rows to a new sheet.",
        hint: "Date Filters > This Month. Copying while filtered takes visible rows only.",
        solution: "Filter the Date column with Date Filters > This Month. Select the visible rows including headers, Ctrl + C, then paste into a blank sheet with Ctrl + V. Compare with the Solution sheet."
      },
      {
        task: "The sheet still shows a funnel icon on the Product column from your earlier work. Clear every filter and then remove filtering completely.",
        hint: "Data > Clear resets all columns; Ctrl + Shift + L removes the arrows.",
        solution: "Data tab > Clear shows all rows again but keeps the arrows. Press Ctrl + Shift + L to remove the filter arrows entirely. The Solution sheet describes both states."
      }
    ],
    quiz: [
      {
        q: "What does applying a filter do to the rows that do not match?",
        options: [
          "Deletes them permanently",
          "Hides them until the filter is removed",
          "Moves them to a new sheet",
          "Greys them out but keeps them fully visible"
        ],
        answer: 1,
        explanation: "Filters only hide rows. The data stays in the file and reappears when you clear the filter — which is why filtering is safe to experiment with."
      },
      {
        q: "Which shortcut turns filtering on or off for the selected data?",
        options: ["Ctrl + F", "Ctrl + Shift + L", "Alt + ==", "Ctrl + T"],
        answer: 1,
        explanation: "Ctrl + Shift + L toggles the filter arrows. Ctrl + F opens Find, Ctrl + T creates a Table, and Alt + = inserts AutoSum."
      },
      {
        q: "You filter Region = West and then Product = Laptops. Which rows stay visible?",
        options: [
          "Rows that are West OR Laptops",
          "Rows that are West AND Laptops",
          "Only the first matching row",
          "All rows, because the second filter replaces the first"
        ],
        answer: 1,
        explanation: "Filters on different columns stack with AND logic: a row must satisfy every active column filter to remain visible."
      },
      {
        q: "Your filtered sheet shows ₹2,64,000 of visible revenue, but a `=SUM()` cell shows ₹3,52,000. Why?",
        options: [
          "SUM has a bug with filters",
          "SUM includes the hidden rows; use SUBTOTAL(9, ...) for visible rows only",
          "The filter is not really applied",
          "The range is too small"
        ],
        answer: 1,
        explanation: "SUM always adds every cell in the range, hidden or not. SUBTOTAL with function number 9 respects the filter and recalculates as the filter changes."
      },
      {
        q: "After copying a filtered selection and pasting it elsewhere, what does the paste contain?",
        options: [
          "All rows, hidden ones included",
          "Only the visible, filtered rows",
          "Only the header row",
          "A link that updates when the filter changes"
        ],
        answer: 1,
        explanation: "When a filter is active, copy and paste operate on visible cells only, which makes filtered extracts a fast and reliable way to split data into new sheets."
      }
    ],
    summary: [
      "Ctrl + Shift + L toggles filter arrows on any data range.",
      "Each dropdown offers tick boxes, a search box and text, number or date filters such as Greater Than, Top 10 and This Month.",
      "Filters on multiple columns combine with AND logic; funnels show which columns are filtered.",
      "`=SUBTOTAL(9, range)` totals only visible rows, so your totals follow the filter.",
      "Copying a filtered range pastes only the visible rows — a fast way to build extracts."
    ],
    downloads: [
      { file: "assets/files/level3/tables-validation.xlsx",
        label: "Tables, filtering and validation workbook",
        note: "Exercise and Solution sheets for sorting, filtering, Excel Tables, validation and conditional formatting" }
    ],
    keywords: ["filter", "ctrl shift l", "text filters", "number filters", "date filters", "top 10", "subtotal", "copy visible cells", "clear filter"]
  },

  {
    id: "l3-excel-tables",
    title: "Excel Tables",
    difficulty: "Beginner",
    duration: 14,
    version: "",
    objectives: [
      "Create an Excel Table with Ctrl + T",
      "Explain why Tables beat plain ranges: auto-expanding, auto-filling formulas",
      "Read and write structured references like =SUM(Sales[Revenue])",
      "Rename a Table, use the totals row and built-in filtering",
      "Convert a Table back to a normal range"
    ],
    intro: "So far you have used plain ranges like `A1:G200`. You copied formulas by hand and froze headers yourself. An Excel Table does this work for you. It grows with your data and fills formulas down automatically.\n\nTables are the one habit that separates fast users from frustrated ones. Start every new report with a Table.",
    keyIdea: "A Table (Ctrl + T) is a smart range that grows by itself and fills formulas down for you.",
    sections: [
      {
        heading: "Creating a Table",
        blocks: [
          { type: "steps", items: [
            "Click any cell inside your data — for example a cell in the Sunrise Traders order list.",
            "Press [[Ctrl]] + [[T]] (or Insert > Table). Excel guesses the surrounding range and shows a small dialog.",
            "Make sure **My table has headers** is ticked (it reads Order ID, Date, Salesperson...). Click OK.",
            "The range gains banded blue rows, filter arrows on every header and a new Table Design tab appears on the ribbon."
          ] },
          { type: "grid", title: "The same order list before and after Ctrl + T — filter arrows appear and rows alternate in colour",
            rows: [
              ["Order ID", "Salesperson", "Region", "Product", "Revenue"],
              ["1001", "Arjun Mehta", "North", "Notebooks", "14400"],
              ["1002", "Priya Nair", "South", "Staplers", "6750"],
              ["1003", "Rahul Verma", "North", "Laptops", "352000"]
            ] },
          { type: "warning", text: "Tables and merged cells do not mix. Unmerge first: select the range, then press Alt + H + M + U. Merged cells break sorting, filtering and Tables." }
        ]
      },
      {
        heading: "Why Tables are essential",
        blocks: [
          { type: "p", text: "A Table fixes the three most repetitive chores of spreadsheet work:" },
          { type: "ul", items: [
            "**Ranges grow automatically.** Type a new order in the row below the Table and the Table swallows it — its filters, its formatting and every formula that references the Table include the new row. No more editing `SUM(G2:G200)` to `SUM(G2:G201)` every week.",
            "**Formulas fill themselves.** Type a formula once in any column of a Table and it copies down the whole column instantly (a \"calculated column\"). Add a new row and the formula is already there waiting for you.",
            "**Headers stay visible and usable.** Filter arrows appear on every column, and column headers are always treated as proper headers, not as data."
          ] },
          { type: "p", text: "In practice, a weekly update becomes simple. Paste new rows at the bottom of the Table. Totals, charts, validation and formatting already include them. That is why every good template is built on Tables." }
        ]
      },
      {
        heading: "Structured references: formulas that read like English",
        blocks: [
          { type: "p", text: "When you reference a Table in a formula, Excel offers its columns by name instead of by letter. These are structured references. If the Table is named Sales, the revenue total is:" },
          { type: "formula", name: "SUM with structured references",
            syntax: "=SUM(TableName[ColumnName])",
            example: "=SUM(Sales[Revenue])",
            explanation: "Adds the entire Revenue column of the Sales table. The range never needs editing, because the Table knows how big it is." },
          { type: "p", text: "Structured references work anywhere a range would — including the functions you learned in Level 2:" },
          { type: "ul", items: [
            "`=SUMIF(Sales[Region], \"North\", Sales[Revenue])` — total revenue for the North region.",
            "`=COUNTIFS(Sales[Product], \"Laptops\", Sales[Units], \">10\")` — count laptop orders of more than 10 units.",
            "`=Sales[@Revenue]` — the Revenue value on the current row (the `@` means \"this row\"), perfect inside a calculated column such as `=[@Units]*120`."
          ] },
          { type: "p", text: "Compare `=SUM(Sales[Revenue])` with `=SUM(Sheet1!G2:G200)`. The first tells every reader exactly what it totals. It never breaks when rows are added. The second is a mystery you must maintain by hand. This also fixes the sorting risks from the last lesson: formulas built on Tables sort and move correctly." }
        ]
      },
      {
        heading: "Naming your Table and using the totals row",
        blocks: [
          { type: "p", text: "Excel names new tables Table1, Table2, Table3 — useless in formulas. Rename immediately:" },
          { type: "steps", items: [
            "Click any cell in the Table so the Table Design tab appears.",
            "On the far left of Table Design, type a name in the Table Name box, for example Sales.",
            "Press [[Enter]]. Every formula that used Table1 now reads Sales automatically."
          ] },
          { type: "warning", text: "Table names must start with a letter, contain no spaces (use SalesData or Sales_Data), and must be unique in the workbook. If the name box looks greyed out, you are in edit mode or the cell is not inside a Table." },
          { type: "p", text: "The **totals row** is a switch, not a formula you type. Tick Total Row on the Table Design tab and a row appears under the Table." },
          { type: "p", text: "Click any cell in that row and pick from a dropdown: Sum, Average, Count. Excel writes a SUBTOTAL formula for you. So the totals row follows filters automatically, just like the filtering lesson showed." },
          { type: "p", text: "Table Design has more. The **banded rows** tick box makes the alternating shading. The gallery lets you pick a style that matches your company look. The shading is automatic logic, not manual fills, so it keeps working as rows are added." }
        ]
      },
      {
        heading: "Filtering comes free, and leaving Tables behind",
        blocks: [
          { type: "p", text: "Every Table comes with filter arrows already on. Everything from the filtering lesson works right away. Tables add one bonus: their dropdowns also offer **Sort by Color** and text shortcuts, and the Table remembers filters per column." },
          { type: "p", text: "Sometimes you need a plain range again — for example, for an old system or a feature that misbehaves inside Tables. Click in the Table, then Table Design > Convert to Range, and confirm." },
          { type: "p", text: "The formatting logic and structured references become normal formatting and normal cell addresses. `=SUM(Sales[Revenue])` is rewritten to a range like `=SUM($G$2:$G$200)`, so nothing shows an error. But you lose the auto-expansion. Convert back only with a good reason." },
          { type: "tip", text: "When a chart, pivot table or validation list points at a Table, it follows the Table as it grows. Point them at `Sales` instead of `G2:G200` and your reports maintain themselves." }
        ]
      }
    ],
    mistakes: [
      {
        wrong: "Leaving every table named Table1, Table2, Query3 and writing `=SUM(Table2[Column1])`.",
        right: "Rename each Table on the Table Design tab with a meaningful name like Sales, Customers or Orders.",
        why: "Meaningful names make formulas readable months later and prevent mistakes when a workbook contains many tables. Renaming also updates existing formulas automatically."
      },
      {
        wrong: "Expecting the Table totals row to show the grand total of every row while a filter is active.",
        right: "Understand that the totals row uses SUBTOTAL by design: it shows the total of visible rows. Clear the filter to see the grand total.",
        why: "This behaviour is intentional — it turns the totals row into a live dashboard — but it surprises people who expected a fixed SUM."
      },
      {
        wrong: "Typing `=SUM(Sales[Revenue])` by hand and getting an error because the table is actually named Sales_Data.",
        right: "Type `=SUM(` and then select the column with the mouse — Excel writes the structured reference for you.",
        why: "Structured references are precise: names, brackets and the @ symbol matter. Selecting with the mouse guarantees correct syntax."
      }
    ],
    shortcuts: [
      { keys: "Ctrl + T", action: "Create a Table from the selected cell's data" },
      { keys: "Ctrl + Shift + L", action: "Toggle the filter arrows of a Table" },
      { keys: "Alt + J + T + A + R", action: "Rename a Table (Table Design tab, Table Name)" },
      { keys: "Tab", action: "Inside the last cell of a Table row: adds a new row instantly" }
    ],
    exercises: [
      {
        task: "On the Tables sheet of the exercise workbook, convert the plain order list into a Table and rename it Sales.",
        hint: "Ctrl + T, tick My table has headers, then Table Design > Table Name.",
        solution: "Click inside the list, press Ctrl + T, confirm the range and header tick, click OK. On Table Design, type Sales in the Table Name box and press Enter. The Solution sheet shows the finished Table."
      },
      {
        task: "Add a Revenue-per-Unit calculated column to the Sales Table without copying any formula down.",
        hint: "If Revenue is in the table, type a structured formula in the first empty column cell and let the Table fill it.",
        solution: "In the first cell of the empty column type =Sales[@Revenue]/Sales[@Units] (or =[@Revenue]/[@Units]) and press Enter. The whole column fills at once. Add a new row and the formula is already waiting there."
      },
      {
        task: "Put a totals row on the Sales Table and switch on the Sum for Revenue.",
        hint: "Table Design > Total Row, then pick Sum in the dropdown of the Revenue cell.",
        solution: "Tick Total Row on the Table Design tab. In the Revenue cell of the totals row choose Sum from the dropdown. Filter the Table to one region and the total follows the filter, because the totals row uses SUBTOTAL."
      },
      {
        task: "Write a formula outside the Table that shows total revenue for the South region, then add a new order row to the bottom of the Table and confirm the total includes it.",
        hint: "SUMIF with structured references; new rows typed directly under the Table join it automatically.",
        solution: "=SUMIF(Sales[Region], \"South\", Sales[Revenue]). Type a new South order in the row below the Table — the Table expands, the formula result updates with no range editing. Check the Solution sheet."
      }
    ],
    quiz: [
      {
        q: "Which shortcut creates an Excel Table around the selected cell?",
        options: ["Ctrl + L only", "Ctrl + T", "Ctrl + R", "Alt + T"],
        answer: 1,
        explanation: "Ctrl + T opens the Create Table dialog (Insert > Table does the same). Tick \"My table has headers\" when your first row contains column names."
      },
      {
        q: "What does the formula =SUM(Sales[Revenue]) do?",
        options: [
          "Adds cells named Sales and Revenue",
          "Adds the entire Revenue column of the table named Sales, including rows added later",
          "Adds only the first row of the Sales table",
          "It is invalid — Tables cannot be used in SUM"
        ],
        answer: 1,
        explanation: "This is a structured reference: the table name followed by a column name in brackets. It stays correct as the Table grows, so you never edit the range again."
      },
      {
        q: "You type a formula in the first cell of an empty column inside a Table and press Enter. What happens?",
        options: [
          "Nothing — you must copy it down yourself",
          "It fills down the entire column automatically as a calculated column",
          "Excel converts the Table back to a range",
          "The formula only applies to that single cell forever"
        ],
        answer: 1,
        explanation: "Calculated columns are a core Table feature: one formula fills the whole column instantly and applies to every future row as well."
      },
      {
        q: "Why does the Table totals row show a different number when you filter the Table?",
        options: [
          "It is broken — convert the Table to a range",
          "It uses SUBTOTAL, so it totals only the visible rows",
          "It ignores numbers above a certain size",
          "Filtering deletes the rows it was summing"
        ],
        answer: 1,
        explanation: "The totals row writes a SUBTOTAL formula, which respects filtering. That makes it a live total of what you can currently see."
      },
      {
        q: "What happens to =SUM(Sales[Revenue]) after you rename the Table from Sales to Orders?",
        options: [
          "It shows a #NAME? error",
          "It keeps the old name and stops working",
          "Excel updates it automatically to =SUM(Orders[Revenue])",
          "It changes to a normal cell reference"
        ],
        answer: 2,
        explanation: "Structured references are live links to the Table object. Renaming a Table updates every formula that points at it — one more reason Tables beat typed ranges."
      }
    ],
    summary: [
      "Ctrl + T turns a range into a Table: banded rows, filter arrows and a Table Design tab.",
      "Tables auto-expand, so new rows are included in every formula, chart and list that points at the Table.",
      "One formula typed in a Table column fills the whole column as a calculated column.",
      "Structured references such as `=SUM(Sales[Revenue])` are readable and never need range maintenance.",
      "The totals row uses SUBTOTAL and follows filters; Convert to Range returns you to a plain range when needed."
    ],
    downloads: [
      { file: "assets/files/level3/tables-validation.xlsx",
        label: "Tables, filtering and validation workbook",
        note: "Exercise and Solution sheets for sorting, filtering, Excel Tables, validation and conditional formatting" }
    ],
    keywords: ["excel table", "ctrl t", "structured references", "calculated column", "totals row", "table name", "banded rows", "convert to range"]
  },

  {
    id: "l3-data-validation",
    title: "Data Validation and Dropdown Lists",
    difficulty: "Beginner",
    duration: 13,
    version: "",
    objectives: [
      "Apply whole number, decimal, date and text-length rules",
      "Create dropdown lists from typed values or from a range",
      "Add input messages and choose the right error alert style",
      "Copy validation to other cells and find invalid data already in the sheet",
      "Explain how dropdowns prevent report errors"
    ],
    intro: "In Level 2 you built reports with SUMIF and VLOOKUP. They have a hidden weakness. If someone types \"north\" instead of North, the report quietly misses that row. No error appears — the numbers are just wrong.\n\nData validation stops the problem at the source. It controls what a person can type: a dropdown of approved values, a number within limits, a date in range. Clean data in, trustworthy reports out.",
    keyIdea: "Validation controls what people can type into a cell, so bad data never enters your sheet.",
    sections: [
      {
        heading: "What data validation is for",
        blocks: [
          { type: "p", text: "You will find it on the Data tab > Data Validation. A validation rule is attached to a cell. Anything typed against the rule is rejected or questioned right away." },
          { type: "p", text: "Two limits to know. It does not check data already in the sheet (you will fix that later in this lesson). And it does not survive a paste over the cell. For every value typed by hand, it works." },
          { type: "p", text: "Think of the Sunrise Traders order log. Four people type orders into one sheet. Validation makes sure Region is always North, South, East or West. Units is always a positive whole number. Order Date is always a real date. Now your SUMIF report adds up correctly, every single day." }
        ]
      },
      {
        heading: "Number, date and text-length rules",
        blocks: [
          { type: "steps", items: [
            "Select the cells the rule should apply to, for example `F2:F500` (Units).",
            "Data > Data Validation. In Settings, open Allow and pick a rule type.",
            "Set the limits — for example Whole number, between 1 and 10000.",
            "Optional: fill the Input Message and Error Alert tabs (next section), then click OK."
          ] },
          { type: "table",
            headers: ["Allow type", "What it controls", "Sunrise Traders example"],
            rows: [
              ["Whole number", "Integers within limits", "Units between 1 and 10000"],
              ["Decimal", "Numbers with decimals", "Discount between 0 and 0.25"],
              ["Date", "Valid dates within a window", "Order Date between 01-Jan-2026 and 31-Dec-2026"],
              ["Text length", "Number of characters", "Product code exactly 8 characters"]
            ] },
          { type: "p", text: "Two small but powerful options. Allow > **List** builds dropdowns (next section). And the Settings tab has an **Ignore blank** tick box. Usually leave it ticked, so empty cells are allowed while the column is being filled." },
          { type: "tip", text: "Text length validation is a simple guard for codes, such as an 8-character SKU. It cannot stop typos that are also 8 characters. But it stops half a code, or a whole sentence, landing in the column." }
        ]
      },
      {
        heading: "Dropdown lists: typed values or a range",
        blocks: [
          { type: "p", text: "The dropdown is the most used validation type. Select the cells, Data > Data Validation, Allow: List, then choose one of two sources:" },
          { type: "ul", items: [
            "**Typed list.** In Source, type the values separated by commas: `North,South,East,West`. Quick, self-contained, no helper cells. The limit is 255 characters total, and if a value itself contains a comma this style breaks — use a range instead.",
            "**Range list.** In Source, point at a range of cells, for example `=$L$2:$L$5` where those cells contain North, South, East, West. The dropdown follows the cells: fix a typo or add a fifth region in `L6` (after extending the range) and every dropdown updates. Anchor the range with `$` so copying does not shift it."
          ] },
          { type: "grid", title: "A dropdown applied to D2:D500 — users click and pick instead of typing",
            rows: [
              ["Order ID", "Region", "Product", "Units"],
              ["1001", "[North v]", "Notebooks", "120"],
              ["1002", "[South v]", "Staplers", "45"]
            ] },
          { type: "p", text: "Build this rule into every template. Keep the list values on a hidden or dedicated sheet called Lists. Name the range, for example RegionList, and use `=RegionList` as the source. Named lists survive copying and sorting far better than raw addresses." },
          { type: "warning", text: "In older Excel versions, a list source cannot point directly at another sheet. The named-range approach is the fix, and it works in every version. Modern Excel accepts cross-sheet ranges, but a named range is still the tidier habit." }
        ]
      },
      {
        heading: "Input messages and error alerts",
        blocks: [
          { type: "p", text: "Validation has two helper tabs that turn a silent rule into a guided one." },
          { type: "p", text: "**Input Message** shows a small yellow note when the cell is selected. For example: \"Choose one of the four regions\". It is a caption for the person filling the sheet. Keep it short — it appears every time the cell is selected." },
          { type: "p", text: "**Error Alert** decides what happens when someone breaks the rule:" },
          { type: "table",
            headers: ["Style", "Behaviour", "When to use"],
            rows: [
              ["Stop (default)", "Rejects the entry; the user must retry or cancel", "Almost always — regions, dates, quantities"],
              ["Warning", "Asks \"Continue?\" and allows the entry anyway", "Rare exceptions that a supervisor can approve"],
              ["Information", "Explains and accepts the entry", "Gentle guidance, not enforcement"]
            ] },
          { type: "p", text: "Write a useful error message. \"Units must be a whole number between 1 and 10000\" teaches. The default message does not." },
          { type: "p", text: "If you untick Show error alert after validation entirely, the rule still exists. But any typing is accepted — you effectively keep only the dropdown arrow." }
        ]
      },
      {
        heading: "Copying validation and catching bad data",
        blocks: [
          { type: "p", text: "Three skills that make validation maintainable in a real file:" },
          { type: "ul", items: [
            "**Copy it.** Copy a validated cell and Paste Special > Validation (Home > Paste dropdown > Paste Special > Validation) onto the target range — formats and values stay untouched, only the rule moves. Extending a Table downward copies validation automatically.",
            "**Audit existing data.** Data > Data Validation dropdown > Circle Invalid Data draws red ovals around every existing value that breaks the selected cells' rules. Fix them (often with the cleaning tools later in this level), then choose Clear Validation Circles.",
            "**Remove or widen it.** Select cells, Data Validation, and click Clear All to remove rules; or select the whole sheet ([[Ctrl]] + [[A]]) before opening the dialog to review every rule in one place."
          ] },
          { type: "warning", text: "Validation does not survive a normal paste. Someone pastes rows from another file over your columns. The pasted values arrive with no rules. They can reintroduce \"north\" and 0-unit orders. Run Circle Invalid Data after every big paste." }
        ]
      },
      {
        heading: "Why dropdowns prevent report errors",
        blocks: [
          { type: "p", text: "Connect this to Level 2. `=SUMIF(D2:D500, \"North\", G2:G500)` misses rows where the region is \"north\", \"N\", \"North \" or \"Nrth\". Each typo survives silently. The row looks normal. The report is short by a few thousand rupees. Nobody notices until the numbers do not match finance." },
          { type: "p", text: "With a dropdown, all four typists pick from the same four values. The criterion \"North\" then matches every North row. The same logic protects VLOOKUP. A lookup works only if names match the lookup table exactly. Validation is cheaper than cleaning: five minutes of setup saves hours of monthly repair." },
          { type: "quote", text: "Reports are only as good as the data entry behind them. Validation is the cheapest data quality tool Excel has." }
        ]
      }
    ],
    mistakes: [
      {
        wrong: "Typing a comma-separated list like `Staplers, red, small` when a value itself should contain a comma.",
        right: "Put the values in cells and point Source at the range (or a named range).",
        why: "In a typed list, commas separate values, so \"Staplers, red\" becomes two items: \"Staplers\" and \" red\" (with a space). Ranges avoid the problem completely."
      },
      {
        wrong: "Assuming validation protects the column after someone pastes data over it.",
        right: "After large pastes, run Data > Data Validation > Circle Invalid Data to find values that broke in.",
        why: "A normal paste replaces values and drops the rules. Validation governs typing, not pasting — auditing afterwards is part of the workflow."
      },
      {
        wrong: "Setting the error alert style to Warning for a critical field like Region.",
        right: "Use Stop for fields that feed reports; reserve Warning for genuine exceptions.",
        why: "A Warning dialog with a yes-click habit is how \"north\" gets back into the data within a week. Stop keeps report criteria reliable."
      }
    ],
    shortcuts: [
      { keys: "Alt + A + V + V", action: "Open the Data Validation dialog" },
      { keys: "Alt + Down Arrow", action: "Open the dropdown of a selected validated cell" },
      { keys: "Ctrl + Alt + V", action: "Paste Special (then choose Validation)" }
    ],
    exercises: [
      {
        task: "On the Validation sheet of the exercise workbook, add a dropdown to the Region column (D2:D500) offering North, South, East and West.",
        hint: "Allow: List. Try the typed version first, then the range version.",
        solution: "Select D2:D500, Data > Data Validation, Allow = List, Source = North,South,East,West, OK. Then improve it: put the four values in cells and use them as the source with $ anchors. Check the Solution sheet."
      },
      {
        task: "Restrict the Units column to whole numbers between 1 and 10000, with a helpful error message.",
        hint: "Allow: Whole number, between. Write your own Stop message.",
        solution: "Select F2:F500, Data Validation, Allow = Whole number, Data = between, Minimum 1, Maximum 10000. On Error Alert choose Stop and type \"Units must be a whole number between 1 and 10000\". Test by typing 0 — Excel refuses it."
      },
      {
        task: "Add an input message to the Region cells so users see guidance before they type.",
        hint: "The Input Message tab of the Data Validation dialog.",
        solution: "With the Region cells selected, Data Validation > Input Message tab: title \"Region\", message \"Choose one of the four regions from the dropdown\". Selecting any Region cell now shows the note."
      },
      {
        task: "The sheet contains some values that were typed before validation existed. Find and list them.",
        hint: "Data > Data Validation dropdown > Circle Invalid Data.",
        solution: "Select the validated ranges, then Data Validation dropdown > Circle Invalid Data. Red ovals mark the invalid values (for example a region typed as \"north\"). Fix them to match the dropdown values, then Clear Validation Circles. The Solution sheet lists what was wrong."
      }
    ],
    quiz: [
      {
        q: "Which validation type creates a dropdown of fixed choices such as the four regions?",
        options: ["Whole number", "List", "Text length", "Date"],
        answer: 1,
        explanation: "Allow = List produces the dropdown. The source is either a typed comma-separated string or — better — a range or named range of cells."
      },
      {
        q: "What is the main advantage of sourcing a dropdown from a range instead of typed values?",
        options: [
          "Range sources load faster",
          "Updating the cells updates every dropdown, and values may contain commas",
          "Typed lists cannot be used in tables",
          "Range sources work without selecting the cells first"
        ],
        answer: 1,
        explanation: "A range source stays linked to the cells: fix or extend the list once and all dropdowns follow. Typed lists are frozen text with a 255-character limit and comma pitfalls."
      },
      {
        q: "Which error alert style should protect a Region column that feeds a SUMIF report?",
        options: ["Stop", "Warning", "Information", "No alert"],
        answer: 0,
        explanation: "Stop rejects invalid entries outright. Warning and Information let bad values through with just a click, which quietly breaks report criteria."
      },
      {
        q: "How do you find values that broke the rules before validation was applied?",
        options: [
          "Sort the column and look by eye",
          "Data > Data Validation > Circle Invalid Data",
          "COUNTIF with a wildcard",
          "It is impossible in Excel"
        ],
        answer: 1,
        explanation: "Circle Invalid Data draws red ovals on existing values that violate the rules of the selected cells, so you can fix historical data after adding validation."
      },
      {
        q: "Why do dropdown lists make SUMIF reports more reliable?",
        options: [
          "SUMIF runs faster on short values",
          "Everyone enters identical values, so the report criteria match every row",
          "Dropdowns prevent blank cells entirely",
          "SUMIF only works on validated columns"
        ],
        answer: 1,
        explanation: "SUMIF needs exact matches. A dropdown guarantees the stored values are exactly the criteria values, so no rows silently drop out of the total."
      }
    ],
    summary: [
      "Data Validation (Data tab) controls what can be typed: numbers, dates, text length and lists.",
      "Dropdowns come from typed values (fast, 255 characters) or from a range (recommended, updatable).",
      "Input messages guide users; Stop alerts enforce rules — use Stop for report-critical columns.",
      "Copy rules with Paste Special > Validation and audit old data with Circle Invalid Data.",
      "Validation protects SUMIF and VLOOKUP from silent mismatches — clean data in, correct reports out."
    ],
    downloads: [
      { file: "assets/files/level3/tables-validation.xlsx",
        label: "Tables, filtering and validation workbook",
        note: "Exercise and Solution sheets for sorting, filtering, Excel Tables, validation and conditional formatting" }
    ],
    keywords: ["data validation", "dropdown list", "input message", "error alert", "stop warning", "circle invalid data", "paste special validation", "named range list"]
  },

  {
    id: "l3-conditional-formatting",
    title: "Conditional Formatting",
    difficulty: "Beginner",
    duration: 15,
    version: "",
    objectives: [
      "Apply highlight rules for numbers, text and dates",
      "Use Top/Bottom, Data Bars, Color Scales and Icon Sets",
      "Manage overlapping rules with priority and Stop If True",
      "Write a formula rule such as =$F2=\"Overdue\" to highlight whole rows"
    ],
    intro: "Conditional formatting is colour that thinks. You do not scan 800 rows for big orders or overdue invoices. Excel colours the cells that match a rule. When the data changes, the colours update themselves.\n\nIt is the visual layer on top of this whole level. Sort the data and the colours follow. Filter it and only visible rows stay coloured. New rows in a Table get coloured too, with zero effort.",
    keyIdea: "Conditional formatting colours cells by rule, and the colours update themselves when the data changes.",
    sections: [
      {
        heading: "How conditional formatting differs from manual fills",
        blocks: [
          { type: "p", text: "A manual fill (the paint-bucket colour) is just paint. It stays even when the value no longer deserves it. You must repaint it by hand when data changes." },
          { type: "p", text: "Conditional formatting is a rule attached to a range. Every time Excel recalculates, it re-checks the rule. The format applies only where the rule is true right now." },
          { type: "p", text: "That difference matters in daily work. Yesterday's red rows turn back to normal automatically once they are fixed. Nobody has to remember to un-highlight anything. The controls live on Home > Conditional Formatting." }
        ]
      },
      {
        heading: "Highlight Cells rules",
        blocks: [
          { type: "steps", items: [
            "Select the range to colour, for example the Revenue column `G2:G500`.",
            "Home > Conditional Formatting > Highlight Cells Rules > Greater Than.",
            "Type the value (20000) and pick a format — the default light red fill with dark red text is a good start.",
            "Click OK. Every cell above 20000 is coloured; change a value and the colour follows instantly."
          ] },
          { type: "table",
            headers: ["Rule", "What it catches", "Sunrise Traders example"],
            rows: [
              ["Greater Than / Less Than", "Numeric thresholds", "Revenue > 20000 = big orders"],
              ["Text that Contains", "Partial text matches", "Product contains \"Stapler\""],
              ["A Date Occurring", "Dates relative to today", "Due date = Yesterday shows overdue payments"],
              ["Duplicate Values", "Repeats in the selected range", "Order ID entered twice (used again in the next part of this level)"]
            ] },
          { type: "grid", title: "Revenue column with a Greater Than 20000 rule applied",
            rows: [
              ["Order ID", "Product", "Revenue"],
              ["1003", "Laptops", "352000"],
              ["1001", "Notebooks", "14400"],
              ["1007", "Laptops", "176000"],
              ["1015", "Printers", "26400"]
            ] },
          { type: "p", text: "In the rendered sheet, rows 2 and 5 get the highlight fill — the ₹3,52,000 laptop order and the ₹26,400 printer order. The notebook and the ₹1,76,000 order stay plain. Change a value and the highlight moves by itself." },
          { type: "tip", text: "Highlight Cells Rules and Top/Bottom rules are the everyday 90 percent. Formula rules (last section) cover the special cases — start simple and only go formula-based when a preset cannot express the condition." }
        ]
      },
      {
        heading: "Top/Bottom rules",
        blocks: [
          { type: "p", text: "Home > Conditional Formatting > Top/Bottom Rules answers ranking questions without sorting:" },
          { type: "ul", items: [
            "**Top 10 Items** — highlight the 10 biggest revenues in the column (the dialog lets you change 10 to any number).",
            "**Top 10% / Bottom 10%** — the biggest or smallest tenth of the data, useful for bonuses or risk lists.",
            "**Above Average** — every order better than the mean, which adapts as the data grows.",
            "**Below Average** — underperformers; combine with a red text format for a quick watch list."
          ] },
          { type: "p", text: "A sales manager gets the view in ten seconds. Select the Revenue column. Apply Top 10 Items with a green fill. The star orders glow, and the sort order is untouched for reading." }
        ]
      },
      {
        heading: "Data Bars, Color Scales and Icon Sets",
        blocks: [
          { type: "p", text: "Three rule families add visuals instead of plain fills:" },
          { type: "ul", items: [
            "**Data Bars** — a bar inside each cell whose length matches the value; an instant in-cell bar chart. Perfect for comparing Revenue or Units down a column while keeping the numbers visible.",
            "**Color Scales** — a gradient across the range, for example red for low through yellow to green for high. A whole column becomes a heat map; great for month-by-month performance grids.",
            "**Icon Sets** — arrows, traffic lights, flags or rating stars next to each value based on thresholds. For example, a three-arrow set on Units shows up-arrows for high quantities and down-arrows for low ones."
          ] },
          { type: "p", text: "All three live in the same Conditional Formatting menu. Each dialog allows threshold tweaks. For example, edit an icon set rule so the green arrow appears only above a target value, not the default percentile split." },
          { type: "warning", text: "Restraint is a professional skill. A sheet with data bars, two colour scales, five icon sets and four fills is harder to read than no formatting at all. Use one or two rules per sheet. Keep meanings consistent — red means problem everywhere in the workbook." }
        ]
      },
      {
        heading: "Managing rules: priority and Stop If True",
        blocks: [
          { type: "p", text: "Real sheets gather many rules. Open Home > Conditional Formatting > Manage Rules. It shows every rule, in priority order from top to bottom. When two rules match one cell, the rule higher in the list wins." },
          { type: "ul", items: [
            "**Reorder with the arrows** — move the rule that should win to the top.",
            "**Edit a rule** to change its condition, format or the range it applies to (the Applies To column).",
            "**Stop If True** — when a higher rule matches and has Stop If True ticked, lower rules are not even evaluated for that cell. Use it, for example, so that \"red = overdue\" rows are never additionally painted by a data bar rule.",
            "**Delete** rules you no longer need; leftover rules on wide ranges slow big files down."
          ] },
          { type: "p", text: "Here is a classic cleanup. Someone selected whole columns like `A:Z` when creating a rule. The Applies To range is then millions of rows. Edit the rule and set Applies To to the real data range. The file gets faster immediately." }
        ]
      },
      {
        heading: "Formula-based rules: highlight the whole row",
        blocks: [
          { type: "p", text: "Presets colour one cell at a time. A formula rule colours anything, based on any logic — the most common use is highlighting an entire row when one column in that row matters." },
          { type: "steps", items: [
            "Select the whole data range **starting at the first data row**, for example `A2:G500`.",
            "Home > Conditional Formatting > New Rule > Use a formula to determine which cells to format.",
            "Enter the formula with the first row of the selection: `=$F2=\"Overdue\"`.",
            "Set a format (light red fill) and click OK."
          ] },
          { type: "p", text: "Why the dollar sign on the column only? The rule runs for every cell in the selection. For cell `C7` the formula becomes `$F7=\"Overdue\"`. The **column is locked** to F, so every cell in a row reads that row's own Status." },
          { type: "p", text: "The row reference stays relative, so the rule slides down correctly. With `$F$2` (both locked), every cell would look only at `F2`. The whole sheet would then light up or not as one block." },
          { type: "formula", name: "Formula rule with AND",
            syntax: "=AND(condition1, condition2)",
            example: "=AND($E2<TODAY(), $F2=\"Unpaid\")",
            explanation: "Highlights rows where the due date (column E) is in the past AND the status (column F) is still Unpaid — the invoices that actually need a phone call today." },
          { type: "grid", title: "Payment tracker with the rule =$F2=\"Overdue\" applied to A2:F6",
            rows: [
              ["Invoice", "Customer", "Amount", "Due Date", "Status"],
              ["INV-2041", "Arjun Mehta", "48500", "05-Feb-2026", "Paid"],
              ["INV-2042", "Priya Nair", "12750", "12-Feb-2026", "Overdue"],
              ["INV-2043", "Rahul Verma", "352000", "20-Feb-2026", "Paid"],
              ["INV-2044", "Sneha Iyer", "8800", "28-Feb-2026", "Overdue"]
            ] },
          { type: "p", text: "Rows 3 and 5 (Priya Nair and Sneha Iyer) colour across the whole row. Their Status column says Overdue. When Sneha pays, the status changes to Paid and her row decolours itself. No repainting, ever." },
          { type: "warning", text: "Formula rules are the easiest place to get the anchor wrong. Check three cells before you trust the rule. The first data row should colour correctly. A row further down should follow its own values. A middle cell like `C10` must look at `F10`, not `F2`." }
        ]
      }
    ],
    mistakes: [
      {
        wrong: "Writing a row-highlight rule as `=$F$2=\"Overdue\"` with both references locked.",
        right: "Lock only the column: `=$F2=\"Overdue\"`, keeping the row relative.",
        why: "With `$F$2`, every cell in the selection evaluates the same single cell F2, so the entire sheet is coloured or not as one block instead of row by row."
      },
      {
        wrong: "Applying the formula rule to `A2:G500` but writing the formula for the wrong first row, like `=$F5=\"Overdue\"`.",
        right: "The formula must describe the first row of the Applies To range — if the range starts at row 2, the formula references row 2.",
        why: "Excel applies the formula relative to the top-left cell of the range. A row offset shifts every highlight down, colouring rows that do not deserve it."
      },
      {
        wrong: "Building a dashboard where one cell matches five overlapping rules with different fills.",
        right: "Open Manage Rules, put the important rule on top, tick Stop If True on it, and delete rules that are no longer needed.",
        why: "Without a clear priority, the visible format is the topmost rule's — often not the one the designer intended — and wide ranges with many rules make large files slow."
      }
    ],
    shortcuts: [
      { keys: "Alt + H + L", action: "Open the Conditional Formatting menu" },
      { keys: "Alt + H + L + C + S", action: "Open Clear Rules > Selected Cells (cleanup)" },
      { keys: "Alt + H + L + R", action: "Open Manage Rules for the selection" }
    ],
    exercises: [
      {
        task: "On the Conditional Formatting sheet of the exercise workbook, highlight every Revenue above ₹20,000 in the order list.",
        hint: "Highlight Cells Rules > Greater Than on the Revenue column.",
        solution: "Select the Revenue cells, Home > Conditional Formatting > Highlight Cells Rules > Greater Than, enter 20000, keep the light red fill, OK. Change one value above and below the threshold to see the colours update live."
      },
      {
        task: "Mark the top 5 orders with a green fill without sorting the list.",
        hint: "Top/Bottom Rules > Top 10 Items — change 10 to 5 in the dialog.",
        solution: "Select the Revenue column, Conditional Formatting > Top/Bottom Rules > Top 10 Items, change the 10 to 5 in the dialog, choose a green fill, OK. The sort order of the sheet stays untouched."
      },
      {
        task: "Add data bars to the Revenue column so the biggest orders are visible at a glance.",
        hint: "Conditional Formatting > Data Bars — pick a solid single colour.",
        solution: "Select the Revenue cells, Conditional Formatting > Data Bars > choose a blue gradient bar. Every cell shows a bar proportional to its value; the ₹3,52,000 laptop order dwarfs the pen orders visually."
      },
      {
        task: "In the payment tracker, highlight the entire row in light red whenever the Status column says Overdue.",
        hint: "New Rule > Use a formula. Select from A2 across, use =$F2=\"Overdue\".",
        solution: "Select A2:G500, Conditional Formatting > New Rule > Use a formula, type =$F2=\"Overdue\", set a light red fill, OK. Only the column F is locked so each row follows its own status. Compare with the Solution sheet."
      }
    ],
    quiz: [
      {
        q: "What makes conditional formatting different from a manual fill colour?",
        options: [
          "It uses nicer colours",
          "It re-evaluates a rule automatically, so colours update when values change",
          "It can only colour numbers",
          "It survives only in Tables"
        ],
        answer: 1,
        explanation: "A manual fill is static paint. Conditional formatting is a rule that Excel re-checks on every change, so highlights appear and disappear as the data deserves."
      },
      {
        q: "Which rule family would you use to show relative sizes inside the cells, like an in-cell bar chart?",
        options: ["Highlight Cells Rules", "Icon Sets", "Data Bars", "Top 10 Items"],
        answer: 2,
        explanation: "Data Bars draw a bar inside each cell proportional to its value — instant visual comparison while the numbers stay visible."
      },
      {
        q: "You want the whole row coloured when Status (column F) says Overdue. Which formula is correct for a rule applied to A2:G500?",
        options: [
          "=$F$2=\"Overdue\"",
          "=F2=\"Overdue\"",
          "=$F2=\"Overdue\"",
          "=F$2=\"Overdue\""
        ],
        answer: 2,
        explanation: "Lock the column (`$F`) so every cell in a row reads its row's Status, keep the row relative so the rule slides down. Fully locking both parts makes the whole sheet depend on one cell."
      },
      {
        q: "Two rules both match a cell with different fills. Which format appears?",
        options: [
          "Both fills blend together",
          "The rule created last wins",
          "The rule higher in the Manage Rules priority list wins",
          "Excel shows a warning dialog"
        ],
        answer: 2,
        explanation: "Rules have top-to-bottom priority in Manage Rules; the topmost rule wins any conflict. Stop If True can also prevent lower rules from being evaluated at all."
      },
      {
        q: "What does the formula =AND($E2<TODAY(), $F2=\"Unpaid\") highlight?",
        options: [
          "Rows where the due date is in the future",
          "Rows where the due date has passed and the status is still Unpaid",
          "Every row with a date in column E",
          "Only the cells in column F"
        ],
        answer: 1,
        explanation: "AND requires both conditions: E is before today (overdue) and F says Unpaid. Together they identify invoices that need action today — and the highlight updates itself every day."
      }
    ],
    summary: [
      "Conditional formatting is rule-driven colour that updates itself; manual fills are static.",
      "Highlight Cells rules cover thresholds, text contains, dates and duplicates; Top/Bottom rules cover rankings.",
      "Data Bars, Color Scales and Icon Sets add instant visuals — use one or two per sheet, not five.",
      "Manage Rules controls priority, Stop If True and the Applies To range; keep ranges tight for speed.",
      "Formula rules like `=$F2=\"Overdue\"` highlight whole rows: lock the column, leave the row relative."
    ],
    downloads: [
      { file: "assets/files/level3/tables-validation.xlsx",
        label: "Tables, filtering and validation workbook",
        note: "Exercise and Solution sheets for sorting, filtering, Excel Tables, validation and conditional formatting" }
    ],
    keywords: ["conditional formatting", "highlight cells", "greater than", "top 10", "data bars", "color scales", "icon sets", "manage rules", "stop if true", "formula rule", "highlight row"]
  }
);
