// level4-2.js — Level 4: Data Analysis and Reporting (lessons 5-8)
window.LEVEL_LESSONS = window.LEVEL_LESSONS || {};

(window.LEVEL_LESSONS[4] = window.LEVEL_LESSONS[4] || []).push(
  {
    id: "l4-kpi-reporting",
    title: "KPI Reporting",
    difficulty: "Intermediate",
    duration: 16,
    version: "",
    objectives: [
      "Decide what makes a number worth reporting as a KPI",
      "Build a KPI summary layout with target, actual and achievement",
      "Write achievement, variance and growth formulas correctly",
      "Color-code status automatically with conditional formatting",
      "Add sparklines to show each month's trend"
    ],
    intro: "A KPI, or key performance indicator, is the small set of numbers a manager checks first. Not twenty numbers: three to five. For a sales team like Sunrise Traders, the core KPI is simple: revenue this month versus revenue target. Everything in this lesson is built around making that one comparison impossible to misread.\n\nYou will work in `kpi-report-data.xlsx`. The Exercise sheet lists the twelve months of 2025 in rows 2 to 13, with the month name in column A, the revenue target in column B and the actual revenue in column C. Your job is to add achievement percentages, automatic status colors and sparkline trends, ending with a page a manager can read in thirty seconds.",
    sections: [
      {
        heading: "What makes a good KPI",
        blocks: [
          { type: "p", text: "Any number can be measured; very few numbers are worth reporting. A good KPI passes five tests." },
          { type: "ul", items: [
            "**Specific**: \"Revenue from Printers in the North\" beats \"sales\".",
            "**Measurable**: it comes from data you actually have, not estimates.",
            "**Has a target**: a number without a target cannot be judged. ₹5,42,000 is neither good nor bad until you know the target was ₹5,00,000.",
            "**Time-bound**: attached to a month, a quarter or a year.",
            "**Few**: three to six KPIs per report. A page of thirty KPIs is a page nobody reads."
          ] },
          { type: "table", headers: ["KPI", "Target", "Actual", "Why it matters"], rows: [
            ["Monthly revenue", "₹5,00,000", "₹5,42,000", "The headline number of the sales team"],
            ["Achievement %", "100%", "108.4%", "Target met or missed, in one glance"],
            ["Units sold", "1,200", "1,315", "Volume behind the revenue"],
            ["MoM growth", "0%", "+4.1%", "Direction of travel versus last month"]
          ] },
          { type: "p", text: "Notice that \"Achievement %\" is itself a KPI built from two others. Derived KPIs like this are the most useful kind, because they combine a number and its target into a single comparable figure." },
          { type: "p", text: "Beware of vanity metrics: numbers that always look good and therefore tell you nothing. \"Total revenue since the company started\" only ever goes up. \"Calls made\" measures effort, not results. A useful test is to imagine both possible outcomes of a month: if the number would not change any decision in either case, it does not belong on the KPI page. Agree the final KPI list with the manager who reads it, in one short conversation, before building anything. KPIs chosen by the reader get read; KPIs chosen because they were easy to calculate get ignored." }
        ]
      },
      {
        heading: "The KPI summary layout",
        blocks: [
          { type: "p", text: "A KPI row always follows the same shape: name, target, actual, achievement, status. Keep that column order consistent across every report you build, and readers never have to relearn your layout." },
          { type: "p", text: "The layout works because it follows the reader's questions in order. \"What are we measuring?\" is answered by the name. \"What should it be?\" by the target. \"What is it?\" by the actual. \"Did we make it?\" by the achievement. \"How bad is the miss?\" by the status color. When you design any new report block, walk through the questions the reader will ask in sequence and give each one a column; layouts built this way feel obvious to read and take minutes to explain." },
          { type: "grid", title: "The KPI layout with formulas, January row",
            rows: [
              ["Month", "Target", "Actual", "Achievement %", "Status"],
              ["Jan", "₹5,00,000", "₹5,42,000", "=C2/B2", "green"],
              ["Feb", "₹5,00,000", "₹4,86,000", "=C3/B3", "amber"],
              ["Mar", "₹5,00,000", "₹5,61,000", "=C4/B4", "green"]
            ] },
          { type: "formula", name: "Achievement percentage",
            syntax: "=Actual/Target",
            example: "=C2/B2",
            explanation: "Divides the actual revenue by the target. Format the cell as a percentage with one decimal. 108.4% means the target was beaten by 8.4%." },
          { type: "p", text: "Type `=C2/B2` in `D2`, format it as a percentage with one decimal place, and fill it down to `D13`. Every month now shows achievement on the same scale, which is exactly what makes months comparable." },
          { type: "p", text: "A percentage column also survives changes that would break absolute comparisons. If management raises the monthly target in July from ₹5,00,000 to ₹5,50,000, the rupee figures in columns B and C stop being comparable with earlier months, but achievement percentages keep their meaning: 103% in March and 103% in August both mean \"slightly over target\". That scale-free quality is why managers quote achievement, not raw revenue, when they compare across months or across teams with different targets." },
          { type: "warning", text: "If `D2` suddenly shows ##### or a date-like value such as 0.5 changing into something odd, the cell format is wrong, not the formula. Press [[Ctrl]] + [[1]], choose Percentage, and set the decimals to 1." }
        ]
      },
      {
        heading: "Status colors with conditional formatting",
        blocks: [
          { type: "p", text: "The status column should never be typed by hand. Manual colors go stale the moment next month's numbers arrive. Conditional formatting applies the color from the value, so the report recolors itself every refresh." },
          { type: "steps", items: [
            "Select `D2:D13`, the achievement column.",
            "On the Home tab, click Conditional Formatting > New Rule.",
            "Choose Format only cells that contain.",
            "Set the first rule: Cell Value greater than or equal to 1, format with green fill. (1 means 100%.)",
            "Add a second rule: Cell Value between 0.9 and 0.9999, format with amber fill.",
            "Add a third rule: Cell Value less than 0.9, format with red fill.",
            "Click OK and check that January turns green and the weaker months turn amber or red."
          ] },
          { type: "p", text: "The thresholds are a business decision, not an Excel decision. Common practice for monthly revenue is 100% and above green, 90% to 99.9% amber, below 90% red. Agree the bands with your manager once and apply them consistently in every report." },
          { type: "p", text: "Rule order matters when bands overlap. Excel evaluates conditional formatting rules from top to bottom and, by default, applies the first rule that matches. Build the three rules in the order green, amber, red, and make the band edges exact: green means greater than or equal to 1, amber means between 0.9 and 0.9999, red means less than 0.9. A value of exactly 90% then falls into amber, never red, because the amber rule catches it. If a rule ever seems to be ignored, open Conditional Formatting > Manage Rules and check the order and the applies-to range." },
          { type: "tip", text: "You can also use Conditional Formatting > Icon Sets for arrows or traffic lights next to the percentage. If you use icon sets, choose Show Icon Only in the rule settings when the number is already visible in the next column." }
        ]
      },
      {
        heading: "Variance and growth",
        blocks: [
          { type: "p", text: "Achievement answers \"did we hit the target?\". Variance answers \"by how much, in rupees?\" and growth answers \"are we moving in the right direction month over month?\". Professional reports usually carry all three." },
          { type: "formula", name: "Variance percentage versus target",
            syntax: "=(Actual-Target)/Target",
            example: "=(C2-B2)/B2",
            explanation: "January: (5,42,000 - 5,00,000) / 5,00,000 = +8.4%. A positive variance means above target. Note that variance % always equals Achievement % minus 1." },
          { type: "formula", name: "Month-over-month growth",
            syntax: "=(ThisMonth-LastMonth)/LastMonth",
            example: "=(C3-C2)/C2",
            explanation: "Compares February's actual with January's actual. Put it in `E3` and fill down; leave `E2` blank or type a dash, because January has no previous month." },
          { type: "p", text: "Keep the signs meaningful: show variance as `+8.4%` or `-12.0%` by formatting with a custom format such as `+0.0%;-0.0%;0.0%`. A manager scanning the column reads plus signs as wins without doing any math." },
          { type: "p", text: "Variance and growth answer different questions, so read them together rather than instead of each other. A month can be at 102% achievement (target met) yet show negative growth versus last month, which means the team still beat its target even though the market softened. Another month can show positive growth while missing the target, because the target itself was raised. A review meeting that only looks at one of the two numbers will celebrate soft months and punish improving ones; the KPI page prevents that by showing both side by side." },
          { type: "warning", text: "If last month's actual is zero, the growth formula divides by zero and shows #DIV/0!. Wrap it as `=IFERROR((C3-C2)/C2,\"\")` so the cell shows blank instead of an error. Never leave error values visible in a report you hand over." }
        ]
      },
      {
        heading: "A monthly KPI table with sparklines",
        blocks: [
          { type: "p", text: "The final touch is a trend column. Put a sparkline in `G2` that shows the actual revenue from January up to the current row, so each row carries its own mini history." },
          { type: "steps", items: [
            "Click `G2` and type the header \"Trend\" in `G1`.",
            "With `G2` selected, go to Insert > Sparklines > Line.",
            "In the Data Range box type `$C$2:C2` and click OK.",
            "Fill the sparkline down to `G13`.",
            "Select the sparklines and use the Sparkline tab to mark the High Point, so the best month is visible at a glance."
          ] },
          { type: "p", text: "The anchored start `$C$2` matters. In `G2` the sparkline covers one month; in `G13` it covers the full January-to-December line. Each row's sparkline ends at that row's month, so you see how the year built up to that point." },
          { type: "grid", title: "The finished monthly KPI table",
            rows: [
              ["Month", "Target", "Actual", "Achievement %", "MoM %", "Status", "Trend"],
              ["Jan", "₹5,00,000", "₹5,42,000", "108.4%", "", "green", "[sparkline]"],
              ["Feb", "₹5,00,000", "₹4,86,000", "97.2%", "-10.3%", "amber", "[sparkline]"],
              ["Jun", "₹5,00,000", "₹4,38,000", "87.6%", "-9.4%", "red", "[sparkline]"]
            ] },
          { type: "p", text: "This is the deliverable: twelve rows, seven columns, every color and every trend generated from the data. When August's numbers arrive, you type them into column C and the entire page updates itself." },
          { type: "p", text: "Keep the sparkline column narrow, about the width of three or four normal columns, so it reads as a trend shape rather than a chart. If the sparklines look cramped or flat, select them and use the Sparkline tab to set the vertical axis minimum to a sensible floor, such as 4,00,000, so month-to-month differences are visible. Like everything else on the page, sparklines print and travel inside the cells, so the KPI table survives being pasted into an email as a picture or printed for a meeting pack." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Typing \"Green\" or coloring status cells by hand each month.", right: "Use conditional formatting driven by the achievement percentage.", why: "Manual colors do not update when the numbers change, so last month's red hangs next to this month's 105%. Conditional formatting recolors automatically." },
      { wrong: "Writing the achievement formula as `=C2-B2` and calling it achievement.", right: "Achievement is `=C2/B2`; the difference `=C2-B2` is the variance in rupees.", why: "Subtraction gives rupees, division gives a comparable percentage. Managers judge performance on percentages because targets differ across months and teams." },
      { wrong: "Leaving #DIV/0! errors in the MoM growth column.", right: "Wrap growth formulas in `=IFERROR((C3-C2)/C2,\"\")`.", why: "A zero or blank previous month divides by zero. An error value in a handed-over report signals carelessness." },
      { wrong: "Mixing units: target in lakhs in one row and rupees in another.", right: "Keep one unit per column and state it once in the header, for example \"Target (₹)\".", why: "A formula divides whatever is stored, not what you meant. Inconsistent units make every derived percentage wrong in a different way." }
    ],
    shortcuts: [
      { keys: "Ctrl + 1", action: "Open Format Cells to set percentage and custom formats" }
    ],
    exercises: [
      { task: "In kpi-report-data.xlsx, add an Achievement % column in `D` with `=C2/B2` filled to `D13`, formatted as a percentage with one decimal.",
        hint: "Type the formula in `D2`, format with [[Ctrl]] + [[1]], then double-click the fill handle to fill down.",
        solution: "Each month shows actual over target. Months above ₹5,00,000 actual show above 100%; weaker months fall below. Compare with the Solution sheet." },
      { task: "Apply the three conditional formatting rules to `D2:D13`: green at 100% or more, amber from 90% to 99.9%, red below 90%.",
        hint: "Home > Conditional Formatting > New Rule > Format only cells that contain. Build three rules and order them top to bottom.",
        solution: "Strong months turn green, middling months amber, and the weakest months red. Your color bands should match the Solution sheet." },
      { task: "Add a MoM growth column in `E` with `=(C3-C2)/C2` starting in `E3`, wrapped in IFERROR.",
        hint: "Use `=IFERROR((C3-C2)/C2,\"\")` and leave `E2` with a dash typed as text.",
        solution: "Growth compares each month with the one before it. Custom-format the column as `+0.0%;-0.0%;0.0%` so wins carry a plus sign." },
      { task: "Add the anchored sparkline column: in `G2` insert a Line sparkline over `$C$2:C2` and fill down to `G13`. Mark the high point.",
        hint: "Insert > Sparklines > Line. The dollar sign on `$C$2` keeps the start fixed while the fill extends the end.",
        solution: "The December row shows the full year's shape and the January row shows a single point. Enable High Point markers from the Sparkline tab." }
    ],
    quiz: [
      { q: "Which formula computes achievement against a monthly target?",
        options: ["=Actual-Target", "=Actual/Target", "=Target/Actual", "=(Actual-Target)/Actual"],
        answer: 1,
        explanation: "Achievement % is actual divided by target, formatted as a percentage. 108% means the target was beaten by 8%." },
      { q: "Why should the status color come from conditional formatting rather than manual fills?",
        options: ["Conditional formatting is faster to apply once", "Manual fills cannot use green", "Conditional formatting updates automatically when the numbers change", "Conditional formatting prints in color"],
        answer: 2,
        explanation: "The color is tied to the value, so next month's numbers recolor the report by themselves. Manual fills go stale immediately." },
      { q: "The achievement percentage for a month is 85%. With the standard bands (green 100%+, amber 90-99.9%, red below 90%), what color should it show?",
        options: ["Green", "Amber", "Red", "No color"],
        answer: 2,
        explanation: "85% is below the 90% amber floor, so it falls in the red band. The exact bands are a business choice, but 85% under these rules is red." },
      { q: "What does the anchored range `$C$2:C2` achieve in a sparkline filled down a column?",
        options: ["It makes the sparkline vertical", "It keeps the start fixed at January while the end extends row by row", "It formats the sparkline color", "It links the sparkline to the target column"],
        answer: 1,
        explanation: "The absolute start `$C$2` never moves, while the relative end `C2` grows with each row, so each row's sparkline shows the year built up to that month." }
    ],
    summary: [
      "A good KPI is specific, measurable, time-bound, has a target, and there are only a few of them.",
      "Achievement is `=Actual/Target` formatted as a percentage; variance is `=(Actual-Target)/Target`.",
      "Status colors belong to conditional formatting, so the report recolors itself when data changes.",
      "Wrap growth formulas in `IFERROR` to avoid #DIV/0! when a previous month is zero.",
      "Anchored sparklines (`$C$2:C2`) give every month a trend line that ends at that month."
    ],
    downloads: [
      { file: "assets/files/level4/kpi-report-data.xlsx",
        label: "KPI report workbook",
        note: "Monthly sales and target data with Exercise and Solution sheets for KPI reporting" }
    ],
    keywords: ["KPI", "KPI reporting", "achievement %", "target vs actual", "variance", "growth %", "conditional formatting status", "sparkline", "monthly KPI"]
  },
  {
    id: "l4-monthly-yearly-reports",
    title: "Monthly and Yearly Reporting",
    difficulty: "Intermediate",
    duration: 18,
    version: "",
    objectives: [
      "Structure a report page with a title block, period, summary and detail",
      "Calculate month-over-month change with a correct formula",
      "Build a year-to-date (YTD) running total",
      "Create a month summary table with SUMIFS on month numbers",
      "Keep raw data on a separate sheet and name sheets sensibly"
    ],
    intro: "Every office has a monthly report ritual: the numbers go from the raw system export to a page the manager signs off. The analysts who do this well are not the ones with the most formulas, but the ones with a clean structure: raw data on one sheet, the report on another, and every number on the report traceable to the data.\n\nYou will build that structure for Sunrise Traders in `dashboard-data.xlsx`. The Data sheet holds the raw transactions for 2025 as a Table named `DataTable`, with columns Date, Region, Salesperson, Product, Units and Revenue and just over 300 rows. If your copy is not formatted as a Table yet, click inside the data and press [[Ctrl]] + [[T]] first.",
    sections: [
      {
        heading: "The structure of a report page",
        blocks: [
          { type: "p", text: "A report page reads top-down in four blocks. Keep each block in its own rows so the reader's eye moves through it in order." },
          { type: "ul", items: [
            "**Title block**: company and report name, for example \"Sunrise Traders — Monthly Sales Report\", in a merged, bold row.",
            "**Period line**: exactly which month the numbers cover, plus who prepared it and when.",
            "**Summary**: three to five headline numbers: total revenue, achievement, best region.",
            "**Detail**: the month-by-month table, and optionally a region or product breakdown."
          ] },
          { type: "p", text: "The period line sounds trivial and saves arguments. \"November\" alone is ambiguous: which November? Write the full period, such as \"Period: 1-30 November 2025\". When a file circulates for months, the period line is how a reader knows the report is current." },
          { type: "p", text: "The detail block does not have to stop at the month table. A second small table beneath it, revenue by region for the latest month, or by product for the year, answers the questions that always follow the headline: where exactly did the growth come from? Keep each detail table small, five to ten rows, and put the most important table first. A report page that ends with two tidy detail tables reads as complete; one that trails into six half-filled tables reads as unfinished." },
          { type: "p", text: "Sheet naming is part of the structure. Name sheets by their job: `Data` for raw data, `Report` for the report page, `Dashboard` for visuals, `Checks` for test totals. Never ship a file with \"Sheet1\" and \"Sheet2\"; the next person reading it has to open every sheet to find out what lives where." },
          { type: "warning", text: "Keep raw data and the report on separate sheets from day one. Mixing them works for one month and collapses by month three, when formulas start pointing at rows that have been sorted, filtered or overwritten." }
        ]
      },
      {
        heading: "Month-over-month change",
        blocks: [
          { type: "p", text: "Month-over-month (MoM) change is the most quoted number in any monthly review: how did this month compare with last month? It is always a percentage of the previous month." },
          { type: "formula", name: "Month-over-month change",
            syntax: "=(ThisMonth-LastMonth)/LastMonth",
            example: "=(C3-C2)/C2",
            explanation: "February versus January: subtract January's revenue from February's, then divide by January's. A positive result is growth, negative is decline." },
          { type: "p", text: "The division by the previous month is what makes the number fair. A rise of ₹40,000 is excellent in a ₹4,00,000 month (+10.0%) and invisible in a ₹24,00,000 month (+1.7%). Never report only the rupee difference." },
          { type: "p", text: "Two practical rules. First, the very first month has no previous month: leave its MoM cell as a dash typed as text. Second, wrap the formula in `IFERROR` so a zero or blank month produces an empty cell rather than #DIV/0!." },
          { type: "formula", name: "MoM change with error protection",
            syntax: "=IFERROR((ThisMonth-LastMonth)/LastMonth, \"\")",
            example: "=IFERROR((C3-C2)/C2, \"\")",
            explanation: "Shows the growth percentage normally, and a blank cell if last month is zero or blank." },
          { type: "p", text: "Interpret MoM in context. Retail and wholesale businesses have seasonal months, so a -6% October in office supplies might simply be the pre-festival lull, not a problem. This is why the YTD column matters next to the MoM column: MoM says what happened this month, YTD says whether the year as a whole is on track. Present both, and a single weak month becomes information instead of alarm." }
        ]
      },
      {
        heading: "Year-to-date totals",
        blocks: [
          { type: "p", text: "A YTD column answers \"how much have we accumulated so far this year?\". It is a running total, and the trick is the anchored range you already met with sparklines." },
          { type: "formula", name: "Year-to-date running total",
            syntax: "=SUM(C$2:C2)",
            example: "=SUM(C$2:C2)",
            explanation: "In row 2 it sums January only. Filled down, the end of the range grows (C$2:C3, C$2:C4, ...) while the anchored start stays on January, giving a cumulative total." },
          { type: "p", text: "Type `=SUM(C$2:C2)` next to January's revenue and fill down. The dollar sign on the row number `2` locks the start; the end of the range moves. That one anchor converts a column of monthly figures into a cumulative story: by June the YTD cell holds the half-year total." },
          { type: "p", text: "Managers read the YTD column against the yearly plan. If the plan is evenly spread, YTD at the end of June should be about 50% of the year; if it is 42%, the conversation at the review meeting writes itself." },
          { type: "p", text: "The same anchored pattern produces a quick forward estimate. Average monthly revenue so far times twelve gives a naive full-year forecast: if YTD through August is ₹36,00,000, the run-rate estimate is `=AVERAGE(C$2:C9)*12`, or ₹54,00,000 for the year. It ignores seasonality, so treat it as a conversation starter rather than a commitment, but managers love a simple \"on current run-rate we land at roughly X\" line, and this formula produces it in one cell." }
        ]
      },
      {
        heading: "Building the month summary with SUMIFS",
        blocks: [
          { type: "p", text: "The raw data is one row per transaction, but the report needs one row per month. The cleanest bridge is a month number: add a helper column to `DataTable` that extracts the month from each date, then let `SUMIFS` match it." },
          { type: "steps", items: [
            "Go to the Data sheet. Type the header \"Month\" in the first empty column next to the table; the Table extends automatically.",
            "In that column type `=MONTH([@Date])` and press Enter. Every row fills with a month number from 1 to 12.",
            "On the Report sheet, put the month numbers 1 to 12 in column A of the detail block.",
            "In the Revenue column, use the SUMIFS formula below and fill down twelve rows."
          ] },
          { type: "formula", name: "Revenue for one month",
            syntax: "=SUMIFS(sum_range, criteria_range, criteria)",
            example: "=SUMIFS(DataTable[Revenue], DataTable[Month], A5)",
            explanation: "Adds the revenue of every transaction whose Month helper equals the number in `A5`. Changing `A5` from 1 to 12 in the fill-down covers the whole year." },
          { type: "grid", title: "The report detail block with formulas",
            rows: [
              ["Month no.", "Month", "Revenue", "MoM %", "YTD"],
              ["1", "Jan", "=SUMIFS(DataTable[Revenue],DataTable[Month],A2)", "", "=SUM(C$2:C2)"],
              ["2", "Feb", "=SUMIFS(DataTable[Revenue],DataTable[Month],A3)", "=(C3-C2)/C2", "=SUM(C$2:C3)"],
              ["3", "Mar", "=SUMIFS(DataTable[Revenue],DataTable[Month],A4)", "=(C4-C3)/C3", "=SUM(C$2:C4)"]
            ] },
          { type: "p", text: "Notice the design: the month number lives in its own cell and the `SUMIFS` reads it. That means the same formula fills down unchanged, and you can re-point the summary at a different month by editing one cell. This \"criteria in a cell\" pattern beats typing the number inside the formula." },
          { type: "tip", text: "For the month name in column B, type the twelve names once and keep them, or use `=TEXT(DATE(2025,A2,1),\"mmm\")` to derive the three-letter name from the month number." },
          { type: "p", text: "The criteria-in-a-cell pattern generalises to every summary you will ever build. Put the region name in a cell instead of typing \"North\" inside the formula, and one `SUMIFS` fills down for all four regions. Put the product name there and the same skeleton summarises products. You type the formula once, and the report becomes a set of parameters with a single engine behind them. This is also what makes the report auditable: a reviewer can click any total, see the formula reference a visible month cell, and verify the criteria in two seconds." },
          { type: "p", text: "The same pattern extends to any breakdown. Revenue by region for one month: `=SUMIFS(DataTable[Revenue], DataTable[Month], A5, DataTable[Region], \"North\")`. Every report table you will ever build is a variation of this formula." }
        ]
      },
      {
        heading: "Keeping raw data safe and sheets named",
        blocks: [
          { type: "p", text: "Three habits keep a monthly workbook alive all year." },
          { type: "ul", items: [
            "**Never edit the Data sheet.** It is the source of truth. Corrections are made deliberately and dated, not patched quietly inside a report.",
            "**One sheet, one job.** `Data` holds rows, `Report` holds the summary, charts live on `Report` or a `Dashboard` sheet.",
            "**Name sheets by content.** `Data`, `Report`, `Settings`, `Checks`. Avoid \"Sheet1\", \"Final\", \"Final v2 REAL\".",
          ] },
          { type: "p", text: "Version history belongs in the file name, not the sheet names: `Sunrise Monthly Report 2025-11.xlsx`. If you need a scratch area for experiments, make a `Checks` sheet and put your test totals there; a `SUM` over the raw table compared with the sum of your report rows is a two-minute sanity check that has saved many analysts from embarrassing meetings." },
          { type: "p", text: "The yearly report is then almost free. The same Report sheet doubles as the yearly view: the twelve-month table plus the YTD column is the annual picture, and a single row of summary cells above it gives the year at a glance, with total revenue `=SUM(C2:C13)`, best month `=MAX(C2:C13)` and average month `=AVERAGE(C2:C13)`. Build the monthly report well once, and the yearly report is the same page seen from a slightly further distance." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Building the month summary with twelve hand-edited formulas, one per month.", right: "Put the month number in a cell and let one `SUMIFS` formula read it, then fill down.", why: "Hand-edited formulas drift: one row ends up summing the wrong month and nobody notices until the meeting. Cell-driven criteria keep every row identical." },
      { wrong: "Calculating MoM change as `=C3-C2` and reporting the rupee difference.", right: "Report the percentage: `=(C3-C2)/C2`.", why: "Rupee differences are not comparable between months of different size. Percentages put every month on the same scale." },
      { wrong: "Writing monthly totals as `=SUM(Data!F2:F30)` with fixed row ranges.", right: "Sum the whole Table columns with `SUMIFS` and a month criterion.", why: "Fixed ranges break silently when rows are added or sorted. Structured references such as `DataTable[Revenue]` always cover the full column." },
      { wrong: "Fixing a wrong revenue figure directly on the Report sheet.", right: "Correct the source row on the Data sheet and let the report formulas pick it up.", why: "A patched report diverges from its data. The next refresh or the next analyst undoes the patch, and the two versions of truth start fighting." }
    ],
    shortcuts: [
      { keys: "Ctrl + T", action: "Format the raw range as a Table so structured references work" }
    ],
    exercises: [
      { task: "In dashboard-data.xlsx, add a Month helper column to `DataTable` with `=MONTH([@Date])`, then build a 12-row summary on the Report sheet with `=SUMIFS(DataTable[Revenue], DataTable[Month], A2)` filled down.",
        hint: "Type the header Month in the first empty column beside the table so the Table extends. Put 1 to 12 in column A of the Report sheet.",
        solution: "You should get twelve monthly revenue totals whose sum equals the full-year revenue. Cross-check with `=SUM(DataTable[Revenue])` on a Checks sheet." },
      { task: "Add the MoM column with `=IFERROR((C3-C2)/C2,\"\")` from row 3 down, and a YTD column with `=SUM(C$2:C2)` filled down.",
        hint: "Start MoM in row 3. The anchor `C$2` must keep the dollar sign when you fill down.",
        solution: "MoM shows growth from February onward, and YTD accumulates to the full-year total in the December row. The December YTD should equal the total revenue." },
      { task: "Build a one-month region breakdown: for month number 11 (November), revenue by each of the four regions.",
        hint: "Use a two-criteria SUMIFS: `=SUMIFS(DataTable[Revenue], DataTable[Month], 11, DataTable[Region], H2)`.",
        solution: "Four rows, one per region, all for November only. Their sum should match November's row in the monthly summary." },
      { task: "Create the title block and period line on the Report sheet, and rename any default sheets to `Data`, `Report` and `Checks`.",
        hint: "Merge cells across the top row for the title, bold it, and put the period and prepared-by lines beneath.",
        solution: "The title reads \"Sunrise Traders — Monthly Sales Report\", with \"Period: 1-30 November 2025\" and your name underneath. Sheet tabs should now read Data, Report and Checks." }
    ],
    quiz: [
      { q: "Which formula computes February's month-over-month change if January's revenue is in `C2` and February's in `C3`?",
        options: ["=C3-C2", "=(C3-C2)/C2", "=(C3-C2)/C3", "=C2/C3"],
        answer: 1,
        explanation: "MoM change divides the difference by the previous month: (this month minus last month) divided by last month. Dividing by C3 would measure against the wrong base." },
      { q: "Why does `=SUM(C$2:C2)` produce a year-to-date total when filled down?",
        options: ["The dollar sign makes Excel sum faster", "The anchored start stays on January while the end grows row by row", "It sums the whole column automatically", "The dollar sign converts the result to currency"],
        answer: 1,
        explanation: "C$2 locks the start row; the relative end C2 expands with the fill. Each row therefore sums January through the current month." },
      { q: "What is the main reason to keep raw data on its own sheet?",
        options: ["It makes the file smaller", "Excel requires one Table per sheet", "It keeps the source untouched so report formulas stay reliable", "Raw data cannot be formatted"],
        answer: 2,
        explanation: "Raw data is the source of truth. Mixing it with the report invites accidental edits, sorts and overwrites that break every formula above it." },
      { q: "In the summary formula `=SUMIFS(DataTable[Revenue], DataTable[Month], A5)`, what is the role of `A5`?",
        options: ["It holds the revenue total", "It holds the month number the formula matches", "It counts the rows", "It stores the sheet name"],
        answer: 1,
        explanation: "The criterion lives in a cell so one identical formula fills down for all twelve months. Edit the cell, and the same formula reports a different month." }
    ],
    summary: [
      "Structure every report top-down: title block, period line, summary, then detail.",
      "MoM change is `=(ThisMonth-LastMonth)/LastMonth`, protected with `IFERROR` and shown as a percentage.",
      "YTD is an anchored running total: `=SUM(C$2:C2)` filled down.",
      "Summarise raw transactions with `SUMIFS` on a Month helper column, with the month number stored in a cell.",
      "Keep raw data on a separate, untouched, well-named sheet; check totals on a Checks sheet before delivering."
    ],
    downloads: [
      { file: "assets/files/level4/dashboard-data.xlsx",
        label: "Dashboard dataset",
        note: "Monthly business data for building the mini dashboard and report; Exercise and Solution sheets" }
    ],
    keywords: ["monthly report", "yearly report", "month over month", "MoM", "YTD", "year to date", "SUMIFS month", "report structure", "sheet naming", "monthly management report"]
  },
  {
    id: "l4-dashboard-design",
    title: "Interactive Dashboard Design",
    difficulty: "Intermediate",
    duration: 20,
    version: "",
    objectives: [
      "Plan a dashboard around an audience and their questions",
      "Apply layout rules: importance top-left, consistent colors, 4-8 visuals",
      "Build KPI cards with live formulas",
      "Combine pivots, PivotCharts and slicers into an interactive page",
      "Give the dashboard a clean, professional look"
    ],
    intro: "A dashboard is a single screen that answers the reader's questions before they have to ask. The difference between a good dashboard and a bad one is almost never the Excel skill involved; it is the planning. Good dashboards are designed on paper first, with a specific reader and three specific questions in mind.\n\nIn this lesson you will build a mini dashboard for Sunrise Traders in `dashboard-data.xlsx`, using the same raw `DataTable` from the reporting lesson, the pivots and slicers from earlier in this level, and a row of KPI cards. By the end you will have a one-page, filter-driven view of the year. Keep the first version deliberately small: four KPI cards, two charts, two slicers and a timeline is already a complete, useful page, and the discipline of keeping it small is itself part of the skill.",
    sections: [
      {
        heading: "Plan before you build",
        blocks: [
          { type: "p", text: "Answer three questions on paper before touching Excel." },
          { type: "p", text: "Why so much planning for one Excel sheet? Because a dashboard is read by people with thirty seconds and one question, and every design decision either serves or wastes those thirty seconds. A planned dashboard puts the answer to the likely question where the eye lands first. An unplanned one grows by accretion, a chart here, a table there, until the reader has to work to find anything, and busy people quietly stop opening it. The plan is short, but it is the whole difference between a tool and a collection of objects." },
          { type: "ol", items: [
            "**Who reads this?** The regional sales head? The owner? A dashboard for the owner leads with total revenue and profit; one for a regional head leads with their region's numbers.",
            "**What three questions should it answer?** For Sunrise Traders: How much revenue this year? How is it trending by month? Which region and product drive it?",
            "**What will they click?** If the audience will slice by region, that decides your slicers up front."
          ] },
          { type: "p", text: "Sketch the layout on paper: a big rectangle for the KPI row, boxes for the charts, a strip for slicers. Five minutes of sketching saves an hour of moving charts around later." },
          { type: "p", text: "Check data readiness as part of the plan. A dashboard is only as good as the table underneath it, so confirm before designing that the raw data covers the full period, has no text in the date column, and is formatted as a Table. If any month is missing or a region name is inconsistent, fix the data first. Designing a beautiful page over a broken foundation is the most expensive mistake in dashboard work, because every correction later costs changes on the finished page." },
          { type: "quote", text: "A dashboard that tries to answer every question answers none of them." }
        ]
      },
      {
        heading: "Layout rules",
        blocks: [
          { type: "p", text: "Keep expectations realistic as you build. The dashboard in this lesson is deliberately small: four KPI cards, two charts, two slicers and a timeline. That is a complete, useful page, and the discipline of keeping it small is itself part of the skill; you can apply the same plan-build-polish cycle to any size of dashboard later." },
          { type: "ul", items: [
            "**Most important goes top-left.** Readers scan from the top-left corner in a Z pattern. Total revenue for the year goes there, not a decorative logo.",
            "**4 to 8 visuals maximum.** A dashboard is a summary, not a report. Four well-chosen charts beat twelve crowded ones.",
            "**One color story.** Same color for the same measure on every chart: revenue always blue, units always gray. A highlight color, such as orange, is reserved for the thing that needs attention.",
            "**Group related things.** Month trend next to the KPI cards it explains; region and product charts side by side for comparison.",
            "**Leave white space.** Empty rows between blocks make the page readable. Crammed dashboards get skimmed and then ignored."
          ] },
          { type: "table", headers: ["Zone", "Content", "Sunrise Traders example"], rows: [
            ["Top strip", "KPI cards: the headline numbers", "Total revenue, best month, average month, achievement"],
            ["Middle left", "The primary trend", "Line PivotChart of revenue by month"],
            ["Middle right", "The primary breakdown", "Column PivotChart of revenue by region"],
            ["Control strip", "Slicers and timeline", "Region, Product slicers and a Date timeline"]
          ] },
          { type: "p", text: "Sizes carry meaning too. The primary chart should be visibly the largest object on the page; supporting charts can be smaller and simpler. Titles belong on every visual, in the same style, so a reader who looks only at titles still gets the full story: the KPI labels, then \"Monthly revenue 2025\", then \"Revenue by region, 2025\". If you can read your dashboard's titles in sequence and reconstruct the business situation, the layout is doing its job." },
          { type: "p", text: "Respect the natural reading order in placement as well as size. A reader moves left to right within a row, so the month trend (the timeline of the story) belongs on the left and the region comparison (the breakdown of that story) on the right. Slicers belong where the hand naturally rests, at the top or the left edge, never floating between charts where they interrupt the flow. When everything is placed on purpose, the page feels calm; when it is not, no amount of formatting fixes it." }
        ]
      },
      {
        heading: "Building blocks: KPI cards",
        blocks: [
          { type: "p", text: "A KPI card is a small block of merged cells: a label in small text above, a big bold number below. The number must be a live formula, never a typed value, so the card updates with the data." },
          { type: "steps", items: [
            "Create a new sheet named `Dashboard` and turn off gridlines (View tab > untick Gridlines).",
            "Select four cells across, for example `B3:D3`, and click Merge & Center. Type the label \"TOTAL REVENUE 2025\" in small gray capitals.",
            "Merge `B4:D4` below it and enter the formula `=SUM(DataTable[Revenue])`.",
            "Format the number large and bold, as ₹ with no decimals.",
            "Repeat for the other cards: best month, average month, and the latest MoM growth."
          ] },
          { type: "grid", title: "KPI card row with live formulas",
            rows: [
              ["TOTAL REVENUE 2025", "", "BEST MONTH", "", "AVG MONTH"],
              ["=SUM(DataTable[Revenue])", "", "=MAX(Report!C2:C13)", "", "=AVERAGE(Report!C2:C13)"],
              ["", "", "", "", ""],
              ["MoM GROWTH (LATEST)", "", "", "", ""],
              ["=Report!D13", "", "", "", ""]
            ] },
          { type: "p", text: "The cards read from the Report sheet's month summary built in the previous lesson, which keeps a single chain of truth: raw table, report summary, dashboard cards. If the Report sheet is named differently in your workbook, adjust the reference, for example `=MAX('Report'!C2:C13)`." },
          { type: "warning", text: "Think twice before pointing a KPI card directly at a pivot cell, such as `='PT Revenue'!B5`. If someone rearranges the pivot, the card silently shows the wrong number. Prefer formulas over `DataTable` or the Report summary, or use `GETPIVOTDATA`, which Excel writes for you when you type `=` and click a pivot cell and which keeps pointing at the right value even if the pivot moves." },
          { type: "p", text: "Style the cards consistently so they read as a family: same height, same gap between them, labels in small gray capitals, numbers large and dark. If a card has a target, add achievement underneath in a smaller line with conditional formatting, so the card carries its own verdict. For example, below a revenue card, a line reading `=Report!D13` formatted as a percentage and colored red below 90% turns four static numbers into four live judgements." }
        ]
      },
      {
        heading: "Interactivity with pivots and slicers",
        blocks: [
          { type: "p", text: "The interactive area of the dashboard is built from the pieces of the slicer lesson: pivots, PivotCharts and connected slicers." },
          { type: "steps", items: [
            "On the Dashboard sheet, insert two PivotTables from `DataTable`, or move existing pivots here: one with months in Rows, one with regions in Rows, both showing Sum of Revenue.",
            "Insert a PivotChart (line) from the month pivot and a PivotChart (column) from the region pivot.",
            "Insert a slicer for Region and one for Product, plus a timeline for Date.",
            "Right-click the Region slicer > Report Connections and tick both pivots so one click filters the whole page.",
            "Position the slicers in a control strip above or beside the charts, and resize them to neat button grids."
          ] },
          { type: "p", text: "Understand the boundary of interactivity: slicers drive pivots and PivotCharts. The KPI cards use ordinary formulas, so they do not respond to slicers. That is usually the right split: the cards stay as stable headline numbers while the pivots below answer the drill-down questions. If you need slicer-aware KPIs, build those cards from pivot values with `GETPIVOTDATA`." },
          { type: "p", text: "Hide the mechanical clutter so the interactivity feels effortless. Right-click each PivotChart and use Hide All Field Buttons on Chart. Shrink slicer captions or remove redundant ones when the button labels are self-explanatory. Then rehearse the main click-paths once: region on and off, a product in and out, a quarter on the timeline. If any path takes more than two clicks, adjust the slicer layout. An interactive page that is awkward to drive gets driven once and then ignored." }
        ]
      },
      {
        heading: "The clean look",
        blocks: [
          { type: "ul", items: [
            "**Remove gridlines** on the dashboard sheet: View tab > untick Gridlines. The page becomes a canvas instead of a spreadsheet.",
            "**One font, three sizes.** For example Calibri throughout: 10 pt labels, 11 pt data, 18-20 pt KPI numbers.",
            "**Align everything.** Hold [[Alt]] while dragging a chart to snap its edges to cell boundaries. Charts that sit one pixel off-grid make the whole page look untidy.",
            "**Mute the fills.** White or very light gray backgrounds, thin borders, and color only in the data itself.",
            "**Delete the decoration.** No heavy chart borders, no clip-art style shapes, no rainbow palettes."
          ] },
          { type: "p", text: "Consistency is the test. Open your finished dashboard and check: does every chart title use the same style? Are all the slicers the same width? Do revenue charts share one color? Fix anything that breaks the pattern, and the dashboard will look professionally made even with modest charts." },
          { type: "p", text: "Test the look the way it will actually be seen. If the dashboard will be projected in a meeting room, view it at 100% zoom from the back of the room's distance: KPI numbers should be readable, chart labels legible, slicer buttons big enough to hit. Dark text on light backgrounds survives projectors far better than light text on dark, and thin gray fonts that look elegant on a laptop disappear on a wall. Design for the room, not the desk." }
        ]
      },
      {
        heading: "Step by step: the mini dashboard",
        blocks: [
          { type: "steps", items: [
            "Open dashboard-data.xlsx and confirm the Data sheet table (`DataTable`) and the Report sheet month summary are in place from the previous lesson.",
            "Add a `Dashboard` sheet, turn gridlines off, and build four KPI cards across the top using the formulas from the KPI cards section.",
            "Below the cards, insert the two pivots and their PivotCharts: monthly trend line on the left, region comparison columns on the right.",
            "Add Region and Product slicers and a Date timeline in a control strip, and connect the Region slicer to both pivots with Report Connections.",
            "Align all objects to the cell grid with [[Alt]] + drag, apply one chart style and one color for revenue everywhere.",
            "Test it: click each slicer button, verify both pivots and charts respond, then clear every filter before saving.",
            "Compare with the Solution sheet and fix any layout or formula differences."
          ] },
          { type: "tip", text: "Protect your work: select the cells a user is allowed to touch (usually none on a dashboard), then Review > Protect Sheet. Viewers can still click slicers and timelines, but cannot move your carefully aligned charts." },
          { type: "p", text: "Treat the first version as a draft to be watched, not finished. After the dashboard has been used in one real meeting, note what people actually asked for: a product breakdown that was missing, a quarter nobody could find, a KPI nobody mentioned. Adjust the page toward the real questions, and by the third monthly cycle you will have a dashboard that fits its audience the way a good report always does: through use, not through theory." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Twelve charts and six slicers on one screen.", right: "Four to eight visuals, one control strip, everything else moved to the report sheet.", why: "A dashboard is a summary for decisions. Past eight visuals, readers stop scanning and the extra charts add noise, not insight." },
      { wrong: "Every chart in a different color scheme.", right: "One color per measure across all charts, one accent color for highlights.", why: "Color should encode meaning. If revenue is blue in one chart and red in the next, the reader subconsciously stops trusting the page." },
      { wrong: "KPI cards that type the number in, such as 2480000 as text.", right: "Live formulas such as `=SUM(DataTable[Revenue])` in every card.", why: "Typed numbers never update, so the dashboard lies quietly after the first data change. A dashboard with static numbers is worse than no dashboard." },
      { wrong: "Building charts first and deciding the questions later.", right: "Plan audience, three questions and a paper sketch before opening Excel.", why: "Design-first dashboards get reworked endlessly because no one agreed on what the page is for. The plan is the specification." }
    ],
    shortcuts: [
      { keys: "Alt + drag", action: "Snap a chart or object edge to the cell grid while moving it" },
      { keys: "Ctrl + T", action: "Ensure the raw data is a Table so all dashboard formulas use structured references" }
    ],
    exercises: [
      { task: "In dashboard-data.xlsx, create a Dashboard sheet with gridlines off and build the four KPI cards: total revenue, best month, average month and latest MoM growth.",
        hint: "Merge cells for each card. Use `=SUM(DataTable[Revenue])`, `=MAX(Report!C2:C13)`, `=AVERAGE(Report!C2:C13)` and `=Report!D13`.",
        solution: "All four numbers should match the Report sheet totals. Format the money cards as ₹ with no decimals and the growth card as a percentage." },
      { task: "Insert the two pivots and PivotCharts (monthly trend line, revenue by region columns) below the KPI cards.",
        hint: "Build both pivots from `DataTable`, then PivotTable Analyze > PivotChart for each.",
        solution: "The line shows the twelve-month trend; the columns compare the four regions. Titles should be specific, such as \"Monthly revenue 2025\"." },
      { task: "Add Region and Product slicers plus a Date timeline, connect the Region slicer to both pivots, and test every button.",
        hint: "PivotTable Analyze > Slicer / Insert Timeline. Right-click the Region slicer > Report Connections.",
        solution: "Clicking West filters both pivots and both charts; the timeline restricts both to the chosen period. Clear all filters when you finish testing." },
      { task: "Polish the layout: align every object to the cell grid, set one font in three sizes, and give revenue charts a single consistent color.",
        hint: "Hold [[Alt]] while dragging to snap to the grid. Use Chart Design styles rather than formatting charts one by one.",
        solution: "The page should read as one designed sheet, not a collection of pasted objects. Compare your proportions with the Solution sheet." }
    ],
    quiz: [
      { q: "Where should the most important element of a dashboard be placed?",
        options: ["Bottom-right, where the eye ends", "Center, for balance", "Top-left, where readers start scanning", "Wherever the biggest chart fits"],
        answer: 2,
        explanation: "Readers scan in a Z pattern starting top-left. The headline number, usually total revenue, goes there." },
      { q: "How many visuals should a dashboard show at most?",
        options: ["As many as there are KPIs", "4 to 8", "Exactly 3", "At least 12 for completeness"],
        answer: 1,
        explanation: "Four to eight well-chosen visuals keep the page scannable. More than that and readers stop engaging with the page." },
      { q: "Which parts of a dashboard respond to slicers?",
        options: ["Everything on the sheet", "Only PivotTables and PivotCharts connected to them", "Only charts, never tables", "Only cells with formulas"],
        answer: 1,
        explanation: "Slicers filter pivots and their PivotCharts via Report Connections. Formula-based KPI cards stay unchanged unless you deliberately build them from pivot values." },
      { q: "Why should a KPI card use a formula like `=SUM(DataTable[Revenue])` instead of a typed number?",
        options: ["Formulas are faster to type", "The card updates automatically when the data changes", "Typed numbers cannot be formatted", "Formulas print better"],
        answer: 1,
        explanation: "A typed number freezes yesterday's truth. The live formula keeps the card correct after every refresh, which is the entire point of a dashboard." }
    ],
    summary: [
      "Plan first: audience, three questions, and a paper sketch before any Excel work.",
      "Layout rules: most important top-left, 4-8 visuals, one color per measure, generous white space.",
      "KPI cards are merged cells with live formulas, chained to the report summary rather than typed values.",
      "Interactivity comes from pivots, PivotCharts, slicers and timelines connected with Report Connections.",
      "Clean look: gridlines off, one font in three sizes, everything aligned to the cell grid, filters cleared before delivery."
    ],
    downloads: [
      { file: "assets/files/level4/dashboard-data.xlsx",
        label: "Dashboard dataset",
        note: "Monthly business data for building the mini dashboard and report; Exercise and Solution sheets" }
    ],
    keywords: ["dashboard", "dashboard design", "KPI cards", "slicer dashboard", "interactive dashboard", "dashboard layout", "GETPIVOTDATA", "gridlines off", "mini dashboard", "Excel dashboard"]
  },
  {
    id: "l4-report-design",
    title: "Professional Report Design",
    difficulty: "Intermediate",
    duration: 16,
    version: "",
    objectives: [
      "Apply consistent fonts, sizes and alignment to a report",
      "Format numbers as currency and thousands correctly",
      "Freeze header rows on screen and repeat them when printing",
      "Set up the page: margins, orientation, scale to fit, print area, headers and footers",
      "Export a report to PDF and run a pre-delivery checklist"
    ],
    intro: "The last mile of analysis is delivery. A manager does not see your `SUMIFS` or your careful pivots; they see the page. Fonts, alignment, number formats and print setup decide whether your month of work looks trustworthy or thrown together. Excel gives you all of these tools in a few menus, and using them takes minutes.\n\nThis lesson finishes the Sunrise Traders report in `dashboard-data.xlsx`. You will take the Report sheet you structured in the monthly reporting lesson and make it presentable on screen, on paper and as a PDF.",
    sections: [
      {
        heading: "Fonts, sizes and alignment",
        blocks: [
          { type: "p", text: "Use one font for the whole report. Calibri or Aptos, the Excel defaults, are fine; a report is not the place for decorative fonts. Vary size, not typeface." },
          { type: "table", headers: ["Element", "Style", "Example"], rows: [
            ["Report title", "16 pt bold", "Sunrise Traders — Monthly Sales Report"],
            ["Section headers", "12 pt bold", "Revenue by region"],
            ["Column headers", "11 pt bold, centered", "Region | Revenue | MoM %"],
            ["Data cells", "10-11 pt regular", "₹6,45,000"],
            ["Notes and footers", "9 pt gray", "Source: DataTable, refreshed 01-Dec-2025"]
          ] },
          { type: "p", text: "Alignment follows a fixed convention that every finance team on earth uses: **text left-aligned, numbers right-aligned, headers centered or left**. Numbers must be right-aligned so their digits line up vertically; a column of right-aligned rupee figures can be compared at a glance, a centered column cannot. Excel aligns numbers right automatically, so mostly your job is not to break it." },
          { type: "p", text: "Zoom consistency is part of alignment too. Set the zoom to 100% while formatting; a page polished at 80% zoom often reveals cramped, overlapping text at full size, and that is how the recipient will first see it. Check the report at 100% before you call it done, and check it once more in Print Preview, where the true print font sizes appear." },
          { type: "p", text: "Bold is for structure, not emphasis. Bold the title, the column headers and the total row; leave the data regular. If everything is bold, nothing stands out." },
          { type: "p", text: "Small spacing decisions add up to the professional feel. Keep row heights uniform so the table breathes evenly; a double-height row in the middle of a table looks like an error. Use consistent column widths within a table, and left-align the report title with the table beneath it rather than centering it, so the page has a single visual edge. Finally, keep decorative touches out: a report needs no fill colors beyond a light header band and the status colors, and every extra element competes with the numbers for attention." }
        ]
      },
      {
        heading: "Number formats",
        blocks: [
          { type: "p", text: "Number formatting is where reports earn their credibility. Press [[Ctrl]] + [[1]] to open Format Cells, and use these standards." },
          { type: "table", headers: ["What", "Format", "Shows as"], rows: [
            ["Revenue and money", "Currency, ₹ symbol, 0 decimals", "₹6,45,000"],
            ["Large numbers in tight tables", "Number with thousands separator, 0 decimals", "6,45,000"],
            ["Percentages", "Percentage, 1 decimal", "26.0%"],
            ["Units", "Number, thousands separator, 0 decimals", "1,315"],
            ["Signed changes", "Custom +0.0%;-0.0%;0.0%", "+8.4% or -12.0%"]
          ] },
          { type: "p", text: "Two habits matter more than the exact formats. First, state the unit once in the header, for example \"Revenue (₹)\", rather than repeating the symbol on every cell when the column is clear. Second, be consistent: the same measure keeps the same format everywhere in the workbook. If revenue is ₹ with no decimals on the Report sheet, it is ₹ with no decimals on the Dashboard too." },
          { type: "tip", text: "The Increase Decimal and Decrease Decimal buttons on the Home tab change only what is displayed, never the stored value. A cell showing ₹6,45,000 still stores 645000, so all your sums stay exact." },
          { type: "p", text: "Do not rely on color alone to carry meaning, because some readers cannot distinguish red from green. Pair every color signal with a second cue: the status column already shows the percentage beside the fill, and a custom number format with plus and minus signs gives variance a symbol as well as a sign. A report that remains readable when printed in black and white is a report that survives every reader and every printer in the building." }
        ]
      },
      {
        heading: "Freeze panes and print titles",
        blocks: [
          { type: "p", text: "On screen, a long report scrolls its headers away after a few rows. Freeze panes keeps the important rows visible while you scroll." },
          { type: "steps", items: [
            "On the Report sheet, click cell `A2` if only row 1 should stay fixed, or click below and to the right of everything that should stay visible for a combined freeze.",
            "Go to View > Freeze Panes > Freeze Panes.",
            "Scroll down: the title row stays put.",
            "To undo later, View > Freeze Panes > Unfreeze Panes."
          ] },
          { type: "p", text: "Freeze panes fix the screen, but printing has the same problem: page 2 of a printed table arrives without headers. **Print titles** solve it." },
          { type: "steps", items: [
            "Go to Page Layout > Print Titles (in the Page Setup group).",
            "Click in the Rows to repeat at top box.",
            "Click row 1 on the sheet, so the box shows `$1:$1`.",
            "Click OK. Every printed page of the sheet now starts with your header row."
          ] },
          { type: "warning", text: "Freeze panes and print titles are separate features: freezing does nothing on paper, and print titles do nothing on screen. A professional report usually needs both." },
          { type: "p", text: "For wide tables, freezing can cover columns as well as rows. Click the cell to the right of the labels column, for example `B2`, then Freeze Panes: row 1 and column A both stay fixed while the rest scrolls. This is the standard setup for any table where you scroll sideways through months. On the Report sheet, with months running down the rows, freezing just the top row is enough; pick the freeze that matches the direction the table grows." }
        ]
      },
      {
        heading: "Page setup: margins, orientation, scale",
        blocks: [
          { type: "p", text: "Open File > Print (or press [[Ctrl]] + [[P]]) and look at the preview. This preview is where most reports reveal their problems: columns spilling onto extra pages, headers missing, rows chopped in half. Fix it all in Page Layout." },
          { type: "ul", items: [
            "**Orientation**: wide tables want Landscape, tall ones Portrait. A month summary with five columns usually fits portrait; a dashboard-style layout needs landscape.",
            "**Margins**: Normal is fine for most reports; Narrow buys you an extra column or two on wide sheets.",
            "**Scale to Fit**: on the Page Layout tab, set Width to 1 page and leave Height automatic. Excel shrinks the printout just enough to fit all columns across one page width.",
            "**Print area**: select exactly what should print, then Page Layout > Print Area > Set Print Area. Stray notes in far columns stop appearing as surprise extra pages. Clear it the same way when it is no longer wanted."
          ] },
          { type: "p", text: "**Headers and footers** carry the page furniture: page numbers, file name, print date. Insert > Header & Footer switches the view; type in the header section, or use the Header & Footer Elements such as Page Number (inserted as the code `&[Page]`), File Name and Current Date. A footer like \"Page &[Page] of &[Pages]\" tells the reader how much is left, which matters on a twelve-page printout." },
          { type: "tip", text: "By default Excel does not print gridlines, which is usually what you want; your cell borders and fills carry the structure. If you ever need the grid itself printed, Page Setup > Sheet > tick Print gridlines." },
          { type: "p", text: "Use Print Preview as a workshop, not a final check. Press [[Ctrl]] + [[P]], study the preview, close it, adjust one thing on the sheet, and look again. For long sheets, View > Page Break Preview shows the page boundaries as blue lines directly on the sheet, so you can drag a break to keep a related table on one page. A report assembled through three or four preview-adjust cycles prints perfectly the first time, every time." }
        ]
      },
      {
        heading: "Exporting to PDF and the pre-delivery checklist",
        blocks: [
          { type: "p", text: "Managers read PDFs. They print them, forward them and annotate them, and the layout cannot be broken by whatever fonts their laptop lacks. Exporting is two clicks, but only if your page setup is already right, because the PDF copies the print layout exactly." },
          { type: "steps", items: [
            "Press [[Ctrl]] + [[P]] and check the preview: correct pages, headers on every page, nothing cut off.",
            "Choose Microsoft Print to PDF as the printer, or use File > Export > Create PDF/XPS Document.",
            "Decide the scope in the Settings list: Print Active Sheets or, if the report spans sheets, Print Entire Workbook.",
            "Click Print (or Publish) and save with a clear name: `Sunrise Monthly Report 2025-11.pdf`.",
            "Open the PDF and page through it once before sending. Always."
          ] },
          { type: "grid", title: "The title block of the finished report",
            rows: [
              ["Sunrise Traders — Monthly Sales Report", "", ""],
              ["Period: 1-30 November 2025", "", "Prepared by: Arjun Mehta"],
              ["Source: DataTable, refreshed 01-Dec-2025", "", "Status: Final"]
            ] },
          { type: "warning", text: "A PDF is a snapshot. If you fix a number after exporting, the PDF does not know; re-export it. Sending a stale PDF next to a corrected workbook is a classic, and entirely avoidable, embarrassment." },
          { type: "p", text: "Name and store the PDF deliberately. A name such as `Sunrise Monthly Report 2025-11.pdf` carries the company, the report and the full period, so it sorts correctly in a shared folder and never needs opening to identify. Save it beside the workbook or in the folder your team already uses for report packs, and if the report goes to several people, send one PDF rather than a zip of workbooks: the PDF keeps your layout, opens in seconds, and cannot be accidentally edited by a reader with good intentions." },
          { type: "p", text: "Before any report leaves your hands, run this checklist. It takes two minutes and catches ninety percent of delivery mistakes." },
          { type: "ul", items: [
            "Numbers: spot-check two or three figures against the source data.",
            "Period and title: the report says which month it covers, in full.",
            "Filters: all slicers and pivot filters are cleared, so totals match the headline.",
            "Errors: no #DIV/0!, #REF! or #N/A visible anywhere.",
            "Formatting: one font, numbers right-aligned, consistent ₹ and % formats.",
            "Screen: freeze panes set, gridlines off on report sheets.",
            "Print: scale to fit 1 page wide, print titles repeating the header row.",
            "PDF: freshly exported, correctly named, and paged through once."
          ] },
          { type: "p", text: "Read the checklist as a story of the last mile: the numbers are right, the page says what it is, the file opens clean, the print survives the printer, and the PDF matches the workbook. Each line exists because someone, somewhere, once shipped a report that failed it. Skippers of the checklist pay for it in the meeting; runners of the checklist develop a reputation for reports that are simply trusted." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Sending the Excel file with no page setup, so the manager's printout spreads one table over four pages.", right: "Set Scale to Fit width 1 page, landscape where needed, and export a PDF yourself.", why: "You control the layout; the recipient's printer does not. A report you have not print-checked is an unreviewed report." },
      { wrong: "Page 2 of the printed report starts with bare numbers and no header row.", right: "Set Page Layout > Print Titles with Rows to repeat at top `$1:$1`.", why: "Without repeated headers, nobody can tell which column is revenue and which is units once the first page is turned." },
      { wrong: "Freezing panes and assuming the printout is fixed too.", right: "Freeze panes for the screen and set print titles for paper; they are separate features.", why: "Freeze panes affect scrolling only. Printers ignore them completely, which is why page 2 arrives headerless despite a frozen top row." },
      { wrong: "Exporting the PDF, then fixing one figure in the workbook and sending both files.", right: "Re-export the PDF after any change, and page through it once more.", why: "The PDF is a snapshot of the moment you exported. A mismatch between PDF and workbook destroys trust in both." }
    ],
    shortcuts: [
      { keys: "Ctrl + P", action: "Open Print preview and settings" },
      { keys: "Ctrl + 1", action: "Open Format Cells for number formats" }
    ],
    exercises: [
      { task: "In dashboard-data.xlsx, style the Report sheet: title at 16 pt bold, column headers 11 pt bold, one font throughout, text left-aligned and numbers right-aligned.",
        hint: "Select cells and set font size on the Home tab. Check the alignment icons: numbers should sit on the right edge of their cells.",
        solution: "The sheet reads with a clear hierarchy: one large title, one style of headers, quiet data rows. Compare with the Solution sheet." },
      { task: "Format the Revenue column as ₹ currency with no decimals and the MoM column with the custom format `+0.0%;-0.0%;0.0%`.",
        hint: "Select the column, press [[Ctrl]] + [[1]], choose Custom, and type the format code.",
        solution: "Revenue reads as clean rupee figures and MoM shows explicit plus and minus signs. Gains are now visible without reading the digits." },
      { task: "Freeze the title row on screen and set print titles so row 1 repeats on every printed page.",
        hint: "View > Freeze Panes for the screen; Page Layout > Print Titles > Rows to repeat at top `$1:$1` for paper.",
        solution: "Scrolling keeps the title visible, and Print Preview shows the header row at the top of every page." },
      { task: "Set the page for printing (portrait, scale to fit 1 page wide, a footer with page numbers) and export the Report sheet as a PDF named `Sunrise Monthly Report 2025-11.pdf`.",
        hint: "Page Layout tab for scale and orientation; Insert > Header & Footer for the footer; Ctrl + P > Microsoft Print to PDF.",
        solution: "The preview shows the whole report on one page width with \"Page 1 of 1\" in the footer. The exported PDF should match the preview exactly." }
    ],
    quiz: [
      { q: "What does the print titles setting do?",
        options: ["Prints the workbook title in big letters", "Repeats chosen rows, such as the header row, at the top of every printed page", "Adds a header and footer to the printout", "Makes the title row bold"],
        answer: 1,
        explanation: "Page Layout > Print Titles > Rows to repeat at top (for example `$1:$1`) repeats those rows on every page, so page 2 still has column headers." },
      { q: "A wide report spills its columns onto a second printed page. What is the quickest fix?",
        options: ["Delete the extra columns", "Set Scale to Fit width to 1 page", "Change the font to 8 pt everywhere", "Print in landscape only"],
        answer: 1,
        explanation: "Scale to Fit with Width 1 page shrinks the printout just enough to fit all columns across one page. Deleting data or shrinking fonts by hand are blunt, fragile alternatives." },
      { q: "Which statement about freeze panes is correct?",
        options: ["They repeat headers on printed pages", "They keep rows visible while scrolling on screen only", "They prevent editing of the header row", "They speed up large workbooks"],
        answer: 1,
        explanation: "Freeze panes are a screen feature (View tab). Printed output needs print titles instead; the two features solve the same problem in different places." },
      { q: "You fix a wrong figure in the workbook after exporting the PDF. What must you do?",
        options: ["Nothing, the PDF updates automatically", "Re-export the PDF and check it again", "Rename the PDF", "Delete the workbook"],
        answer: 1,
        explanation: "A PDF is a snapshot of the print layout at export time. Any change to the workbook after that leaves the PDF stale until you re-export it." }
    ],
    summary: [
      "One font, a clear size hierarchy, bold only for structure, and the classic alignment: text left, numbers right.",
      "State units in the header and keep the same number format for the same measure everywhere.",
      "Freeze panes fix the screen; print titles fix the paper; a professional report uses both.",
      "Page setup is part of the deliverable: orientation, scale to fit 1 page wide, print area, headers and footers.",
      "Export to PDF, page through it once, and run the pre-delivery checklist before anything leaves your desk."
    ],
    downloads: [
      { file: "assets/files/level4/dashboard-data.xlsx",
        label: "Dashboard dataset",
        note: "Monthly business data for building the mini dashboard and report; Exercise and Solution sheets" }
    ],
    keywords: ["report design", "professional report", "freeze panes", "print titles", "scale to fit", "print area", "headers and footers", "export PDF", "number formats", "pre-delivery checklist"]
  }
);
