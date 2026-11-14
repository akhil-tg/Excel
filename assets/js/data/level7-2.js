// Excel Lab — Level 7 (Practical Projects), part 2 of 2
// The remaining five project lessons (no level metadata here)

window.LEVEL_LESSONS = window.LEVEL_LESSONS || {};

(window.LEVEL_LESSONS[7] = window.LEVEL_LESSONS[7] || []).push(
  {
    id: "l7-project-expenses",
    title: "Project: Monthly Expense Tracker",
    difficulty: "Intermediate",
    duration: 35,
    version: "",
    objectives: [
      "Build an expense log that accepts only valid categories from a dropdown",
      "Total spending by category and by month with SUMIF",
      "Compare actual spending against budgets and compute the variance",
      "Flag over-budget categories automatically with IF and conditional formatting",
      "Reconcile the summaries against the log so the numbers can be trusted"
    ],
    intro: "Every office watches money leave. In this project you build the expense tracker for Sunrise Traders' head office: a log where every payment is typed once — date, category, description, amount — and two summary sheets that instantly show where the money went, month by month and category by category, against the budgets management approved.\n\nThe tracker covers July to October 2026 and uses six categories: Rent, Salaries, Electricity, Transport, Marketing and Miscellaneous. The tools are all from Levels 2 to 6 — validation, SUMIF, TEXT, IF, conditional formatting — but the project teaches a discipline that tools alone cannot: a small set of checks that prove the summaries agree with the log before anyone sees them.",
    sections: [
      {
        heading: "What You Will Build",
        blocks: [
          { type: "p", text: "Three sheets with a clear division of labour. The **Expense Log** is the only sheet anyone types into, and its Category column only accepts the six approved values — that one dropdown is what makes every summary reliable. The **Category Summary** totals each category across all four months, sets it against its budget, and computes the variance. The **Monthly Summary** shows how total spending moved month to month and what share each month carried." },
          { type: "ul", items: [
            "**Expense Log** — about 25 payments from July to October 2026, with a Month helper column.",
            "**Category Summary** — actual vs budget per category, variance, and an Over Budget / Within Budget status with colour.",
            "**Monthly Summary** — total per month and each month's share of the whole."
          ] },
          { type: "p", text: "The budgets are for the full four-month period, since the summary covers the whole log: Rent ₹1,60,000, Salaries ₹4,80,000, Electricity ₹48,000, Transport ₹60,000, Marketing ₹80,000 and Miscellaneous ₹40,000 — ₹8,08,000 in all. By the end you will see that actual spending ran to ₹8,50,900, with two categories over budget." }
        ]
      },
      {
        heading: "Skills You Will Use",
        blocks: [
          { type: "ul", items: [
            "**Data validation** (Level 3) — the category dropdown on the log.",
            "**SUMIF** (Level 4) — totals by category and by month.",
            "**TEXT with dates** (Level 2) — the Month helper column, exactly as in the sales tracker.",
            "**IF and conditional formatting** (Levels 4 and 6) — the budget status flag.",
            "**Percentage of total** (Level 2) — each month's share of total spend."
          ] },
          { type: "p", text: "If you built the sales tracker in Project 5, this project will feel like a cousin — same helper column, same summary pattern, one new element: comparing against a budget instead of a target." }
        ]
      },
      {
        heading: "Open the Exercise File",
        blocks: [
          { type: "p", text: "Download **Expense Tracker — exercise file**. Inside:" },
          { type: "ul", items: [
            "**Expense Log** — headers in row 1: Date, Category, Description, Amount, Month. About 25 payment rows are filled, but the Category column is free-typed text with no validation yet.",
            "**Category Summary** — headers only: Category, Actual, Budget, Variance, Status.",
            "**Monthly Summary** — headers only: Month, Total, % of Total."
          ] },
          { type: "table", headers: ["Column", "Expense Log column", "Example"],
            rows: [
              ["A", "Date", "01-Jul-26"],
              ["B", "Category (dropdown)", "Salaries"],
              ["C", "Description", "Staff salaries — July"],
              ["D", "Amount (₹)", "118000"],
              ["E", "Month = TEXT helper", "Jul-26"]
            ] },
          { type: "tip", text: "Click `E2` and read `=TEXT(A2,\"mmm-yy\")`. Same trick as the sales tracker: dates become stable labels like `Jul-26` that the monthly SUMIF can match exactly." }
        ]
      },
      {
        heading: "Build Section 1: Guard the Log",
        blocks: [
          { type: "p", text: "One dropdown protects every number downstream. A category typed as \"electricity\" or \"Mkt\" would silently vanish from its summary, so we make that impossible." },
          { type: "steps", items: [
            "On **Expense Log**, select `B2:B100` — the filled rows plus generous headroom for the rest of the year.",
            "Open **Data → Data Validation → List** and set the source to `Rent,Salaries,Electricity,Transport,Marketing,Miscellaneous`. Click OK.",
            "Press `[[Alt]] + [[Down]]` on `B2` to confirm the dropdown opens with six options.",
            "Scan the existing 25 rows for any category that does not match the six approved spellings — fix any you find by picking from the dropdown. The summaries in the next section are only as good as this check."
          ] },
          { type: "warning", text: "Apply the validation to empty rows too, not just the 25 filled ones. The first expense of November should inherit the dropdown without you remembering to add it." }
        ]
      },
      {
        heading: "Build Section 2: The Category Summary and Budget Variance",
        blocks: [
          { type: "p", text: "Six rows, four formulas each, filled once. This is the sheet the manager signs." },
          { type: "steps", items: [
            "On **Category Summary**, type the six categories in `A2:A7` — exact spellings from the dropdown list.",
            "Type the four-month budgets in `C2:C7`: `160000`, `480000`, `48000`, `60000`, `80000`, `40000`.",
            "In `B2` type `=SUMIF('Expense Log'!$B$2:$B$100,$A2,'Expense Log'!$D$2:$D$100)` and fill down — actual spending per category.",
            "In `D2` type `=$C2-$B2` and fill down — the variance: budget minus actual.",
            "In `E2` type `=IF($B2>$C2,\"Over Budget\",\"Within Budget\")` and fill down.",
            "Select `E2:E7` and apply **Highlight Cells Rules → Equal To** with the value `Over Budget` in light red fill; add green for `Within Budget`.",
            "Select `D2:D7` and add a second rule with **Highlight Cells Rules → Less Than** and the value `0`, also in light red — negative variances glow even if someone reads only the numbers column.",
            "In `A9` type `Total` and in `B9` type `=SUM($B$2:$B$7)`; in `C9` type `=SUM($C$2:$C$7)`."
          ] },
          { type: "grid", title: "Category Summary — what you should see",
            rows: [
              ["Category", "Actual (₹)", "Budget (₹)", "Variance (₹)", "Status", ""],
              ["Rent", "160000", "160000", "0", "Within Budget", ""],
              ["Salaries", "462000", "480000", "18000", "Within Budget", ""],
              ["Electricity", "51200", "48000", "-3200", "Over Budget", ""],
              ["Transport", "55000", "60000", "5000", "Within Budget", ""],
              ["Marketing", "86500", "80000", "-6500", "Over Budget", ""],
              ["Miscellaneous", "36200", "40000", "3800", "Within Budget", ""],
              ["Total", "850900", "808000", "-42900", "", ""]
            ] },
          { type: "tip", text: "Read the variance column with its sign, not its colour: positive means money left over, negative means overspend. Rent's variance of exactly 0 is a real result, not a mistake — actual spending matched the budget to the rupee." }
        ]
      },
      {
        heading: "Build Section 3: The Monthly Summary and the Reconciliation Check",
        blocks: [
          { type: "p", text: "The monthly view answers a different question — not where the money went, but when. And one small check at the end proves the whole file is honest." },
          { type: "steps", items: [
            "On **Monthly Summary**, type `Jul-26`, `Aug-26`, `Sep-26`, `Oct-26` in `A2:A5` — matching the log's helper column format exactly.",
            "In `B2` type `=SUMIF('Expense Log'!$E$2:$E$100,$A2,'Expense Log'!$D$2:$D$100)` and fill down to `B5`.",
            "In `C2` type `=$B2/SUM($B$2:$B$5)` and fill down, then format `C2:C5` as percentage with 1 decimal.",
            "In `A7` type `Grand Total (₹)` and in `B7` type `=SUM($B$2:$B$5)`.",
            "Now the check: in `A9` type `Reconciliation` and in `B9` type `=SUM('Expense Log'!$D$2:$D$100)-SUM('Category Summary'!$B$2:$B$7)`. It must read 0.",
            "Format the money columns with thousands separators, bold the headers, and border the blocks."
          ] },
          { type: "grid", title: "Monthly Summary — what you should see",
            rows: [
              ["Month", "Total (₹)", "% of Total", "", ""],
              ["Jul-26", "207650", "24.4%", "", ""],
              ["Aug-26", "212100", "24.9%", "", ""],
              ["Sep-26", "216200", "25.4%", "", ""],
              ["Oct-26", "214950", "25.3%", "", ""],
              ["Grand Total (₹)", "850900", "100%", "", ""],
              ["Reconciliation", "0", "", "", ""]
            ] },
          { type: "formula", name: "The reconciliation check",
            syntax: "=SUM(all log amounts) - SUM(all category totals)",
            example: "=SUM('Expense Log'!$D$2:$D$100)-SUM('Category Summary'!$B$2:$B$7)",
            explanation: "Every rupee in the log must land in exactly one category. If the difference is not 0, some log row carries a category the summary does not know — a typo, a space, a new label. This one cell is how professional files stay trustworthy." },
          { type: "tip", text: "The reconciliation reading 0 is not luck — it is the dropdown working. Try typing a seventh category into a spare log row (bypass validation with a paste) and watch the check jump. That is the exact alarm you want." }
        ]
      },
      {
        heading: "Level Up Your System",
        blocks: [
          { type: "p", text: "The tracker is complete: guarded log, budget comparison, monthly trend and a self-check. The exercises below are the upgrades accounts teams ask for first." }
        ]
      },
      {
        heading: "Compare With the Solution",
        blocks: [
          { type: "steps", items: [
            "Open **Expense Tracker — completed solution** and compare the Category Summary first: actuals, budgets, variances and the two red rows.",
            "Open **Data → Data Validation** on the solution's log and confirm the list covers the spare rows down to row 100.",
            "Click the solution's variance rule set under **Conditional Formatting → Manage Rules** and compare with yours — status colour and negative-variance colour are two separate rules.",
            "Check the solution's reconciliation cell reads 0, then paste a dummy log row with the category \"Stationery\" and watch it break. Undo.",
            "Compare monthly totals: your four figures and the grand total ₹8,50,900 must match, then close without saving."
          ] },
          { type: "p", text: "If your category actuals differ from the solution, the cause lives in the log: a category spelling the dropdown does not know, or a log amount that slipped a digit. The reconciliation cell will point you there — that is its job." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Defining the variance as =B2-C2 (actual minus budget) while labelling the column \"Variance\".", right: "=$C2-$B2 — budget minus actual, so positive means under budget", why: "Both formulas are arithmetic; the difference is convention. A column labelled Variance that reads negative when the team undershot will be misread at the review meeting — pick a sign convention and keep it consistent across every report you build." },
      { wrong: "Typing monthly criteria as \"Jul-2026\" while the helper column produces \"Jul-26\".", right: "Match the helper exactly: =TEXT(A2,\"mmm-yy\") produces Jul-26, so use Jul-26", why: "SUMIF compares text letter for letter. A format mismatch returns 0 for every month with no error anywhere — the numbers simply vanish." },
      { wrong: "Computing each month's share as =B2/SUM(B2:B5) without locking.", right: "=$B2/SUM($B$2:$B$5)", why: "Filled down, the unlocked denominator shrinks row by row — the last month divides by itself and reads 100%. Every share above it is inflated too." },
      { wrong: "Applying the category dropdown only to B2:B26, where data exists today.", right: "Select B2:B100 so the validation covers future rows", why: "November's first expense gets typed into a free cell, one unexpected category slips into the log, and the reconciliation check — not the eye — is what catches it." }
    ],
    shortcuts: [
      { keys: "Alt + Down", action: "Open the category dropdown in a log row" },
      { keys: "Ctrl + 1", action: "Format money and percentage columns" },
      { keys: "F4", action: "Lock a range while writing a formula" }
    ],
    exercises: [
      { task: "Show the name of the category with the highest actual spend, without sorting.", hint: "MAX over the actual column, then INDEX/MATCH to convert the amount into a name.", solution: "=INDEX($A$2:$A$7,MATCH(MAX($B$2:$B$7),$B$2:$B$7,0))" },
      { task: "Count how many categories are over budget.", hint: "COUNTIF over the status column.", solution: "=COUNTIF($E$2:$E$7,\"Over Budget\")" },
      { task: "Add an average monthly spend figure under the Monthly Summary.", hint: "One AVERAGE over the four month totals.", solution: "=AVERAGE($B$2:$B$5)" }
    ],
    quiz: [
      { q: "Which formula totals Marketing's actual spend on the Category Summary?",
        options: [
          "=SUMIF('Expense Log'!$D$2:$D$100,\"Marketing\",'Expense Log'!$B$2:$B$100)",
          "=SUMIF('Expense Log'!$B$2:$B$100,$A5,'Expense Log'!$D$2:$D$100)",
          "=SUM('Expense Log'!$D$2:$D$100,\"Marketing\")",
          "=COUNTIF('Expense Log'!$B$2:$B$100,\"Marketing\")"
        ],
        answer: 1,
        explanation: "SUMIF's order is: the range to test (categories), the criteria (the label cell), then the range to add (amounts). Using the label cell $A5 instead of typed text keeps the formula fillable down the whole summary." },
      { q: "Electricity shows actual ₹51,200 against a budget of ₹48,000. What do the variance and status show?",
        options: [
          "Variance ₹3,200, Within Budget",
          "Variance -₹3,200, Over Budget",
          "Variance ₹99,200, Within Budget",
          "Variance -₹3,200, Within Budget"
        ],
        answer: 1,
        explanation: "Variance = budget − actual = 48,000 − 51,200 = −3,200. The status formula tests actual > budget, so it returns Over Budget, and both the status rule and the negative-variance rule light it red." },
      { q: "What is the reconciliation cell actually checking?",
        options: [
          "That every budget equals its actual",
          "That the log's grand total equals the sum of the category totals — every expense landed in a known category",
          "That the four monthly totals are equal",
          "That the dropdown list has six items"
        ],
        answer: 1,
        explanation: "It is a two-sided count of the same money. A difference other than zero means a log row carries a category outside the summary's six labels — the silent failure the dropdown is meant to prevent." },
      { q: "Why does the monthly summary need the helper column at all?",
        options: [
          "To make the file look organised",
          "Because SUMIF needs an exact text match, and the helper turns every date into a stable label like Jul-26",
          "Because Excel cannot calculate with dates",
          "To allow more than 25 rows"
        ],
        answer: 1,
        explanation: "Dates are serial numbers with display formats — reliable for people, awkward as SUMIF criteria. One TEXT column per log row converts them into plain labels that match exactly, every time." },
      { q: "Sep-26 shows ₹2,16,200 of the ₹8,50,900 total. Roughly what should the % of Total cell read?",
        options: [
          "About 39% — 216200/550000",
          "About 25% — 216200/850900",
          "About 4% — 216200 divided by six categories",
          "100% — the largest month is always 100%"
        ],
        answer: 1,
        explanation: "Share of total is the month divided by the grand total of all four months, with the denominator locked: =$B4/SUM($B$2:$B$5). The four shares must add to 100%." }
    ],
    summary: [
      "A dropdown on the category column is the single control that keeps every summary honest.",
      "SUMIF per category and per month — same tool, two criteria columns: Category (B) and Month helper (E).",
      "Variance has a convention: budget minus actual, positive means under budget — state it and keep it.",
      "Two conditional formatting rules do the alarm work: the status text and the negative variance.",
      "A reconciliation cell that must read 0 turns silent data errors into visible ones."
    ],
    downloads: [
      { file: "assets/files/level7/expense-tracker-project.xlsx", label: "Expense Tracker — exercise file", note: "Expense log sample and a blank summary layout" },
      { file: "assets/files/level7/expense-tracker-project-solution.xlsx", label: "Expense Tracker — completed solution", note: "Finished tracker with category and monthly summaries" }
    ],
    keywords: ["expense tracker excel", "SUMIF category total", "budget vs actual variance", "monthly expense summary", "TEXT month helper", "over budget conditional formatting", "reconciliation check", "expense log dropdown"]
  },
  {
    id: "l7-project-invoice",
    title: "Project: Invoice Generator",
    difficulty: "Intermediate",
    duration: 35,
    version: "",
    objectives: [
      "Build a reusable invoice with a header block and an item table",
      "Compute line amounts, 18% and 12% GST, and a rounded grand total",
      "Pull customer details automatically with INDEX/MATCH from a customer master",
      "Keep blank invoice lines clean instead of showing zeros and errors",
      "Set up the sheet as a print-ready one-page invoice"
    ],
    intro: "Invoices are the most repeated document in any trading business, and most offices still type them from scratch each time. In this project you build the reusable invoice of Sunrise Traders: pick a customer ID, type up to ten item lines, and the customer's name, address and GST number fill themselves from the customer master while the amounts, GST and grand total compute to the rupee — ready to print.\n\nThe invoice uses the INDEX/MATCH pattern from Level 5 and one number everyone in India knows: GST, the Goods and Services Tax added to most sales. Stationery items carry 12% GST, electronics like printers and laptops carry 18%. Everything works in Excel 2010 and later; no new functions appear, but the discipline here is different — every formula must behave well when the invoice is half-empty, because a reusable template is empty most of the time.",
    sections: [
      {
        heading: "What You Will Build",
        blocks: [
          { type: "p", text: "Three sheets. The **Customer Master** holds each buyer once — ID, name, address, city, GST number. The **Invoice** is the printable document itself: a company header, a customer block that fills from one typed ID, a ten-line item table, and totals. The **Data Entry** sheet is the office record: one row per invoice raised, with the customer's name looked up automatically so the record never has to be retyped." },
          { type: "ul", items: [
            "**Customer Master** — six customers with GSTIN-format numbers.",
            "**Invoice** — header block, customer block, item table rows 13 to 22, subtotal, GST, grand total.",
            "**Data Entry** — an invoice register with automatic customer names."
          ] },
          { type: "p", text: "By the end, raising a fresh invoice means: type the customer ID, type ten or fewer item lines, print. Everything else — names, GST, rounding, totals — is already built." }
        ]
      },
      {
        heading: "Skills You Will Use",
        blocks: [
          { type: "ul", items: [
            "**INDEX and MATCH** (Level 5) — the customer lookups.",
            "**IFERROR** (Level 5) — blank instead of #N/A while the template is empty.",
            "**ROUND** (Level 2) — every money figure in whole rupees.",
            "**SUM** (Level 2) — the totals block.",
            "**Print area and page setup** (Level 2) — the one-page print-ready layout."
          ] },
          { type: "p", text: "The genuinely new skill is template thinking: writing formulas that stay quiet in an empty row. `=IF(B13=\"\",\"\",...)` is a small pattern with a big payoff in every reusable document you will ever build." }
        ]
      },
      {
        heading: "Open the Exercise File",
        blocks: [
          { type: "p", text: "Download **Invoice Generator — exercise file**. Inside:" },
          { type: "ul", items: [
            "**Customer Master** — complete: C001 Sharma Book House (Chennai), C002 Gupta Stationers (Delhi), C003 Mehra Office Supply (Bengaluru), C004 Nair Traders (Mumbai), C005 Kaur Enterprises (Chandigarh), C006 Reddy & Sons (Hyderabad), each with a GST number.",
            "**Invoice** — company header and labels laid out, but every value cell is empty and the item table has no formulas.",
            "**Data Entry** — headers only: Invoice No, Date, Customer ID, Customer Name, Grand Total."
          ] },
          { type: "table", headers: ["Cells", "Invoice area", "You build"],
            rows: [
              ["B4 / F4", "Customer ID input / Invoice No", "Input cells"],
              ["B6–B10, D6–D10", "Customer name, address, city, GSTIN", "INDEX/MATCH formulas"],
              ["Row 12", "Item table headers: S.No, Description, Qty, Rate, Amount, GST %, GST, Total", "Headers exist"],
              ["Rows 13–22", "Ten item lines", "Five formulas per line"],
              ["Rows 23–24", "Subtotal, Total GST, Grand Total", "Three SUM formulas"]
            ] },
          { type: "tip", text: "The customer block input cell is `B4` on the Invoice sheet. Every lookup in the customer block points there with `$B$4` — typed once, locked, reused by six formulas." }
        ]
      },
      {
        heading: "Build Section 1: The Customer Block Fills Itself",
        blocks: [
          { type: "p", text: "Type an ID, and four cells write themselves. This is the same one-MATCH-many-INDEX pattern as the payslip." },
          { type: "steps", items: [
            "On **Invoice**, type `C003` into `B4` and `INV-1042` into `F4`.",
            "In `B6` (Name) type `=IFERROR(INDEX('Customer Master'!$B$2:$B$7,MATCH($B$4,'Customer Master'!$A$2:$A$7,0)),\"\")`.",
            "In `B7` (Address) reuse the same formula, changing only the INDEX range to `'Customer Master'!$C$2:$C$7`.",
            "In `B8` (City) use the INDEX range `'Customer Master'!$D$2:$D$7`.",
            "In `E8` (GSTIN) use the INDEX range `'Customer Master'!$E$2:$E$7`.",
            "Fill `B6:B8` right or retype for the Ship To block if your layout has one — the same MATCH serves every customer field.",
            "Clear `B4` and confirm all four cells go blank instead of showing #N/A. Type `C003` back."
          ] },
          { type: "grid", title: "Customer block after step 6",
            rows: [
              ["Sunrise Traders", "", "", "", "Invoice No:", "INV-1042"],
              ["12, MG Road, Bengaluru", "", "", "", "Date:", "14-Nov-26"],
              ["", "", "", "", "", ""],
              ["Customer ID:", "C003", "", "", "", ""],
              ["Name:", "Mehra Office Supply", "", "GSTIN:", "", "29AABCM9012P1Z8"],
              ["Address:", "5 MG Road", "", "City:", "", "Bengaluru"]
            ] },
          { type: "formula", name: "Customer lookup with IFERROR",
            syntax: "=IFERROR(INDEX(return_range, MATCH(lookup_value, lookup_range, 0)), \"\")",
            example: "=IFERROR(INDEX('Customer Master'!$B$2:$B$7,MATCH($B$4,'Customer Master'!$A$2:$A$7,0)),\"\")",
            explanation: "MATCH finds the position of the ID typed in B4; INDEX returns the matching field; IFERROR swallows the #N/A that appears while the template is empty or the ID is wrong. A blank is the polite result for an unfinished invoice." }
        ]
      },
      {
        heading: "Build Section 2: The Item Table",
        blocks: [
          { type: "p", text: "Ten lines, five formulas, one fill-down. Each line computes its own amount, its own GST at its own rate, and stays invisible until a description exists." },
          { type: "steps", items: [
            "In `A13` type `=IF($B13=\"\",\"\",ROW()-12)` and fill down to `A22` — automatic serial numbers that only appear on used lines.",
            "In `E13` type `=IF($B13=\"\",\"\",ROUND($C13*$D13,0))` and fill down — the line amount before tax.",
            "In `G13` type `=IF($B13=\"\",\"\",ROUND($E13*$F13,0))` and fill down — the GST for the line, at the rate typed in F.",
            "In `H13` type `=IF($B13=\"\",\"\",$E13+$G13)` and fill down — the line total including GST.",
            "Fill the first four used lines to match a real sale to Mehra Office Supply: `Notebooks` 100 at ₹45 with 12% in F; `Ball Pens (box)` 300 at ₹15 with 12%; `Staplers` 10 at ₹180 with 12%; `Printers` 2 at ₹6,800 with 18%; `Laptop` 1 at ₹43,500 with 18%.",
            "Apply thousands separators to `D13:H22` and centre `A13:A22` and `F13:F22`."
          ] },
          { type: "table", headers: ["Row", "Description", "Qty", "Rate (₹)", "Amount (₹)", "GST %", "GST (₹)", "Total (₹)"],
            rows: [
              ["13", "Notebooks", "100", "45", "4500", "12%", "540", "5040"],
              ["14", "Ball Pens (box)", "300", "15", "4500", "12%", "540", "5040"],
              ["15", "Staplers", "10", "180", "1800", "12%", "216", "2016"],
              ["16", "Printers", "2", "6800", "13600", "18%", "2448", "16048"],
              ["17", "Laptop", "1", "43500", "43500", "18%", "7830", "51330"],
              ["18–22", "(empty — formulas silent)", "", "", "", "", "", ""]
            ] },
          { type: "warning", text: "Type the GST rate as a percentage (12% or 18%) in column F, not as 12 or 18. The formula multiplies the amount by the cell's value — `12` would mean twelve times the amount, not twelve percent." }
        ]
      },
      {
        heading: "Build Section 3: Totals, GST and the Print-Ready Finish",
        blocks: [
          { type: "p", text: "Three sums close the invoice, and a page setup makes it a document instead of a worksheet." },
          { type: "steps", items: [
            "In `D23` type `Subtotal` and in `E23` type `=ROUND(SUM($E$13:$E$22),0)`.",
            "In `G23` type `Total GST` and in `H23` type `=ROUND(SUM($G$13:$G$22),0)`.",
            "In `D24` type `Grand Total` and in `H24` type `=$E$23+$H$23`. Bold both rows; give `H24` a light fill and a thick top border.",
            "Select `A1:H24`, then **Page Layout → Print Area → Set Print Area**.",
            "Still in **Page Layout**, set Orientation to Portrait and choose **Fit Sheet on One Page** under scaling.",
            "Press `[[Ctrl]] + [[P]]` and preview: one clean page, no gridlines, no stray empty-row zeros. Cancel the print."
          ] },
          { type: "grid", title: "The totals block",
            rows: [
              ["", "", "", "Subtotal", "67900", "", "Total GST", "11574"],
              ["", "", "", "Grand Total", "", "", "", "79474"]
            ] },
          { type: "tip", text: "Check the arithmetic once by hand: amounts 4,500 + 4,500 + 1,800 + 13,600 + 43,500 = ₹67,900 subtotal. GST at 540 + 540 + 216 + 2,448 + 7,830 = ₹11,574. Grand total ₹79,474. If your cells disagree with these three numbers, re-read the formulas in rows 13 to 17 before anything else." }
        ]
      },
      {
        heading: "Build Section 4: The Data Entry Record",
        blocks: [
          { type: "p", text: "The printed invoice leaves the building; the office keeps a register. This small sheet records every invoice raised and pulls the customer name on its own." },
          { type: "steps", items: [
            "On **Data Entry**, type `INV-1042` in `A2`, `14-Nov-26` in `B2`, and `C003` in `C2`.",
            "In `D2` type `=IFERROR(INDEX('Customer Master'!$B$2:$B$7,MATCH($C2,'Customer Master'!$A$2:$A$7,0)),\"\")` and fill down to row 20 — the name fills itself per row.",
            "In `E2` type the grand total ₹79474 by hand, straight off the printed invoice.",
            "Format columns and add borders; select `A1` and press `[[Ctrl]] + [[Shift]] + [[L]]` if you want the register filterable.",
            "Raise one more invoice to test the whole system: clear the Invoice sheet, type `C001`, two lines (Pens 200 at ₹15, Notebooks 50 at ₹45), confirm subtotal ₹12,750, GST ₹1,530, grand total ₹14,280 — then record it on Data Entry as INV-1043."
          ] },
          { type: "tip", text: "Note what the register deliberately does not contain: amounts recomputed from the invoice sheet. A register row is a snapshot of a document that was already printed and sent — typing the final total keeps history stable even though the invoice template gets reused and cleared." }
        ]
      },
      {
        heading: "Level Up Your System",
        blocks: [
          { type: "p", text: "The generator is complete and tested on two customers. The exercises below take it toward a full billing module." }
        ]
      },
      {
        heading: "Compare With the Solution",
        blocks: [
          { type: "steps", items: [
            "Open **Invoice Generator — completed solution** and compare the customer block with the ID `C003` in place.",
            "Click the solution's `A13`, `E13`, `G13` and `H13` and compare each formula with yours — in particular that all four are wrapped in the same `IF($B13=\"\",\"\",...)` guard.",
            "Compare the totals: subtotal ₹67,900, total GST ₹11,574, grand total ₹79,474.",
            "Clear the customer ID on the solution and confirm the whole invoice goes blank — no errors, no zeros — then re-enter `C003`.",
            "Open the solution's print preview and compare margins and scaling with yours, then close without saving."
          ] },
          { type: "p", text: "The likeliest difference in learner versions is a GST line computed from the wrong cell — GST belongs on the line's Amount (column E) times the rate (column F), never on Qty times Rate directly." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Writing the GST formula as =IF($B13=\"\",\"\",ROUND($C13*$D13*$F13,0)) — multiplying quantity, rate and the rate column together.", right: "GST = line amount × rate: =IF($B13=\"\",\"\",ROUND($E13*$F13,0))", why: "GST applies to the taxable value — the line amount — not to quantity times rate times rate again. Computing from C and D also breaks the moment a line uses a discount or a rate override." },
      { wrong: "Typing the GST rate as 12 or 18 in column F.", right: "Type 12% or 18% so the cell holds 0.12 or 0.18", why: "The formula multiplies by the cell's stored value. A plain 12 means twelve times, turning a ₹4,500 line into ₹54,000 of GST — a visible but alarming error that survives every check that only looks at totals." },
      { wrong: "Leaving the lookups unguarded: =INDEX('Customer Master'!$B$2:$B$7,MATCH($B$4,...)) with no IFERROR.", right: "=IFERROR(INDEX(...),\"\") on every customer and item formula", why: "A reusable template sits empty most of the time. Without the guard, a fresh invoice opens covered in #N/A — and an ID typo in a busy office shows raw errors on a document being printed for a customer." },
      { wrong: "Computing the grand total as =SUM($H$13:$H$22) while the subtotal and GST rows already sum their columns.", right: "=$E$23+$H$23 — reuse the two computed blocks", why: "Both are correct mathematics, but summing H double-counts nothing and adding E23+H23 makes the printed audit trail visible: subtotal, plus GST, equals grand total, each traceable on the page." }
    ],
    shortcuts: [
      { keys: "Ctrl + P", action: "Preview the printed invoice" },
      { keys: "F4", action: "Lock the MATCH and INDEX ranges while editing" },
      { keys: "Ctrl + 1", action: "Format the money columns with thousands separators" }
    ],
    exercises: [
      { task: "Add a line under the grand total that reads the amount with the rupee symbol, like ₹ 79,474 only.", hint: "Concatenate a symbol with TEXT(number, \"#,##0\").", solution: "=\"₹ \"&TEXT($H$24,\"#,##0\")&\" only\"" },
      { task: "Make the invoice number advance automatically: when Data Entry has n−1 rows, the Invoice sheet proposes INV-10 followed by the next number.", hint: "COUNTA over the register's invoice column, then build the text.", solution: "=\"INV-10\"&(COUNTA('Data Entry'!$A$2:$A$20)+1)" },
      { task: "Add a CGST/SGST split under the totals: two rows showing half the GST each.", hint: "Total GST lives in H23; each half is that value divided by 2.", solution: "CGST: =ROUND($H$23/2,0) and SGST: =ROUND($H$23/2,0)" }
    ],
    quiz: [
      { q: "The line Notebooks 100 × ₹45 with 12% GST — what does column G (GST) show?",
        options: [
          "₹54,000 — 100 × 45 × 12",
          "₹540 — the line amount ₹4,500 × 12%",
          "₹45 — the rate × 12%",
          "₹5,040 — the full line total"
        ],
        answer: 1,
        explanation: "GST applies to the taxable value: the line amount 4,500 × 0.12 = 540. The ₹5,040 figure is the line total including GST, which lives in column H." },
      { q: "Why is every item-table formula wrapped as =IF($B13=\"\",\"\",...)?",
        options: [
          "To make the sheet calculate faster",
          "So unused lines show nothing — no zeros, no #N/A — and the invoice prints clean at any size",
          "Because IFERROR cannot be used twice in one sheet",
          "To stop anyone editing the formulas"
        ],
        answer: 1,
        explanation: "A reusable invoice is empty until filled. The guard makes a half-used template look like a finished short invoice rather than a wall of zeros and errors." },
      { q: "Which formula correctly pulls the customer's GSTIN into E8?",
        options: [
          "=VLOOKUP($B$4,'Customer Master'!$E$2:$E$7,5,FALSE)",
          "=IFERROR(INDEX('Customer Master'!$E$2:$E$7,MATCH($B$4,'Customer Master'!$A$2:$A$7,0)),\"\")",
          "=INDEX('Customer Master'!$A$2:$A$7,MATCH($B$4,'Customer Master'!$E$2:$E$7,0))",
          "=MATCH('Customer Master'!$E$2:$E$7,$B$4,0)"
        ],
        answer: 1,
        explanation: "MATCH finds the ID's position in the master's code column; INDEX returns the value at that position from the GSTIN column; IFERROR keeps the cell blank while the ID is empty. Option 3 swaps the two ranges and returns IDs instead of GSTINs." },
      { q: "An invoice has subtotal ₹67,900 and total GST ₹11,574. What does the grand total formula show?",
        options: [
          "₹79,474",
          "₹56,326",
          "₹67,900",
          "₹11,574"
        ],
        answer: 0,
        explanation: "Grand total = subtotal + GST = 67,900 + 11,574 = ₹79,474, rounded to whole rupees as every printed document should be." },
      { q: "What is the Data Entry sheet's role in the system?",
        options: [
          "It computes the GST rates for the invoice",
          "It stores one row per invoice raised, with the customer name looked up automatically",
          "It backs up the Customer Master",
          "It prints the invoice"
        ],
        answer: 1,
        explanation: "The invoice template is reused and cleared, so history cannot live there. The register keeps one immutable row per invoice — number, date, ID, auto-filled name, final total." }
    ],
    summary: [
      "A reusable invoice is one input cell, one lookup pattern reused four times, and ten guarded item lines.",
      "Guard every formula with IF(cell=\"\",\"\",...) so an empty template stays empty instead of showing errors.",
      "GST is charged on the line amount at the line's own rate — 12% for stationery, 18% for electronics.",
      "Totals build visibly: subtotal, total GST, grand total — each its own traceable cell.",
      "The Data Entry register keeps one row per invoice; the invoice template itself is never a record."
    ],
    downloads: [
      { file: "assets/files/level7/invoice-generator.xlsx", label: "Invoice Generator — exercise file", note: "Customer master and a blank invoice layout to complete" },
      { file: "assets/files/level7/invoice-generator-solution.xlsx", label: "Invoice Generator — completed solution", note: "Working invoice that auto-fills customer details and totals" }
    ],
    keywords: ["invoice generator excel", "GST invoice formula", "INDEX MATCH customer lookup", "IFERROR blank template", "print area invoice", "CGST SGST split", "invoice item table", "grand total ROUND", "customer master"]
  },
  {
    id: "l7-project-purchase",
    title: "Project: Purchase Management System",
    difficulty: "Intermediate",
    duration: 40,
    version: "",
    objectives: [
      "Build a PO register with supplier, status and payment tracking in one row",
      "Pull supplier details from a supplier master with VLOOKUP",
      "Track order status with validated Open / Received / Cancelled values",
      "Compute the balance due per PO with an IF that handles cancelled orders",
      "Summarise pending POs and supplier-wise dues with COUNTIF and SUMIFS"
    ],
    intro: "Buying is half of any trading business, and it is easy to lose track of: orders that were placed but never arrived, goods received but never paid for, suppliers owed money in three different places. In this project you build the purchase system for Sunrise Traders — a register where each purchase order is one row carrying its supplier, its status and its payment position, plus a summary that answers the two questions purchase managers live by: what is still open, and who do we owe?\n\nThe tools are familiar — validation, VLOOKUP, COUNTIF, SUMIFS, IF — but this project teaches something new in how they are arranged: three pieces of state (order status, payment status, amount paid) living in one row, with formulas that respect the business logic, such as a cancelled order never owing money. Everything works in Excel 2010 and later.",
    sections: [
      {
        heading: "What You Will Build",
        blocks: [
          { type: "p", text: "Three sheets. The **Supplier Master** holds each supplier once — ID, name, city, GST number and payment terms. The **PO Register** is one row per purchase order: who, what, how much, whether it has arrived, and how much of it has been paid. The **Summary** sheet counts the register by status, totals the open value, and builds a supplier-wise ledger of orders, payments and balance due." },
          { type: "ul", items: [
            "**Supplier Master** — five suppliers with payment terms of 15, 30 or 45 days.",
            "**PO Register** — six November 2026 orders, with dropdowns on supplier and status.",
            "**Summary** — status block, pending-PO figures, and a supplier-wise dues table."
          ] },
          { type: "p", text: "The three status ideas, kept simple: **Open** means ordered but not yet received. **Received** means goods are in. **Cancelled** means the order died — and a cancelled order owes nothing and receives nothing, a rule the balance formula will enforce. Payment status runs Unpaid, Part Paid, Paid, with the amount actually paid recorded alongside." }
        ]
      },
      {
        heading: "Skills You Will Use",
        blocks: [
          { type: "ul", items: [
            "**Data validation** (Level 3) — supplier dropdown and status dropdown.",
            "**VLOOKUP** (Level 5) — supplier names from the master.",
            "**COUNTIF and SUMIF** (Level 4) — the status block.",
            "**SUMIFS** (Level 4) — supplier-wise dues with two conditions.",
            "**IF** (Level 4) — the balance-due rule for cancelled orders.",
            "**Conditional formatting** (Level 6) — status colours and overdue money."
          ] },
          { type: "p", text: "Nothing new to learn, one new way to think: this register is a small database. Each row is self-contained — any question about a PO is answered by formulas reading that row plus the master." }
        ]
      },
      {
        heading: "Open the Exercise File",
        blocks: [
          { type: "p", text: "Download **Purchase Management System — exercise file**. Inside:" },
          { type: "ul", items: [
            "**Supplier Master** — complete: S001 Kaviraj Papers Ltd. (Rajkot), S002 Metro Plastics Co. (Delhi), S003 TechnoSource India (Pune), S004 Balaji Traders (Chennai), S005 WinEdge Solutions (Noida), with GST numbers and payment terms.",
            "**PO Register** — six PO rows filled for date, supplier ID, item, quantity and rate; the Name, Status and payment columns are blank or empty formulas.",
            "**Summary** — empty except a title."
          ] },
          { type: "table", headers: ["Column", "PO Register column", "Example"],
            rows: [
              ["A", "PO No", "PO-1001"],
              ["B", "Date", "01-Nov-26"],
              ["C", "Sup ID (dropdown)", "S001"],
              ["D", "Supplier Name (formula)", "Kaviraj Papers Ltd."],
              ["E", "Item", "A4 Paper Reams"],
              ["F", "Qty", "200"],
              ["G", "Rate (₹)", "95"],
              ["H", "PO Value (₹) = Qty × Rate", "19000"],
              ["I", "Status (dropdown)", "Received"],
              ["J", "Received Date", "05-Nov-26"],
              ["K", "Payment Status (dropdown)", "Paid"],
              ["L", "Amount Paid (₹)", "19000"],
              ["M", "Balance Due (₹) (formula)", "0"]
            ] },
          { type: "tip", text: "The six sample orders tell a complete story: two received (one fully paid, one part-paid), two open and unpaid, one received but unpaid, and one cancelled. Build the formulas against all six and every business situation is already tested." }
        ]
      },
      {
        heading: "Build Section 1: Complete the Register Rows",
        blocks: [
          { type: "p", text: "Names, values and the balance rule — four formulas written once in row 2 and filled down." },
          { type: "steps", items: [
            "On **PO Register**, in `D2` type `=IF($C2=\"\",\"\",VLOOKUP($C2,'Supplier Master'!$A$2:$B$6,2,FALSE))` and fill down to row 30 — the supplier name appears as soon as an ID is picked.",
            "In `H2` type `=IF($C2=\"\",\"\",ROUND($F2*$G2,0))` and fill down — the order value.",
            "Select `C2:C30` and apply **Data Validation → List** with source `'Supplier Master'!$A$2:$A$6`; select `I2:I30` with source `Open,Received,Cancelled`; select `K2:K30` with source `Unpaid,Part Paid,Paid`.",
            "Complete the sample data so rows 2 to 7 read: PO-1001 Received 05-Nov-26, Paid, 19000; PO-1002 Received 08-Nov-26, Part Paid, 20000; PO-1003 Open, Unpaid, 0; PO-1004 Open, Unpaid, 0; PO-1005 Received 14-Nov-26, Unpaid, 0; PO-1006 Cancelled with blank dates, Unpaid, 0.",
            "In `M2` type `=IF($I2=\"Cancelled\",0,$H2-$L2)` and fill down — the balance due.",
            "Apply thousands separators to `G2:H30` and `L2:M30`, and bold the header row."
          ] },
          { type: "formula", name: "Balance due with the cancellation rule",
            syntax: "=IF(status=\"Cancelled\", 0, PO_value - amount_paid)",
            example: "=IF($I2=\"Cancelled\",0,$H2-$L2)",
            explanation: "A cancelled order has no liability, so its balance is forced to 0 no matter what the arithmetic says. Business rules belong inside formulas exactly like this — otherwise the summary quietly counts money nobody will ever pay or receive." },
          { type: "warning", text: "Check PO-1006 in row 7: value ₹12,000, cancelled, paid ₹0 — balance must read 0, not 12,000. If it reads 12,000, the IF is comparing the wrong cell or the status cell holds a variant like \"cancelled \" with a space." }
        ]
      },
      {
        heading: "Build Section 2: Status Colours That Tell the Story",
        blocks: [
          { type: "p", text: "A register of fourteen columns is scanned, not read. Colour carries the meaning at a glance." },
          { type: "steps", items: [
            "Select `I2:I30` and add three **Highlight Cells Rules → Equal To** rules: `Open` in amber, `Received` in green, `Cancelled` in grey.",
            "Select `K2:K30` and add two rules: `Unpaid` in light red, `Paid` in green (Part Paid stays neutral).",
            "Select `M2:M30` and add **Highlight Cells Rules → Greater Than** with value `0` in light red — any money owed stands out immediately.",
            "Scan the register: PO-1003, PO-1004 and PO-1005 should each glow red in column M with balances ₹30,000, ₹18,000 and ₹51,000."
          ] },
          { type: "tip", text: "Notice what you did NOT format: the item, the date, the supplier. Formatting earns its place only where a decision is made — status, payment, money owed. Decorative colour is noise that a real register cannot afford." }
        ]
      },
      {
        heading: "Build Section 3: The Summary Sheet",
        blocks: [
          { type: "p", text: "Two blocks: how the orders stand, and what each supplier is owed. Both read the register; neither is ever typed." },
          { type: "steps", items: [
            "On **Summary**, build the status block: headers `Status`, `Count`, `Value (₹)` in `A3:C3`; type `Open`, `Received`, `Cancelled` in `A4:A6`.",
            "In `B4` type `=COUNTIF('PO Register'!$I$2:$I$30,$A4)` and fill down; in `C4` type `=SUMIF('PO Register'!$I$2:$I$30,$A4,'PO Register'!$H$2:$H$30)` and fill down.",
            "In `A8` type `Total Pending Value (Open POs)` and in `C8` type `=SUMIF('PO Register'!$I$2:$I$30,\"Open\",'PO Register'!$H$2:$H$30)`.",
            "Build the supplier ledger from row 11: headers in `A11:G11` — Sup ID, Supplier, POs, Total Value, Open Value, Paid, Balance Due.",
            "Type `S001` to `S005` in `A12:A16`; in `B12` type `=IF($A12=\"\",\"\",VLOOKUP($A12,'Supplier Master'!$A$2:$B$6,2,FALSE))` and fill down.",
            "In `C12` type `=COUNTIF('PO Register'!$C$2:$C$30,$A12)`; in `D12` `=SUMIF('PO Register'!$C$2:$C$30,$A12,'PO Register'!$H$2:$H$30)`; in `E12` `=SUMIFS('PO Register'!$H$2:$H$30,'PO Register'!$C$2:$C$30,$A12,'PO Register'!$I$2:$I$30,\"Open\")`; in `F12` `=SUMIF('PO Register'!$C$2:$C$30,$A12,'PO Register'!$L$2:$L$30)`; in `G12` `=SUMIF('PO Register'!$C$2:$C$30,$A12,'PO Register'!$M$2:$M$30)` — then fill `C12:G16` down."
          ] },
          { type: "grid", title: "Supplier-wise ledger — what you should see",
            rows: [
              ["Sup ID", "Supplier", "POs", "Total Value (₹)", "Open Value (₹)", "Paid (₹)", "Balance Due (₹)"],
              ["S001", "Kaviraj Papers Ltd.", "2", "31000", "0", "19000", "0"],
              ["S002", "Metro Plastics Co.", "1", "30000", "30000", "0", "30000"],
              ["S003", "TechnoSource India", "1", "48000", "0", "20000", "28000"],
              ["S004", "Balaji Traders", "1", "18000", "18000", "0", "18000"],
              ["S005", "WinEdge Solutions", "1", "51000", "0", "0", "51000"]
            ] },
          { type: "formula", name: "Open value per supplier with SUMIFS",
            syntax: "=SUMIFS(sum_range, criteria_range1, criteria1, criteria_range2, criteria2)",
            example: "=SUMIFS('PO Register'!$H$2:$H$30,'PO Register'!$C$2:$C$30,$A12,'PO Register'!$I$2:$I$30,\"Open\")",
            explanation: "Adds order values where two conditions both hold: the row belongs to this supplier AND its status is Open. Kaviraj Papers shows 0 because both its orders are Received or Cancelled — the money question changes with status, and SUMIFS is how you ask it." },
          { type: "p", text: "Trace one line to prove the sheet: TechnoSource India has one order, ₹48,000, received, of which ₹20,000 is paid — so open value 0 and balance due ₹28,000. If your ledger says otherwise, the mismatch is in column I, K or L of that PO's register row, not in the summary formula." }
        ]
      },
      {
        heading: "Level Up Your System",
        blocks: [
          { type: "p", text: "The purchase system is complete. The exercises below add the pieces a purchase head would request after the first review meeting." }
        ]
      },
      {
        heading: "Compare With the Solution",
        blocks: [
          { type: "steps", items: [
            "Open **Purchase Management System — completed solution** and compare the register row by row, starting with the balance column.",
            "Click the solution's `M7` (the cancelled PO) and confirm the IF reads the status cell `I7`, not the payment cell.",
            "Open **Data → Data Validation** on all three dropdown columns and compare the sources with yours.",
            "Compare the status block: Open should show 2 orders worth ₹48,000, and the supplier ledger should match the grid shown in this lesson.",
            "Change PO-1003's status to Received on the solution copy and watch its open value move from the supplier ledger, then close without saving."
          ] },
          { type: "p", text: "The most common learner error lives in the supplier ledger's SUMIF columns: the criteria range must be the register's supplier column `C`, while the sum ranges are `H`, `L` and `M` depending on the question. Mixing those up produces ledger rows that look plausible and owe the wrong amounts." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Computing balance due as =$H2-$L2 with no status check.", right: "=IF($I2=\"Cancelled\",0,$H2-$L2)", why: "Arithmetic does not know the PO was cancelled. Without the rule, the register and every supplier ledger shows a ₹12,000 liability that will never exist — the kind of figure that wrecks a payment run." },
      { wrong: "Using SUMIF with two conditions: =SUMIF(register C:C, supplier, register H:H, status=\"Open\").", right: "Use SUMIFS: =SUMIFS(H range, C range, supplier, I range, \"Open\")", why: "SUMIF takes exactly one criteria pair. The moment a question needs two — this supplier AND this status — the function changes to SUMIFS, whose argument order (sum range first) also changes." },
      { wrong: "Typing status values by hand: \"open\", \"OPEN \", \"Recd\".", right: "Pick from the dropdown: Open, Received, Cancelled", why: "COUNTIF and the colour rules match exact text. Variants silently drop out of the status block, so the pending count shrinks while the register looks full." },
      { wrong: "Leaving the Received Date filled on a PO whose status is still Open.", right: "Clear the received date until the status says Received — the columns must agree", why: "Status is the master fact and the date is its evidence. A row claiming Received evidence with Open status makes the register untrustworthy, and no formula can repair a contradiction." }
    ],
    shortcuts: [
      { keys: "Alt + Down", action: "Open the supplier or status dropdown in a register row" },
      { keys: "F4", action: "Lock a lookup range while writing the formula" },
      { keys: "Ctrl + 1", action: "Format the money columns" }
    ],
    exercises: [
      { task: "Add a cell showing how many POs are still open and unpaid for money.", hint: "Two conditions again — status Open and payment status Unpaid — so COUNTIFS.", solution: "=COUNTIFS('PO Register'!$I$2:$I$30,\"Open\",'PO Register'!$K$2:$K$30,\"Unpaid\")" },
      { task: "Add a Total Balance Due line under the supplier ledger.", hint: "SUM over the balance column of the ledger.", solution: "=SUM($G$12:$G$16)" },
      { task: "Flag suppliers with any balance due: in column H of the ledger, show \"Pay Now\" or \"Clear\".", hint: "One IF reading the balance cell in column G.", solution: "=IF($G12>0,\"Pay Now\",\"Clear\")" }
    ],
    quiz: [
      { q: "PO-1006 is cancelled with a value of ₹12,000 and ₹0 paid. What does the balance formula show?",
        options: [
          "₹12,000",
          "₹0 — the IF forces cancelled orders to a zero balance",
          "#N/A because there is no received date",
          "-₹12,000"
        ],
        answer: 1,
        explanation: "The business rule is inside the formula: =IF($I2=\"Cancelled\",0,$H2-$L2). A cancelled order has no liability regardless of the arithmetic behind it." },
      { q: "Which formula correctly totals the open order value for supplier S004 in the ledger?",
        options: [
          "=SUMIFS('PO Register'!$H$2:$H$30,'PO Register'!$C$2:$C$30,$A15,'PO Register'!$I$2:$I$30,\"Open\")",
          "=SUMIF('PO Register'!$C$2:$C$30,$A15,'PO Register'!$H$2:$H$30,\"Open\")",
          "=SUMIFS('PO Register'!$C$2:$C$30,$A15,'PO Register'!$H$2:$H$30,\"Open\")",
          "=SUMIF('PO Register'!$I$2:$I$30,\"Open\",'PO Register'!$H$2:$H$30)"
        ],
        answer: 0,
        explanation: "Two conditions — this supplier AND status Open — require SUMIFS, with the sum range (order values) first, then each criteria range with its criteria. Option 4 sums every supplier's open orders, not one supplier's." },
      { q: "The status block shows Open with a count of 2 and a value of ₹48,000. Which two POs make that up?",
        options: [
          "PO-1002 and PO-1005",
          "PO-1003 and PO-1004 — ₹30,000 plus ₹18,000",
          "PO-1003 and PO-1006",
          "PO-1001 and PO-1003"
        ],
        answer: 1,
        explanation: "Open means ordered but not received: PO-1003 (Ball Pen Refills, ₹30,000) and PO-1004 (Stapler Bodies, ₹18,000). Every other order is Received or Cancelled." },
      { q: "Why does the supplier ledger count Kaviraj Papers' cancelled PO in the POs column but not in the Open Value column?",
        options: [
          "Because COUNTIF counts every matching row while SUMIFS adds only rows meeting both conditions",
          "Because cancelled POs are hidden from the register",
          "Because the cancelled PO has no value",
          "Because SUMIFS ignores the supplier criteria"
        ],
        answer: 0,
        explanation: "The count asks one question (whose rows are these?) so both register rows match. The open-value SUMIFS asks two (theirs AND status Open), and the cancelled row fails the second." },
      { q: "A team member types \"recieved\" into the Status column. What happens?",
        options: [
          "Excel corrects the spelling automatically",
          "The row is accepted but vanishes from every count and colour rule that matches \"Received\"",
          "The balance formula shows #N/A",
          "The COUNTIF returns an error message"
        ],
        answer: 1,
        explanation: "Text matching is exact and silent. The dropdown validation exists precisely so this cannot happen — and it is why status columns should never be free-typed." }
    ],
    summary: [
      "One register row carries the whole story of an order: supplier, value, status, payment, balance.",
      "Business rules live inside formulas — the IF that zeroes cancelled orders is the clearest example.",
      "One condition means SUMIF or COUNTIF; two conditions mean SUMIFS or COUNTIFS, with the sum range first.",
      "Colour only where a decision happens: status, payment, money owed.",
      "The supplier ledger is five SUMIF/SUMIFS questions per supplier, all reading the register — never typed."
    ],
    downloads: [
      { file: "assets/files/level7/purchase-management-system.xlsx", label: "Purchase Management System — exercise file", note: "PO register sample and blank summary sheets" },
      { file: "assets/files/level7/purchase-management-system-solution.xlsx", label: "Purchase Management System — completed solution", note: "Finished system with status tracking and summaries" }
    ],
    keywords: ["purchase order register excel", "PO status tracking", "supplier master VLOOKUP", "SUMIFS supplier ledger", "balance due formula", "pending PO summary", "payment status tracking", "COUNTIFS open orders", "conditional formatting status"]
  },
  {
    id: "l7-project-dashboard",
    title: "Project: Interactive Sales Dashboard",
    difficulty: "Intermediate",
    duration: 40,
    version: "",
    objectives: [
      "Design a dashboard with KPI cards that compute themselves",
      "Drive the whole screen from one month-selector dropdown",
      "Build month-specific cards with two-pair SUMIFS",
      "Add data bars and a column chart so numbers read at a glance",
      "Test the dashboard by changing one cell and watching everything move"
    ],
    intro: "The sales tracker from Project 5 answers questions one at a time — you read a summary, then a performance sheet, then a region block. A dashboard answers them all at once, on one screen, and updates when you change one cell.\n\nIn this project you build that screen for Sunrise Traders' sales meeting. One dropdown selects the month, and every card on the sheet — revenue, share of quarter, top performer, region split — recomputes for that month. The data underneath is the same Sales Log from Project 5: six salespeople, four regions, five products, three months of invoices.\n\nEverything here works in Excel 2010 and later. The dashboard introduces no new functions; it is an arrangement discipline — one input cell, locked ranges, and formulas that read the log through generous headroom ranges.",
    sections: [
      {
        heading: "What You Will Build",
        blocks: [
          { type: "p", text: "One new sheet, added to the Project 5 workbook. The **Dashboard** has four zones: a title row with the month selector, a KPI card row, a **Month Focus** block that changes with the selector, and a static block that always shows the full quarter." },
          { type: "ul", items: [
            "**Title and selector** — the month dropdown lives in one cell, `B3`, validated to `Aug-26`, `Sep-26`, `Oct-26`.",
            "**KPI cards** — Total Revenue, Team Achievement, Best Month, Top Performer. These describe the whole quarter and never change with the selector.",
            "**Month Focus** — revenue for the selected month, its share of the quarter, growth on the previous month, and product-wise revenue for that month.",
            "**Quarter view** — the six-person leaderboard with data bars, region totals, and the monthly revenue column chart from Project 5."
          ] },
          { type: "p", text: "The underlying numbers this dashboard shows, so you can check yourself at every step: monthly revenue is 3,20,000 (Aug-26), 3,45,000 (Sep-26) and 3,98,500 (Oct-26) — 10,63,500 in total, against a team target of 10,10,000. The exercise file's Sales Log is built so every SUMIFS you write lands on exactly these figures." }
        ]
      },
      {
        heading: "Skills You Will Use",
        blocks: [
          { type: "ul", items: [
            "**Data Validation** (Level 3) — the month dropdown that drives the whole sheet.",
            "**SUMIFS with two criteria pairs** (Level 4) — month AND region, month AND product.",
            "**INDEX/MATCH** (Level 5) — the Best Month and Top Performer cards.",
            "**COUNTIF on a computed column** (Level 4) — the 90%-club counter.",
            "**Data bars and a column chart** (Level 6) — the visual layer."
          ] },
          { type: "p", text: "The one new *idea* — not a function — is the input cell. A dashboard is a machine with exactly one lever, and every formula on the sheet points at that lever's cell. Two levers sound friendlier and behave twice as badly: half the cards answer one month, half answer the other, and nobody trusts the screen again." }
        ]
      },
      {
        heading: "Open the Exercise File",
        blocks: [
          { type: "p", text: "Download **Interactive Sales Dashboard — exercise file**. It is the completed Project 5 workbook plus an empty Dashboard sheet, so your tracker and your dashboard share one log." },
          { type: "ul", items: [
            "**Sales Log** — filled as in Project 5, with the Month helper in column I.",
            "**Monthly Summary** and **Performance** — already complete, exactly as Project 5's solution.",
            "**Dashboard** — empty except for a title in `A1`. You build everything."
          ] },
          { type: "table", headers: ["Card", "Reads", "Shows"], rows: [
            ["Total Revenue", "Monthly Summary B7", "₹10,63,500"],
            ["Team Achievement", "Performance totals", "105%"],
            ["Best Month", "INDEX/MATCH on summary", "Oct-26"],
            ["Top Performer", "INDEX/MATCH on rank 1", "Divya Menon"]
          ] },
          { type: "tip", text: "Keep the Project 5 sheets untouched. A dashboard that edits its own sources is not a dashboard — it is a second copy of the data, and the two will drift." }
        ]
      },
      {
        heading: "Build Section 1: The Selector and the KPI Cards",
        blocks: [
          { type: "steps", items: [
            "On **Dashboard**, type `Sunrise Traders — Sales Dashboard` in `A1`, bold, size 14.",
            "In `A3` type `Month:` and select `B3`. Data, Data Validation, Allow: List, Source: `Aug-26,Sep-26,Oct-26`. Pick `Oct-26` to start.",
            "KPI card row — headers in `A5:D5`: `Total Revenue`, `Team Achievement`, `Best Month`, `Top Performer`.",
            "In `A6` type `=Monthly Summary!$B$7`. In `B6` type `=Performance!$C$9/Performance!$B$9` and format as percentage with 0 decimals.",
            "In `C6` type `=INDEX('Monthly Summary'!$A$2:$A$4,MATCH(MAX('Monthly Summary'!$B$2:$B$4),'Monthly Summary'!$B$2:$B$4,0))`.",
            "In `D6` type `=INDEX(Performance!$A$2:$A$7,MATCH(1,Performance!$E$2:$E$7,0))`.",
            "Give the four cards a light fill and a thick outside border so they read as cards, not cells."
          ] },
          { type: "formula", name: "Top Performer card",
            syntax: "=INDEX(names, MATCH(1, rank_range, 0))",
            example: "=INDEX(Performance!$A$2:$A$7,MATCH(1,Performance!$E$2:$E$7,0))",
            explanation: "MATCH finds where rank equals 1 inside the locked rank column, and INDEX returns the name at that position. When Divya Menon's 118% achievement puts her at rank 1, the card shows her name — and if next quarter's log puts someone else at rank 1, the card follows without being touched." },
          { type: "p", text: "The quarter's numbers, for self-checking: achievement card 105% (10,63,500 actual against 10,10,000 target), best month Oct-26 at 3,98,500, top performer Divya Menon. If any card disagrees, the error is in the card's formula, not the data." }
        ]
      },
      {
        heading: "Build Section 2: The Month Focus Block",
        blocks: [
          { type: "p", text: "This block answers the meeting's real question — how is the selected month doing? — and it is the part that moves when the dropdown moves." },
          { type: "steps", items: [
            "Headers in `A8:C8`: `Measure`, `Value`, `Against Quarter`.",
            "In `A9` type `Revenue — selected month`. In `B9` type `=SUMIFS('Sales Log'!$H$2:$H$200,'Sales Log'!$I$2:$I$200,$B$3)`. In `C9` type `=$B$9/$A$6` and format as percentage.",
            "In `A10` type `Growth on previous month`. In `B10` type `=IF(MATCH($B$3,'Monthly Summary'!$A$2:$A$4,0)=1,\"—\",INDEX('Monthly Summary'!$C$2:$C$4,MATCH($B$3,'Monthly Summary'!$A$2:$A$4,0)))` and format as percentage.",
            "Product block from row 12: headers `Product`, `Selected Month` in `A12:B12`; the five product names below.",
            "In `B13` type `=SUMIFS('Sales Log'!$H$2:$H$200,'Sales Log'!$D$2:$D$200,$A13,'Sales Log'!$I$2:$I$200,$B$3)` and fill down through the five products."
          ] },
          { type: "formula", name: "Two-pair SUMIFS",
            syntax: "=SUMIFS(sum_range, range1, criteria1, range2, criteria2)",
            example: "=SUMIFS('Sales Log'!$H$2:$H$200,'Sales Log'!$D$2:$D$200,$A13,'Sales Log'!$I$2:$I$200,$B$3)",
            explanation: "The sum range comes first; every criteria pair after it narrows the rows. Here a row counts only if its Product matches the block AND its Month helper matches the selector — that is how one formula serves five product rows and three selectable months at once." },
          { type: "p", text: "With `Oct-26` selected you should read: revenue 3,98,500, which is 37.5% of the quarter; growth 15.5% on September; and the product block led by Laptops. Click the dropdown, choose `Aug-26`, and the block flips to 3,20,000, no growth line (there is no July), and a different product leader. Nothing else on the sheet moved — that is the design working." }
        ]
      },
      {
        heading: "Build Section 3: The Quarter View and the Visual Layer",
        blocks: [
          { type: "steps", items: [
            "Leaderboard block from row 20: in `A20` type `=Performance!A2` and fill across to `F20` and down to row 25, so the block mirrors the Performance sheet by reference.",
            "Select the Actual column of the mirrored block and apply **Conditional Formatting, Data Bars** — Sneha Iyer's 2,66,000 gets the longest bar, Rahul Verma's 1,14,000 the shortest.",
            "Region block from row 28: `North`, `South`, `East`, `West` in `A29:A32`, and in `B29` type `=SUMIFS('Sales Log'!$H$2:$H$200,'Sales Log'!$C$2:$C$200,$A29)` and fill down.",
            "Copy the monthly revenue chart from Monthly Summary onto the Dashboard, or insert a fresh one from `Monthly Summary` `A2:B4`, and place it beside the region block.",
            "In `A34` type `90%-club` and in `B34` type `=COUNTIF(Performance!$D$2:$D$7,\">=90%\")` — all six qualify this quarter."
          ] },
          { type: "grid", title: "Region totals — what you should see", rows: [
            ["Region", "Revenue (₹)"],
            ["North", "294000"],
            ["South", "267500"],
            ["East", "258000"],
            ["West", "244000"]
          ] },
          { type: "p", text: "The four regions add back to 10,63,500 — the same total the KPI card shows. North leads and West trails, and the gap between them (50,000) is roughly one month's worth of Stationery Kits. That is the kind of sentence a dashboard exists to produce." },
          { type: "tip", text: "The mirrored leaderboard uses formulas like `=Performance!A2`, not retyped names. When next quarter's targets change, the dashboard's leaderboard follows the Performance sheet with zero edits — that is the whole difference between a dashboard and a screenshot." }
        ]
      },
      {
        heading: "Test Against the Solution",
        blocks: [
          { type: "steps", items: [
            "Open **Interactive Sales Dashboard — completed solution** and put the two workbooks side by side.",
            "Compare all four KPI cards: 10,63,500 / 105% / Oct-26 / Divya Menon.",
            "Select `Sep-26` on both selectors and compare the Month Focus block: 3,45,000, growth 7.8%, share 32.4%.",
            "Check the dropdown rejects a typed value like `Nov-26` — validation should refuse it.",
            "Change one log row's amount by 100 on your copy, watch every card move, then press Ctrl + Z. The cards must all return."
          ] },
          { type: "p", text: "The most common learner gap is a Month Focus card that reads the whole log instead of the selector — the revenue figure stays at 10,63,500 for every month. Trace the formula: it must test column I (the Month helper) against `$B$3`, not against a typed month." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Putting the selected-month revenue at 10,63,500 for every month because the card sums the whole log.", right: "Test the Month helper column against the selector: =SUMIFS(...,'Sales Log'!$I$2:$I$200,$B$3)", why: "Without the month criteria the card ignores the dropdown entirely — the dashboard looks interactive but reports the same number forever." },
      { wrong: "Typing `Aug-2026` as a dropdown option while the log's helper produces `Aug-26`.", right: "Match the helper's TEXT format exactly: mmm-yy produces Aug-26", why: "SUMIFS compares text letter for letter. A mismatched option returns 0 with no error, and the card looks simply broken." },
      { wrong: "Building the leaderboard by retyping the six names and figures onto the Dashboard.", right: "Mirror with formulas: =Performance!A2 filled across and down", why: "Retyped blocks are frozen screenshots. When targets change next quarter the dashboard keeps flattering last quarter's numbers." },
      { wrong: "Adding a second selector — a region dropdown — next to the month dropdown.", right: "One lever per dashboard until the audience actually asks for two; then give the new lever its own clearly labelled block.", why: "With two selectors and mixed cards, some numbers answer the month lever and some the region lever, and nobody can tell which is which." }
    ],
    shortcuts: [
      { keys: "Alt + Down Arrow", action: "Open the month dropdown without the mouse" },
      { keys: "F9", action: "Recalculate after editing the log to see every card move" },
      { keys: "Ctrl + 1", action: "Format the card values and percentage cells" },
      { keys: "Ctrl + Z", action: "Undo a test edit on the log — the cards must return to their values" }
    ],
    exercises: [
      { task: "Build the KPI card row and confirm all four cards against the quarter's numbers.", hint: "The cards read Monthly Summary and Performance; only the Top Performer card needs INDEX/MATCH on rank 1.", solution: "Total Revenue 10,63,500; Team Achievement 105%; Best Month Oct-26; Top Performer Divya Menon." },
      { task: "Complete the Month Focus block and verify it for all three months.", hint: "Revenue card =SUMIFS on the Month helper vs $B$3; the growth card must show — for Aug-26.", solution: "Aug-26: 3,20,000, growth —, share 30.1%. Sep-26: 3,45,000, growth 7.8%, share 32.4%. Oct-26: 3,98,500, growth 15.5%, share 37.5%." },
      { task: "Add a month-AND-region card: revenue of the selected month in the North region.", hint: "Two criteria pairs: Region column C against a North cell, Month helper column I against $B$3.", solution: "=SUMIFS('Sales Log'!$H$2:$H$200,'Sales Log'!$C$2:$C$200,\"North\",'Sales Log'!$I$2:$I$200,$B$3)" }
    ],
    quiz: [
      { q: "Why should a dashboard have exactly one input cell?",
        options: [
          "Because Data Validation allows only one per sheet",
          "So every card answers the same question and the screen can never contradict itself",
          "Because formulas are slower with two inputs",
          "So the dashboard fits one printed page"
        ],
        answer: 1,
        explanation: "One lever means one interpretation. Two selectors split the sheet's answer in half, and the audience cannot tell which cards follow which lever." },
      { q: "The Top Performer card reads =INDEX(Performance!$A$2:$A$7,MATCH(1,Performance!$E$2:$E$7,0)). What happens next quarter when someone else ranks first?",
        options: [
          "The card must be retyped",
          "The card updates automatically because it reads the rank column, not a name",
          "The card shows #N/A until edited",
          "The card keeps showing Divya Menon forever"
        ],
        answer: 1,
        explanation: "The card follows rank 1 wherever it goes. That is why dashboards read summary sheets by formula — they survive the data changing underneath them." },
      { q: "Selected month is Oct-26 (3,98,500) and the quarter total is 10,63,500. What share does the card show?",
        options: ["37.5%", "105.3%", "15.5%", "26.7%"],
        answer: 0,
        explanation: "398500/1063500 = 37.5%. The share divides the selected month by the quarter total — never by the target." },
      { q: "A product block row shows 0 for Notebooks while the log clearly holds notebook sales. What do you check first?",
        options: [
          "Whether Excel needs a restart",
          "Letter-by-letter match between the block's product name and the log's Product column",
          "The chart's data range",
          "The number format of the block"
        ],
        answer: 1,
        explanation: "SUMIFS is silent when text does not match exactly — a trailing space or a singular/plural difference returns 0 with no error anywhere." },
      { q: "Why does the leaderboard block mirror the Performance sheet with =Performance!A2 instead of retyping it?",
        options: [
          "Retyping is slower",
          "So the dashboard follows the Performance sheet automatically when targets or results change",
          "Formulas cannot reference other sheets",
          "Retyping breaks Data Validation"
        ],
        answer: 1,
        explanation: "A mirrored block is a live view; a retyped block is a screenshot. The screenshot is wrong the first time the source sheet changes." }
    ],
    summary: [
      "One dashboard sheet, one validated month selector in B3, and every card pointed at it.",
      "KPI cards describe the quarter and never move with the selector; the Month Focus block always does.",
      "Two-pair SUMIFS (month AND product, month AND region) is the engine of month-specific cards.",
      "Mirrored blocks and INDEX/MATCH cards follow their source sheets automatically.",
      "Test by changing the selector and one log row — every card must move and return."
    ],
    downloads: [
      { file: "assets/files/level7/sales-dashboard.xlsx", label: "Interactive Sales Dashboard — exercise file", note: "Project 5 workbook plus an empty Dashboard sheet to build on" },
      { file: "assets/files/level7/sales-dashboard-solution.xlsx", label: "Interactive Sales Dashboard — completed solution", note: "Finished dashboard with KPI cards, Month Focus block, leaderboard and charts" }
    ],
    keywords: ["sales dashboard excel", "KPI cards", "month selector dropdown", "SUMIFS two criteria", "INDEX MATCH top performer", "data bars", "dashboard design", "single input cell"]
  },
  {
    id: "l7-project-office-mis",
    title: "Project: Complete Office MIS",
    difficulty: "Intermediate",
    duration: 45,
    version: "",
    objectives: [
      "Assemble every Level 7 system into one workbook with one MIS sheet",
      "Show revenue and expenses month by month on a single page",
      "Compute the month's net position with formulas, not opinions",
      "Run a reconciliation checklist that proves every block",
      "Close the period and hand over a file the next person can run"
    ],
    intro: "Seven projects in, you own the pieces: an employee system, attendance, payroll, inventory, a sales tracker, an expense tracker, invoices and a purchase register. Each one works alone. An office needs them to work together — and it needs one page where the owner can read the month without opening six files.\n\nThis final project builds that page: the Complete Office MIS. It pulls the sales tracker's revenue and the expense tracker's spending onto one sheet, computes the net position month by month, and finishes with a reconciliation checklist and a closing routine. Nothing new is introduced — the project is the discipline of connecting what you already built.\n\nEverything works in Excel 2010 and later.",
    sections: [
      {
        heading: "What You Will Build",
        blocks: [
          { type: "p", text: "The exercise file carries the finished Project 5 and Project 6 systems plus one empty **MIS** sheet and one empty **Close** sheet. You build the last two." },
          { type: "ul", items: [
            "**MIS** — a headline block, a month-by-month revenue vs expense table with net position, a leaderboard mirror, and the expense category summary.",
            "**Close** — the reconciliation checklist: one subtraction per check, every one must read 0.",
            "The logs, summaries and masters stay exactly as the earlier projects left them."
          ] },
          { type: "p", text: "The numbers the finished MIS shows: revenue of 10,63,500 across Aug-26 to Oct-26 against a team target of 10,10,000, and operating expenses of 8,50,900 across Jul-26 to Oct-26 against an 8,08,000 budget. The month table pairs the two where they overlap." }
        ]
      },
      {
        heading: "Skills You Will Use",
        blocks: [
          { type: "ul" , items: [
            "**Cross-sheet references** (Level 7 so far) — every MIS figure points at a summary sheet, never at a retyped number.",
            "**SUMIFS by month label** (Projects 5 and 6) — the Month helper columns both logs already carry.",
            "**INDEX/MATCH** (Project 9) — the headline cards.",
            "**The zero test** (Project 6) — the Close sheet's whole philosophy.",
            "**Save As discipline** (Level 5 capstone) — closing a period without breaking it."
          ] },
          { type: "p", text: "Two logs, two different Month helpers, one MIS. The sales log covers three months; the expense log covers four. The MIS must respect that asymmetry rather than pretend both cover the same window — a July expense row has no revenue twin, and the net column says so honestly." }
        ]
      },
      {
        heading: "Open the Exercise File",
        blocks: [
          { type: "p", text: "Download **Complete Office MIS — exercise file**. Inside:" },
          { type: "ul", items: [
            "**Sales Log, Monthly Summary, Performance, Dashboard** — the finished Project 9 system.",
            "**Expense Log, Category Summary, Monthly Summary** — the finished Project 6 system, on its own sheets.",
            "**MIS** and **Close** — empty. This is your work."
          ] },
          { type: "table", headers: ["Sheet group", "Owner project", "Feeds the MIS"], rows: [
            ["Sales Log + Monthly Summary + Performance", "Project 5 and 9", "Revenue by month, target, leaderboard"],
            ["Expense Log + Category Summary", "Project 6", "Expenses by month, category, budget variance"],
            ["MIS", "this project", "The one-page read"],
            ["Close", "this project", "The proof the MIS is honest"]
          ] },
          { type: "tip", text: "Sheet order matters for reading: put MIS first, then Close, then the systems. The owner opens the file and lands on the page that matters." }
        ]
      },
      {
        heading: "Build Section 1: The MIS Headline and the Month Table",
        blocks: [
          { type: "steps", items: [
            "On **MIS**, type `Sunrise Traders — Office MIS` in `A1`, bold, size 14, and `Aug–Oct 2026` in `A2`.",
            "Headline cards in `A4:C5`: `Revenue` `='Monthly Summary'!$B$7`; `Expenses` `=Category Summary!$B$9`; `Net Position` `=B4-C4`.",
            "Below the cards, a second headline line: `Vs target` `=B4/1010000` and `Vs budget` `=C4/808000`, both formatted as percentage.",
            "Month table headers in `A8:D8`: `Month`, `Revenue`, `Expenses`, `Net`.",
            "In `A9:A12` type `Aug-26`, `Sep-26`, `Oct-26`, and `Jul-26` last with a note — or keep Jul out and note it below.",
            "In `B9` type `=SUMIFS('Sales Log'!$H$2:$H$200,'Sales Log'!$I$2:$I$200,$A9)`; in `C9` type `=SUMIFS('Expense Log'!$D$2:$D$100,'Expense Log'!$E$2:$E$100,$A9)`; in `D9` type `=B9-C9`. Fill all three down."
          ] },
          { type: "grid", title: "MIS month table — what you should see", rows: [
            ["Month", "Revenue (₹)", "Expenses (₹)", "Net (₹)"],
            ["Aug-26", "320000", "212100", "107900"],
            ["Sep-26", "345000", "216200", "128800"],
            ["Oct-26", "398500", "214950", "183550"],
            ["Jul-26", "—", "207650", "—"]
          ] },
          { type: "p", text: "Read the table the way the owner will: October earned the most (3,98,500) and spent the least of the revenue months (2,14,950), so it carries the best net at 1,83,550. July shows an expense with no revenue because the sales log starts in August — the dash is honest, and the note under the table explains it in one line." },
          { type: "formula", name: "Net position",
            syntax: "=revenue - expenses",
            example: "=B9-C9",
            explanation: "One subtraction per row, computed live. Net position is the number every meeting orbit, and computing it by formula means a late invoice or a late voucher moves it without anyone remembering to re-type it." }
        ]
      },
      {
        heading: "Build Section 2: The Detail Blocks",
        blocks: [
          { type: "steps", items: [
            "Leaderboard block: mirror the Performance sheet with `=Performance!A2` filled across to `F2` and down to row 7, and apply data bars to the Actual column.",
            "Expense category block: mirror Category Summary rows 2 to 9 the same way — categories, actual, budget, variance, status.",
            "Add the two summary lines under the blocks: total revenue `=SUM(B9:B11)` over the three revenue months, and the 90%-club counter from Project 9.",
            "Format every money column with thousands separators; bold the block headers; border each block so the page reads in panels.",
            "Do one final pass for typed numbers: every cell on the MIS must be either a label or a formula. If you find a number you typed, replace it with the formula that computes it."
          ] },
          { type: "p", text: "The expense block will show exactly what Project 6 produced: Rent on budget at 1,60,000, Salaries 4,62,000 within budget, Electricity 51,200 and Marketing 86,500 over, and the total at 8,50,900 — 42,900 over the 8,08,000 budget. The two over-budget heads are the ones the owner will ask about first, and the MIS presents them without comment; the story belongs in the notes." },
          { type: "warning", text: "The two Month helpers use the same format on purpose: both logs produce `mmm-yy` labels like `Aug-26`. If the expense log's helper ever reads `Aug-2026`, its MIS column silently shows 0 for every month. When a block reads zero, compare helper formats before you doubt the data." }
        ]
      },
      {
        heading: "Build Section 3: The Close Sheet",
        blocks: [
          { type: "steps", items: [
            "On **Close**, type `Reconciliation Checklist` in `A1` and headers in `A3:C3`: `Check`, `Formula result`, `Verdict`.",
            "Check 1 — revenue: `='Monthly Summary'!$B$7-SUM(MIS!$B$9:$B$11)` — the summary's three-month total against the MIS month table's revenue column.",
            "Check 2 — expenses: `=Category Summary!$B$9-MIS!$C$12` where `C12` is the expense total across all four months.",
            "Check 3 — categories: `=SUM('Expense Log'!$D$2:$D$100)-Category Summary!$B$9` (Project 6's own check, restated here).",
            "Check 4 — leaderboard: `=Performance!$C$9-SUM('Sales Log'!$H$2:$H$200)`.",
            "In `C4:C7` type `=IF(B4=0,\"PASS\",\"CHECK\")` and fill down, then conditionally format: PASS green, CHECK red.",
            "Save As `Sunrise-Traders-MIS-CLOSED.xlsx` — this file is now the period's permanent record."
          ] },
          { type: "grid", title: "Close sheet — what you should see", rows: [
            ["Check", "Result", "Verdict"],
            ["Revenue: summary total vs MIS month table", "0", "PASS"],
            ["Expenses: category total vs MIS month table", "0", "PASS"],
            ["Expense log vs category summary", "0", "PASS"],
            ["Sales log vs performance actuals", "0", "PASS"]
          ] },
          { type: "p", text: "Four subtractions, four zeros, one verdict per row. The Close sheet is small, and that is the point: a page this simple gets checked every single time, which is exactly why the numbers on the MIS can be trusted every single time." }
        ]
      },
      {
        heading: "Hand Over and Close the Period",
        blocks: [
          { type: "steps" , items: [
            "Read the MIS top to bottom as the owner would: headline, month table, leaderboard, expense block, notes.",
            "Run the Close checklist one last time and watch all four verdicts read PASS.",
            "Save As the dated CLOSED file, then copy the whole workbook to start the next period — masters and structure stay, transaction logs clear.",
            "Write a five-line handover note: which sheets are inputs, which are computed, when the checklist runs, who owns the dropdown lists, where the backups live.",
            "Never edit a CLOSED file again; corrections happen in the live file and the next close."
          ] },
          { type: "p", text: "The handover note is part of the system. A workbook only one person understands is a risk dressed as an asset. Five honest lines — inputs, outputs, checks, lists, backups — turn your MIS into something the office can run after you move on." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Retyping revenue or expense totals onto the MIS because the formulas looked long.", right: "Reference the summary sheets and the logs through locked, generous ranges.", why: "One typed number breaks the chain the Close sheet is built to prove, and the checklist fails the first time the source moves." },
      { wrong: "Assuming both logs cover the same months and forcing July into the revenue table.", right: "Respect each log's real window and show an honest dash where a month has no data.", why: "A July revenue of 0 looks like a terrible sales month; a dash with a one-line note reads as the truth — the sales log starts in August." },
      { wrong: "Treating the Close sheet as a one-time build.", right: "Run all four checks at every close, before the file is archived.", why: "The checklist is only worth what it catches. Skipping it once teaches the office that the MIS is believed, not verified." },
      { wrong: "Editing the CLOSED file to fix a found error.", right: "Correct the live workbook, re-run the checklist, and close again with a new dated file.", why: "A CLOSED file that changes is not a record. The archive's value is that nobody — including you — can quietly rewrite history." }
    ],
    shortcuts: [
      { keys: "Ctrl + Page Down", action: "Walk from MIS to Close to each system while building" },
      { keys: "F9", action: "Recalculate the whole chain after any correction" },
      { keys: "F12", action: "Save As the dated CLOSED file at period end" },
      { keys: "Ctrl + `", action: "Show formulas — the fastest way to hunt for a typed number hiding on the MIS" }
    ],
    exercises: [
      { task: "Build the MIS headline and month table and verify every figure against the source systems.", hint: "Revenue cells read the sales Monthly Summary; expense cells read the expense log by month label; net is one subtraction.", solution: "Aug-26: 3,20,000 / 2,12,100 / 1,07,900. Sep-26: 3,45,000 / 2,16,200 / 1,28,800. Oct-26: 3,98,500 / 2,14,950 / 1,83,550. Headline: revenue 10,63,500 (105.3% of target), expenses 8,50,900 (105.3% of budget), net position 2,12,600 across the full window." },
      { task: "Complete the Close checklist and make all four verdicts read PASS.", hint: "Each check is a subtraction between a summary total and the MIS or log figure it must equal.", solution: "All four read 0 and PASS: revenue, expense, category-to-log, and leaderboard-to-log." },
      { task: "Produce the handover: name the input sheets, the computed sheets, and write the one-line note for July's dash.", hint: "Inputs are the two logs and the masters; everything else is computed; the note explains the sales log starts in August.", solution: "Inputs: Sales Log, Expense Log, both masters. Computed: all summaries, Dashboard, MIS, Close. July note: \"Expense log covers Jul–Oct; the sales log starts Aug-26, so July shows expenses with no revenue.\""}
    ],
    quiz: [
      { q: "Where does the MIS's revenue total come from?",
        options: [
          "Typed once a month from the sales log",
          "A cross-sheet reference to the sales Monthly Summary",
          "A PivotTable rebuilt each time",
          "The Dashboard's KPI card, retyped"
        ],
        answer: 1,
        explanation: "The MIS quotes summaries, and summaries quote logs. Every link in that chain is a formula, which is why the Close checklist can prove the page with four subtractions." },
      { q: "July-26 shows 2,07,650 of expenses and no revenue in the month table. Why?",
        options: [
          "The sales data is missing and must be restored",
          "The sales log covers Aug–Oct while the expense log covers Jul–Oct, so July has no revenue row",
          "July's revenue was zero",
          "The revenue SUMIFS is broken"
        ],
        answer: 1,
        explanation: "Two logs, two windows. The dash with a one-line note is honest reporting; a forced 0 would read as a disastrous sales month." },
      { q: "What makes the Close sheet's verdicts trustworthy?",
        options: [
          "Green conditional formatting",
          "They are computed subtractions that must read exactly 0 — not opinions or spot checks",
          "They were PASS last month",
          "The workbook is protected"
        ],
        answer: 1,
        explanation: "Each verdict is =IF(difference=0,\"PASS\",\"CHECK\") over a real subtraction. Formatting just makes the result visible; the zero is the proof." },
      { q: "An error is found in the CLOSED file a week after the period was closed. What is the correct move?",
        options: [
          "Edit the CLOSED file quietly",
          "Correct the live workbook, re-run the Close checklist, and close the period again with a new dated file",
          "Delete the CLOSED file",
          "Add a comment to the CLOSED file and leave the numbers"
        ],
        answer: 1,
        explanation: "The archive must never change — that is what makes it a record. Corrections happen in the live file and produce a new, later close." },
      { q: "Why is the handover note part of the project, not an extra?",
        options: [
          "Because marking requires it",
          "Because a workbook only one person can run is a risk; the note makes the system transferable",
          "Because Excel cannot store sheet descriptions",
          "Because auditors demand prose"
        ],
        answer: 1,
        explanation: "Inputs, computed sheets, checks, lists, backups — five lines that let the next person run the month without you. That is the difference between a file and a system." }
    ],
    summary: [
      "One MIS page: headline cards, month table with net position, leaderboard mirror, expense category block.",
      "Every MIS cell is a label or a formula — Ctrl + ` hunts down any typed number.",
      "The two logs have different windows and the MIS respects that with honest dashes and notes.",
      "The Close sheet proves the page with four subtractions that must each read 0.",
      "Close means Save As a dated CLOSED file, then start the next period in a fresh copy."
    ],
    downloads: [
      { file: "assets/files/level7/office-mis-system.xlsx", label: "Complete Office MIS — exercise file", note: "Finished Project 5–9 systems plus empty MIS and Close sheets to build" },
      { file: "assets/files/level7/office-mis-system-solution.xlsx", label: "Complete Office MIS — completed solution", note: "Finished MIS and Close sheets with all checks passing" }
    ],
    keywords: ["office MIS excel", "management information system", "month table revenue expense", "net position formula", "reconciliation checklist", "period close", "handover documentation", "cross sheet MIS"]
  },
  // __APPEND__
);
