// level5-2.js — Level 5: Real-World Office Applications (lessons 8-13, no metadata)
window.LEVEL_LESSONS = window.LEVEL_LESSONS || {};
(window.LEVEL_LESSONS[5] = window.LEVEL_LESSONS[5] || []).push(
  {
    id: "l5-sales-management",
    title: "Sales Management and Tracking",
    difficulty: "Intermediate",
    duration: 25,
    version: "",
    objectives: [
      "Maintain a sales log with one row per invoice",
      "Produce daily and monthly totals with date-based SUMIFS",
      "Compare target versus actual for every salesperson",
      "Measure month-on-month growth correctly",
      "Find the top performer and customer-wise totals"
    ],
    intro: "Sales is the one department where every manager asks for numbers daily. The way to survive that is a disciplined sales log plus a handful of formulas that rebuild every report automatically.\n\nSunrise Traders' four salespeople — Arjun Mehta, Priya Nair, Rahul Verma and Sneha Iyer — log every invoice as one row. In this lesson you will turn their October log into daily totals, monthly totals, target tracking, growth rates and a top-performer board.",
    keyIdea: "One row per invoice in a disciplined log lets formulas rebuild every sales report automatically.",
    sections: [
      {
        heading: "The Sales Log: One Row Per Invoice",
        blocks: [
          { type: "grid", title: "Sales Log sheet", rows: [
            ["Date", "Invoice No", "Customer", "Product", "Qty", "Rate", "Total", "Salesperson"],
            ["02-Oct-2026", "INV-2026-301", "Mehta Book House", "Notebook", "500", "60", "=E2*F2", "Arjun Mehta"],
            ["03-Oct-2026", "INV-2026-302", "Northstar Retail", "Printer", "10", "10500", "=E3*F3", "Priya Nair"],
            ["06-Oct-2026", "INV-2026-303", "Bright Office Co.", "Pen", "2000", "20", "=E4*F4", "Rahul Verma"],
            ["09-Oct-2026", "INV-2026-304", "Shree Ganesh Traders", "Stapler", "400", "175", "=E5*F5", "Sneha Iyer"]
          ] },
          { type: "ul", items: [
            "**Total** is always `=Qty * Rate`. A typed total hides an error in quantity or rate.",
            "**Invoice No** is unique forever and matches the invoice you will build in the billing lesson.",
            "**Salesperson** should be a Data Validation dropdown of the four names, spelled identically every time."
          ] },
          { type: "warning", text: "The salesperson column is a magnet for inconsistency: \"Arjun\", \"Arjun M\", \"Arjun Mehta \". SUMIFS matches exact text, so every variant splits one person's sales into two. A dropdown fixes it at entry." }
        ]
      },
      {
        heading: "Daily and Monthly Totals",
        blocks: [
          { type: "formula", name: "Total for one day", syntax: "=SUMIFS(totals, dates, DATE(year, month, day))", example: "=SUMIFS($G$2:$G$100, $A$2:$A$100, DATE(2026,10,2))", explanation: "Adds invoice totals only where the date equals 2 October 2026. DATE builds a real date, so the criteria compares date against date, not date against text." },
          { type: "formula", name: "Total for one month", syntax: "=SUMIFS(totals, dates, \">=\"&DATE(2026,10,1), dates, \"<\"&DATE(2026,11,1))", example: "=SUMIFS($G$2:$G$100, $A$2:$A$100, \">=\"&DATE(2026,10,1), $A$2:$A$100, \"<\"&DATE(2026,11,1))", explanation: "Two conditions bracket the month: on or after 1 October, and before 1 November. This catches every October invoice no matter which day it fell on." },
          { type: "p", text: "October 2026 revenue for Sunrise Traders comes to 30,58,500. The bracket trick matters because dates are stored as numbers, so `>=` and `<` do the clean filtering that a text criteria like \"October\" never could." },
          { type: "tip", text: "Build a small month summary with one SUMIFS per month and one growth row under it. At month end you change two date arguments, and the whole management report refreshes itself." }
        ]
      },
      {
        heading: "Target vs Actual per Salesperson",
        blocks: [
          { type: "grid", title: "Target vs actual, October 2026", rows: [
            ["Salesperson", "Target", "Actual", "Achievement %"],
            ["Arjun Mehta", "800000", "=SUMIFS($G$2:$G$100, $H$2:$H$100, A2, $A$2:$A$100, \">=\"&DATE(2026,10,1), $A$2:$A$100, \"<\"&DATE(2026,11,1))", "=D2/C2"],
            ["Priya Nair", "1000000", "=SUMIFS($G$2:$G$100, $H$2:$H$100, A3, $A$2:$A$100, \">=\"&DATE(2026,10,1), $A$2:$A$100, \"<\"&DATE(2026,11,1))", "=D3/C3"],
            ["Rahul Verma", "600000", "=SUMIFS($G$2:$G$100, $H$2:$H$100, A4, $A$2:$A$100, \">=\"&DATE(2026,10,1), $A$2:$A$100, \"<\"&DATE(2026,11,1))", "=D4/C4"],
            ["Sneha Iyer", "500000", "=SUMIFS($G$2:$G$100, $H$2:$H$100, A5, $A$2:$A$100, \">=\"&DATE(2026,10,1), $A$2:$A$100, \"<\"&DATE(2026,11,1))", "=D5/C5"]
          ] },
          { type: "formula", name: "Achievement %", syntax: "=Actual / Target", example: "=D2/C2", explanation: "Formats as a percentage. 100% means the target was met exactly; anything above is a beat." },
          { type: "table", headers: ["Salesperson", "Target", "Actual", "Achievement"], rows: [
            ["Arjun Mehta", "800000", "925000", "115.6%"],
            ["Priya Nair", "1000000", "1084000", "108.4%"],
            ["Rahul Verma", "600000", "582000", "97.0%"],
            ["Sneha Iyer", "500000", "467500", "93.5%"]
          ] },
          { type: "p", text: "One glance now tells the story: two salespeople beat their targets, two fell just short. Adding conditional formatting below 100% turns this table into a monthly review screen with no extra effort." }
        ]
      },
      {
        heading: "Growth vs Last Month",
        blocks: [
          { type: "formula", name: "Growth %", syntax: "=(this_month - last_month) / last_month", example: "=(3058500-2684500)/2684500", explanation: "Compares this month's revenue with the previous month's. Format as a percentage; 0.139 shows as 13.9%." },
          { type: "p", text: "September revenue was 26,84,500 and October was 30,58,500, so growth is 13.9%. In the sheet, compute both months with the bracketed SUMIFS and then divide the difference by the previous month." },
          { type: "warning", text: "If last month's total is zero, this division produces a division-by-zero error. Guard it: `=IF(B4=0, \"n/a\", (C4-B4)/B4)`. A new territory or a shutdown month makes this guard essential." },
          { type: "p", text: "One honest caution: growth alone flatters. October was strong, but compare growth against the target plan, not just against last month — otherwise one festival season makes every month look like success." }
        ]
      },
      {
        heading: "Top Performer and Customer-wise Totals",
        blocks: [
          { type: "formula", name: "Top performer", syntax: "=INDEX(names, MATCH(MAX(actuals), actuals, 0))", example: "=INDEX($A$2:$A$5, MATCH(MAX($D$2:$D$5), $D$2:$D$5, 0))", explanation: "MAX finds the highest actual, MATCH finds its position, INDEX returns the salesperson's name from that same position in the names column." },
          { type: "p", text: "October's board reads: Priya Nair tops with 10,84,000, Arjun Mehta follows with 9,25,000. The formula updates itself whenever the log changes — worth it, because naming a winner by hand at the review meeting is exactly how disputes start." },
          { type: "formula", name: "Customer-wise total", syntax: "=SUMIFS(totals, customers, customer)", example: "=SUMIFS($G$2:$G$100, $C$2:$C$100, \"Mehta Book House\")", explanation: "Adds every invoice total for one customer across all months and salespeople. Build one row per customer to see who really buys." },
          { type: "table", headers: ["Customer", "Oct revenue"], rows: [
            ["Mehta Book House", "960000"],
            ["Northstar Retail", "771500"],
            ["Bright Office Co.", "745000"],
            ["Shree Ganesh Traders", "582000"]
          ] },
          { type: "p", text: "Mehta Book House leads customers this month. Cross-check: the four customer totals add to 30,58,500, matching the monthly total exactly. Numbers that reconcile two ways are numbers you can defend in a meeting." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Writing date criteria as text, like \"02/10/2026\" or \"October\".", right: "Use real dates in the criteria: DATE(2026,10,1) with >= and < brackets.", why: "Text criteria compare against the stored date number and match nothing. The bracket of two DATE values catches the whole month reliably." },
      { wrong: "Computing growth as this month divided by last month.", right: "Subtract first: (this - last) / last.", why: "3058500/2684500 is 1.139, displayed as 113.9%. Growth is the change, so the base month must be subtracted before dividing." },
      { wrong: "Logging the salesperson as \"Arjun\" in some rows and \"Arjun Mehta\" in others.", right: "Use a dropdown list of full names on the salesperson column.", why: "SUMIFS matches exact text. Every spelling variant quietly splits one person's totals into two smaller ones." },
      { wrong: "Typing invoice totals by hand.", right: "Total = Qty x Rate as a formula on every row.", why: "Typed totals disagree with quantities and rates, and every report built on them inherits the error." }
    ],
    shortcuts: [
      { keys: "Alt + Down", action: "Pick a customer or salesperson from a dropdown while logging" },
      { keys: "Ctrl + Shift + L", action: "Filter the log to one salesperson or one product" },
      { keys: "Ctrl + D", action: "Copy the previous row's formulas down to a new invoice" }
    ],
    exercises: [
      { task: "In the sales-records workbook, compute the October 2026 total revenue with a bracketed SUMIFS and a single-day total for 2 October.", hint: "One condition pair with DATE(2026,10,2); the month needs >= 1-Oct and < 1-Nov.", solution: "October total is 30,58,500. The 2 October daily total matches the Solution sheet." },
      { task: "Build the target vs actual table for all four salespeople with Achievement % and highlight anyone below 100%.", hint: "SUMIFS with three criteria pairs: salesperson plus the two date brackets. Achievement = actual/target.", solution: "Arjun 115.6%, Priya 108.4%, Rahul 97.0%, Sneha 93.5% — Rahul and Sneha fall short." },
      { task: "Find October's top performer with MAX, MATCH and INDEX, then compute customer-wise totals and confirm they reconcile to the month total.", hint: "INDEX(names, MATCH(MAX(actuals), actuals, 0)). Customer totals should sum to 30,58,500.", solution: "Top performer: Priya Nair with 10,84,000. Customer totals: 9,60,000 + 7,71,500 + 7,45,000 + 5,82,000 = 30,58,500." }
    ],
    quiz: [
      { q: "Which formula correctly totals all invoices of October 2026?", options: ["=SUMIF($A:$A, \"October\", $G:$G)", "=SUMIFS($G$2:$G$100, $A$2:$A$100, \">=\"&DATE(2026,10,1), $A$2:$A$100, \"<\"&DATE(2026,11,1))", "=SUM($G$2:$G$100, DATE(2026,10,1))", "=COUNTIFS($A:$A, \"Oct\")"], answer: 1, explanation: "A month is a range of dates, so SUMIFS with two conditions — on or after the 1st, before the next month's 1st — brackets it exactly. Text criteria like October match nothing against real dates." },
      { q: "Target 10,00,000, actual 10,84,000. What does Achievement % show?", options: ["8.4%", "92.3%", "108.4%", "1,08,400"], answer: 2, explanation: "Achievement = actual / target = 1084000/1000000 = 108.4%. Anything above 100% means the target was beaten." },
      { q: "September revenue 26,84,500; October 30,58,500. Which is the correct growth calculation?", options: ["=3058500/2684500", "=(3058500-2684500)/2684500", "=3058500-2684500", "=2684500/3058500"], answer: 1, explanation: "Growth is the change divided by the base: 374000/2684500 = 13.9%. Dividing without subtracting gives 113.9%, which is a ratio, not growth." },
      { q: "Which combination finds the top performer's name from the actuals table?", options: ["=MAX(D2:D5) alone", "=INDEX(names, MATCH(MAX(actuals), actuals, 0))", "=VLOOKUP(MAX(D2:D5), names, 1)", "=LARGE(names, 1)"], answer: 1, explanation: "MAX finds the value, MATCH locates its position, and INDEX returns the name at that position. MAX alone returns a number, not a person." },
      { q: "Why should the Salesperson column use a dropdown of exact full names?", options: ["It makes the log look neat", "SUMIFS matches exact text, so variants split one person's totals", "Dropdowns speed up typing", "Filters require dropdowns"], answer: 1, explanation: "Arjun and Arjun Mehta are different strings to Excel. Every variant divides one salesperson's numbers into two, breaking targets and top-performer logic." }
    ],
    summary: [
      "One row per invoice with Total = Qty x Rate and a unique invoice number.",
      "Monthly totals come from SUMIFS bracketed by two DATE values: >= month start, < next month start.",
      "Achievement % = actual / target, per salesperson, with conditional formatting below 100%.",
      "Growth % = (this month - last month) / last month, guarded against a zero base.",
      "MAX + MATCH + INDEX names the top performer; customer-wise SUMIFS must reconcile to the month total."
    ],
    downloads: [
      { file: "assets/files/level5/sales-records.xlsx", label: "Sales records workbook", note: "Exercise and Solution sheets — sales log with targets" }
    ],
    keywords: ["sales tracker", "sales log", "monthly sales total", "target vs actual", "growth percentage", "top performer index match", "customer wise sales"]
  },
  {
    id: "l5-budget-preparation",
    title: "Monthly Budget Preparation",
    difficulty: "Intermediate",
    duration: 22,
    version: "",
    objectives: [
      "Set up a monthly budget with one row per expense head",
      "Compare budget against actual with a variance column",
      "Mark every head Over, Under or On budget automatically",
      "Track income against target and compute achievement",
      "Close the month with a total variance you can explain"
    ],
    intro: "A budget is a promise written as numbers. Before October began, Sunrise Traders sat down and wrote down what each expense head should cost and what sales should bring in. October is now over, so the promise can be checked against reality.\n\nIn this lesson you will build the budget sheet for October 2026: the expense heads, the budget column, the actual column and the variance column that turns two numbers into one sentence \u2014 over, under or on budget.",
    keyIdea: "Budget vs actual is checked with one variance column that says over, under or on.",
    sections: [
      {
        heading: "A Budget Is a Plan with a Difference Column",
        blocks: [
          { type: "p", text: "A budget sheet has only three working columns: what we planned, what actually happened, and the difference. Everything else is decoration. The difference column is called the **variance**, and reading variances is the whole point of the exercise." },
          { type: "ul", items: [
            "**One row per expense head.** Not one row per bill \u2014 bills live in the expense log.",
            "**Same head names every month.** \"Marketing\" in October and \"Advertising\" in November breaks every month-on-month comparison.",
            "**Budget typed once, actuals pulled by formula** from the payroll sheet and the expense log wherever possible.",
            "**Variance always a formula:** `=Actual - Budget`. Never typed, never mental arithmetic."
          ] },
          { type: "tip", text: "Prepare the budget in the last week of the previous month, and freeze it. A budget that keeps changing after the month starts is not a budget \u2014 it is a wish list." }
        ]
      },
      {
        heading: "The Expense Heads of Sunrise Traders",
        blocks: [
          { type: "p", text: "Eight heads cover the office. Salaries come from the payroll sheet (net paid out), and every other head is either a fixed commitment (rent) or comes from the expense log during the month." },
          { type: "grid", title: "Budget sheet \u2014 October 2026", rows: [
            ["Expense Head", "Budget", "Actual", "Variance", "Status"],
            ["Office Rent", "30000", "30000", "=C2-B2", "=IF(D2>0, \"Over\", IF(D2<0, \"Under\", \"On budget\"))"],
            ["Salaries (Net)", "240760", "240760", "=C3-B3", "=IF(D3>0, \"Over\", IF(D3<0, \"Under\", \"On budget\"))"],
            ["Electricity", "8500", "9100", "=C4-B4", "=IF(D4>0, \"Over\", IF(D4<0, \"Under\", \"On budget\"))"],
            ["Internet and Phone", "2500", "2650", "=C5-B5", "=IF(D5>0, \"Over\", IF(D5<0, \"Under\", \"On budget\"))"],
            ["Transport and Delivery", "6000", "5400", "=C6-B6", "=IF(D6>0, \"Over\", IF(D6<0, \"Under\", \"On budget\"))"],
            ["Stationery", "3000", "3200", "=C7-B7", "=IF(D7>0, \"Over\", IF(D7<0, \"Under\", \"On budget\"))"],
            ["Marketing", "8000", "7600", "=C8-B8", "=IF(D8>0, \"Over\", IF(D8<0, \"Under\", \"On budget\"))"],
            ["Miscellaneous", "5000", "5900", "=C9-B9", "=IF(D9>0, \"Over\", IF(D9<0, \"Under\", \"On budget\"))"],
            ["Total", "=SUM(B2:B9)", "=SUM(C2:C9)", "=C10-B10", "=IF(D10>0, \"Over\", IF(D10<0, \"Under\", \"On budget\"))"]
          ] },
          { type: "p", text: "The budget column totals 3,03,760. The actual column totals 3,04,610. The company spent 850 more than planned \u2014 about 0.3% \u2014 which is a good month. But the total alone hides the story: some heads overspent and others underspent, and the individual variances are what you act on." }
        ]
      },
      {
        heading: "Variance: Over, Under and Conditional Formatting",
        blocks: [
          { type: "formula", name: "Variance", syntax: "=actual - budget", example: "=C4-B4", explanation: "For an expense head, a positive variance means you spent more than planned (over budget) and a negative variance means you spent less (under budget). The sign carries the meaning, so never use ABS() here." },
          { type: "steps", items: [
            "Select the variance cells D2:D10.",
            "Home, Conditional Formatting, Highlight Cells Rules, Greater Than.",
            "Enter 0 and pick Light Red Fill \u2014 for expenses, red means over budget.",
            "Add a second rule: Less Than 0 with Light Green Fill \u2014 under budget is good news for expenses.",
            "Heads with a variance of exactly 0 stay unformatted: they are on budget."
          ] },
          { type: "warning", text: "The sign flips for income. For the sales line, spending more than budget is good. If one sheet mixes expense rows and income rows, a single red rule will highlight the wrong rows. Keep income on its own block with its own rule, or use a formula-based rule that checks the row type." }
        ]
      },
      {
        heading: "Income Against Target",
        blocks: [
          { type: "p", text: "The income side of the budget has one big line for October: sales. The budget set a sales target of 29,10,000 \u2014 the same target the four salespeople carry individually, which you will meet again in the MIS lesson." },
          { type: "grid", title: "Income block \u2014 October 2026", rows: [
            ["Income Head", "Budget", "Actual", "Variance", "Achievement %"],
            ["Sales", "2910000", "3058500", "=C13-B13", "=ROUND(C13/B13*100, 1)"]
          ] },
          { type: "p", text: "Actual sales beat target by 1,48,500, and achievement is 105.1%. For income, a positive variance is favourable \u2014 the opposite of the expense rows above \u2014 which is exactly why the two blocks are kept separate." },
          { type: "formula", name: "Achievement percentage", syntax: "=ROUND(actual / budget * 100, 1)", example: "=ROUND(C13/B13*100, 1)", explanation: "One decimal place is enough for a monthly review. 105.1% reads instantly; 105.08620689% does not. The ROUND also makes the stored value match the displayed value." }
        ]
      },
      {
        heading: "Explain Every Big Variance",
        blocks: [
          { type: "p", text: "A variance without a reason is just a number. Sunrise Traders uses one simple rule: any head off budget by more than 5% of its own budget needs one written line of explanation in a Notes column." },
          { type: "table", headers: ["Head", "Variance", "% of head budget", "Explanation line"], rows: [
            ["Electricity", "+600", "7.1%", "Extra festival-week working days in the shop"],
            ["Internet and Phone", "+150", "6.0%", "One extra data top-up"],
            ["Stationery", "+200", "6.7%", "Bulk pen order for November invoicing"],
            ["Miscellaneous", "+900", "18.0%", "Emergency printer repair"],
            ["Transport and Delivery", "-600", "10.0%", "Two deliveries combined per trip"],
            ["Marketing", "-400", "5.0%", "One newspaper insert postponed to November"]
          ] },
          { type: "p", text: "Six lines of notes turn a variance report into a review meeting. Without them, the meeting spends its time re-deriving numbers instead of making decisions." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Typing the variance as a number: \"600 over\".", right: "Use =Actual - Budget so the variance recomputes when the actual changes.", why: "A typed variance goes stale the moment anyone edits the actual column, and the sheet starts contradicting itself." },
      { wrong: "Colouring every positive variance red on a sheet that also holds income.", right: "Keep expense rows and income rows in separate blocks with separate rules.", why: "For expenses, over budget is bad; for income, over budget is good. One shared rule highlights the wrong rows in one of the two blocks." },
      { wrong: "Renaming heads month to month: \"Marketing\" becomes \"Advertising\".", right: "Fix the head list once and reuse it every month.", why: "SUMIFS, filters and month-on-month comparisons match exact text, so renamed heads silently vanish from the totals." },
      { wrong: "Rounding with Decrease Decimal instead of ROUND.", right: "Wrap computed money in ROUND so stored and displayed values agree.", why: "Budget totals that differ from payroll and expense logs by a few paise destroy trust in the sheet." }
    ],
    shortcuts: [
      { keys: "Alt + =", action: "AutoSum the budget or actual column in one keystroke" },
      { keys: "Ctrl + Shift + %", action: "Apply percent format to the achievement cells" },
      { keys: "Ctrl + 1", action: "Open Format Cells for number formats and alignment" },
      { keys: "Ctrl + Shift + L", action: "Toggle filters to review the over-budget heads quickly" }
    ],
    exercises: [
      { task: "In the budget-tracker workbook, complete the variance and status columns for all eight expense heads and find the total variance for October.", hint: "Variance = Actual - Budget; total budget 3,03,760 and total actual 3,04,610.", solution: "Total variance is +850 (over budget, about 0.3%). Overs: Electricity +600, Internet +150, Stationery +200, Miscellaneous +900 (1,850 combined). Unders: Transport -600, Marketing -400 (1,000 combined)." },
      { task: "Compute the sales variance and achievement for October against the 29,10,000 target.", hint: "Actual 30,58,500 from the sales log; achievement = actual / budget.", solution: "Variance +1,48,500 favourable; achievement 105.1%." },
      { task: "Apply the 5% rule: list every head that needs a written explanation, with its variance as a percentage of its own budget.", hint: "Compare each variance to that head's budget, not to the total.", solution: "Electricity +7.1%, Internet and Phone +6.0%, Stationery +6.7%, Miscellaneous +18.0%, Transport and Delivery -10.0%, Marketing -5.0% \u2014 all six need explanation lines." }
    ],
    quiz: [
      { q: "For an expense head, the variance is =Actual - Budget. A positive variance means:", options: ["The head was under budget", "The head was over budget", "The budget was wrong", "Nothing \u2014 sign has no meaning"], answer: 1, explanation: "Actual minus budget positive means more was spent than planned: over budget. The same positive sign on an income row would be good news, which is why the two blocks stay separate." },
      { q: "Which formula marks the Status column correctly?", options: ["=IF(B2>0, \"Over\", \"Under\")", "=IF(D2>0, \"Over\", IF(D2<0, \"Under\", \"On budget\"))", "=IF(D2=0, \"Over\", \"Under\")", "=OVER(D2)"], answer: 1, explanation: "Three cases exist: over, under and exactly on budget. The nested IF tests positive, negative and zero in that order." },
      { q: "Budget total 3,03,760, actual total 3,04,610. What is the total variance?", options: ["+850, about 0.3% over", "+850, about 3% over", "-850, under budget", "+1,850, about 0.6% over"], answer: 0, explanation: "304610 - 303760 = 850, and 850 / 303760 = 0.28%, so about 0.3% over \u2014 a well-controlled month." },
      { q: "Sales actual 30,58,500 against a 29,10,000 budget. Which statement is right?", options: ["Variance -1,48,500, below target", "Variance +1,48,500, achievement 105.1%", "Achievement 94.9%", "Variance +1,48,500 \u2014 over budget in the bad sense"], answer: 1, explanation: "3058500 - 2910000 = 148500 favourable, and 3058500/2910000 = 105.1%. On income rows a positive variance is favourable." },
      { q: "Why must expense head names stay identical every month?", options: ["It looks tidy", "Filters and SUMIFS match exact text, so renamed heads vanish from comparisons", "Excel forbids renaming", "Shorter names are faster"], answer: 1, explanation: "\"Marketing\" and \"Advertising\" are different strings to Excel. Month-on-month totals silently drop the renamed head unless you fix the name list." }
    ],
    summary: [
      "One row per expense head; budget typed once, actuals pulled from payroll and the expense log.",
      "Variance = Actual - Budget, always a formula; status from a nested IF.",
      "For expenses red means over; for income red means under \u2014 keep the blocks separate.",
      "Achievement % = ROUND(actual / budget * 100, 1); October sales hit 105.1%.",
      "Any head off budget by more than 5% of its own budget needs a one-line explanation."
    ],
    downloads: [
      { file: "assets/files/level5/budget-tracker.xlsx", label: "Budget tracker workbook", note: "Exercise and Solution sheets \u2014 October 2026 budget vs actual with heads and income block" }
    ],
    keywords: ["monthly budget", "budget vs actual", "variance analysis", "over budget", "achievement percentage", "expense heads", "conditional formatting"]
  },
  {
    id: "l5-expense-tracker",
    title: "Office Expense and Reimbursement Tracker",
    difficulty: "Intermediate",
    duration: 24,
    version: "",
    objectives: [
      "Log every office expense with one row per voucher",
      "Keep categories and names clean with Data Validation dropdowns",
      "Total each category with SUMIFS",
      "Work out who must be reimbursed how much",
      "Close the month with a reconciliation check"
    ],
    intro: "Every office spends small amounts every day: diesel for the delivery van, a ream of paper, courier charges, tea for a visiting customer. Nobody plans these in advance, everybody pays somehow, and by month-end nobody can say where the money went.\n\nThe fix is an expense log with discipline: one row per voucher, categories from a dropdown, and three small SUMIFS reports on top. In this lesson you will build that log for November 2026 and end up knowing exactly who must be reimbursed what.",
    keyIdea: "Log each voucher in one row with a category dropdown, and the summaries build themselves.",
    sections: [
      {
        heading: "One Row Per Voucher",
        blocks: [
          { type: "p", text: "A voucher is one payment, given one number, forever. At Sunrise Traders November vouchers run EXP-2026-031 onwards, continuing from October. The log has eight columns and no merged cells, no blank rows and no total row in the middle." },
          { type: "grid", title: "Expense log \u2014 November 2026 (first half)", rows: [
            ["Date", "Voucher No", "Paid By", "Category", "Description", "Amount", "Mode", "Bill?"],
            ["02-Nov-2026", "EXP-2026-031", "Vikram Singh", "Stationery", "Printer paper 10 reams", "3200", "Cash", "Y"],
            ["03-Nov-2026", "EXP-2026-032", "Divya Menon", "Travel", "Bus fare - bank visit", "480", "Cash", "N"],
            ["05-Nov-2026", "EXP-2026-033", "Sneha Iyer", "Fuel", "Delivery van diesel", "4600", "Card", "Y"],
            ["07-Nov-2026", "EXP-2026-034", "Vikram Singh", "Meals", "Staff tea and biscuits", "850", "Cash", "Y"],
            ["09-Nov-2026", "EXP-2026-035", "Rahul Verma", "Repairs", "Printer service", "1500", "UPI", "Y"],
            ["10-Nov-2026", "EXP-2026-036", "Vikram Singh", "Stationery", "Pens and registers", "1750", "Cash", "Y"],
            ["12-Nov-2026", "EXP-2026-037", "Divya Menon", "Travel", "Courier to customer", "350", "UPI", "Y"],
            ["14-Nov-2026", "EXP-2026-038", "Sneha Iyer", "Fuel", "Delivery van diesel", "4400", "Card", "Y"],
            ["17-Nov-2026", "EXP-2026-039", "Vikram Singh", "Repairs", "Tube light replacement", "620", "Cash", "Y"],
            ["19-Nov-2026", "EXP-2026-040", "Divya Menon", "Postage", "Speed post invoices", "260", "Cash", "Y"]
          ] },
          { type: "ul", items: [
            "**Amount** is a plain number: 3200, never \"Rs. 3,200\". Text amounts silently drop out of every SUMIFS.",
            "**Mode** records how the money moved: Cash, UPI or Card. It tells you how much cash to arrange.",
            "**Bill?** is Y or N. A missing bill is a follow-up task, and the Y/N filter finds them all in one click."
          ] },
          { type: "warning", text: "The Paid By column is where reimbursement money disappears. \"Vikram\", \"Vikram S\" and \"Vikram Singh \" are three different people as far as SUMIFS is concerned. A dropdown of exact full names keeps one person's spending in one place." }
        ]
      },
      {
        heading: "Dropdowns That Keep the Log Clean",
        blocks: [
          { type: "steps", items: [
            "List the categories on a Lists sheet: Travel, Stationery, Fuel, Meals, Repairs, Postage, Other.",
            "Select the Category column range on the log.",
            "Data, Data Validation, Allow: List, Source: =Lists!$A$2:$A$8.",
            "Repeat for Paid By (the six employee names) and Mode (Cash, UPI, Card).",
            "Set Bill? validation to List with only Y and N."
          ] },
          { type: "p", text: "The dropdown does two jobs at once. It makes entry faster \u2014 one keypress instead of typing \u2014 and it makes reporting possible, because SUMIFS can only total what it can match exactly. Seven categories are enough; add a new one only when a genuinely new kind of spending appears." },
          { type: "tip", text: "Press Alt + Down Arrow in a validated cell to open its dropdown without touching the mouse. Ctrl + ; drops today's date into the Date column." }
        ]
      },
      {
        heading: "Category Totals with SUMIFS",
        blocks: [
          { type: "formula", name: "Total for one category", syntax: "=SUMIFS(sum_range, criteria_range, criteria)", example: "=SUMIFS($F$2:$F$11, $D$2:$D$11, \"Fuel\")", explanation: "Adds the Amount column only where Category equals Fuel. The two ranges must be the same size, and both start at row 2 \u2014 including the header row would try to add the word \"Amount\"." },
          { type: "grid", title: "Category summary \u2014 November (first half)", rows: [
            ["Category", "Total", "Share of spend"],
            ["Stationery", "=SUMIFS($F$2:$F$11, $D$2:$D$11, \"Stationery\")", "=ROUND(B2/$B$8*100, 1)"],
            ["Fuel", "=SUMIFS($F$2:$F$11, $D$2:$D$11, \"Fuel\")", "=ROUND(B3/$B$8*100, 1)"],
            ["Repairs", "=SUMIFS($F$2:$F$11, $D$2:$D$11, \"Repairs\")", "=ROUND(B4/$B$8*100, 1)"],
            ["Travel", "=SUMIFS($F$2:$F$11, $D$2:$D$11, \"Travel\")", "=ROUND(B5/$B$8*100, 1)"],
            ["Meals", "=SUMIFS($F$2:$F$11, $D$2:$D$11, \"Meals\")", "=ROUND(B6/$B$8*100, 1)"],
            ["Postage", "=SUMIFS($F$2:$F$11, $D$2:$D$11, \"Postage\")", "=ROUND(B7/$B$8*100, 1)"],
            ["Total", "=SUM(B2:B7)", "=ROUND(B8/$B$8*100, 1)"]
          ] },
          { type: "p", text: "Fuel leads with 9,000 \u2014 half of all spending, and no surprise for a company that delivers. Stationery follows at 4,950. The share column turns each total into a percentage of the 18,010 grand total, which is the sentence managers actually ask for: \"fuel is half our expense\"." }
        ]
      },
      {
        heading: "Who Gets Reimbursed",
        blocks: [
          { type: "p", text: "Employees paid from their own pockets wait for reimbursement. One more SUMIFS \u2014 by Paid By instead of Category \u2014 produces the payment list." },
          { type: "grid", title: "Reimbursement sheet \u2014 November (first half)", rows: [
            ["Employee", "To Reimburse"],
            ["Sneha Iyer", "=SUMIFS($F$2:$F$11, $C$2:$C$11, \"Sneha Iyer\")"],
            ["Vikram Singh", "=SUMIFS($F$2:$F$11, $C$2:$C$11, \"Vikram Singh\")"],
            ["Rahul Verma", "=SUMIFS($F$2:$F$11, $C$2:$C$11, \"Rahul Verma\")"],
            ["Divya Menon", "=SUMIFS($F$2:$F$11, $C$2:$C$11, \"Divya Menon\")"],
            ["Total", "=SUM(B2:B5)"]
          ] },
          { type: "p", text: "Sneha 9,000 (both fuel fills on the company card \u2014 confirm with the mode column before paying), Vikram 6,420, Rahul 1,500 and Divya 1,090. The four amounts add back to exactly 18,010, the same as the log total and the category total. That equality is not luck; it is the check." },
          { type: "formula", name: "The reconciliation check", syntax: "=log_total - category_total", example: "=SUM($F$2:$F$11)-B8", explanation: "Every report built from the same log must total the same. If log total minus report total is not zero, rows are missing, duplicated or filtered out of one of the ranges." }
        ]
      },
      {
        heading: "Modes, Bills and the Month-End Close",
        blocks: [
          { type: "p", text: "Two more one-click views keep the office honest. Mode-wise, Cash is 7,160, Card 9,000 and UPI 1,850 \u2014 useful for planning the petty-cash refill. Bill-wise, filter Bill? = N: only the 480 bus fare lacks a bill this half-month, and Divya should produce it before reimbursement day." },
          { type: "steps", items: [
            "Check every voucher has a number, a category, a name from the dropdown and an amount.",
            "Filter Bill? = N and chase the missing bills.",
            "Confirm the three totals agree: log, category sheet and reimbursement sheet.",
            "Copy the reimbursement sheet into the payment instruction for the accountant.",
            "Lock the half-month's rows (or the whole sheet) so closed vouchers cannot be edited."
          ] }
        ]
      }
    ],
    mistakes: [
      { wrong: "Typing categories freely: \"Fuel\", \"fuel \", \"Diesel\".", right: "Use the Data Validation dropdown of exact category names.", why: "SUMIFS matches exact text, so every variant splits the category total and the reconciliation check stops being zero." },
      { wrong: "Entering the amount as \"Rs. 4,600\" or \"4,600/-\".", right: "Type the plain number 4600 and format the display as currency.", why: "Text amounts are ignored by SUM and SUMIFS, so totals come out short with no error message to warn you." },
      { wrong: "Reusing or skipping voucher numbers when a voucher is cancelled.", right: "Never reuse a number; mark a cancelled voucher with a status note instead.", why: "Gaps you can explain; duplicate vouchers you cannot audit. Both break the voucher register's purpose." },
      { wrong: "Adding a total row inside the log between data rows.", right: "Keep totals on a separate summary sheet or below the data with a blank row between.", why: "A total row inside the range gets added into SUMIFS results, doubling the totals." }
    ],
    shortcuts: [
      { keys: "Alt + Down Arrow", action: "Open the dropdown list in the current validated cell" },
      { keys: "Ctrl + ;", action: "Insert today's date in the Date column" },
      { keys: "Ctrl + Shift + L", action: "Toggle filters, e.g. to list every Bill? = N voucher" },
      { keys: "Alt + =", action: "AutoSum the Amount column at the bottom of the log" }
    ],
    exercises: [
      { task: "In the expense-tracker workbook, build the category summary with SUMIFS and confirm the grand total.", hint: "Six categories; both ranges span rows 2 to 11.", solution: "Stationery 4,950; Fuel 9,000; Repairs 2,120; Travel 830; Meals 850; Postage 260. Grand total 18,010, matching the log total." },
      { task: "Build the reimbursement sheet per employee and verify it reconciles.", hint: "SUMIFS by the Paid By column, then compare the total to the log total.", solution: "Vikram Singh 6,420; Divya Menon 1,090; Sneha Iyer 9,000; Rahul Verma 1,500. The four add to 18,010 \u2014 reconciliation passes." },
      { task: "Produce mode-wise totals and list the vouchers missing bills.", hint: "SUMIFS by Mode; filter Bill? = N.", solution: "Cash 7,160 (6 vouchers), Card 9,000 (2), UPI 1,850 (2). Only EXP-2026-032, Divya's 480 bus fare, has no bill." }
    ],
    quiz: [
      { q: "Why must the expense log have one row per voucher?", options: ["It makes the sheet longer", "SUMIFS, filters and reconciliation assume that shape", "It saves paper", "Excel requires a voucher column"], answer: 1, explanation: "Every report in this lesson \u2014 category, employee, mode \u2014 is a SUMIFS over the log. A log with two vouchers in one row or totals inside the data breaks all of them at once." },
      { q: "Which formula totals the Stationery category?", options: ["=SUMIF(D2:D11, \"Stationery\")", "=SUMIFS($F$2:$F$11, $D$2:$D$11, \"Stationery\")", "=SUM($F$2:$F$11) + \"Stationery\"", "=COUNTIFS($F$2:$F$11, \"Stationery\")"], answer: 1, explanation: "SUMIFS takes the sum range first, then each criteria range with its criteria. SUM would add every voucher, and COUNTIFS counts rows, not rupees." },
      { q: "The reimbursement sheet totals 18,010 and the log totals 17,530. What does that mean?", options: ["The reimbursement sheet is wrong \u2014 delete it", "480 of vouchers are missing from one of the two ranges", "Excel rounded the difference", "The numbers must be text"], answer: 1, explanation: "Two reports from the same log must agree. A difference means rows are missing, duplicated or the ranges cover different rows \u2014 find it before anyone is paid." },
      { q: "What happens to \"Rs. 4,600\" typed into the Amount column?", options: ["SUMIFS reads it correctly", "Excel converts it to 4600 automatically", "It is treated as text and silently excluded from totals", "Excel shows an error"], answer: 2, explanation: "A rupee sign and commas make the entry text. Every sum skips it quietly, so totals come out short with no warning \u2014 the most expensive silent failure in an expense log." },
      { q: "Why is Bill? stored as Y/N instead of a note?", options: ["Notes take more space", "Y/N can be filtered in one click to find every missing bill", "Excel cannot store notes", "Y/N looks professional"], answer: 1, explanation: "A filter on Bill? = N lists every voucher awaiting a bill instantly. Free-text notes cannot be filtered or counted." }
    ],
    summary: [
      "One row per voucher with a unique number that is never reused.",
      "Dropdowns for Category, Paid By and Mode keep SUMIFS matching possible.",
      "Category summary, reimbursement sheet and mode totals are all single-pair SUMIFS.",
      "Three reports must total the same as the log: 18,010 here \u2014 that equality is the reconciliation check.",
      "Filter Bill? = N to chase missing bills before reimbursement day."
    ],
    downloads: [
      { file: "assets/files/level5/expense-tracker.xlsx", label: "Expense tracker workbook", note: "Exercise and Solution sheets \u2014 November voucher log with category and reimbursement reports" }
    ],
    keywords: ["expense tracker", "reimbursement", "petty cash", "voucher register", "SUMIFS category", "expense log", "reconciliation"]
  },
  {
    id: "l5-invoice-billing",
    title: "Billing and Invoice Creation",
    difficulty: "Intermediate",
    duration: 26,
    version: "",
    objectives: [
      "Build a printable invoice with a proper bill-to block",
      "Pull product rates from the product master with VLOOKUP",
      "Compute discount, GST and the invoice total with correct rounding",
      "Number invoices so no two months ever collide",
      "Feed every invoice into the sales log as one row"
    ],
    intro: "The sales log you built earlier records invoices after the fact. This lesson is where those invoices are born. A good invoice does three jobs at once: it tells the customer exactly what to pay, it gives your accountant a tax document, and it feeds your sales log one clean row.\n\nSunrise Traders invoices with 18% GST on office supplies. In this lesson you will build the invoice layout, pull rates from the product master, apply discount and tax with correct rounding, and connect the result back to the log.",
    keyIdea: "An invoice informs the customer, feeds the accounts, and adds one clean row to your log.",
    sections: [
      {
        heading: "Invoice Numbering That Never Collides",
        blocks: [
          { type: "p", text: "October's sales log used INV-2026-301 onwards. November continues the same series: the first November invoice is INV-2026-350, the second INV-2026-351, and so on. One series, one year, one gapless sequence." },
          { type: "ul", items: [
            "**The number carries the year**, so 2027 starts a fresh series at INV-2027-001 without any risk of collision.",
            "**Numbers are issued in order, never reused.** A cancelled invoice keeps its number and gets a \"CANCELLED\" note \u2014 the gap itself is the audit trail.",
            "**One number per invoice, one invoice per number.** Never split one delivery across two numbers to dodge a customer's purchase-order limit."
          ] },
          { type: "warning", text: "Two people issuing invoices from two copies of the file is how duplicate numbers happen. One person issues numbers, or the number cell is generated from a counter sheet \u2014 never typed from memory." }
        ]
      },
      {
        heading: "The Bill-To Block and the Item Table",
        blocks: [
          { type: "p", text: "An invoice has two parts. The bill-to block says who pays; the item table says for what. Keep each in its own area of the sheet, with the item table as a clean, bordered grid." },
          { type: "grid", title: "Invoice INV-2026-350 \u2014 Northstar Retail, 05-Nov-2026", rows: [
            ["Item", "Qty", "Rate", "Amount"],
            ["HP Laser Printer", "2", "10500", "=B7*C7"],
            ["Ink Cartridge", "10", "850", "=B8*C8"],
            ["A4 Paper (box of 5)", "5", "320", "=B9*C9"],
            ["Subtotal", "", "", "=SUM(D7:D9)"],
            ["Discount", "", "", "1100"],
            ["Taxable Value", "", "", "=D10-D11"],
            ["GST @ 18%", "", "", "=ROUND(D12*0.18, 2)"],
            ["Invoice Total", "", "", "=ROUND(D12+D13, 2)"]
          ] },
          { type: "p", text: "Walk the numbers: subtotal =SUM(D7:D9) is 31,100; less the 1,100 discount leaves a taxable value of 30,000; GST of 18% on 30,000 is exactly 5,400; and the customer pays 35,400. Every figure the customer sees is a formula except the discount, which is a business decision typed by a human." }
        ]
      },
      {
        heading: "The Product Master and VLOOKUP",
        blocks: [
          { type: "p", text: "Rates should never be typed on an invoice. They live once, in the product master, and every invoice looks them up. When the Notebook rate changes in December, every December invoice is correct and no October invoice changes." },
          { type: "grid", title: "Products sheet \u2014 the master", rows: [
            ["Product", "Rate"],
            ["Notebook", "60"],
            ["Ball Pen (box of 100)", "20"],
            ["Stapler", "175"],
            ["HP Laser Printer", "10500"],
            ["A4 Paper (box of 5)", "320"],
            ["Ink Cartridge", "850"]
          ] },
          { type: "formula", name: "Rate lookup", syntax: "=VLOOKUP(lookup_value, table_array, col_index_num, FALSE)", example: "=VLOOKUP(A7, Products!$A$2:$B$7, 2, FALSE)", explanation: "Finds the item name from the invoice in the product master and returns its rate. FALSE demands an exact match \u2014 approximate matching against product names invents prices." },
          { type: "tip", text: "Lock the master range with $ signs ($A$2:$B$7) before copying the formula down the item table. Without the dollars, each copied row looks one row further down the master and starts returning #N/A or, worse, the wrong product's rate." }
        ]
      },
      {
        heading: "Discount, GST and Rounding",
        blocks: [
          { type: "formula", name: "GST after discount", syntax: "=ROUND(taxable_value * rate, 2)", example: "=ROUND(D12*0.18, 2)", explanation: "GST applies to the taxable value \u2014 the subtotal after discount \u2014 never to the pre-discount subtotal. Taxing 31,100 instead of 30,000 overcharges the customer 198 in GST alone." },
          { type: "formula", name: "The grand total", syntax: "=ROUND(taxable_value + gst, 2)", example: "=ROUND(D12+D13, 2)", explanation: "ROUND with 2 digits makes the stored value match the printed invoice to the paisa. Without it, a total of 35399.99999998 prints as 35,400.00 while your accounts carry 35,399.99 \u2014 a one-paise hole in every reconciliation." },
          { type: "p", text: "Order matters and the layout already enforces it: subtotal, discount, taxable value, GST, total. Each row feeds the next, so an auditor can trace any figure on the invoice back to quantities and rates in two steps." }
        ]
      },
      {
        heading: "From Invoice to Sales Log",
        blocks: [
          { type: "p", text: "The moment an invoice is issued, it becomes one row in the sales log: date, invoice number, customer, product lines and the total. The log total must equal the sum of the invoice grand totals \u2014 that is the billing reconciliation." },
          { type: "grid", title: "Sales log \u2014 November so far", rows: [
            ["Date", "Invoice No", "Customer", "Total"],
            ["05-Nov-2026", "INV-2026-350", "Northstar Retail", "35400"],
            ["08-Nov-2026", "INV-2026-351", "Mehta Book House", "80240"],
            ["Total", "", "", "=SUM(D2:D3)"]
          ] },
          { type: "p", text: "Two invoices, 1,15,640 so far for November. If the log ever disagrees with the invoice file, one of the two was edited without the other \u2014 find it the same day, not at month-end." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Typing the rate on the invoice from memory.", right: "Look up every rate with VLOOKUP from the product master.", why: "A retyped rate disagrees with the master the day after a price revision, and every affected invoice must be found and reissued." },
      { wrong: "Computing GST on the subtotal before discount.", right: "Apply 18% to the taxable value: =ROUND((Subtotal - Discount) * 0.18, 2).", why: "The tax law taxes what the customer actually pays. On a 1,100 discount that mistake overcharges 198 in GST and invites a query." },
      { wrong: "Leaving the total unrounded and trusting the cell format.", right: "Wrap the total in ROUND(x, 2) so stored and printed values agree.", why: "Formats change only the display. A total stored as 35399.99999998 disagrees with the accounts by a paisa and the reconciliation never nets to zero." },
      { wrong: "Issuing numbers like INV-350, INV-350A, INV-351-revised.", right: "One gapless numeric series per year: INV-2026-350, 351, 352...", why: "Suffixes and revisions make duplicates undetectable and void the invoice series as an audit trail." }
    ],
    shortcuts: [
      { keys: "Ctrl + P", action: "Open print preview to check the invoice fits one page" },
      { keys: "Ctrl + 1", action: "Format the rate and amount columns as currency" },
      { keys: "Ctrl + D", action: "Copy the item row's formulas down for the next line item" },
      { keys: "F4", action: "Toggle $ locking on a selected master range while editing the VLOOKUP" }
    ],
    exercises: [
      { task: "In the billing-records workbook, verify invoice INV-2026-350 line by line: subtotal, taxable value, GST and total.", hint: "Subtotal = 2\u00d710500 + 10\u00d7850 + 5\u00d7320; GST applies after the 1,100 discount.", solution: "Subtotal 31,100; taxable value 30,000; GST 5,400; invoice total 35,400." },
      { task: "Build invoice INV-2026-351 for Mehta Book House: 500 Notebooks at 60 and 2,000 Ball Pen boxes at 20, with a 2,000 discount.", hint: "Pull both rates from the product master; discount comes off before GST.", solution: "Subtotal 70,000; discount 2,000; taxable value 68,000; GST 12,240; total 80,240." },
      { task: "Add both November invoices to the sales log and prove the log reconciles to the invoice file.", hint: "One row per invoice; log total minus invoice total must be zero.", solution: "35,400 + 80,240 = 1,15,640, and the log's SUM over the two rows equals the same figure \u2014 reconciliation passes." }
    ],
    quiz: [
      { q: "The Amount cell of an invoice line should contain:", options: ["The typed line total", "=Qty * Rate", "=Rate / Qty", "A link to the customer's purchase order"], answer: 1, explanation: "Amount is always Qty times Rate. A typed total hides an error in either input and never updates when one is corrected." },
      { q: "GST of 18% applies to:", options: ["The subtotal before discount", "The taxable value after discount", "The quantity", "The rate"], answer: 1, explanation: "Discount comes off first; the tax applies to what the customer actually pays. GST on the pre-discount subtotal overcharges the customer." },
      { q: "Why look rates up with VLOOKUP instead of typing them?", options: ["VLOOKUP is faster to type", "The product master is the single source; a revision flows to every new invoice automatically", "Typed rates are not allowed by Excel", "VLOOKUP rounds automatically"], answer: 1, explanation: "One price list, many invoices. Typed rates freeze at typing time and drift from the master the moment any rate changes." },
      { q: "Subtotal 31,100, discount 1,100, GST 18%. What does the customer pay?", options: ["35,398", "36,498", "35,400", "30,540"], answer: 2, explanation: "Taxable value 30,000; GST 5,400; total 35,400. Option B taxes the pre-discount subtotal \u2014 the classic overcharge." },
      { q: "What enters the sales log when an invoice is issued?", options: ["Every line item as a separate row", "One row per invoice, with the total equal to the invoice grand total", "Only the customer name and date", "Nothing until the customer pays"], answer: 1, explanation: "One row per invoice keeps the log a log. Its total must reconcile to the sum of invoice grand totals \u2014 that equality is the billing check." }
    ],
    summary: [
      "One gapless invoice series per year; cancelled invoices keep their number with a note.",
      "Bill-to block plus an item table where Amount = Qty \u00d7 Rate, never typed.",
      "Rates come from the product master by exact-match VLOOKUP with locked ranges.",
      "Discount first, then GST on the taxable value, with ROUND(x, 2) on the tax and total.",
      "Every invoice becomes one row in the sales log; log total must equal invoice total."
    ],
    downloads: [
      { file: "assets/files/level5/billing-records.xlsx", label: "Billing records workbook", note: "Exercise and Solution sheets \u2014 product master and the two November invoices to build" }
    ],
    keywords: ["invoice", "billing", "GST 18 percent", "taxable value", "discount", "VLOOKUP product master", "invoice numbering"]
  },
  {
    id: "l5-mis-reports",
    title: "Management Reports and MIS",
    difficulty: "Intermediate",
    duration: 25,
    version: "",
    objectives: [
      "Assemble the month's key numbers on a single MIS sheet",
      "Build person-wise and product-wise sales reports with SUMIFS",
      "Track target achievement for every salesperson",
      "Report office expenses by head",
      "Reconcile every report back to its source log"
    ],
    intro: "MIS stands for Management Information System, but in a small office it means something simpler: one sheet the owner can open on the 1st and read the whole month from. Not ten files, not a folder of screenshots \u2014 one page of numbers that all reconcile.\n\nIn this lesson you will build the October 2026 MIS for Sunrise Traders: sales by person, sales by product, payroll and office expense, every block computed by formula from the logs you built in the earlier lessons.",
    keyIdea: "One MIS page, computed by formula from your logs, reads the whole month at a glance.",
    sections: [
      {
        heading: "MIS: One Sheet That Answers the Month",
        blocks: [
          { type: "p", text: "Before any block is built, agree on the questions the MIS must answer. For Sunrise Traders there are four: How were sales against target? Which products and which people drove them? What did salaries cost? What did the office itself cost?" },
          { type: "grid", title: "MIS \u2014 October 2026, headline block", rows: [
            ["Measure", "Value", "Against", "Reading"],
            ["Total sales", "3058500", "Target 2910000", "=ROUND(B2/2910000*100, 1)"],
            ["Gross payroll", "271600", "Budget head", "Matches payroll sheet"],
            ["Net paid out", "240760", "Bank transfer", "Matches payslip totals"],
            ["Office expenses", "304610", "Budget 303760", "=ROUND(B5/303760*100, 1)"]
          ] },
          { type: "p", text: "Sales finished at 105.1% of target and office spending at 100.6% of budget \u2014 the whole month in two percentages. The rest of the MIS explains those two numbers." }
        ]
      },
      {
        heading: "Person-wise Sales Report",
        blocks: [
          { type: "formula", name: "Actual per salesperson", syntax: "=SUMIFS(totals, salesperson_range, name)", example: "=SUMIFS($G$2:$G$100, $H$2:$H$100, \"Arjun Mehta\")", explanation: "Sums the invoice totals of the October sales log for one salesperson. Because the log's Salesperson column is a validated dropdown, every row matches exactly." },
          { type: "grid", title: "Sales by person \u2014 October 2026", rows: [
            ["Salesperson", "Target", "Actual", "Achievement %", "Status"],
            ["Arjun Mehta", "735000", "=SUMIFS('Sales Log'!$G$2:$G$100, 'Sales Log'!$H$2:$H$100, \"Arjun Mehta\")", "=ROUND(C2/B2*100, 1)", "=IF(D2>=100, \"On target\", \"Below\")"],
            ["Priya Nair", "1000000", "=SUMIFS('Sales Log'!$G$2:$G$100, 'Sales Log'!$H$2:$H$100, \"Priya Nair\")", "=ROUND(C3/B3*100, 1)", "=IF(D3>=100, \"On target\", \"Below\")"],
            ["Rahul Verma", "750000", "=SUMIFS('Sales Log'!$G$2:$G$100, 'Sales Log'!$H$2:$H$100, \"Rahul Verma\")", "=ROUND(C4/B4*100, 1)", "=IF(D4>=100, \"On target\", \"Below\")"],
            ["Sneha Iyer", "425000", "=SUMIFS('Sales Log'!$G$2:$G$100, 'Sales Log'!$H$2:$H$100, \"Sneha Iyer\")", "=ROUND(C5/B5*100, 1)", "=IF(D5>=100, \"On target\", \"Below\")"],
            ["Total", "=SUM(B2:B5)", "=SUM(C2:C5)", "=ROUND(C6/B6*100, 1)", ""]
          ] },
          { type: "p", text: "Arjun 8,50,000 and Priya 10,84,000 finished on target; Rahul 7,27,200 and Sneha 3,97,300 finished below. The four actuals add to exactly 30,58,500 \u2014 the sales log total. If they do not, a name variant or a missing row is hiding somewhere." }
        ]
      },
      {
        heading: "Product-wise Sales Report",
        blocks: [
          { type: "formula", name: "Total per product", syntax: "=SUMIFS(totals, product_range, product)", example: "=SUMIFS($G$2:$G$100, $E$2:$E$100, \"Notebook\")", explanation: "Same pattern as the person report, one column to the left. One SUMIFS per product row \u2014 no need for PivotTables until the product list outgrows the page." },
          { type: "grid", title: "Sales by product \u2014 October 2026", rows: [
            ["Product", "Total", "Share"],
            ["Notebook", "=SUMIFS($G$2:$G$100, $E$2:$E$100, \"Notebook\")", "=ROUND(B2/$B$8*100, 1)"],
            ["A4 Paper (box of 5)", "=SUMIFS($G$2:$G$100, $E$2:$E$100, \"A4 Paper (box of 5)\")", "=ROUND(B3/$B$8*100, 1)"],
            ["HP Laser Printer", "=SUMIFS($G$2:$G$100, $E$2:$E$100, \"HP Laser Printer\")", "=ROUND(B4/$B$8*100, 1)"],
            ["Ink Cartridge", "=SUMIFS($G$2:$G$100, $E$2:$E$100, \"Ink Cartridge\")", "=ROUND(B5/$B$8*100, 1)"],
            ["Ball Pen (box of 100)", "=SUMIFS($G$2:$G$100, $E$2:$E$100, \"Ball Pen (box of 100)\")", "=ROUND(B6/$B$8*100, 1)"],
            ["Stapler", "=SUMIFS($G$2:$G$100, $E$2:$E$100, \"Stapler\")", "=ROUND(B7/$B$8*100, 1)"],
            ["Total", "=SUM(B2:B7)", ""]
          ] },
          { type: "p", text: "Notebooks lead at 12,40,000 (40.5% of the month), A4 Paper follows at 6,80,000, and printers plus cartridges together add 8,80,000 \u2014 hardware drives the big single invoices. The shares only mean something because the six totals add back to 30,58,500." }
        ]
      },
      {
        heading: "Payroll and Office Expense Blocks",
        blocks: [
          { type: "p", text: "The bottom half of the MIS quotes two sheets you have already built. Payroll: gross 2,71,600 booked as salary expense, net 2,40,760 actually paid. Office expense: the budget sheet's actual column, 3,04,610 against a 3,03,760 budget." },
          { type: "grid", title: "Office expense by head \u2014 October 2026", rows: [
            ["Head", "Budget", "Actual", "Variance"],
            ["Office Rent", "30000", "30000", "=C2-B2"],
            ["Electricity", "8500", "9100", "=C3-B3"],
            ["Internet and Phone", "2500", "2650", "=C4-B4"],
            ["Transport and Delivery", "6000", "5400", "=C5-B5"],
            ["Stationery", "3000", "3200", "=C6-B6"],
            ["Marketing", "8000", "7600", "=C7-B7"],
            ["Miscellaneous", "5000", "5900", "=C8-B8"],
            ["Operating total", "=SUM(B2:B8)", "=SUM(C2:C8)", "=C9-B9"]
          ] },
          { type: "p", text: "Salaries stay on their own line because they are contracted, not discretionary. The operating heads total 63,000 budgeted and 63,850 actual (+850) \u2014 this is the block where management can actually change something next month." }
        ]
      },
      {
        heading: "The Rule: Every Block Reconciles",
        blocks: [
          { type: "steps", items: [
            "Person report total = product report total = sales log total = 30,58,500.",
            "MIS salary expense = payroll gross = 2,71,600; MIS net = bank transfer = 2,40,760.",
            "Office expense block = budget sheet actual column = 3,04,610.",
            "Run all three checks on the 1st of the next month, before the MIS is circulated.",
            "Only then add the commentary lines \u2014 notes on a sheet that fails reconciliation are notes about wrong numbers."
          ] },
          { type: "warning", text: "Never type a report number into the MIS \"just this once\". A retyped number looks right today and is silently wrong next month when the source changes. If a figure cannot yet be computed by formula, put the formula in anyway and let it show 0 or an error until the source is ready \u2014 errors get fixed, stale numbers get believed." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Retyping the month's totals into the MIS by hand.", right: "Compute every block with SUMIFS and VLOOKUP from the logs.", why: "Typed numbers are frozen at typing time. When the sales log gains a late invoice, the MIS lies about it until someone remembers to retype." },
      { wrong: "Circulating the MIS without running the reconciliation checks.", right: "Prove person = product = log total before the sheet leaves your desk.", why: "A MIS that reconciles earns trust for a year. One MIS that contradicts the accounts destroys trust for a year." },
      { wrong: "Using date criteria like \"October\" against the log.", right: "Bracket the month with two DATE values: >= DATE(2026,10,1) and < DATE(2026,11,1).", why: "The log stores real dates, so text criteria match nothing and the report quietly shows zero." },
      { wrong: "Hiding source rows or columns inside the MIS sheet to tidy it up.", right: "Keep sources on their own sheets and reference them across sheets.", why: "Hidden rows still feed SUMIFS, so the sheet looks clean while carrying data nobody can see or audit." }
    ],
    shortcuts: [
      { keys: "Ctrl + Page Down", action: "Move to the next sheet when collecting numbers from the logs" },
      { keys: "Ctrl + Page Up", action: "Move back to the MIS sheet to place the formula" },
      { keys: "F9", action: "Recalculate the whole workbook after a late correction" },
      { keys: "Alt + F1", action: "Insert a column chart of the person-wise report in one step" }
    ],
    exercises: [
      { task: "In the mis-reports workbook, complete the person-wise table: actuals by SUMIFS, achievement to one decimal, and status.", hint: "Targets: 7,35,000 / 10,00,000 / 7,50,000 / 4,25,000. Achievement = ROUND(actual/target*100, 1).", solution: "Arjun 8,50,000 (115.6%), Priya 10,84,000 (108.4%), Rahul 7,27,200 (97.0%), Sneha 3,97,300 (93.5%). Total 30,58,500 at 105.1%. Arjun and Priya on target; Rahul and Sneha below." },
      { task: "Build the product-wise report and prove it reconciles to the sales log.", hint: "One SUMIFS per product; the six totals must sum to the log total.", solution: "Notebook 12,40,000; A4 Paper 6,80,000; Printer 5,20,000; Ink Cartridge 3,60,000; Pen 1,58,500; Stapler 1,00,000 \u2014 total 30,58,500, equal to the log." },
      { task: "Complete the office-expense block and state the two headline percentages of the month.", hint: "Variance = actual - budget; compare each total to its own benchmark.", solution: "Operating heads (incl. rent): 63,000 budget, 63,850 actual (+850). Whole-month readings: sales 105.1% of target, office expense 100.6% of budget (3,04,610 vs 3,03,760)." }
    ],
    quiz: [
      { q: "Where should the MIS's person-wise actuals come from?", options: ["Typed from the sales log once a month", "SUMIFS over the sales log, recomputed live", "The payroll sheet", "The budget sheet"], answer: 1, explanation: "Formulas stay correct when the source changes. Typed numbers are frozen and go stale with every late invoice." },
      { q: "Which formula computes achievement percentage for the MIS?", options: ["=C2-B2", "=C2/B2", "=ROUND(C2/B2*100, 1)", "=PERCENT(C2, B2)"], answer: 2, explanation: "Achievement is actual divided by target as a percentage, rounded to one decimal so stored and displayed values agree." },
      { q: "The person-wise actuals total 30,58,500 but the sales log totals 30,60,000. What happened?", options: ["The MIS is more accurate than the log", "1,500 of rows are missing or mismatched in the MIS ranges \u2014 reconcile before circulating", "Excel dropped the difference", "The log has a typo in its total"], answer: 1, explanation: "Two reports from one source must agree exactly. A 1,500 gap means a row is missing, a name variant went unmatched, or a range is short \u2014 find it first." },
      { q: "Which criteria correctly restrict a SUMIFS to October 2026?", options: [">= DATE(2026,10,1) and < DATE(2026,11,1)", "The word \"October\"", "Month = 10 on its own", "> 01/10/2026 only"], answer: 0, explanation: "A month is a range of dates. The lower bracket includes the 1st; the upper bracket excludes the 1st of the next month. Text like \"October\" matches nothing against real dates." },
      { q: "Why must the office-expense block agree with the budget sheet's actual column?", options: ["It looks consistent", "Both must describe the same October \u2014 a mismatch means one sheet was edited without the other", "Excel requires it", "The budget is always larger"], answer: 1, explanation: "Same month, same heads, same source \u2014 two views of one truth. Any difference is an editing error somewhere, and the check exists to catch it." }
    ],
    summary: [
      "One MIS sheet with a headline block, person-wise sales, product-wise sales, payroll and office expense.",
      "Every actual is a SUMIFS over a log; nothing is retyped.",
      "Person report total = product report total = log total = 30,58,500 for October.",
      "Payroll block quotes gross 2,71,600 and net 2,40,760; expense block matches the budget's actual column 3,04,610.",
      "Run all reconciliation checks before the MIS is circulated \u2014 commentary comes after proof."
    ],
    downloads: [
      { file: "assets/files/level5/mis-reports.xlsx", label: "MIS reports workbook", note: "Exercise and Solution sheets \u2014 October MIS blocks to complete with SUMIFS" }
    ],
    keywords: ["MIS report", "management report", "sales by person", "sales by product", "achievement percentage", "reconciliation", "monthly report"]
  },
  {
    id: "l5-office-capstone",
    title: "Office Data Management Capstone",
    difficulty: "Intermediate",
    duration: 30,
    version: "",
    objectives: [
      "Link employee, attendance, payroll, sales and expense sheets into one system",
      "Use Emp IDs, invoice numbers and vouchers as the keys between sheets",
      "Run month-end reconciliation checks across every sheet",
      "Protect master data from accidental edits",
      "Close a month and start the next one without breaking history"
    ],
    intro: "This lesson ties Level 5 together. You have built the pieces one at a time \u2014 employee master, attendance, leave, payroll, inventory, purchase, sales, budget, expense log, invoices and MIS. Now they become one system: a single workbook where every sheet has a job, every sheet is reachable by a key, and the month closes with a checklist instead of a panic.\n\nThe capstone workbook holds October 2026 closed and November 2026 open, exactly as an office would.",
    keyIdea: "The capstone joins every Level 5 sheet into one workbook where each sheet has a job.",
    sections: [
      {
        heading: "One Workbook, Many Sheets, One Set of Keys",
        blocks: [
          { type: "p", text: "A system is a set of sheets that reference each other through keys. The key rule at Sunrise Traders: every sheet can be traced back to the Employee Master, the Product Master or a numbered document." },
          { type: "table", headers: ["Sheet", "Its key", "Feeds"], rows: [
            ["Employee Master", "Emp ID (EMP-001...)", "Attendance, Leave, Payroll"],
            ["Attendance (per month)", "Emp ID + the month sheet", "Payroll working days"],
            ["Leave Log", "Req No (LV-...)", "Leave balance sheet"],
            ["Sales Log", "Invoice No (INV-2026-...)", "MIS, budget income"],
            ["Expense Log", "Voucher No (EXP-2026-...)", "Reimbursement, budget actuals"],
            ["Product Master", "Product name", "Invoices, inventory"],
            ["MIS", "Formulas only \u2014 no typed numbers", "The owner's monthly read"]
          ] },
          { type: "p", text: "Notice what the MIS row says: formulas only. The MIS owns no data of its own. Every number on it is borrowed through a formula, which is why it can never disagree with the sheets underneath it." }
        ]
      },
      {
        heading: "How Data Flows Through the Month",
        blocks: [
          { type: "ul", items: [
            "**Daily:** every invoice adds one row to the sales log; every office payment adds one voucher row to the expense log.",
            "**Weekly:** attendance is marked P/A/L/H; the expense log's missing bills are chased.",
            "**On event:** leave requests go into the Leave Log and get approved or rejected there, not in chat messages.",
            "**Month-end:** payroll runs from the master and attendance; the budget's actual column fills from payroll and the expense log; the MIS recompute and the reconciliation checklist runs.",
            "**On the 1st:** a fresh copy of the workbook starts the new month; the closed copy is archived and never edited again."
          ] },
          { type: "p", text: "The flow has one direction: documents create log rows, logs feed reports, reports feed decisions. Nothing flows backwards \u2014 nobody \"fixes\" the MIS by typing into it." }
        ]
      },
      {
        heading: "The Month-End Reconciliation Checklist",
        blocks: [
          { type: "p", text: "Closing October means proving five equalities. Each is one subtraction that must show zero." },
          { type: "grid", title: "October close \u2014 checklist", rows: [
            ["Check", "Formula", "Result"],
            ["Sales log total = MIS person report total", "=3058500-3058500", "0 \u2014 pass"],
            ["Sales log total = MIS product report total", "=3058500-3058500", "0 \u2014 pass"],
            ["Payroll gross = MIS salary expense", "=271600-271600", "0 \u2014 pass"],
            ["Payroll net = budget salaries actual", "=240760-240760", "0 \u2014 pass"],
            ["Budget expense total = expense actuals", "=304610-304610", "0 \u2014 pass"]
          ] },
          { type: "p", text: "Then the two partial-month items are noted, not closed: November's two invoices (INV-2026-350 and 351) total 1,15,640 so far, and the expense log's first half stands at 18,010 with one bill still to arrive. A close that says what is still open is more honest than one that pretends everything is done." },
          { type: "formula", name: "The zero test", syntax: "=source_total - report_total", example: "=SUM('Sales Log'!$G$2:$G$100)-C6", explanation: "Every reconciliation is a subtraction that must equal zero. Format the check cell with a custom format that shows \"PASS\" for 0 and the difference otherwise: 0\" \\u2014 pass\";0.00\" \\u2014 CHECK\";0.00\" \\u2014 CHECK\"." }
        ]
      },
      {
        heading: "Protection, Validation and Backups",
        blocks: [
          { type: "steps", items: [
            "Protect the Employee Master and Product Master sheets (Review, Protect Sheet) with a password the office actually remembers.",
            "Keep Data Validation on every keyed column, so new rows inherit the dropdowns.",
            "Unlock only the cells meant for typing (Format Cells, Protection, uncheck Locked) before protecting a transaction sheet.",
            "Save a dated backup at each month-end close: Sunrise-Traders-2026-10-CLOSED.xlsx.",
            "Store the backup somewhere the daily machine cannot wipe it \u2014 a second drive or cloud folder counts."
          ] },
          { type: "warning", text: "Protection is not security against a determined person; it is a fence against accidents. A sort gone wrong, a stray Delete on column F, a paste that overwrites formulas \u2014 protection and validation stop all of these without slowing down daily work." }
        ]
      },
      {
        heading: "Starting November Without Breaking October",
        blocks: [
          { type: "steps", items: [
            "Save the closed October workbook as Sunrise-Traders-2026-10-CLOSED.xlsx and never open it for editing again.",
            "Copy the file and rename it Sunrise-Traders-2026-11.xlsx.",
            "In the new copy, clear the transaction rows of the sales log and expense log \u2014 masters, budget structure and MIS layout stay.",
            "Move the budget sheet to a fresh November budget column with October actuals frozen beside it for reference.",
            "Continue the invoice and voucher series \u2014 November starts at INV-2026-350 and EXP-2026-031, exactly where October stopped."
          ] },
          { type: "p", text: "This is why the series numbers carry the year and never restart mid-year: the archived October file and the live November file form one continuous story. Six months from now, \"what did we sell to Northstar Retail in October?\" has one authoritative answer, in one closed file that nobody has touched since." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Editing the Employee Master after the month's payroll has run.", right: "Record the change with an effective date and apply it from the next month's payroll.", why: "A mid-month salary edit silently rewrites closed payslips and the archived payroll stops matching what was actually paid." },
      { wrong: "Deleting the sales log or expense log after the MIS is built.", right: "Logs are the system's memory \u2014 archive them, never delete them.", why: "The MIS recomputes from its sources. Delete a source and every report on it turns to zero or #REF! with no undo." },
      { wrong: "Keeping employee data, sales, expenses and reports on one giant sheet.", right: "One sheet per job, linked by keys.", why: "A single mixed sheet cannot be filtered, protected or reconciled \u2014 and one accidental sort scrambles everything at once." },
      { wrong: "Restarting invoice numbers every month: INV-2026-301 in October and again in November.", right: "One continuous series per year across all months.", why: "Duplicate invoice numbers make duplicate payments undetectable and destroy the audit trail the numbering exists to create." }
    ],
    shortcuts: [
      { keys: "Ctrl + Page Down", action: "Walk the sheets while running the reconciliation checklist" },
      { keys: "F12", action: "Save As the closed workbook under its dated CLOSED name" },
      { keys: "Ctrl + S", action: "Save after every checklist item that shows a zero" },
      { keys: "Ctrl + Z", action: "Undo an accidental edit \u2014 before protection, this is the only fence" }
    ],
    exercises: [
      { task: "In the capstone workbook, run the full October reconciliation checklist and show every check as zero.", hint: "Five subtractions: sales twice, payroll twice, budget expenses once.", solution: "All five pass: 30,58,500 = person report = product report; gross 2,71,600 = salary expense; net 2,40,760 = salaries actual; budget total 3,04,610 = expense actuals." },
      { task: "Post November's two invoices (INV-2026-350: 35,400 and INV-2026-351: 80,240) to the sales log and reconcile the partial month.", hint: "One row per invoice; log total minus invoice total = 0.", solution: "Log rows sum to 1,15,640, equal to the two invoice totals \u2014 the November-so-far check passes." },
      { task: "Prepare the November workbook: close October under its CLOSED name, clear the transaction logs, and set the November series starts.", hint: "Copy, rename, clear transaction rows only; masters and structure stay.", solution: "Sunrise-Traders-2026-11.xlsx carries the masters, an empty sales log and expense log, and continues at INV-2026-350 / EXP-2026-031." }
    ],
    quiz: [
      { q: "Which key connects the Payroll sheet to the Employee Master?", options: ["The employee's name", "The Emp ID", "The row number", "The salary amount"], answer: 1, explanation: "Names can repeat and row numbers shift on every sort; the Emp ID is unique forever. Payroll looks Basic up by Emp ID with VLOOKUP." },
      { q: "What must every MIS block do before the sheet is circulated?", options: ["Look colourful", "Reconcile to its source log \u2014 every check subtracts to zero", "Fit one printed page", "Include last year's numbers"], answer: 1, explanation: "Reconciliation is the MIS's proof of correctness. Notes and commentary only matter once the numbers are proven." },
      { q: "Why protect the Employee Master sheet?", options: ["To hide salaries from employees", "To stop accidental edits that would silently rewrite payroll and history", "Because Excel files need passwords", "To make sorting impossible"], answer: 1, explanation: "One stray keystroke in the master corrupts every sheet that looks up to it. Protection is the fence that keeps accidents out." },
      { q: "How does November begin without damaging October?", options: ["Delete October's rows from the live file", "Archive October as a CLOSED copy, then work in a fresh copy with cleared transaction logs and kept masters", "Keep both months in one log forever", "Email October's file to the owner"], answer: 1, explanation: "The closed file is October's permanent record; the fresh copy starts November clean while the masters, structure and series continue." },
      { q: "October's budget actual expense total is 3,04,610. Where must this same figure appear?", options: ["Only in the budget sheet", "In the budget sheet, the MIS expense block and nowhere else disagreeing", "In every sheet, retyped", "In the expense log's total row"], answer: 1, explanation: "The budget's actual column and the MIS quote the same October from the same sources. If either shows something else, one of them was edited by hand." }
    ],
    summary: [
      "One workbook, one sheet per job, all sheets linked by keys: Emp ID, invoice number, voucher number.",
      "Data flows one way \u2014 documents to logs, logs to reports \u2014 and the MIS owns no typed numbers.",
      "Month-end is a checklist of subtractions that must each show zero.",
      "Protect masters, validate keyed columns, and save a dated CLOSED backup at every close.",
      "Start the new month in a fresh copy: transactions cleared, masters kept, series continued."
    ],
    downloads: [
      { file: "assets/files/level5/office-data-capstone.xlsx", label: "Office data capstone workbook", note: "Exercise and Solution sheets \u2014 the linked Sunrise Traders system with October closed and November open" }
    ],
    keywords: ["office system", "linked sheets", "reconciliation checklist", "month-end close", "sheet protection", "capstone", "data management"]
  }
);
