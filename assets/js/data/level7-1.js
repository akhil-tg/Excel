// Excel Lab — Level 7 (Practical Projects), part 1 of 2
// Level metadata + the first five project lessons

window.LEVEL_META = window.LEVEL_META || {};
window.LEVEL_LESSONS = window.LEVEL_LESSONS || {};

window.LEVEL_META[7] = {
  id: 7,
  title: "Practical Projects",
  icon: "folder",
  difficulty: "Advanced",
  subtitle: "Ten complete office systems you build yourself, step by step",
  description: "Apply everything: build complete, working office systems — employee management, payroll, inventory, invoicing and professional dashboards — with sample data and solutions."
};

(window.LEVEL_LESSONS[7] = window.LEVEL_LESSONS[7] || []).push(
  {
    id: "l7-project-employee",
    title: "Project: Employee Management System",
    difficulty: "Intermediate",
    duration: 40,
    version: "",
    objectives: [
      "Build an employee master that other sheets can trust, with dropdowns for Department and Status",
      "Create attendance and salary summaries that fill themselves with VLOOKUP",
      "Build a department dashboard using COUNTIF and SUMIF",
      "Apply conditional formatting so problems stand out on their own",
      "Connect four sheets into one system where a single edit updates everything"
    ],
    intro: "You have learned the tools. Now use them on a real job. In this project you build the HR system of Sunrise Traders, a wholesale office supplies company. It has six employees: Arjun Mehta, Priya Nair, Rahul Verma, Sneha Iyer, Vikram Singh and Divya Menon. The finished system has four sheets. HR types each person once in the Employee Master. Then an Attendance Summary and a Salary Summary fill themselves. A small dashboard answers the questions a manager asks: how many people sit in each department, who is still active, and what the monthly salary bill looks like.\n\nThis project introduces no new functions. Everything comes from Levels 2 to 6: data validation, VLOOKUP, COUNTIF, SUMIF and conditional formatting. Your job is to combine them so they work together. That is exactly what real office Excel work is. The exercise workbook has a starter file and a completed solution, so you can check every step. Everything here works in Excel 2010 and later.",
    keyIdea: "You build a four-sheet HR system where every sheet fills itself from one employee master.",
    sections: [
      {
        heading: "What You Will Build",
        blocks: [
          { type: "p", text: "The finished system does four things. HR types each employee once in the Employee Master. The other sheets pull the name automatically after that. Department and Status cells accept only dropdown values. So nobody can type \"Sels\" or \"activee\" and silently break the counts. The Attendance Summary and Salary Summary calculate themselves. You just enter IDs and day counts. And the dashboard shows every total at one glance." },
          { type: "ul", items: [
            "**Employee Master** — one row per employee: ID, name, department, designation, join date, status and basic salary, with dropdowns on Department and Status.",
            "**Attendance Summary** — days present, absent and on leave per employee, with an automatic attendance percentage.",
            "**Salary Summary** — basic salary pulled from the master, plus HRA at 40% and a gross figure.",
            "**Dashboard** — headcount and salary cost per department, active headcount, and total payroll, all live."
          ] },
          { type: "p", text: "This is the pattern behind almost every small-office Excel system. One master sheet holds clean data. Summary sheets look values up. A dashboard counts and adds. Once you build it here, you can rebuild it for any business." }
        ]
      },
      {
        heading: "Skills You Will Use",
        blocks: [
          { type: "ul", items: [
            "**Data validation dropdowns** (Level 3) — restrict Department and Status to fixed lists.",
            "**VLOOKUP** (Level 5) — pull names, departments and salaries from the master into the summaries.",
            "**COUNTIF and SUMIF** (Level 4) — count and total by department and by status on the dashboard.",
            "**Absolute references** (Level 4) — lock lookup ranges so they survive the fill handle.",
            "**Conditional formatting** (Level 6) — highlight resignations and poor attendance automatically."
          ] },
          { type: "p", text: "If any of these feel shaky, open the matching level lesson for a refresher before you start. The project moves quickly because it assumes you know where the buttons are." }
        ]
      },
      {
        heading: "Open the Exercise File",
        blocks: [
          { type: "p", text: "Download **Employee Management System — exercise file** at the bottom of this lesson and open it. You will find four sheets:" },
          { type: "ul", items: [
            "**Employee Master** — six employees already filled in, but no dropdowns and no formatting yet.",
            "**Attendance Summary** — only the header row exists. You will build the rest.",
            "**Salary Summary** — only the header row exists.",
            "**Dashboard** — completely empty. You will design it from scratch."
          ] },
          { type: "tip", text: "Keep the solution file closed while you build. Peek only when you are stuck, and then close it again — comparing finished work teaches far more than copying it." }
        ]
      },
      {
        heading: "Build Section 1: Lock In the Employee Master",
        blocks: [
          { type: "p", text: "The master is the foundation. If wrong values get in here, every other sheet silently goes wrong too, so we start by making the two text columns impossible to mistype." },
          { type: "steps", items: [
            "On the **Employee Master** sheet, select `C2:C11` — six filled rows plus four spare rows for future hires.",
            "Go to **Data → Data Validation → Settings tab**, choose **List**, and in Source type `Sales,Stores,Accounts,HR` (commas between values, no spaces). Click OK.",
            "Select `F2:F11`, repeat the same steps, and use the source `Active,On Leave,Resigned`.",
            "Press `[[Alt]] + [[Down]]` on any cell in either column to confirm the dropdown opens.",
            "Select `E2:E11`, press `[[Ctrl]] + [[1]]`, choose **Date**, pick the format **14-Mar-12**, and click OK so all join dates read the same way.",
            "Save the workbook."
          ] },
          { type: "grid", title: "Employee Master after setup (four spare rows sit below E106)",
            rows: [
              ["Emp ID", "Name", "Department", "Designation", "Join Date", "Status", "Basic Salary (₹)"],
              ["E101", "Arjun Mehta", "Sales", "Sales Executive", "15-Mar-21", "Active", "32000"],
              ["E102", "Priya Nair", "Accounts", "Accountant", "02-Jun-20", "Active", "38000"],
              ["E103", "Rahul Verma", "Stores", "Store Keeper", "10-Jan-22", "Active", "26000"],
              ["E104", "Sneha Iyer", "HR", "HR Executive", "21-Jul-21", "On Leave", "30000"],
              ["E105", "Vikram Singh", "Sales", "Sales Manager", "05-Sep-19", "Active", "55000"],
              ["E106", "Divya Menon", "Sales", "Sales Executive", "18-Feb-23", "Resigned", "28000"]
            ] },
          { type: "warning", text: "Apply the dropdowns down to row 11, not just row 7. The first new hire past row 7 needs a dropdown too. One mistyped department there throws off every count on the dashboard." }
        ]
      },
      {
        heading: "Build Section 2: Attendance Summary Fills Itself",
        blocks: [
          { type: "p", text: "On this sheet HR types only the ID and three day counts each month. Names and percentages compute themselves." },
          { type: "steps", items: [
            "On **Attendance Summary**, type the IDs `E101` to `E106` into `A2:A7`.",
            "In `B2` type `=VLOOKUP($A2,'Employee Master'!$A$2:$B$7,2,FALSE)` and fill it down to `B7`. Press `[[F4]]` while the range is selected to add the dollar signs.",
            "In `C2:E7` enter the day counts for the month. Use these sample figures: Arjun 25/0/0, Priya 24/0/1, Rahul 23/2/0, Sneha 20/0/5, Vikram 25/0/0, Divya 22/1/2 (present/absent/leave).",
            "In `F2` type `=C2+D2+E2` and fill down — the total days each employee was on the rolls.",
            "In `G2` type `=C2/F2`, fill down, then select `G2:G7`, press `[[Ctrl]] + [[1]]` and format as **Percentage** with 0 decimals."
          ] },
          { type: "grid", title: "Attendance Summary — what you should see",
            rows: [
              ["Emp ID", "Name", "Present", "Absent", "Leave", "Total Days", "Attendance %"],
              ["E101", "Arjun Mehta", "25", "0", "0", "25", "100%"],
              ["E102", "Priya Nair", "24", "0", "1", "25", "96%"],
              ["E103", "Rahul Verma", "23", "2", "0", "25", "92%"],
              ["E104", "Sneha Iyer", "20", "0", "5", "25", "80%"],
              ["E105", "Vikram Singh", "25", "0", "0", "25", "100%"],
              ["E106", "Divya Menon", "22", "1", "2", "25", "88%"]
            ] },
          { type: "formula", name: "VLOOKUP into another sheet",
            syntax: "=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])",
            example: "=VLOOKUP($A2,'Employee Master'!$A$2:$B$7,2,FALSE)",
            explanation: "Finds the ID from A2 in column A of the Employee Master and returns the value from the 2nd column of the range — the name. FALSE forces an exact match. The dollar signs lock the range so filling down does not slide it off the data." },
          { type: "tip", text: "Type the ID in A2 first. Get the VLOOKUP right before filling down. If B2 shows `#N/A`, the ID does not exist in the master. Check for a typo or a stray space." }
        ]
      },
      {
        heading: "Build Section 3: Salary Summary with Two Lookups",
        blocks: [
          { type: "p", text: "This sheet proves why the master is worth the effort: the salary numbers here update the moment HR edits a basic salary." },
          { type: "steps", items: [
            "On **Salary Summary**, again type `E101` to `E106` in `A2:A7`.",
            "In `B2` use the same VLOOKUP as before and fill down to `B7`.",
            "In `C2` type `=VLOOKUP($A2,'Employee Master'!$A$2:$G$7,3,FALSE)` and fill down — this pulls the department (column 3 of the wider range).",
            "In `D2` type `=VLOOKUP($A2,'Employee Master'!$A$2:$G$7,7,FALSE)` and fill down — column 7 is the basic salary.",
            "In `E2` type `=ROUND($D2*40%,0)` and fill down. This is HRA — House Rent Allowance, a standard part of Indian salaries, commonly set at 40% of basic for private companies. ROUND keeps the amount in whole rupees.",
            "In `F2` type `=$D2+$E2` and fill down for the gross figure."
          ] },
          { type: "grid", title: "Salary Summary — what you should see",
            rows: [
              ["Emp ID", "Name", "Department", "Basic (₹)", "HRA (₹)", "Gross (₹)", ""],
              ["E101", "Arjun Mehta", "Sales", "32000", "12800", "44800", ""],
              ["E102", "Priya Nair", "Accounts", "38000", "15200", "53200", ""],
              ["E103", "Rahul Verma", "Stores", "26000", "10400", "36400", ""],
              ["E104", "Sneha Iyer", "HR", "30000", "12000", "42000", ""],
              ["E105", "Vikram Singh", "Sales", "55000", "22000", "77000", ""],
              ["E106", "Divya Menon", "Sales", "28000", "11200", "39200", ""]
            ] },
          { type: "p", text: "Two lookups, two column numbers, one locked range. Notice only the third argument changes between `C2` and `D2`. That is the whole trick behind salary and payroll sheets. The Payroll Calculator project extends this with PF and tax." }
        ]
      },
      {
        heading: "Build Section 4: Dashboard-lite",
        blocks: [
          { type: "p", text: "Now the payoff. The dashboard counts and totals without you touching a single employee record." },
          { type: "steps", items: [
            "On the **Dashboard** sheet, type `Sunrise Traders — HR Snapshot` in `A1` and make it bold, size 14.",
            "In row 3 type the headers: `A3` = `Department`, `B3` = `Headcount`, `C3` = `Basic Salary (₹)`.",
            "In `A4:A7` type `Sales`, `Stores`, `Accounts`, `HR`.",
            "In `B4` type `=COUNTIF('Employee Master'!$C$2:$C$11,$A4)` and fill down to `B7`.",
            "In `C4` type `=SUMIF('Employee Master'!$C$2:$C$11,$A4,'Employee Master'!$G$2:$G$11)` and fill down.",
            "In `A9` type `Active Employees` and in `B9` type `=COUNTIF('Employee Master'!$F$2:$F$11,\"Active\")`.",
            "In `A10` type `Monthly Basic Payroll (₹)` and in `B10` type `=SUM('Salary Summary'!$D$2:$D$11)`.",
            "Add light borders around `A3:C7` and bold the row 3 headers."
          ] },
          { type: "grid", title: "Dashboard — what you should see",
            rows: [
              ["Sunrise Traders — HR Snapshot", "", "", ""],
              ["", "", "", ""],
              ["Department", "Headcount", "Basic Salary (₹)", ""],
              ["Sales", "3", "115000", ""],
              ["Stores", "1", "26000", ""],
              ["Accounts", "1", "38000", ""],
              ["HR", "1", "30000", ""],
              ["", "", "", ""],
              ["Active Employees", "4", "", ""],
              ["Monthly Basic Payroll (₹)", "209000", "", ""]
            ] },
          { type: "formula", name: "SUMIF across sheets",
            syntax: "=SUMIF(range, criteria, [sum_range])",
            example: "=SUMIF('Employee Master'!$C$2:$C$11,$A4,'Employee Master'!$G$2:$G$11)",
            explanation: "Adds the basic salary (column G) of only the rows whose department (column C) matches the label in A4. Notice the range and the sum_range sit on the other sheet, while the criteria comes from this one — SUMIF handles mixed locations without any problem." }
        ]
      },
      {
        heading: "Build Section 5: Make Problems Shout with Conditional Formatting",
        blocks: [
          { type: "p", text: "Numbers are easy to scroll past. Colour is not. Three small rules turn the system into something a manager can scan in five seconds." },
          { type: "steps", items: [
            "On **Employee Master**, select `F2:F11`. Choose **Home → Conditional Formatting → Highlight Cells Rules → Equal To**, type `Resigned`, pick **Light Red Fill with Dark Red Text**, and click OK.",
            "On **Attendance Summary**, select `G2:G7` and add another **Highlight Cells Rules → Less Than** rule with the value `0.9` (90%) in light red. Sneha Iyer's 80% should turn red immediately.",
            "On the **Dashboard**, select `B4:B7` and apply **Conditional Formatting → Data Bars** so headcount gets a visible bar.",
            "Test everything: change Divya Menon's status to `Active` on the master. Her red highlight disappears, Active Employees flips to 5, and the Sales headcount bar stays at 3 — but the Salary Summary never changes, because she is still on the rolls. Change her department to `Stores` and watch the dashboard move both bars."
          ] },
          { type: "tip", text: "That last test is the real lesson. You edited one cell and three sheets responded. If something did not move, check two things. Make sure your COUNTIF and SUMIF ranges go down to row 11. And make sure you locked them with dollar signs." }
        ]
      },
      {
        heading: "Level Up Your System",
        blocks: [
          { type: "p", text: "The workbook you just built is a working system. The upgrade exercises below push it closer to a real HR file. Each one appears in the workbook panel with a hint and a worked solution." }
        ]
      },
      {
        heading: "Compare With the Solution",
        blocks: [
          { type: "steps", items: [
            "Open **Employee Management System — completed solution** and click through the four sheets next to your own.",
            "On the master, click any Department cell and press `[[Alt]] + [[Down]]` — the dropdown should open. Open **Data → Data Validation** to compare the source lists word for word.",
            "Click `B2` on the solution's Attendance Summary and read the formula in the formula bar. Compare it to yours: sheet name in single quotes, range locked, FALSE at the end.",
            "On the solution dashboard, click `B4` and `C4` and check both COUNTIF and SUMIF ranges reach row 11.",
            "Change one employee's department on the solution master and confirm the dashboard moves, then close without saving so the solution stays clean."
          ] },
          { type: "p", text: "If your numbers match the solution everywhere, your system is correct. If one sheet differs, the mismatch is almost always in a locked range or a dropdown list — check those two things first." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Writing the lookup as =VLOOKUP($A2,'Employee Master'!A2:B7,2,FALSE) with no dollar signs, then filling down.", right: "=VLOOKUP($A2,'Employee Master'!$A$2:$B$7,2,FALSE)", why: "Without locked references the range slides down as you fill — by row 7 it is looking in A8:B13, which is empty, so every name below the first few shows #N/A." },
      { wrong: "Building the dashboard counts with =COUNTIF(C2:C11,\"Sales\").", right: "=COUNTIF('Employee Master'!$C$2:$C$11,$A4)", why: "C2:C11 on the Dashboard sheet is empty, so the count is 0. COUNTIF looks on the sheet the formula sits on unless you point it at the Employee Master sheet explicitly." },
      { wrong: "Applying the Department dropdown only to C2:C7.", right: "Select C2:C11 so four spare rows carry the validation too", why: "The first new hire past row 7 gets a free-typing cell, and one misspelled department breaks the COUNTIF totals without any error message." },
      { wrong: "Leaving the Attendance % column as =C2/F2 with General formatting, showing 0.92.", right: "Format G2:G7 as Percentage with 0 decimals so it reads 92%", why: "The formula is correct but unreadable as a fraction. Percentage formatting is a display change — the underlying value stays right either way." }
    ],
    shortcuts: [
      { keys: "Alt + Down", action: "Open the dropdown list in a validated cell" },
      { keys: "F4", action: "Toggle absolute ($ locked) references on the selected range while editing a formula" },
      { keys: "Ctrl + 1", action: "Open Format Cells to set date and percentage formats" }
    ],
    exercises: [
      { task: "On the Dashboard, add a row showing the average basic salary of the Stores department in one formula.", hint: "AVERAGEIF has exactly the same shape as SUMIF: range, criteria, average_range.", solution: "=AVERAGEIF('Employee Master'!$C$2:$C$11,\"Stores\",'Employee Master'!$G$2:$G$11)" },
      { task: "Count how many employees have attendance below 90% this month.", hint: "COUNTIF accepts typed criteria like \"<90%\" — no helper column needed.", solution: "=COUNTIF('Attendance Summary'!$G$2:$G$7,\"<90%\")" },
      { task: "Show the name of the highest-paid employee without sorting the master.", hint: "Find the biggest basic salary with MAX, then match it back to a name with INDEX and MATCH.", solution: "=INDEX('Employee Master'!$B$2:$B$7,MATCH(MAX('Employee Master'!$G$2:$G$7),'Employee Master'!$G$2:$G$7,0))" }
    ],
    quiz: [
      { q: "Which formula correctly counts active employees for the dashboard?",
        options: [
          "=COUNTIF('Employee Master'!$F$2:$F$11,\"Active\")",
          "=COUNT('Employee Master'!$F$2:$F$11)",
          "=COUNTA(\"Active\")",
          "=SUMIF('Employee Master'!$F$2:$F$11,\"Active\")"
        ],
        answer: 0,
        explanation: "COUNTIF counts cells that meet a criterion and needs the sheet-qualified range plus the criterion in quotes. COUNT counts numbers, not text values like \"Active\", and SUMIF adds values rather than counting rows." },
      { q: "Why must the master range be locked as $A$2:$G$7 inside the Salary Summary lookups?",
        options: [
          "Because VLOOKUP refuses to work without dollar signs",
          "So the formula can be copied across columns without the lookup range sliding",
          "Because dollar signs convert the result to rupees",
          "So Excel searches the whole sheet instead of a range"
        ],
        answer: 1,
        explanation: "Locked references survive the fill handle and sideways copies. Without them the range drifts into empty rows and the lookups return #N/A or wrong values." },
      { q: "Which SUMIF correctly totals the basic salary of the Sales department?",
        options: [
          "=SUMIF('Employee Master'!$G$2:$G$11,\"Sales\",'Employee Master'!$C$2:$C$11)",
          "=SUM('Employee Master'!$G$2:$G$11,\"Sales\")",
          "=SUMIF('Employee Master'!$C$2:$C$11,\"Sales\",'Employee Master'!$G$2:$G$11)",
          "=SUMIF(\"Sales\",'Employee Master'!$C$2:$C$11,'Employee Master'!$G$2:$G$11)"
        ],
        answer: 2,
        explanation: "SUMIF's order is fixed: first the range to test (department), then the criteria, then the range to add (salary). Option 1 tests the salary column against \"Sales\", which never matches." },
      { q: "You add a new hire in row 8 with department typed as \"Sels\". What happens on the dashboard?",
        options: [
          "Excel shows an error message immediately",
          "The Sales headcount drops by one",
          "Nothing visibly breaks, but no department count includes the new hire",
          "The COUNTIF formula returns #N/A"
        ],
        answer: 2,
        explanation: "This is the classic silent failure: \"Sels\" matches no dropdown value and no dashboard label, so every count simply misses that row. That is why the validation dropdowns must cover the spare rows too." }
    ],
    summary: [
      "One clean master sheet with validated dropdowns is the heart of every small office system.",
      "VLOOKUP with a locked, sheet-qualified range lets summary sheets fill themselves from the master.",
      "COUNTIF counts rows that match a label; SUMIF totals the matching values — same three-part idea.",
      "Conditional formatting turns a wall of numbers into signals: red for resigned, red for low attendance, bars for headcount.",
      "Always test the system by editing the master and watching every other sheet respond."
    ],
    downloads: [
      { file: "assets/files/level7/employee-management-system.xlsx", label: "Employee Management System — exercise file", note: "Starter workbook with sample employee data and empty sheets to build" },
      { file: "assets/files/level7/employee-management-system-solution.xlsx", label: "Employee Management System — completed solution", note: "The finished system with all formulas and formatting in place" }
    ],
    keywords: ["employee management system excel", "data validation dropdown", "VLOOKUP across sheets", "COUNTIF dashboard", "SUMIF department total", "attendance summary", "salary summary HRA", "conditional formatting project"]
  },
  {
    id: "l7-project-attendance",
    title: "Project: Attendance Management System",
    difficulty: "Intermediate",
    duration: 40,
    version: "",
    objectives: [
      "Build a full-month attendance register using the codes P, A, L, WO and H",
      "Make per-employee totals and attendance percentages calculate with COUNTIF",
      "Shade weekend columns automatically with a conditional formatting formula",
      "Build a daily totals strip and a per-employee summary sheet",
      "Keep a 30-column register readable with freezing and colour rules"
    ],
    intro: "Most Indian offices still use paper attendance registers. They are slow to total and hard to analyse. In this project you build the Excel version for Sunrise Traders. Each employee's month becomes one row of single-letter codes. Totals, percentages and weekend shading then appear on their own.\n\nThe system has two sheets. The **Attendance Register** has a column for every day of November 2025. The **Summary** condenses the register into one line per employee. You will use COUNTIF heavily. You will also meet one new idea: conditional formatting driven by a formula. Everything else comes from Levels 3 to 6. Everything here works in Excel 2010 and later.",
    keyIdea: "You build a full-month attendance register that totals codes and shades weekends by itself.",
    sections: [
      {
        heading: "What You Will Build",
        blocks: [
          { type: "p", text: "The register works like the paper one your office knows. But here the maths is done. HR types a code into each day cell. `P` means present, `A` absent, `L` paid leave, `WO` week off and `H` holiday. The right-hand columns instantly show how many days each code was used, how many days are payable, and the attendance percentage." },
          { type: "p", text: "A strip under the register counts how many people were present on each individual day. That is the number an office manager actually needs at 9:05 am." },
          { type: "ul", items: [
            "**Attendance Register** — 6 employees × 30 day columns, with automatic totals in columns AG to AM.",
            "**Weekend shading** — the five Sunday columns glow grey so nobody marks a week off by mistake.",
            "**Daily totals strip** — present and absent counts for every date, row by row.",
            "**Summary sheet** — one clean line per employee pulled straight from the register."
          ] },
          { type: "p", text: "November 2025 has 30 days and five Sundays: the 2nd, 9th, 16th, 23rd and 30th. Days 1 to 14 are already filled in the exercise file. They are your model. You complete the second half of the month yourself." }
        ]
      },
      {
        heading: "Skills You Will Use",
        blocks: [
          { type: "ul", items: [
            "**Data validation** (Level 3) — a P/A/L/WO/H dropdown on every day cell so no stray codes get in.",
            "**COUNTIF** (Level 4) — the engine behind every total on this sheet.",
            "**Mixed references like `$C2` and `C$2`** (Level 4) — the key to filling formulas across a wide register.",
            "**Conditional formatting** (Level 6) — both simple rules and one formula-based rule.",
            "**Cross-sheet references** (Level 5) — the Summary sheet reads the register directly."
          ] },
          { type: "p", text: "The one genuinely new idea is a conditional formatting rule whose condition is a formula. Everything else should feel familiar. This project applies known tools to a much wider grid than you have used before." }
        ]
      },
      {
        heading: "Open the Exercise File",
        blocks: [
          { type: "p", text: "Download **Attendance Management System — exercise file** and open it. You will find:" },
          { type: "ul", items: [
            "**Attendance Register** — headers in row 1: Emp ID in `A1`, Name in `B1`, the day numbers `1` to `30` in `C1:AF1`, and the total headers in `AG1:AM1`. Rows 2 to 7 hold the six employees. Days 1 to 14 are already coded, including `WO` on Sundays 2 and 9.",
            "**Summary** — empty except for a title in `A1`.",
            "Row 1 day numbers are real numbers, not text — the weekend shading formula depends on that."
          ] },
          { type: "table", headers: ["Column", "Content", "Example"],
            rows: [
              ["A", "Emp ID", "E101"],
              ["B", "Employee name", "Arjun Mehta"],
              ["C to AF", "Day 1 to day 30 — one column per day", "P, A, L, WO or H"],
              ["AG", "Present count", "23"],
              ["AH", "Absent count", "1"],
              ["AI", "Leave count", "0"],
              ["AJ", "Week-off count", "5"],
              ["AK", "Holiday count", "0"],
              ["AL", "Working days = 30 minus WO minus H", "25"],
              ["AM", "Attendance %", "96%"]
            ] },
          { type: "tip", text: "Select cell `C2` and choose **View → Freeze Panes → Freeze Panes**. The names stay visible while you scroll through the month — essential in a register this wide." }
        ]
      },
      {
        heading: "Build Section 1: Finish the Month and Guard the Codes",
        blocks: [
          { type: "p", text: "First complete the coding, then lock the door so only valid codes can ever be typed." },
          { type: "steps", items: [
            "Study the filled fortnight: everyone is `P` except Sneha Iyer (row 5), who is on `L` from day 3 to day 7, and Rahul Verma (row 4), who is `A` on day 11. Sundays 2 and 9 are `WO` for all.",
            "Fill days 15 to 30 for every employee with `P`, and enter `WO` on the 16th, 23rd and 30th — the remaining Sundays.",
            "Select `C2:AF7`, open **Data → Data Validation**, choose **List**, and set the source to `P,A,L,WO,H`. Click OK.",
            "Press `[[Ctrl]] + [[Enter]]` trick check: select any empty block of day cells, type `P`, and press `[[Ctrl]] + [[Enter]]` to fill the whole selection at once."
          ] },
          { type: "warning", text: "COUNTIF matches text exactly, so \"P \" with a trailing space is not a P. The dropdown list prevents this completely — which is exactly why real registers always use one." }
        ]
      },
      {
        heading: "Build Section 2: Automatic Totals in Columns AG to AM",
        blocks: [
          { type: "p", text: "Here is where mixed references earn their keep. You will write seven formulas once, in row 2, and then fill the whole block down." },
          { type: "steps", items: [
            "In `AG2` type `=COUNTIF($C2:$AF2,\"P\")`.",
            "In `AH2` type `=COUNTIF($C2:$AF2,\"A\")`, in `AI2` the same with `\"L\"`, in `AJ2` with `\"WO\"`, and in `AK2` with `\"H\"`.",
            "In `AL2` type `=30-$AJ2-$AK2` — the payable working days: the month length minus week-offs minus holidays.",
            "In `AM2` type `=AG2/AL2`, then format `AM2:AM7` as **Percentage** with 0 decimals.",
            "Select `AG2:AM2` and fill it down to row 7. Every employee now has a complete score line."
          ] },
          { type: "formula", name: "Row-wise COUNTIF with locked columns",
            syntax: "=COUNTIF($C2:$AF2, \"P\")",
            example: "=COUNTIF($C2:$AF2,\"P\")",
            explanation: "The dollar signs before C and AF lock the columns but leave the row free. That means you can copy the formula safely both across the totals row and down the six employees, and the day range never slides off the register." },
          { type: "p", text: "Check your row 2. Arjun Mehta should read 25 present, 0 absent, 0 leave, 5 week-offs, 0 holidays, 25 working days, 100%. Sneha Iyer in row 5 reads 20 present, 5 leave, 80%. If a number is off, look for stray spaces in that employee's row. A code typed without the dropdown is the almost-always cause." }
        ]
      },
      {
        heading: "Build Section 3: Weekend Shading and Absence Highlights",
        blocks: [
          { type: "p", text: "Simple rules highlight values; formula-based rules highlight positions. You will use both." },
          { type: "steps", items: [
            "Select `C2:AF7`. Choose **Home → Conditional Formatting → New Rule → Use a formula to determine which cells to format**.",
            "In the formula box type `=OR(C$1=2,C$1=9,C$1=16,C$1=23,C$1=30)` — that is the top-left cell of your selection, `C$1`, compared against the five Sunday day numbers.",
            "Click **Format**, choose a light grey fill, and click OK twice. The five Sunday columns shade immediately.",
            "With `C2:AF7` still selected, add a second rule: **Highlight Cells Rules → Equal To**, value `A`, **Light Red Fill with Dark Red Text**.",
            "In **Conditional Formatting → Manage Rules**, drag the grey weekend rule **above** the red absence rule so weekends win, then click OK."
          ] },
          { type: "formula", name: "Conditional formatting formula rule",
            syntax: "=OR(condition1, condition2, …)",
            example: "=OR(C$1=2,C$1=9,C$1=16,C$1=23,C$1=30)",
            explanation: "Excel evaluates this formula for every cell in the selection, adjusting the reference the same way it adjusts a copied formula. Because the row is locked with C$1, every cell in a column looks at its own day number in row 1 — and lights up when that day is a Sunday." },
          { type: "tip", text: "Type the formula for the top-left cell of the selection. Use the right mix of $ locks. This one rule idea powers hundreds of real-world formats: overdue dates, weekend columns, duplicate rows." }
        ]
      },
      {
        heading: "Build Section 4: Daily Totals Strip and the Summary Sheet",
        blocks: [
          { type: "p", text: "Two final pieces: a per-day headcount under the register, and a clean per-employee summary sheet." },
          { type: "steps", items: [
            "In `B9` type `Present per day`, and in `B10` type `Absent per day`. Bold both labels.",
            "In `C9` type `=COUNTIF(C$2:C$7,\"P\")` and fill it right across to `AF9`. The dollar signs lock the rows while the column moves.",
            "In `C10` type `=COUNTIF(C$2:C$7,\"A\")` and fill right too. Day 11 should show 5 present and 1 absent — Rahul Verma's absence.",
            "On the **Summary** sheet, build headers in row 2: `Emp ID`, `Name`, `Present`, `Absent`, `Leave`, `Working Days`, `Attendance %` in `A2:G2`.",
            "In `A3` type `='Attendance Register'!A2` and fill across to `G3`, then down to row 8, so each summary cell points straight at the matching register cell.",
            "Format `G3:G8` as percentage, add borders, and bold the header row."
          ] },
          { type: "grid", title: "Summary sheet — what you should see",
            rows: [
              ["Emp ID", "Name", "Present", "Absent", "Leave", "Working Days", "Attendance %"],
              ["E101", "Arjun Mehta", "25", "0", "0", "25", "100%"],
              ["E102", "Priya Nair", "25", "0", "0", "25", "100%"],
              ["E103", "Rahul Verma", "24", "1", "0", "25", "96%"],
              ["E104", "Sneha Iyer", "20", "0", "5", "25", "80%"],
              ["E105", "Vikram Singh", "25", "0", "0", "25", "100%"],
              ["E106", "Divya Menon", "25", "0", "0", "25", "100%"]
            ] },
          { type: "p", text: "Notice the Summary uses plain references like `='Attendance Register'!AG2` instead of re-counting. One source of truth, many views. This is the same principle as the employee system. It keeps the two sheets forever in sync." }
        ]
      },
      {
        heading: "Level Up Your System",
        blocks: [
          { type: "p", text: "The register is complete and self-checking. The exercises below add the touches real HR teams ask for within a week of using a file like this." }
        ]
      },
      {
        heading: "Compare With the Solution",
        blocks: [
          { type: "steps", items: [
            "Open **Attendance Management System — completed solution** and put your two sheets side by side with its two.",
            "Click `AG2` on the solution register and read the COUNTIF. Confirm your dollar signs sit in the same places — before C and AF, not before the row numbers.",
            "Open **Conditional Formatting → Manage Rules** on the solution register to see the weekend formula exactly as it is stored.",
            "Compare the daily strip: your day 11 figures must read 5 present and 1 absent.",
            "On the Summary, click `C3` and confirm it references the register, not a re-count, then close the solution without saving."
          ] },
          { type: "p", text: "The most common difference is one Sunday column left unshaded. That usually means the rule was typed for the wrong top-left cell. That single mistake is worth more than an hour of reading about mixed references." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Typing =COUNTIF(C2:AF2,\"P\") in AG2 and copying it right to AM2.", right: "=COUNTIF($C2:$AF2,\"P\") so the day range stays put when copied across", why: "With relative columns, copying right shifts the range to D2:AG2 and then E2:AH2 — each totals column counts a different slice of the month, and every number is quietly wrong." },
      { wrong: "Typing the working-days formula as =30-AJ2.", right: "=30-$AJ2-$AK2 so holidays are deducted too", why: "The system is built for any month. The day a December file brings in an H code for Christmas, the single-deduction version overstates payable days and every attendance percentage drops." },
      { wrong: "Writing the weekend rule as =OR(C1=2,C1=9,...) with both row and column relative.", right: "=OR(C$1=2,C$1=9,C$1=16,C$1=23,C$1=30) with the row locked", why: "Excel applies the rule from the selection's top-left cell and adjusts relative parts. With the row unlocked, cells in row 3 start checking row 2 — which holds employee codes, not day numbers — and the shading lands in the wrong places." },
      { wrong: "Entering a week off as \"w.o\" or \"W/O\" by hand.", right: "Pick WO from the dropdown in every day cell", why: "COUNTIF matches the exact text \"WO\". Any variant is simply not counted, the week-off total drops below 5, and the working-days and percentage figures inflate silently." }
    ],
    shortcuts: [
      { keys: "Ctrl + Enter", action: "Fill the whole selected range with the code you just typed" },
      { keys: "Alt + Down", action: "Open the P/A/L/WO/H dropdown in a day cell" },
      { keys: "Ctrl + 1", action: "Format the percentage column" }
    ],
    exercises: [
      { task: "Add a total at the bottom of the register counting every L in the month across all employees.", hint: "One COUNTIF over the whole two-dimensional block of day cells.", solution: "=COUNTIF($C$2:$AF$7,\"L\")" },
      { task: "In column AN of the register, flag employees with perfect attendance in one formula.", hint: "Compare the present count with the working-days count — no need to touch the percentage.", solution: "=IF(AG2=AL2,\"Perfect\",\"Check\")" },
      { task: "Highlight any day on which more than one employee was absent.", hint: "Select C10:AF10 and add a conditional formatting rule using Highlight Cells Rules → Greater Than.", solution: "Select C10:AF10, then Conditional Formatting → Highlight Cells Rules → Greater Than → 1, with a light red fill." }
    ],
    quiz: [
      { q: "Which formula correctly counts Rahul Verma's present days in row 4?",
        options: [
          "=COUNTIF($C$4:$AF$4,\"P\")",
          "=COUNTIF($C2:$AF2,\"P\")",
          "=COUNT($C4:$AF4,\"P\")",
          "=SUMIF($C4:$AF4,\"P\")"
        ],
        answer: 0,
        explanation: "The range must point at row 4, his row. COUNT cannot take a criteria argument, and SUMIF without a sum_range only counts matching cells, not adds them — COUNTIF is the right tool for counting codes." },
      { q: "Why is the formula written as =COUNTIF($C2:$AF2,\"P\") with dollar signs before C and AF?",
        options: [
          "To make the formula run faster",
          "So it can be copied both across the totals columns and down the employee rows without the day range sliding",
          "Because COUNTIF requires absolute references by definition",
          "To lock the sheet against edits"
        ],
        answer: 1,
        explanation: "The columns are locked but the row is free. Copying right keeps the day columns fixed; copying down moves the row so each employee counts their own row." },
      { q: "Sneha Iyer has 20 present, 5 leave and 5 week-offs in a 30-day month with no holidays. What is her attendance percentage?",
        options: [
          "83%",
          "80%",
          "67%",
          "100%"
        ],
        answer: 1,
        explanation: "Working days = 30 − 5 WO − 0 H = 25. Attendance = 20 ÷ 25 = 80%. Week-offs are not days she could have attended, so they leave both sides of the fraction." },
      { q: "How do the five Sunday columns get their grey shading?",
        options: [
          "Manually painting each column with the fill bucket",
          "A conditional formatting formula rule: =OR(C$1=2,C$1=9,C$1=16,C$1=23,C$1=30)",
          "A COUNTIF in row 1",
          "Excel shades weekends automatically when dates are detected"
        ],
        answer: 1,
        explanation: "The formula rule checks each cell's own day number in row 1 against the five Sundays. Automatic detection does not exist for plain numbers, and manual shading never survives a new month." },
      { q: "In the daily strip, C9 =COUNTIF(C$2:C$7,\"P\") is filled right to AF9. What does the formula in J9 count?",
        options: [
          "Present codes in column C only",
          "Present codes in column J — day 8 — for all six employees",
          "All codes in rows 2 to 7",
          "Nothing — the formula breaks when copied"
        ],
        answer: 1,
        explanation: "The rows are locked with C$2:C$7 but the column is relative, so filling right moves the counted column to J while the employee rows stay put." }
    ],
    summary: [
      "A register is just a wide grid plus one counting formula per code — COUNTIF does all the work.",
      "Lock the columns in row totals ($C2:$AF2) and the rows in column totals (C$2:C$7); mixed references are the whole skill.",
      "Working days = month length minus week-offs minus holidays; attendance % = present ÷ working days.",
      "Formula-based conditional formatting can shade positions, not just values — the weekend columns are the classic case.",
      "Keep one source of truth: the Summary sheet references the register instead of recounting it."
    ],
    downloads: [
      { file: "assets/files/level7/attendance-management-system.xlsx", label: "Attendance Management System — exercise file", note: "Starter register with attendance codes filled for two weeks" },
      { file: "assets/files/level7/attendance-management-system-solution.xlsx", label: "Attendance Management System — completed solution", note: "Finished register with all totals, percentages and formatting" }
    ],
    keywords: ["attendance register excel", "P A L WO H codes", "COUNTIF attendance percentage", "weekend column shading", "conditional formatting formula rule", "mixed references $C2", "monthly attendance sheet", "daily present count"]
  },
  {
    id: "l7-project-payroll",
    title: "Project: Payroll Calculator",
    difficulty: "Intermediate",
    duration: 40,
    version: "",
    objectives: [
      "Build a payroll sheet that computes HRA, PF, gross, tax and net pay from the employee master",
      "Write a simplified slab-based tax formula with a nested IF",
      "Round every money figure to whole rupees with ROUND",
      "Create a payslip for any single employee using INDEX with MATCH",
      "Know when XLOOKUP replaces the INDEX/MATCH pattern"
    ],
    intro: "Payroll is the calculation office Excel was born for. In this project you turn the Sunrise Traders employee master into a working monthly payroll. For each of the six employees, the sheet computes HRA at 40% of basic. It also computes PF at 12% of basic, a simplified slab tax, and a net pay rounded to the rupee. Then you build a printable payslip. Type an employee's ID and their full breakdown appears.\n\nThe payslip uses INDEX with MATCH to pull each figure. That is the lookup pair from Level 5, and it works in any direction. On Microsoft 365 or Excel 2021, you will also see how the single XLOOKUP function does the same job. The slab tax is the only genuinely new pattern. It is just a nested IF from Level 4 with rupee amounts inside.",
    keyIdea: "You build a payroll calculator that pays staff with HRA, PF, slab tax and a payslip.",
    sections: [
      {
        heading: "What You Will Build",
        blocks: [
          { type: "p", text: "Three sheets, one flow. HR never retypes a salary. The **Payroll** sheet reads basic pay from the **Employee Master**. It applies the standard Indian salary components and produces one net figure per person. The **Payslip** sheet is a formatted view for one employee. Type `E101` in one cell and the whole slip fills in, ready to print and sign." },
          { type: "ul", items: [
            "**Employee Master** — the six Sunrise Traders employees with their basic salaries (already filled).",
            "**Payroll** — one row per employee: Basic, HRA 40%, PF 12%, Gross, slab Tax, Net.",
            "**Payslip** — a print-ready salary slip driven by a single Emp ID input cell."
          ] },
          { type: "p", text: "The Indian terms, simply: **HRA** (House Rent Allowance) covers rent. Many private companies set it at 40% of basic. **PF** (Provident Fund) is retirement savings deducted at 12% of basic. **Gross pay** is Basic + HRA + PF before tax. **Net pay** is what actually reaches the bank. This project uses simplified rules so the pattern is clear. A real payroll adds exemptions, a PF wage ceiling and state-level detail." }
        ]
      },
      {
        heading: "Skills You Will Use",
        blocks: [
          { type: "ul", items: [
            "**VLOOKUP** (Level 5) — pull basic salary and names from the master.",
            "**Nested IF** (Level 4) — the tax slab calculation.",
            "**ROUND** (Level 2) — keep every amount in whole rupees.",
            "**INDEX and MATCH** (Level 5) — the payslip lookups.",
            "**Absolute references and the fill handle** (Level 4) — one payroll formula, filled down six times."
          ] },
          { type: "p", text: "Everything here is revision in disguise. If a component confuses you, it is the Excel part that needs a refresher, not the payroll part." }
        ]
      },
      {
        heading: "Open the Exercise File",
        blocks: [
          { type: "p", text: "Download **Payroll Calculator — exercise file**. Inside:" },
          { type: "ul", items: [
            "**Employee Master** — complete, with basics of ₹32,000, ₹38,000, ₹26,000, ₹30,000, ₹55,000 and ₹28,000 for E101 to E106.",
            "**Payroll** — headers only: Emp ID, Name, Basic, HRA, PF, Gross, Tax, Net in `A1:H1`.",
            "**Payslip** — a formatted frame with labels but empty value cells and one input cell for the Emp ID."
          ] },
          { type: "tip", text: "Note the master's layout: Emp ID in A, Name in B, Department in C, Designation in D, Join Date in E, Status in F, Basic Salary in G, rows 2 to 7. Every lookup column number in this project depends on that order." }
        ]
      },
      {
        heading: "Build Section 1: The Payroll Engine",
        blocks: [
          { type: "p", text: "Six formulas per employee, written once in row 2 and filled down to row 7. Get row 2 right and the rest is a fill handle." },
          { type: "steps", items: [
            "On **Payroll**, type `E101` to `E106` in `A2:A7`.",
            "In `B2` type `=VLOOKUP($A2,'Employee Master'!$A$2:$B$7,2,FALSE)` and fill down.",
            "In `C2` type `=VLOOKUP($A2,'Employee Master'!$A$2:$G$7,7,FALSE)` and fill down — column 7 of the wide range is Basic Salary.",
            "In `D2` type `=ROUND($C2*40%,0)` — HRA — and fill down.",
            "In `E2` type `=ROUND($C2*12%,0)` — PF — and fill down.",
            "In `F2` type `=$C2+$D2+$E2` — Gross — and fill down.",
            "Format `C2:F7` as numbers with thousands separators, and bold the header row."
          ] },
          { type: "grid", title: "Payroll after the first five columns",
            rows: [
              ["Emp ID", "Name", "Basic (₹)", "HRA (₹)", "PF (₹)", "Gross (₹)", "", ""],
              ["E101", "Arjun Mehta", "32000", "12800", "3840", "48640", "", ""],
              ["E102", "Priya Nair", "38000", "15200", "4560", "57760", "", ""],
              ["E103", "Rahul Verma", "26000", "10400", "3120", "39520", "", ""],
              ["E104", "Sneha Iyer", "30000", "12000", "3600", "45600", "", ""],
              ["E105", "Vikram Singh", "55000", "22000", "6600", "83600", "", ""],
              ["E106", "Divya Menon", "28000", "11200", "3360", "42560", "", ""]
            ] },
          { type: "warning", text: "Use column 7, not column 6, for Basic. Column 6 of the master is Status — the giveaway that something is wrong is a payroll where Basic reads \"Active\"." }
        ]
      },
      {
        heading: "Build Section 2: The Slab Tax Formula",
        blocks: [
          { type: "p", text: "India's income tax is charged in slabs. Different rates apply to different slices of income, not one rate on the whole amount. We use a simple monthly version so you can see the machinery. No tax applies to the first ₹25,000 of monthly gross. The slice from ₹25,001 to ₹50,000 is taxed at 5%. Everything above ₹50,000 is taxed at 10%." },
          { type: "steps", items: [
            "In `G2` type the formula below and fill it down to `G7`.",
            "Format `G2:G7` the same way as the money columns.",
            "Spot-check `G5` (Sneha Iyer, gross ₹45,600): only her slice above 25,000 is taxed, so expect ₹1,030."
          ] },
          { type: "formula", name: "Simplified slab tax",
            syntax: "=IF(F2<=25000, 0, IF(F2<=50000, (F2-25000)*5%, 1250+(F2-50000)*10%))",
            example: "=IF(F2<=25000,0,IF(F2<=50000,(F2-25000)*5%,1250+(F2-50000)*10%))",
            explanation: "The outer IF catches anyone in the zero slab. The inner IF splits the middle from the top slab. The 1250 is fixed because the full middle slab (50,000 − 25,000) at 5% is ₹1,250 — anyone above ₹50,000 owes that whole middle slab plus 10% of their excess. Check Vikram Singh: 1250 + (83600 − 50000) × 10% = 1250 + 3360 = ₹4,610." },
          { type: "warning", text: "Order matters in nested IFs. Excel stops at the first TRUE condition, so the check `<=50000` must come after `<=25000`. Swap the two and everyone above ₹25,000 gets the middle-slab rate, including the big earners who owe 10%." }
        ]
      },
      {
        heading: "Build Section 3: Net Pay and the Payroll Bottom Line",
        blocks: [
          { type: "p", text: "Two cells finish the payroll — and they are the two the accountant actually signs off." },
          { type: "steps", items: [
            "In `H2` type `=ROUND($F2-$G2,0)` — net pay, what reaches the bank — and fill down to `H7`.",
            "In `A9` type `Total Net Payroll (₹)` and in `H9` type `=SUM($H$2:$H$7)`.",
            "Select `A1:H1`, bold it, and give the range a light fill; put borders under row 7.",
            "Sanity-check one employee by hand: Arjun Mehta — Gross 48,640, tax (48,640 − 25,000) × 5% = ₹1,182, net ₹47,458."
          ] },
          { type: "grid", title: "The completed payroll",
            rows: [
              ["Emp ID", "Name", "Basic (₹)", "HRA (₹)", "PF (₹)", "Gross (₹)", "Tax (₹)", "Net (₹)"],
              ["E101", "Arjun Mehta", "32000", "12800", "3840", "48640", "1182", "47458"],
              ["E102", "Priya Nair", "38000", "15200", "4560", "57760", "1638", "56122"],
              ["E103", "Rahul Verma", "26000", "10400", "3120", "39520", "726", "38794"],
              ["E104", "Sneha Iyer", "30000", "12000", "3600", "45600", "1030", "44570"],
              ["E105", "Vikram Singh", "55000", "22000", "6600", "83600", "4610", "78990"],
              ["E106", "Divya Menon", "28000", "11200", "3360", "42560", "878", "41682"]
            ] },
          { type: "tip", text: "Why ROUND everywhere? Percentages produce paise — tax of ₹1,116.50 on some basic would put half-rupees into a bank transfer and look wrong on a printed slip. ROUND(value, 0) snaps every figure to whole rupees before it moves downstream." }
        ]
      },
      {
        heading: "Build Section 4: The Payslip with INDEX and MATCH",
        blocks: [
          { type: "p", text: "The payslip shows one employee. Instead of six VLOOKUPs, you build one MATCH that finds the employee's row, and INDEX calls that reuse it — the Level 5 pattern that reads cleanly and never needs a column number recount." },
          { type: "steps", items: [
            "On **Payslip**, the Emp ID input cell is `B4`. Type `E101` there.",
            "In the Name cell (`B6`) type `=INDEX('Payroll'!$B$2:$B$7,MATCH($B$4,'Payroll'!$A$2:$A$7,0))`.",
            "For Department (`D6`), Designation (`B7`), Join Date (`D7`), Basic (`B8`), HRA (`D8`), PF (`B9`), Gross (`D9`), Tax (`B10`) and Net (`D10`), use the same pattern, changing only the INDEX range and the column it points at — for example Basic is `=INDEX('Payroll'!$C$2:$C$7,MATCH($B$4,'Payroll'!$A$2:$A$7,0))`.",
            "Type `E105` into `B4` and confirm the slip now shows Vikram Singh: Basic ₹55,000, Gross ₹83,600, Tax ₹4,610, Net ₹78,990.",
            "Wrap every lookup in IFERROR so an empty or wrong ID shows a friendly blank instead of `#N/A`: `=IFERROR(INDEX(...),\"\")`.",
            "Select `A1:D12` and set **Page Layout → Print Area → Set Print Area**. Press `[[Ctrl]] + [[P]]` to preview a clean one-page slip."
          ] },
          { type: "formula", name: "INDEX + MATCH payslip lookup",
            syntax: "=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))",
            example: "=INDEX('Payroll'!$B$2:$B$7,MATCH($B$4,'Payroll'!$A$2:$A$7,0))",
            explanation: "MATCH finds which row of the payroll holds the ID in B4 — say position 5 — and INDEX returns the value from that position of the name column. Change the INDEX range to C, D or H and the same MATCH fetches any payroll field for that employee." },
          { type: "tip", text: "On Microsoft 365 or Excel 2021 you can write the same thing once with XLOOKUP: `=XLOOKUP($B$4,'Payroll'!$A$2:$A$7,'Payroll'!$B$2:$B$7)`. INDEX/MATCH is worth learning anyway — it runs in every Excel version, and most office files you will inherit use it." }
        ]
      },
      {
        heading: "Level Up Your System",
        blocks: [
          { type: "p", text: "The payroll and payslip are complete. The upgrade exercises below stretch both — including the annual view that HR prepares in March for Form 16 season." }
        ]
      },
      {
        heading: "Compare With the Solution",
        blocks: [
          { type: "steps", items: [
            "Open **Payroll Calculator — completed solution** and compare the payroll column by column with yours.",
            "Click `G5` on the solution and read the nested IF in the formula bar. Compare it with yours condition by condition — the order of the slab checks is the part people get wrong.",
            "On the payslip, confirm every value cell wraps its lookup in IFERROR, then clear `B4` and confirm the slip goes blank instead of showing #N/A.",
            "Type `E103` into `B4` on your payslip and verify your figures match the solution's row for Rahul Verma: net ₹38,794.",
            "Check the print preview matches: one page, no stray gridlines, all twelve label cells visible."
          ] },
          { type: "p", text: "If your payroll matches but the payslip does not, it is almost always the MATCH range — it must be exactly `'Payroll'!$A$2:$A$7` for every field, while only the INDEX range changes." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Writing the tax formula as =IF(F2<=50000,(F2-25000)*5%,IF(F2<=25000,0,...)) with the 50,000 check first.", right: "=IF(F2<=25000,0,IF(F2<=50000,(F2-25000)*5%,1250+(F2-50000)*10%)) — lowest slab first", why: "Nested IFs stop at the first TRUE condition. Testing 50,000 first means a ₹30,000 gross is caught by the middle slab and the zero slab can never fire for anyone above ₹25,000." },
      { wrong: "Looking up Basic with =VLOOKUP($A2,'Employee Master'!$A$2:$G$7,6,FALSE).", right: "Column index 7 — Basic Salary is the seventh column of the master", why: "Column 6 is Status, so the payroll would try to compute 40% of the word \"Active\" and return #VALUE! — or, with looser layouts, silently pay people the wrong amount." },
      { wrong: "Computing net pay as =F2-G2 and leaving it unrounded.", right: "=ROUND($F2-$G2,0)", why: "Slab percentages create half-rupees. A net of ₹46,213.50 printed on a signed payslip, or sent to a bank, is the kind of small wrongness that destroys trust in the whole file." },
      { wrong: "Leaving the 0 out of MATCH: =MATCH($B$4,'Payroll'!$A$2:$A$7).", right: "=MATCH($B$4,'Payroll'!$A$2:$A$7,0)", why: "Without 0, MATCH runs an approximate match on unsorted IDs and can land on the wrong row — the payslip then shows another employee's salary with complete confidence." }
    ],
    shortcuts: [
      { keys: "F4", action: "Lock a selected range with dollar signs while editing a formula" },
      { keys: "Ctrl + P", action: "Open print preview to check the payslip layout" },
      { keys: "Ctrl + 1", action: "Open Format Cells for the money and date columns" }
    ],
    exercises: [
      { task: "Add an Annual Gross column (I) to the payroll showing each employee's yearly gross in one formula.", hint: "Gross is monthly; the year has 12 months.", solution: "=ROUND($F2*12,0)" },
      { task: "Add a Tax Slab label column (J) that reads No tax, 5% slab or 10% slab.", hint: "One more nested IF reading column F, the gross — reuse the same thresholds as the tax formula.", solution: "=IF(F2<=25000,\"No tax\",IF(F2<=50000,\"5% slab\",\"10% slab\"))" },
      { task: "On the payslip, add an Amount in Words line that shows net pay with the rupee symbol and two labels — for example ₹ 47,458 only.", hint: "Concatenate the symbol with the formatted number using TEXT.", solution: "=\"₹ \"&TEXT($D$10,\"#,##0\")&\" only\"" }
    ],
    quiz: [
      { q: "Vikram Singh's monthly gross is ₹83,600. What does the simplified slab formula compute as his tax?",
        options: [
          "₹4,180 — 5% of his whole gross",
          "₹1,250 — only the middle slab",
          "₹4,610 — the full middle slab plus 10% of the amount above ₹50,000",
          "₹8,360 — 10% of his whole gross"
        ],
        answer: 2,
        explanation: "Slabs are charged on slices, not on the whole amount: the first ₹25,000 is free, the next ₹25,000 is taxed at 5% (₹1,250), and only the ₹33,600 above ₹50,000 is taxed at 10% (₹3,360)." },
      { q: "In the payroll, which formula computes PF correctly?",
        options: [
          "=ROUND($C2*12%,0)",
          "=ROUND($F2*12%,0)",
          "=ROUND($C2+12%,0)",
          "=$C2*12"
        ],
        answer: 0,
        explanation: "PF is 12% of Basic, which lives in column C — not of Gross (F). ROUND with 0 digits keeps the deduction in whole rupees." },
      { q: "Why does the payslip use INDEX + MATCH instead of plain VLOOKUP?",
        options: [
          "It runs faster than VLOOKUP",
          "Each field only needs its own INDEX range, while one MATCH finds the employee's row — and the pattern also works when the return column sits left of the lookup column",
          "VLOOKUP cannot read from another sheet",
          "INDEX + MATCH needs no locked references"
        ],
        answer: 1,
        explanation: "With VLOOKUP you must recount a column number for every field; with INDEX/MATCH you swap only the return range. And if the payslip ever needs a column left of the ID, VLOOKUP simply cannot go there." },
      { q: "You are on Microsoft 365. Which XLOOKUP replaces the payslip's Name lookup?",
        options: [
          "=XLOOKUP('Payroll'!$B$2:$B$7,$B$4)",
          "=XLOOKUP($B$4,'Payroll'!$A$2:$A$7,'Payroll'!$B$2:$B$7)",
          "=XLOOKUP($B$4,'Payroll'!$B$2:$B$7,'Payroll'!$A$2:$A$7)",
          "=XLOOKUP(MATCH($B$4),'Payroll'!$B$2:$B$7)"
        ],
        answer: 1,
        explanation: "XLOOKUP's order is lookup value, then where to search, then what to return. Option 3 has the two ranges swapped and would look for names to return IDs." },
      { q: "Arjun Mehta's gross is ₹48,640. If his tax had come out as ₹1,182.50, what is the fix?",
        options: [
          "Format the cell to hide decimals — the hidden half-rupee never reaches the bank",
          "Change the slab rate from 5% to 4%",
          "Wrap the tax formula in ROUND: =ROUND(tax,0) so the stored value itself is a whole rupee",
          "Nothing — paise in payroll are normal"
        ],
        answer: 2,
        explanation: "Number formatting only changes what is displayed; the stored value keeps the paise and every downstream sum inherits them. ROUND changes the actual value, which is what payroll needs." }
    ],
    summary: [
      "A payroll is five percentage-and-sum formulas per row, filled down from one correct row 2.",
      "Slab tax is a nested IF charged on slices — test the lowest threshold first, and remember the fixed amount owed by every completed slab.",
      "ROUND every money figure to 0 digits so stored values, not just displays, are whole rupees.",
      "INDEX + MATCH lets one MATCH position feed every field on a single-employee payslip.",
      "XLOOKUP shortens the pattern on Microsoft 365, but INDEX/MATCH keeps your file usable in every Excel."
    ],
    downloads: [
      { file: "assets/files/level7/payroll-calculator.xlsx", label: "Payroll Calculator — exercise file", note: "Employee data and a blank payroll structure to complete" },
      { file: "assets/files/level7/payroll-calculator-solution.xlsx", label: "Payroll Calculator — completed solution", note: "Finished payroll with working formulas and a sample payslip" }
    ],
    keywords: ["payroll calculator excel", "HRA 40 percent basic", "PF 12 percent", "slab tax nested IF", "INDEX MATCH payslip", "XLOOKUP alternative", "ROUND whole rupees", "gross and net salary", "salary slip excel"]
  },
  {
    id: "l7-project-inventory",
    title: "Project: Inventory Management System",
    difficulty: "Intermediate",
    duration: 45,
    version: "",
    objectives: [
      "Link an item master to inward and outward registers using validated item codes",
      "Compute live closing stock for every item with SUMIFS",
      "Flag low stock automatically with IF and conditional formatting",
      "Produce a stock value report that finance can use directly",
      "Understand why every stock figure in the system traces back to one master"
    ],
    intro: "Every shop and warehouse runs on one question: how much do we have right now? In this project you build the stock system for Sunrise Traders, which buys and sells office supplies — Notebooks, Pens, Staplers, Printers and Laptops. Goods come **inward** from suppliers and go **outward** to customers through two registers, and a **Current Stock** sheet keeps a live balance for every item, warns you when stock falls below the reorder level, and totals the money value of everything on the shelves.\n\nThe heart of the project is one formula pattern: Closing Stock = Opening + Inward − Outward, where the inward and outward amounts come from SUMIFS scanning the registers. Everything else — validation, lookups, conditional formatting — is revision from Levels 3 to 6. Nothing here needs a newer Excel than 2010.",
    keyIdea: "Closing stock is always Opening + Inward − Outward, computed live with SUMIFS.",
    sections: [
      {
        heading: "What You Will Build",
        blocks: [
          { type: "p", text: "The finished system has four sheets. Staff type goods movements into two simple registers — a date, an item code, a quantity. The **Current Stock** sheet does everything else: it looks up each item's opening stock and rate from the master, adds everything received, subtracts everything sold, shows the closing balance, values it in rupees, and turns the status cell red the moment an item drops to its reorder level." },
          { type: "ul", items: [
            "**Item Master** — five items with code, name, unit, opening stock, cost rate and reorder level.",
            "**Inward** — goods received from suppliers: date, item code, quantity, rate, value.",
            "**Outward** — goods sold or issued: the same columns.",
            "**Current Stock** — the live position per item, with reorder alerts and a stock value report."
          ] },
          { type: "p", text: "This is the exact structure a real storekeeper's Excel file uses before a company grows into proper software. Get the pattern here and every ERP you meet later will feel familiar." }
        ]
      },
      {
        heading: "Skills You Will Use",
        blocks: [
          { type: "ul", items: [
            "**SUMIFS** (Level 4) — total the quantity received and issued for each item.",
            "**VLOOKUP** (Level 5) — pull names, opening stock, rates and reorder levels from the master.",
            "**Data validation** (Level 3) — item-code dropdowns in both registers.",
            "**IF and conditional formatting** (Levels 4 and 6) — the reorder alert.",
            "**Arithmetic and formatting** (Level 2) — closing stock and stock value."
          ] },
          { type: "p", text: "SUMIFS is the one skill to have sharp before starting. You used it for multi-condition totals in Level 4; here it scans whole registers with a single condition, which is the most common way it appears in office work." }
        ]
      },
      {
        heading: "Open the Exercise File",
        blocks: [
          { type: "p", text: "Download **Inventory Management System — exercise file**. Inside:" },
          { type: "ul", items: [
            "**Item Master** — complete: `ITM001` Notebooks, `ITM002` Pens, `ITM003` Staplers, `ITM004` Printers, `ITM005` Laptops, with opening stock, rates and reorder levels filled.",
            "**Inward** and **Outward** — headers in row 1 and the columns ready; the November 2026 entries are already listed.",
            "**Current Stock** — only the header row exists. This is where you build."
          ] },
          { type: "table", headers: ["Column", "Item Master", "Inward / Outward registers"],
            rows: [
              ["A", "Item Code (ITM001…)", "Date"],
              ["B", "Item Name", "Item Code (with dropdown)"],
              ["C", "Unit (Nos / Units)", "Item Name (formula)"],
              ["D", "Opening Stock", "Quantity"],
              ["E", "Rate (₹)", "Rate (₹) (formula)"],
              ["F", "Reorder Level", "Value = Qty × Rate"]
            ] },
          { type: "tip", text: "The registers deliberately keep the item code in column B, not A, because dates come first in real registers. Your SUMIFS criteria must point at column B — check this twice when writing them." }
        ]
      },
      {
        heading: "Build Section 1: Wire Up the Registers",
        blocks: [
          { type: "p", text: "Both registers get the same two upgrades: a dropdown on the item code, and formulas that fill in the name and rate from the master. Staff then type only four things per movement." },
          { type: "steps", items: [
            "On **Inward**, select `B2:B50`, open **Data → Data Validation → List**, and click inside the Source box. Then go to the Item Master sheet and select `A2:A6` — Excel writes the reference `'Item Master'!$A$2:$A$6` for you. Click OK.",
            "In `C2` type `=IF($B2=\"\",\"\",VLOOKUP($B2,'Item Master'!$A$2:$B$6,2,FALSE))` and fill down to row 50 — the name appears the moment a code is chosen.",
            "In `E2` type `=IF($B2=\"\",\"\",VLOOKUP($B2,'Item Master'!$A$2:$E$6,5,FALSE))` and fill down — the cost rate comes from the master too.",
            "In `F2` type `=IF($B2=\"\",\"\",ROUND($D2*$E2,0))` and fill down — the value of the movement.",
            "Repeat steps 1 to 4 on the **Outward** sheet.",
            "Check the November entries: for example, Inward row 3 shows ITM001, Notebooks, 150 units at ₹40, value ₹6,000."
          ] },
          { type: "formula", name: "Guarded VLOOKUP",
            syntax: "=IF(cell=\"\", \"\", VLOOKUP(cell, table, col, FALSE))",
            example: "=IF($B2=\"\",\"\",VLOOKUP($B2,'Item Master'!$A$2:$B$6,2,FALSE))",
            explanation: "Empty register rows produce empty lookups instead of a wall of #N/A. The IF checks the code cell first and only runs the VLOOKUP when a code exists. This is standard practice in every well-built register." }
        ]
      },
      {
        heading: "Build Section 2: Current Stock — the Live Balance",
        blocks: [
          { type: "p", text: "Now the engine room. One row per item, and five formulas that together answer \"how many do we have and what are they worth?\"" },
          { type: "steps", items: [
            "On **Current Stock**, type the codes `ITM001` to `ITM005` in `A2:A6`.",
            "In `B2` type `=VLOOKUP($A2,'Item Master'!$A$2:$B$6,2,FALSE)` and fill down — the item name.",
            "In `C2` type `=VLOOKUP($A2,'Item Master'!$A$2:$D$6,4,FALSE)` and fill down — opening stock.",
            "In `D2` type `=SUMIFS(Inward!$D$2:$D$50,Inward!$B$2:$B$50,$A2)` and fill down — everything received.",
            "In `E2` type `=SUMIFS(Outward!$D$2:$D$50,Outward!$B$2:$B$50,$A2)` and fill down — everything issued.",
            "In `F2` type `=$C2+$D2-$E2` and fill down — the closing stock: opening plus inward minus outward."
          ] },
          { type: "grid", title: "Current Stock after step 6",
            rows: [
              ["Item Code", "Item Name", "Opening", "Inward", "Outward", "Closing", "", ""],
              ["ITM001", "Notebooks", "200", "150", "120", "230", "", ""],
              ["ITM002", "Pens", "500", "300", "460", "340", "", ""],
              ["ITM003", "Staplers", "80", "20", "75", "25", "", ""],
              ["ITM004", "Printers", "15", "5", "8", "12", "", ""],
              ["ITM005", "Laptops", "10", "2", "4", "8", "", ""]
            ] },
          { type: "formula", name: "SUMIFS scanning a register",
            syntax: "=SUMIFS(sum_range, criteria_range1, criteria1)",
            example: "=SUMIFS(Inward!$D$2:$D$50,Inward!$B$2:$B$50,$A2)",
            explanation: "Adds the quantities in the register's column D — but only from rows whose item code in column B matches this row's code. One formula serves all five items because the criteria is a cell reference, not typed text." }
        ]
      },
      {
        heading: "Build Section 3: Reorder Alerts and the Stock Value Report",
        blocks: [
          { type: "p", text: "A stock figure nobody acts on is decoration. Two columns turn the balance into decisions: a status flag and a rupee value." },
          { type: "steps", items: [
            "In `G2` type `=VLOOKUP($A2,'Item Master'!$A$2:$E$6,5,FALSE)` and fill down — the cost rate.",
            "In `H2` type `=ROUND($F2*$G2,0)` and fill down — the stock value: closing units times cost rate.",
            "In `I2` type `=VLOOKUP($A2,'Item Master'!$A$2:$F$6,6,FALSE)` and fill down — the reorder level.",
            "In `J2` type `=IF($F2<=$I2,\"REORDER\",\"OK\")` and fill down.",
            "Select `J2:J6`, then **Home → Conditional Formatting → Highlight Cells Rules → Equal To**, type `REORDER`, and pick **Light Red Fill with Dark Red Text**.",
            "In `A8` type `Total Stock Value (₹)` and in `H8` type `=SUM($H$2:$H$6)`."
          ] },
          { type: "grid", title: "The finished Current Stock sheet",
            rows: [
              ["Item Code", "Item Name", "Opening", "Inward", "Outward", "Closing", "Rate (₹)", "Value (₹)"],
              ["ITM001", "Notebooks", "200", "150", "120", "230", "40", "9200"],
              ["ITM002", "Pens", "500", "300", "460", "340", "12", "4080"],
              ["ITM003", "Staplers", "80", "20", "75", "25", "150", "3750"],
              ["ITM004", "Printers", "15", "5", "8", "12", "6500", "78000"],
              ["ITM005", "Laptops", "10", "2", "4", "8", "42000", "336000"],
              ["", "", "", "", "", "", "Total", "431030"]
            ] },
          { type: "warning", text: "Staplers close at 25 units against a reorder level of 30, so `J4` must read REORDER and glow red. If your sheet shows OK, re-read the formula: the test is `<=`, not `<`, because a stock exactly at the reorder level must also be flagged." }
        ]
      },
      {
        heading: "Build Section 4: Test the System Like a Storekeeper",
        blocks: [
          { type: "p", text: "A system is not finished until you have tried to break it. These tests take two minutes and teach you how the pieces move together." },
          { type: "steps", items: [
            "On **Outward**, add a new row: date `23-Nov-2026`, item code `ITM003` from the dropdown, quantity `10`. Watch Staplers drop to 15 on Current Stock — still REORDER.",
            "On **Inward**, add a receipt: date `24-Nov-2026`, code `ITM003`, quantity `60`. Staplers climb to 75 and the status flips to OK without you touching the Current Stock sheet.",
            "Type an item code that is not in the master, like `ITM009`, into a register row. The name and rate cells go blank and the movement never affects stock — exactly what you want.",
            "Change the reorder level of Notebooks in the master from 100 to 240. The status flips to REORDER instantly, proving the alert is live, not hardcoded.",
            "Undo both master edits and the two test register rows, or delete them, so the file matches the exercise state again."
          ] },
          { type: "tip", text: "That last test is the deepest lesson: the reorder level lives in the master and is only looked up. Parameters belong in one place, formulas reference them — never bury thresholds inside formulas." }
        ]
      },
      {
        heading: "Level Up Your System",
        blocks: [
          { type: "p", text: "The stock system is complete and live. The exercises below add the reports a purchase manager would ask for in week one." }
        ]
      },
      {
        heading: "Compare With the Solution",
        blocks: [
          { type: "steps", items: [
            "Open **Inventory Management System — completed solution** and compare the Current Stock sheet cell by cell with yours.",
            "Click `D2` on the solution and read the SUMIFS: the sum range and criteria range must both sit on the Inward sheet, and the criteria must be `$A2` — the code cell, not typed text.",
            "Confirm the solution's Stapler row shows closing 25 and status REORDER, and that its total stock value is ₹4,31,030.",
            "Open Data Validation on the solution's registers and check the source range points at the master's code column.",
            "Add a test inward row on the solution copy, watch the stock move, then close without saving."
          ] },
          { type: "p", text: "If your closing stocks differ, add up one item's movements by hand — the register rows are short enough to check on fingers. The mismatch is nearly always a criteria range pointing at the wrong column, or a register row typed without the dropdown." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Writing the inward total as =SUMIFS(Inward!$B$2:$B$50,$A2,Inward!$D$2:$D$50).", right: "=SUMIFS(Inward!$D$2:$D$50,Inward!$B$2:$B$50,$A2) — the sum range comes first", why: "SUMIFS is unlike SUMIF: the range to add is always the first argument, and every pair after it is a criteria range plus its criteria. Swapped arguments either error out or add the wrong column." },
      { wrong: "Typing opening stock into Current Stock as plain numbers instead of looking it up.", right: "=VLOOKUP($A2,'Item Master'!$A$2:$D$6,4,FALSE) so the master stays the single source", why: "Typed numbers never update. When a stocktake corrects the opening balance in the master, your report keeps showing the old figure — the classic way two sheets start disagreeing." },
      { wrong: "Writing the alert as =IF($F2<$I2,\"REORDER\",\"OK\").", right: "=IF($F2<=$I2,\"REORDER\",\"OK\")", why: "An item sitting exactly at its reorder level has, by definition, reached it. With a strict less-than, that item slips through unflagged until it goes one unit lower." },
      { wrong: "Typing item codes into register rows by hand instead of using the dropdown.", right: "Pick the code from the validation list in column B", why: "One mistyped code creates a movement that matches no item: stock silently overstates, and SUMIFS finds no error to report because nothing is technically broken." }
    ],
    shortcuts: [
      { keys: "Alt + Down", action: "Open the item-code dropdown in a register row" },
      { keys: "F4", action: "Lock lookup ranges while writing the formulas" },
      { keys: "Ctrl + Z", action: "Undo a test entry when rehearsing the system" }
    ],
    exercises: [
      { task: "Add a column K reporting the total inward value each item has cost this month.", hint: "SUMIFS again — this time the sum range is column F of the Inward sheet, not column D.", solution: "=SUMIFS(Inward!$F$2:$F$50,Inward!$B$2:$B$50,$A2)" },
      { task: "On a small report block, count how many items are currently flagged REORDER.", hint: "COUNTIF over the status column.", solution: "=COUNTIF($J$2:$J$6,\"REORDER\")" },
      { task: "Show the name of the item holding the highest stock value, without sorting.", hint: "Find the biggest value with MAX in column H, then match it back to a name with INDEX and MATCH.", solution: "=INDEX($B$2:$B$6,MATCH(MAX($H$2:$H$6),$H$2:$H$6,0))" }
    ],
    quiz: [
      { q: "Which formula correctly computes the closing stock of Notebooks in row 2?",
        options: [
          "=$D2+$E2-$C2",
          "=$C2+$D2-$E2",
          "=SUM($C2:$E2)",
          "=$C2-$D2+$E2"
        ],
        answer: 1,
        explanation: "Closing = Opening (C) + Inward (D) − Outward (E). Option 1 adds and subtracts in the wrong order and would count outward stock as a gain." },
      { q: "In =SUMIFS(Inward!$D$2:$D$50,Inward!$B$2:$B$50,$A2), what does the middle pair of arguments do?",
        options: [
          "Defines the rows to add — the quantities",
          "Sets the criteria range (the register's item codes) and the criteria (this row's item code)",
          "Sorts the register before summing",
          "Limits the sum to the first 50 rows only"
        ],
        answer: 1,
        explanation: "SUMIFS reads: sum this first range, but only where that criteria range equals this criteria. The pair `Inward!$B$2:$B$50, $A2` selects the rows belonging to one item." },
      { q: "Staplers close at exactly 30 units and their reorder level is 30. Which status formula flags it correctly?",
        options: [
          "=IF($F2<$I2,\"REORDER\",\"OK\")",
          "=IF($F2<=$I2,\"REORDER\",\"OK\")",
          "=IF($F2>$I2,\"REORDER\",\"OK\")",
          "=IF($F2=$I2,\"OK\",\"REORDER\")"
        ],
        answer: 1,
        explanation: "Reaching the reorder level is the trigger to order, so the comparison must include equality. Option 4 flags everything except the exact-boundary case — backwards." },
      { q: "Why does the stock value use the rate from the Item Master rather than a rate typed into the Current Stock sheet?",
        options: [
          "Because VLOOKUP is faster than typing",
          "So a single price correction in the master instantly revalues the whole report",
          "Because rates must always be numbers",
          "So the value column cannot be edited by accident"
        ],
        answer: 1,
        explanation: "Valuation parameters belong in one place. Typed rates freeze yesterday's prices into the report and guarantee a disagreement the day the cost changes." },
      { q: "A register row has the code \"ITM003 \" with a trailing space. What does the Current Stock sheet do?",
        options: [
          "Shows #N/A on the stock row",
          "Ignores the movement — its quantity never reaches the stock figures",
          "Rounds the quantity down",
          "Converts the code automatically"
        ],
        answer: 1,
        explanation: "SUMIFS matches text exactly, and \"ITM003 \" is not \"ITM003\". Nothing errors — the movement just silently vanishes from the totals, which is why the dropdown validation matters." }
    ],
    summary: [
      "Every stock system is three sheets: a master, movement registers, and a live balance sheet that reads them.",
      "Closing Stock = Opening + SUMIFS(inward) − SUMIFS(outward) — one formula pattern, five items.",
      "In SUMIFS the sum range comes first; every criteria pair after it selects which rows to include.",
      "Reorder alerts come from a parameter in the master plus a simple IF with conditional formatting.",
      "Test a system by adding movements and changing parameters — a good file updates everywhere without being touched."
    ],
    downloads: [
      { file: "assets/files/level7/inventory-management-system.xlsx", label: "Inventory Management System — exercise file", note: "Item master and sample inward/outward entries to build on" },
      { file: "assets/files/level7/inventory-management-system-solution.xlsx", label: "Inventory Management System — completed solution", note: "Finished system with live stock calculations and alerts" }
    ],
    keywords: ["inventory management excel", "inward outward register", "SUMIFS stock calculation", "closing stock formula", "reorder level alert", "stock value report", "item master", "current stock sheet"]
  },
  {
    id: "l7-project-sales-tracking",
    title: "Project: Sales Tracking System",
    difficulty: "Intermediate",
    duration: 40,
    version: "",
    objectives: [
      "Turn a raw sales log into a monthly revenue summary with SUMIFS",
      "Compare each salesperson against a target and compute achievement %",
      "Rank the team and find the best month with RANK and INDEX/MATCH",
      "Filter the log and build a simple revenue chart",
      "Design a tracking file that keeps working as the log grows"
    ],
    intro: "Sunrise Traders sells to offices across four regions — North, South, East and West — through six salespeople. Every invoice lands in one log sheet, and this project turns that log into a management tool: monthly revenue with growth, each salesperson's achievement against target, a ranking of the team, and a chart the sales meeting can put on a screen.\n\nThis is the summary-layer project: no new Excel machinery at all, but two habits that separate working trackers from broken ones. First, a helper column converts messy daily dates into clean month labels so SUMIFS has something solid to match. Second, every formula scans a generous range like rows 2 to 200, so the tracker keeps working long after you stop maintaining it.",
    keyIdea: "Helper columns and generous formula ranges keep a tracker working as the log grows.",
    sections: [
      {
        heading: "What You Will Build",
        blocks: [
          { type: "p", text: "Three sheets. The **Sales Log** is the only sheet anyone types into — date, invoice number, region, product, salesperson, quantity and rate, with the amount computed. The **Monthly Summary** condenses the log into one revenue line per month plus growth. The **Performance** sheet puts each salesperson against their three-month target and ranks the team." },
          { type: "ul", items: [
            "**Sales Log** — about 30 rows covering August, September and October 2026, with a Month helper column.",
            "**Monthly Summary** — revenue per month, growth %, and the best month.",
            "**Performance** — target vs actual per salesperson, achievement %, rank, and a status flag."
          ] },
          { type: "p", text: "The log covers three months: August, September and October 2026. The exercise file contains the log fully typed — roughly thirty invoices across all four regions and all five products — so you spend your time building the summary layer, which is the part managers actually look at." }
        ]
      },
      {
        heading: "Skills You Will Use",
        blocks: [
          { type: "ul", items: [
            "**SUMIF and SUMIFS** (Level 4) — revenue by month, by salesperson and by region.",
            "**TEXT with dates** (Level 2) — the month helper column.",
            "**Absolute references and generous ranges** (Level 4) — formulas that survive new data.",
            "**RANK, MAX and INDEX/MATCH** (Level 5) — the leaderboard and best-month cells.",
            "**AutoFilter and a basic column chart** (Levels 2 and 6) — a quick look and a quick picture."
          ] },
          { type: "p", text: "Nothing here is new — that is the point. A tracking system is a disciplined arrangement of tools you already own." }
        ]
      },
      {
        heading: "Open the Exercise File",
        blocks: [
          { type: "p", text: "Download **Sales Tracking System — exercise file**. Inside:" },
          { type: "ul", items: [
            "**Sales Log** — headers in row 1: Date, Invoice No, Region, Product, Salesperson, Qty, Rate, Amount, Month. Rows 2 to about 31 are filled; Amount and Month are formulas already.",
            "**Monthly Summary** — headers only: Month, Revenue, Growth %.",
            "**Performance** — headers only: Salesperson, Target, Actual, Achievement %, Rank, Status."
          ] },
          { type: "table", headers: ["Column", "Sales Log column", "Example"],
            rows: [
              ["A", "Date", "03-Aug-26"],
              ["B", "Invoice No", "INV-2001"],
              ["C", "Region", "North"],
              ["D", "Product", "Notebooks"],
              ["E", "Salesperson", "Arjun Mehta"],
              ["F", "Qty", "50"],
              ["G", "Rate (₹)", "45"],
              ["H", "Amount (₹) = Qty × Rate", "2250"],
              ["I", "Month = TEXT helper", "Aug-26"]
            ] },
          { type: "tip", text: "Click `I2` on the log and read the helper formula `=TEXT(A2,\"mmm-yy\")`. That single formula turns every date into a stable text label like `Aug-26` — the key that lets every summary formula in this project use simple text criteria instead of fragile date arithmetic." }
        ]
      },
      {
        heading: "Build Section 1: The Monthly Revenue Summary",
        blocks: [
          { type: "p", text: "One row per month, one SUMIFS per row, and a growth column that shows whether the business is moving forward." },
          { type: "steps", items: [
            "On **Monthly Summary**, type `Aug-26`, `Sep-26`, `Oct-26` in `A2:A4`. They must match the helper column exactly — same spelling, same order.",
            "In `B2` type `=SUMIFS('Sales Log'!$H$2:$H$200,'Sales Log'!$I$2:$I$200,$A2)` and fill down to `B4`.",
            "In `C3` type `=($B3-$B2)/$B2` and fill down to `C4`, then format `C3:C4` as percentage with 1 decimal. Leave `C2` blank — there is no month before August to grow from.",
            "In `A6` type `Best Month` and in `B6` type `=INDEX($A$2:$A$4,MATCH(MAX($B$2:$B$4),$B$2:$B$4,0))`.",
            "In `A7` type `Total 3-Month Revenue (₹)` and in `B7` type `=SUM($B$2:$B$4)`.",
            "Format `B2:B7` with thousands separators, bold the headers, and add borders."
          ] },
          { type: "formula", name: "Monthly revenue with SUMIFS",
            syntax: "=SUMIFS(sum_range, criteria_range, criteria)",
            example: "=SUMIFS('Sales Log'!$H$2:$H$200,'Sales Log'!$I$2:$I$200,$A2)",
            explanation: "Adds the Amount column of the log, but only rows whose Month helper reads the label in A2. The ranges reach row 200 while the data ends near row 31 — that headroom is deliberate: paste next month's invoices below the log and every summary updates with no edits at all." },
          { type: "warning", text: "If a revenue cell shows 0 while the log clearly has sales that month, compare the two texts letter by letter: `Aug-26` in the summary does not match `Aug-2026` in the helper. The criteria and the helper must come from the same TEXT format string." }
        ]
      },
      {
        heading: "Build Section 2: Target vs Actual and the Team Leaderboard",
        blocks: [
          { type: "p", text: "Targets make a tracker honest. The Performance sheet compares each salesperson's actual three-month revenue against their target, ranks the team, and flags anyone below 90%." },
          { type: "steps", items: [
            "On **Performance**, type the six names in `A2:A7`: Arjun Mehta, Priya Nair, Rahul Verma, Sneha Iyer, Vikram Singh, Divya Menon.",
            "Type the three-month targets in `B2:B7`: `150000`, `200000`, `120000`, `250000`, `130000`, `160000`.",
            "In `C2` type `=SUMIF('Sales Log'!$E$2:$E$200,$A2,'Sales Log'!$H$2:$H$200)` and fill down — each person's total sales by name.",
            "In `D2` type `=$C2/$B2` and fill down, then format as percentage with 0 decimals.",
            "In `E2` type `=RANK($D2,$D$2:$D$7,0)` and fill down — 0 means rank from the highest.",
            "In `F2` type `=IF($D2>=1,\"Target Met\",IF($D2>=0.9,\"Close\",\"Below\"))` and fill down.",
            "Select `F2:F7` and apply conditional formatting: Equal To `Target Met` in green, Equal To `Below` in light red."
          ] },
          { type: "grid", title: "Performance sheet layout",
            rows: [
              ["Salesperson", "Target (₹)", "Actual (₹)", "Achievement %", "Rank", "Status", ""],
              ["Arjun Mehta", "150000", "—", "—", "—", "—", ""],
              ["Priya Nair", "200000", "—", "—", "—", "—", ""],
              ["Rahul Verma", "120000", "—", "—", "—", "—", ""],
              ["Sneha Iyer", "250000", "—", "—", "—", "—", ""],
              ["Vikram Singh", "130000", "—", "—", "—", "—", ""],
              ["Divya Menon", "160000", "—", "—", "—", "—", ""]
            ] },
          { type: "formula", name: "RANK",
            syntax: "=RANK(number, ref, [order])",
            example: "=RANK($D2,$D$2:$D$7,0)",
            explanation: "Returns where D2 stands inside the whole achievement column: 1 for the best when order is 0 (descending). The ref range is locked so all six cells rank against the same set. Ties share the same rank, which is realistic and fine." }
        ]
      },
      {
        heading: "Build Section 3: Region Insight, Filters and the Chart",
        blocks: [
          { type: "p", text: "Two quick wins finish the tracker: a region block that shows where the money comes from, and a chart for the sales meeting screen." },
          { type: "steps", items: [
            "On **Monthly Summary**, build a small region block from row 10: headers `Region` and `Revenue` in `A10:B10`, then `North`, `South`, `East`, `West` in `A11:A14`.",
            "In `B11` type `=SUMIFS('Sales Log'!$H$2:$H$200,'Sales Log'!$C$2:$C$200,$A11)` and fill down to `B14`.",
            "On the **Sales Log**, click any header cell and press `[[Ctrl]] + [[Shift]] + [[L]]` to switch on AutoFilter. Filter Product to `Printers` to see every printer sale at once, then clear the filter.",
            "On **Monthly Summary**, select `A2:B4`, choose **Insert → Charts → Column**, and pick the plain clustered column.",
            "Click the chart title and type `Monthly Revenue (₹)`. Move the chart to `D2`, size it to cover roughly `D2:J14`, and keep the region block visible below.",
            "Save, then add a fake November row to the log — say a ₹9,000 Notebook sale in the West — and confirm the chart, the region block and the Performance sheet all move. Delete the test row."
          ] },
          { type: "tip", text: "The chart is built from the summary cells, not the raw log — that is a rule, not a preference. Charts read clean aggregated ranges far better, and any new data only changes three summary cells instead of redrawing from thirty log rows." }
        ]
      },
      {
        heading: "Level Up Your System",
        blocks: [
          { type: "p", text: "The tracker is complete. The exercises below push it toward the interactive dashboard you will build properly in Project 9." }
        ]
      },
      {
        heading: "Compare With the Solution",
        blocks: [
          { type: "steps", items: [
            "Open **Sales Tracking System — completed solution** and compare the Monthly Summary first: your three revenue figures and the total must match exactly.",
            "Click `B2` on the solution and check your SUMIFS ranges both stop at row 200 — the single most common shortcoming in learner versions is a range ending at row 31.",
            "Compare the Performance sheet: ranks, status flags and the conditional colours.",
            "On the solution's log, switch AutoFilter on and off and confirm your sheet behaves the same.",
            "Compare the chart placement and title, then close the solution without saving."
          ] },
          { type: "p", text: "One nuance to notice in the solution: the two `East` region figures and the `Rahul Verma` figures come from different columns of the log — region lives in column C, salesperson in E. Summaries differ by which column the SUMIFS tests, and mixing those up produces plausible-looking wrong numbers." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Typing the criteria as \"Aug-2026\" while the log's helper column produces \"Aug-26\".", right: "Use the same TEXT format on both sides: =TEXT(A2,\"mmm-yy\") and criteria \"Aug-26\"", why: "SUMIFS compares text exactly. A mismatch returns 0 with no error at all — the most confusing failure in tracking sheets because everything looks correct and adds up to nothing." },
      { wrong: "Scanning only =SUMIFS('Sales Log'!$H$2:$H$31,...) because that is where the data ends today.", right: "Extend ranges to $H$2:$H$200 so future rows are included automatically", why: "Next month's invoices land in rows 32 onward and never appear in any summary. The tracker looks alive but reports only the past." },
      { wrong: "Computing growth as =($B2-$B3)/$B3.", right: "=($B3-$B2)/$B2 — later month minus earlier month, over the earlier month", why: "Growth is the change relative to where you started. Reversed, it reports the previous month's change against the wrong base and flips the sign." },
      { wrong: "Ranking with =RANK($D2,D2:D7,0) without locking the ref range.", right: "=RANK($D2,$D$2:$D$7,0)", why: "As the formula fills down, the comparison window shrinks — by row 7 it is ranking one number against itself and everyone gets rank 1." }
    ],
    shortcuts: [
      { keys: "Ctrl + Shift + L", action: "Toggle AutoFilter on the sales log" },
      { keys: "F4", action: "Lock a range while writing a formula" },
      { keys: "Ctrl + 1", action: "Format the revenue and percentage columns" }
    ],
    exercises: [
      { task: "On the Monthly Summary, add a Product block showing total revenue for Laptops across all three months.", hint: "SUMIFS again, testing the Product column D of the log.", solution: "=SUMIFS('Sales Log'!$H$2:$H$200,'Sales Log'!$D$2:$D$200,\"Laptops\")" },
      { task: "Show the top salesperson's name without sorting or reading the ranks yourself.", hint: "Find rank 1's achievement with MATCH, then convert the position to a name with INDEX over the name column.", solution: "=INDEX($A$2:$A$7,MATCH(1,$E$2:$E$7,0))" },
      { task: "Add a 90%-club counter: how many salespeople achieved at least 90% of target?", hint: "COUNTIF can test a computed percentage column directly.", solution: "=COUNTIF($D$2:$D$7,\">=90%\")" }
    ],
    quiz: [
      { q: "Why does every summary formula scan rows 2 to 200 when the log only has about 30 rows?",
        options: [
          "To make the workbook open faster",
          "So new invoices pasted below the log are picked up automatically, with no formula edits",
          "Because SUMIFS requires at least 200 rows",
          "To avoid absolute references"
        ],
        answer: 1,
        explanation: "The headroom is the design. Blank rows contribute nothing to a SUMIFS, so generous ranges cost nothing today and save every future month from manual maintenance." },
      { q: "What does the Month helper column contain, and why does it exist?",
        options: [
          "Serial date numbers, for faster calculation",
          "Text labels like Aug-26 produced by =TEXT(A2,\"mmm-yy\"), giving SUMIFS a clean exact-match key",
          "Month names typed by hand each row",
          "The day of the month, for daily totals"
        ],
        answer: 1,
        explanation: "Dates carry hidden time and serial-number complexity that makes criteria fragile. One TEXT column turns every date into a stable label that both people and SUMIFS can match exactly." },
      { q: "Priya Nair's actual sales are ₹1,90,000 against a target of ₹2,00,000. What does the status formula show?",
        options: [
          "Target Met",
          "Close",
          "Below",
          "#N/A"
        ],
        answer: 1,
        explanation: "Achievement is 95%, which fails the >=100% test but passes the >=90% test, so the middle branch of the nested IF returns \"Close\"." },
      { q: "Which formula finds the best month by revenue without sorting?",
        options: [
          "=MAX($B$2:$B$4)",
          "=INDEX($A$2:$A$4,MATCH(MAX($B$2:$B$4),$B$2:$B$4,0))",
          "=RANK($B$2:$B$4,1)",
          "=SUMIF($B$2:$B$4,MAX($B$2:$B$4))"
        ],
        answer: 1,
        explanation: "MAX finds the biggest revenue, MATCH turns it into a row position, and INDEX converts that position into the month label. MAX alone returns a number, not the month's name." },
      { q: "You need revenue for the East region. Which column of the log must the SUMIFS criteria test?",
        options: [
          "Column D — Product",
          "Column C — Region",
          "Column E — Salesperson",
          "Column I — Month"
        ],
        answer: 1,
        explanation: "Region lives in column C. Testing the product column instead returns the Notebook or Laptop total — a plausible number that answers the wrong question." }
    ],
    summary: [
      "A tracking system is a log everyone types into, plus a summary layer nobody edits.",
      "A TEXT helper column converts dates into exact-match labels — the hinge every SUMIFS turns on.",
      "Scan generous ranges (rows 2 to 200) so the tracker survives months of new data without edits.",
      "Targets turn a report into a scoreboard: achievement %, RANK and a status flag with conditional colour.",
      "Build charts from summary cells, never from the raw log."
    ],
    downloads: [
      { file: "assets/files/level7/sales-tracking-system.xlsx", label: "Sales Tracking System — exercise file", note: "Three months of sales data and a blank summary structure" },
      { file: "assets/files/level7/sales-tracking-system-solution.xlsx", label: "Sales Tracking System — completed solution", note: "Finished tracker with summaries, targets and charts" }
    ],
    keywords: ["sales tracking excel", "SUMIFS monthly revenue", "TEXT month helper column", "target vs actual", "RANK leaderboard", "growth percentage", "AutoFilter sales log", "sales performance sheet", "region summary"]
  },
);
