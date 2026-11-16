// level5-1.js — Level 5: Real-World Office Applications (metadata + lessons 1-7)
window.LEVEL_META = window.LEVEL_META || {};
window.LEVEL_LESSONS = window.LEVEL_LESSONS || {};

window.LEVEL_META[5] = {
  id: 5,
  title: "Real-World Office Applications",
  icon: "briefcase",
  difficulty: "Intermediate",
  subtitle: "Use Excel the way offices actually do: HR, payroll, inventory, sales and billing",
  description: "Build the spreadsheets every office needs: employee databases, attendance, payroll, inventory, purchase and sales trackers, budgets and invoices."
};

(window.LEVEL_LESSONS[5] = window.LEVEL_LESSONS[5] || []).push(
  {
    id: "l5-employee-database",
    title: "Building an Employee Database",
    difficulty: "Intermediate",
    duration: 20,
    version: "",
    objectives: [
      "Lay out an employee master sheet that behaves like a proper database",
      "Add dropdown lists with Data Validation to keep Department and Status clean",
      "Give every employee a consistent, unique ID",
      "Find any record in seconds using filters",
      "Count headcount by department with COUNTIF"
    ],
    intro: "Who works here? Which team, what salary, since when? Small companies keep this in a notebook. Growing companies keep it in Excel.\n\nIn this lesson you build the employee master for Sunrise Traders, an office supplies firm. Six employees work there today: Arjun Mehta, Priya Nair, Rahul Verma, Sneha Iyer, Vikram Singh and Divya Menon. This one sheet will feed attendance, leave, payroll and reports. Build it well once, and every later lesson becomes easy.",
    keyIdea: "One clean master table of employees feeds every other HR sheet you build later.",
    sections: [
      {
        heading: "Why a Database, Not Just a List",
        blocks: [
          { type: "p", text: "A list is anything you type. A database is a list with rules. The rules let Excel answer questions fast. Ask \"How many people are in Sales?\" or \"Who joined this year?\" Excel counts, without you re-reading every row." },
          { type: "ul", items: [
            "**One header row** at the top, and data from row 2 downwards. Filters and formulas expect this shape.",
            "**No merged cells** anywhere in the data area. A merged cell belongs to one column only, so sorting, filtering and ranges all break around it.",
            "**One employee per row.** Never put two people in one row to \"save space\".",
            "**No blank rows or columns** inside the data. A blank row silently splits your table in two, and `COUNTIF` on the lower part stops working.",
            "**One value per cell.** Phone and email get their own columns; do not write \"98450 12345 / 98450 67890\" in one cell."
          ] },
          { type: "p", text: "These are not style preferences. `COUNTIF`, `VLOOKUP`, filters and sorting all assume this shape. One merged header cell, added \"to make it look nice\", can ruin the whole sheet." }
        ]
      },
      {
        heading: "The Columns of the Employee Master",
        blocks: [
          { type: "table", headers: ["Column", "Content", "Example"], rows: [
            ["A - Emp ID", "Unique code, text", "EMP-001"],
            ["B - Name", "Full name", "Arjun Mehta"],
            ["C - Department", "From a fixed list", "Sales"],
            ["D - Designation", "Job title", "Sales Executive"],
            ["E - Joining Date", "Real date, not text", "12-Apr-2021"],
            ["F - Salary", "Monthly basic salary in rupees", "25000"],
            ["G - Phone", "Digits only", "9845012345"],
            ["H - Email", "Lowercase", "arjun.mehta@sunrisetrd.com"],
            ["I - Status", "Active or Inactive", "Active"]
          ] },
          { type: "grid", title: "The Sunrise Traders employee master", rows: [
            ["Emp ID", "Name", "Department", "Designation", "Joining Date", "Salary", "Status"],
            ["EMP-001", "Arjun Mehta", "Sales", "Sales Executive", "12-Apr-2021", "25000", "Active"],
            ["EMP-002", "Priya Nair", "Sales", "Sales Manager", "03-Jan-2020", "40000", "Active"],
            ["EMP-003", "Rahul Verma", "Accounts", "Accountant", "21-Jul-2022", "28000", "Active"],
            ["EMP-004", "Sneha Iyer", "Stores", "Store Keeper", "15-Aug-2023", "22000", "Active"],
            ["EMP-005", "Vikram Singh", "Admin", "Office Administrator", "09-Feb-2022", "24000", "Active"],
            ["EMP-006", "Divya Menon", "HR", "HR Executive", "27-Sep-2022", "26000", "Active"]
          ] },
          { type: "tip", text: "Type the Joining Date as a real date (12-Apr-2021), not text. Real dates sort, filter by year, and work in formulas like years of service. Text dates just sit there, looking right and failing everywhere else." }
        ]
      },
      {
        heading: "Dropdowns That Keep Data Clean",
        blocks: [
          { type: "p", text: "Six people typing in the same column will write \"Sales\", \"sales\", \"Sales Dept\" and \"SLS\". Every spelling breaks `COUNTIF`. The fix is Data Validation. It turns the column into a dropdown list." },
          { type: "steps", items: [
            "Select the Department cells, for example `C2:C200`.",
            "Go to the Data tab and click Data Validation.",
            "Set Allow to List.",
            "In Source, type: `Sales,Accounts,Stores,Admin,HR` (comma separated, no spaces needed).",
            "Make sure In-cell dropdown is ticked and click OK.",
            "Repeat for the Status column with the list `Active,Inactive`."
          ] },
          { type: "p", text: "From now on, users pick from the list instead of typing. Wrong entries are rejected with a message, so the column stays consistent forever." },
          { type: "warning", text: "Data Validation only checks what a user types. If someone pastes data over the cells, validation is bypassed. That is why masters also need the protection steps later in this lesson." }
        ]
      },
      {
        heading: "ID Format: Small Rule, Big Payoff",
        blocks: [
          { type: "ul", items: [
            "Pick one pattern and never change it: `EMP-001`, `EMP-002`, and so on.",
            "Zero-pad to the same length. `EMP-001` sorts correctly; `EMP-1`, `EMP-10`, `EMP-2` sorts alphabetically into nonsense.",
            "Store the ID as text, never as a number, so Excel does not strip the leading zeros.",
            "Never reuse an ID, even after an employee leaves. Old payslips and leave records must still point to exactly one person."
          ] },
          { type: "p", text: "The ID is how the other sheets in this level reference people. Attendance, leave and payroll will look people up by ID, not by name, because names can repeat but IDs cannot." }
        ]
      },
      {
        heading: "Finding Records Fast",
        blocks: [
          { type: "p", text: "Once the master is shaped correctly, Excel's own tools do the searching for you." },
          { type: "steps", items: [
            "Click any cell in the master and press [[Ctrl]] + [[Shift]] + [[L]] to turn on filters.",
            "To see only the Sales team, open the filter on the Department header and tick Sales only.",
            "To find one person by name, open the Name filter and use Text Filters, Contains.",
            "To find everyone who joined after 2022, filter Joining Date by Date Filters, After.",
            "Press [[Ctrl]] + [[Shift]] + [[L]] again to clear the filters when done."
          ] },
          { type: "tip", text: "Freeze the header row with View, Freeze Panes, Freeze Top Row so headings stay visible while you scroll through hundreds of employees." }
        ]
      },
      {
        heading: "Headcount and Protecting the Data",
        blocks: [
          { type: "p", text: "The most-asked question in any HR meeting is headcount. With a clean Department column, one formula answers it:" },
          { type: "formula", name: "COUNTIF", syntax: "=COUNTIF(range, criteria)", example: "=COUNTIF($C$2:$C$200, \"Sales\")", explanation: "Counts how many employees have the word Sales in the Department column. The dollar signs lock the range so you can copy the formula down a summary table." },
          { type: "grid", title: "A small headcount summary built with COUNTIF", rows: [
            ["Department", "Headcount", ""],
            ["Sales", "=COUNTIF($C$2:$C$200, \"Sales\")", ""],
            ["Accounts", "=COUNTIF($C$2:$C$200, \"Accounts\")", ""],
            ["Stores", "=COUNTIF($C$2:$C$200, \"Stores\")", ""],
            ["Admin", "=COUNTIF($C$2:$C$200, \"Admin\")", ""],
            ["HR", "=COUNTIF($C$2:$C$200, \"HR\")", ""]
          ] },
          { type: "p", text: "Total active employees is just as simple: `=COUNTIF($I$2:$I$200, \"Active\")` returns 6 for Sunrise Traders today." },
          { type: "p", text: "Finally, protect what you built. When an employee leaves, do not delete their row. Change Status to Inactive instead. The row still answers old payroll and leave questions. Headcount stays correct too, because it counts only Active employees." },
          { type: "ul", items: [
            "Use Review, Protect Sheet with a password so formulas and headers cannot be edited by accident.",
            "Keep a backup copy of the master before any bulk change, such as a salary revision.",
            "Let only one or two people edit the master. Everyone else should be reading a filtered view or a copy."
          ] }
        ]
      }
    ],
    mistakes: [
      { wrong: "Merging the header cells A1:I1 into one big \"Employee Master\" title.", right: "Keep the title on a separate row above the table, or skip it and let the sheet tab name speak.", why: "Merged cells break sorting, filtering and range-based formulas. One merged cell can make a whole column unsortable." },
      { wrong: "Typing departments freely as \"Sales\", \"sales \" and \"Sales Dept\".", right: "Use a Data Validation dropdown with the fixed list Sales, Accounts, Stores, Admin, HR.", why: "COUNTIF matches exact text. \"sales \" with a trailing space is a different string and gets miscounted." },
      { wrong: "Deleting the row of an employee who has resigned.", right: "Set Status to Inactive and keep the row.", why: "Old attendance, leave and payroll records reference that ID. Deleting the row breaks history and headcount math." },
      { wrong: "Entering IDs as numbers like 1, 2, 3 without padding.", right: "Store EMP-001 as text with zero padding.", why: "Unpadded numbers sort as 1, 10, 2 and lose the leading zeros that make the code readable and sortable." }
    ],
    shortcuts: [
      { keys: "Ctrl + Shift + L", action: "Turn filters on or off for the whole table" },
      { keys: "Ctrl + T", action: "Convert the range into a structured Excel Table" },
      { keys: "Ctrl + Down", action: "Jump to the last row of the master" }
    ],
    exercises: [
      { task: "In the employee-records workbook, add Data Validation dropdowns to the Department and Status columns on the Exercise sheet, with lists Sales, Accounts, Stores, Admin, HR and Active, Inactive.", hint: "Select the column cells first, then Data tab, Data Validation, Allow: List.", solution: "Both columns now show a dropdown arrow when selected, and typing anything outside the list is rejected." },
      { task: "Build a small summary that counts headcount for each of the five departments using COUNTIF, then count Active employees.", hint: "Lock the ranges with dollar signs, for example $C$2:$C$200, so you can copy the formula down.", solution: "Sales 2, Accounts 1, Stores 1, Admin 1, HR 1, Active 6. Compare with the Solution sheet." },
      { task: "Filter the Exercise sheet to show only employees who joined before 2023, and note who they are.", hint: "Ctrl + Shift + L, then Date Filters on the Joining Date column.", solution: "Arjun Mehta (2021), Priya Nair (2020), Rahul Verma (2022), Vikram Singh (2022), Divya Menon (2022). Only Sneha Iyer joined in 2023." }
    ],
    quiz: [
      { q: "Why should an employee master have no merged cells?", options: ["They make the file too large", "They break sorting, filtering and range-based formulas", "They stop colours from printing", "They hide the column letters"], answer: 1, explanation: "A merged cell really belongs to just one column. Excel's data tools get confused about which column a row belongs to, so sorting and filtering fail around the merge." },
      { q: "Which formula counts how many employees work in Sales?", options: ["=SUM(C2:C7, \"Sales\")", "=COUNT(C2:C7, \"Sales\")", "=COUNTIF($C$2:$C$200, \"Sales\")", "=COUNTA($C$2:$C$200)"], answer: 2, explanation: "COUNTIF counts cells that meet a text condition. SUM adds numbers, COUNT counts numbers only, and COUNTA counts all non-empty cells." },
      { q: "An employee resigns. What is the correct action in the master?", options: ["Delete the row completely", "Change Status to Inactive and keep the row", "Move the row to the bottom of the sheet", "Leave the row but blank out the name"], answer: 1, explanation: "Keeping the row preserves history for old payslips and leave records. Headcount formulas count only Active employees, so totals stay correct." },
      { q: "Which feature stops users typing \"Sales Dept\" instead of \"Sales\"?", options: ["Conditional Formatting", "Freeze Panes", "Data Validation with a List", "COUNTIF"], answer: 2, explanation: "Data Validation with a List turns the column into a dropdown and rejects anything not on the list, keeping every entry consistent." }
    ],
    summary: [
      "A database is a list with rules: one header row, no merged cells, no blank rows, one record per row.",
      "Data Validation dropdowns keep Department and Status consistent so COUNTIF works.",
      "Use zero-padded text IDs like EMP-001 and never reuse them.",
      "Count headcount with COUNTIF on the Department and Status columns.",
      "When someone leaves, set Status to Inactive instead of deleting the row."
    ],
    downloads: [
      { file: "assets/files/level5/employee-records.xlsx", label: "Employee records workbook", note: "Exercise and Solution sheets — employee master data ready for the database lesson" }
    ],
    keywords: ["employee database", "employee master", "data validation dropdown", "headcount countif", "employee id format", "hr excel sheet"]
  },
  {
    id: "l5-attendance",
    title: "Attendance Management System",
    difficulty: "Intermediate",
    duration: 25,
    version: "",
    objectives: [
      "Build a monthly attendance register with short codes instead of words",
      "Count P, A and L days per employee with COUNTIF",
      "Calculate attendance percentage the correct way",
      "Shade weekends automatically with conditional formatting",
      "Summarise the whole month on a separate summary sheet"
    ],
    intro: "HR keeps one register per month. Who was present, absent, on leave, on holiday, or on a week off? Doing this by hand wastes hours and invites errors. In Excel, a register plus five `COUNTIF` formulas gives the full picture in seconds.\n\nSunrise Traders marks attendance with single-letter codes. P is present, A is absent, L is leave, H is holiday and WO is week off. In this lesson you build the October 2026 register for six employees and turn it into a monthly summary.",
    keyIdea: "Mark each day with one short code, and formulas count the whole month for you.",
    sections: [
      {
        heading: "Why Codes Beat Words",
        blocks: [
          { type: "table", headers: ["Code", "Meaning", "Counts as a working day?"], rows: [
            ["P", "Present", "Yes"],
            ["A", "Absent", "Yes"],
            ["L", "Approved leave", "Yes"],
            ["H", "Public holiday (e.g. 2-Oct for Gandhi Jayanti)", "No"],
            ["WO", "Week off (Sunday, 2nd and 4th Saturday)", "No"]
          ] },
          { type: "ul", items: [
            "COUNTIF needs exact text. \"Present\", \"present \" and \"P\" are three different strings to Excel, but one letter per cell means one possible spelling.",
            "One keypress fills a cell. For 31 days times 30 employees, that is thousands of extra keypresses saved every month.",
            "Single letters line up neatly in narrow day columns, so the register stays readable."
          ] },
          { type: "warning", text: "COUNTIF is not case-sensitive, so \"p\" and \"P\" both match. The real danger is variants like \"Pr\" or \"Present\" — they count nothing. Pick the codes above and never invent new ones mid-month." }
        ]
      },
      {
        heading: "Building the Monthly Register",
        blocks: [
          { type: "p", text: "One sheet per month. Employees go down the rows, days go across the columns. Sunrise Traders uses a layout like this:" },
          { type: "grid", title: "Attendance register, October 2026 (day columns continue up to 31 in column AG)", rows: [
            ["Month: October 2026", "", "", "", ""],
            ["Emp ID", "Name", "1", "2", "3"],
            ["EMP-001", "Arjun Mehta", "P", "P", "H"],
            ["EMP-002", "Priya Nair", "P", "P", "H"],
            ["EMP-003", "Rahul Verma", "P", "P", "H"]
          ] },
          { type: "p", text: "Columns A and B hold the ID and name. Columns C to AG hold days 1 to 31. C is day 1, so day 31 lands in column AG. Columns AH to AK hold the totals: P count, A count, L count and attendance percentage. Row 1 is a note cell with the month. Row 2 is the header. Employees start in row 3." },
          { type: "p", text: "Leave the weekend columns blank and shade them. Then nobody is tempted to mark P on a Sunday. Holidays get an H inside the working-day area. October 2026 has four Sundays (4, 11, 18, 25), two week-off Saturdays (10 and 24, the 2nd and 4th) and one holiday (2 October)." },
          { type: "steps", items: [
            "Select the day area, for example `C3:AG8`.",
            "Open Home, Conditional Formatting, New Rule, Use a formula.",
            "For Sundays, enter: `=WEEKDAY(DATE(2026,10,C$2),1)=1`",
            "Add a second rule with: `=OR(C$2=10,C$2=24)` for the 2nd and 4th Saturdays.",
            "Give both rules a light grey fill and click OK."
          ] },
          { type: "p", text: "The formula `DATE(2026,10,C$2)` builds each day's real date from row 2. `WEEKDAY(...,1)` returns 1 for Sundays. The shading moves with the day numbers. Next month, change only the day-number row and the `DATE` formula." }
        ]
      },
      {
        heading: "Counting Each Code with COUNTIF",
        blocks: [
          { type: "formula", name: "COUNTIF (per employee row)", syntax: "=COUNTIF(day_range, \"P\")", example: "=COUNTIF(C3:AG3, \"P\")", explanation: "Counts the P codes across one employee's row. In the totals column AH, this gives days present. Copy it down for every employee, and repeat with \"A\" and \"L\" for the other totals." },
          { type: "grid", title: "The totals columns at the right end of the register", rows: [
            ["Emp ID", "Name", "P", "A", "L", "Attendance %"],
            ["EMP-001", "Arjun Mehta", "=COUNTIF(C3:AG3, \"P\")", "=COUNTIF(C3:AG3, \"A\")", "=COUNTIF(C3:AG3, \"L\")", "=AH3/(AH3+AI3+AJ3)"],
            ["EMP-002", "Priya Nair", "=COUNTIF(C4:AG4, \"P\")", "=COUNTIF(C4:AG4, \"A\")", "=COUNTIF(C4:AG4, \"L\")", "=AH4/(AH4+AI4+AJ4)"]
          ] },
          { type: "p", text: "For October 2026, every employee's totals should add up to 24 working days. That is 31 calendar days minus 4 Sundays, minus 2 week-off Saturdays, minus 1 holiday. Arjun Mehta shows 20 P, 2 A and 2 L, so his P + A + L is exactly 24. This is your built-in check. If any row does not add up to 24, a day was marked wrong or a weekend column was filled by mistake." }
        ]
      },
      {
        heading: "Attendance Percentage Done Right",
        blocks: [
          { type: "formula", name: "Attendance %", syntax: "=P / (P + A + L)", example: "=AH3/(AH3+AI3+AJ3)", explanation: "Divides days present by days the employee was expected to work. Format the cell as a percentage." },
          { type: "p", text: "The denominator is the part everyone gets wrong. Working days are only the days marked P, A or L. Not 31, not 30, not \"the days in the month\". H and WO columns are not working days. Keep them out of the division." },
          { type: "table", headers: ["Employee", "P", "A", "L", "Attendance %"], rows: [
            ["Arjun Mehta", "20", "2", "2", "83.3%"],
            ["Priya Nair", "22", "0", "2", "91.7%"],
            ["Rahul Verma", "24", "0", "0", "100%"],
            ["Sneha Iyer", "21", "1", "2", "87.5%"],
            ["Vikram Singh", "23", "1", "0", "95.8%"],
            ["Divya Menon", "20", "3", "1", "83.3%"]
          ] },
          { type: "warning", text: "Dividing P by 31 punishes employees for weekends and holidays they never had to work. If your percentage looks too low, check that you are not dividing by the full month." }
        ]
      },
      {
        heading: "The Monthly Summary Sheet",
        blocks: [
          { type: "p", text: "Managers do not want to read a 31-column grid. Give them a summary sheet that pulls the totals out of the register:" },
          { type: "grid", title: "Summary sheet, October 2026", rows: [
            ["Emp ID", "Name", "P", "A", "L", "Attendance %"],
            ["EMP-001", "Arjun Mehta", "=COUNTIF('Oct 2026'!$C3:$AG3, \"P\")", "=COUNTIF('Oct 2026'!$C3:$AG3, \"A\")", "=COUNTIF('Oct 2026'!$C3:$AG3, \"L\")", "=C3/(C3+D3+E3)"],
            ["EMP-002", "Priya Nair", "=COUNTIF('Oct 2026'!$C4:$AG4, \"P\")", "=COUNTIF('Oct 2026'!$C4:$AG4, \"A\")", "=COUNTIF('Oct 2026'!$C4:$AG4, \"L\")", "=C4/(C4+D4+E4)"]
          ] },
          { type: "p", text: "The sheet name 'Oct 2026' contains a space, so it sits inside single quotes in the formula. The dollars lock the day columns C to AG but leave the row relative. One formula per column fills down neatly. Keep the summary rows in the same employee order as the register. Or use the lookup skills from Level 4 to pull each row by Emp ID instead." },
          { type: "tip", text: "Keep one register file per month (or one sheet tab per month) and an identical summary on each. At year end, twelve monthly summaries line up perfectly for annual review discussions." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Writing \"Present\", \"Absent\" and \"On leave\" as words in the day cells.", right: "Use the fixed single-letter codes P, A, L, H, WO.", why: "COUNTIF matches exact text. Word variants break your totals, and wide words make the day grid unreadable." },
      { wrong: "Computing attendance as P divided by 31.", right: "Divide P by P + A + L, the days the employee was actually expected to work.", why: "Sundays, week offs and holidays are not working days. Dividing by the full month understates everyone's attendance." },
      { wrong: "Marking a Sunday with A because the person \"did not come\".", right: "Leave weekend columns blank and shaded, and use WO or H only where the rules say so.", why: "An A on a non-working day inflates the A count and makes P + A + L exceed the true working days, breaking your own check." },
      { wrong: "Typing each employee's totals by hand at month end.", right: "Let the COUNTIF totals and the summary sheet calculate automatically.", why: "Manual totals drift out of date the moment any code changes, and nobody notices until payroll is wrong." }
    ],
    shortcuts: [
      { keys: "Ctrl + Shift + L", action: "Filter the register to one employee or one department" },
      { keys: "Ctrl + R", action: "Fill the selected cell's code to the right, cell by cell" },
      { keys: "Ctrl + D", action: "Copy the row's totals formulas down to the next employee" }
    ],
    exercises: [
      { task: "In the attendance-records workbook, write the P, A and L COUNTIF totals for every employee row on the Exercise register.", hint: "The day columns run from C to AG. Build the P total in one row, then copy it across and down.", solution: "Every employee's P + A + L should total 24 working days for October 2026. Match your answers against the Solution sheet." },
      { task: "Add the attendance percentage column with the correct denominator and format it as a percentage.", hint: "Use =AH3/(AH3+AI3+AJ3) in the register, or =C3/(C3+D3+E3) on the summary sheet.", solution: "Arjun Mehta 83.3%, Priya Nair 91.7%, Rahul Verma 100%, Sneha Iyer 87.5%, Vikram Singh 95.8%, Divya Menon 83.3%." },
      { task: "Apply conditional formatting so Sundays and the 2nd and 4th Saturdays are shaded in the day area.", hint: "Two formula rules: =WEEKDAY(DATE(2026,10,C$2),1)=1 and =OR(C$2=10,C$2=24).", solution: "Columns for days 4, 11, 18, 25 (Sundays) and days 10, 24 (Saturdays) turn grey automatically." }
    ],
    quiz: [
      { q: "Why do offices use codes like P and A instead of writing Present and Absent?", options: ["Codes look more professional", "COUNTIF needs exact text, and one letter has no spelling variants", "Excel cannot count words", "Codes save file space"], answer: 1, explanation: "COUNTIF matches exact strings. Words invite variants like Present and present with a space, which count as different text and break totals. Single letters keep entries and counts consistent." },
      { q: "Which formula counts Arjun Mehta's present days in the register (row 3, days in C3:AG3)?", options: ["=SUM(C3:AG3)", "=COUNT(C3:AG3)", "=COUNTIF(C3:AG3, \"P\")", "=COUNTA(C3:AG3)"], answer: 2, explanation: "COUNTIF counts only the cells containing the code P. SUM would try to add letters, COUNT counts numbers, and COUNTA counts every non-empty cell including H and WO." },
      { q: "How should attendance percentage be calculated?", options: ["P divided by 31", "P divided by the number of filled cells", "P divided by (P + A + L)", "(P - A) divided by L"], answer: 2, explanation: "The denominator is the days the employee was expected to work, which is exactly the days marked P, A or L. Holidays and week offs must not dilute the percentage." },
      { q: "For October 2026 (31 days, 4 Sundays, 2 week-off Saturdays, 1 holiday), how many working days does each full-time employee have?", options: ["31", "27", "26", "24"], answer: 3, explanation: "31 - 4 Sundays - 2 Saturdays - 1 holiday = 24. Every employee's P + A + L should equal 24, which is a quick self-check on the register." }
    ],
    summary: [
      "One register per month: employees down the rows, days 1-31 across the columns.",
      "Codes P, A, L, H, WO keep entries countable; words create spelling chaos.",
      "Count each code per row with COUNTIF; totals live in the columns after day 31.",
      "Attendance % = P divided by (P + A + L), never P divided by 31.",
      "Shade weekends with conditional formatting and summarise the month on a separate sheet."
    ],
    downloads: [
      { file: "assets/files/level5/attendance-records.xlsx", label: "Attendance records workbook", note: "Exercise and Solution sheets — a filled month of attendance codes to summarize" }
    ],
    keywords: ["attendance register", "attendance sheet", "P A L H WO codes", "countif attendance", "attendance percentage formula", "monthly attendance summary"]
  },
  {
    id: "l5-leave-management",
    title: "Leave Management System",
    difficulty: "Intermediate",
    duration: 20,
    version: "",
    objectives: [
      "Set up leave types and yearly entitlements for every employee",
      "Track entitlement, taken and balance as three linked numbers",
      "Record requests in a leave log that calculates days automatically",
      "Update balances automatically with SUMIFS",
      "Flag over-limit requests with an IF check"
    ],
    intro: "Leave is where handwritten office systems break first. Someone takes three casual leave days in March, two more in September. By December, nobody knows how many are left. Excel fixes this with one request log. Formulas keep every balance updated automatically.\n\nSunrise Traders gives each employee three types of leave per year. You will build the entitlement table, the request log and the balance sheet. Nothing in it needs manual updating.",
    keyIdea: "Log every request once, and formulas keep every leave balance updated automatically.",
    sections: [
      {
        heading: "Leave Types and Entitlements",
        blocks: [
          { type: "table", headers: ["Type", "Typical yearly entitlement", "What it is for"], rows: [
            ["Casual Leave", "12 days", "Short personal matters, usually taken a day or two at a time"],
            ["Sick Leave", "7 days", "Illness, often needs a note for longer spells"],
            ["Earned Leave", "15 days", "Planned holidays, usually earned by service and taken in blocks"]
          ] },
          { type: "p", text: "Entitlements are a company policy decision. The numbers above are simply Sunrise Traders' policy. Store them in one table, in one place. The balance formula will read them from this table. Never scatter them across emails and sticky notes." },
          { type: "grid", title: "Leave balances sheet — the three numbers that matter", rows: [
            ["Emp ID", "Name", "Casual Allowed", "Sick Allowed", "Earned Allowed", "Casual Taken", "Casual Balance"],
            ["EMP-001", "Arjun Mehta", "12", "7", "15", "=SUMIFS('Leave Log'!$G$2:$G$50, 'Leave Log'!$B$2:$B$50, A2, 'Leave Log'!$D$2:$D$50, \"Casual\")", "=C2-F2"],
            ["EMP-002", "Priya Nair", "12", "7", "15", "=SUMIFS('Leave Log'!$G$2:$G$50, 'Leave Log'!$B$2:$B$50, A3, 'Leave Log'!$D$2:$D$50, \"Casual\")", "=C3-F3"]
          ] },
          { type: "p", text: "Entitlement, taken and balance: three columns, three numbers, one rule — `Balance = Allowed - Taken`. Nothing in this sheet is typed except the entitlements. Everything else is calculated, so the sheet is always current." }
        ]
      },
      {
        heading: "The Leave Request Log",
        blocks: [
          { type: "p", text: "Every request, approved or not, gets one row in the log. This log is the single source of truth; the balance sheet just reads it." },
          { type: "grid", title: "Leave Log sheet", rows: [
            ["Req No", "Emp ID", "Name", "Leave Type", "From", "To", "Days", "Status"],
            ["LV-013", "EMP-001", "Arjun Mehta", "Sick", "21-Sep-2026", "22-Sep-2026", "=F2-E2+1", "Approved"],
            ["LV-014", "EMP-002", "Priya Nair", "Casual", "05-Nov-2026", "07-Nov-2026", "=F3-E3+1", "Approved"],
            ["LV-015", "EMP-004", "Sneha Iyer", "Earned", "02-Dec-2026", "09-Dec-2026", "=F4-E4+1", "Pending"]
          ] },
          { type: "formula", name: "Days (inclusive)", syntax: "=To - From + 1", example: "=F2-E2+1", explanation: "A leave from 5-Nov to 7-Nov covers three days, not two. Both endpoints are included, so add 1. Format both date cells as real dates, or this arithmetic silently fails." },
          { type: "ul", items: [
            "Give every request a running number: LV-001, LV-002, and so on, so you can discuss requests without saying names.",
            "Keep the Status column with a dropdown: Pending, Approved, Rejected.",
            "Only Approved rows should count against balances — SUMIFS can filter on Status too, which you will do in the exercise."
          ] }
        ]
      },
      {
        heading: "Updating Balances with SUMIFS",
        blocks: [
          { type: "formula", name: "SUMIFS", syntax: "=SUMIFS(sum_range, criteria_range1, criteria1, criteria_range2, criteria2)", example: "=SUMIFS('Leave Log'!$G$2:$G$50, 'Leave Log'!$B$2:$B$50, A2, 'Leave Log'!$D$2:$D$50, \"Casual\")", explanation: "Adds the Days column from the Leave Log, but only for rows where the Emp ID matches this row AND the leave type is Casual. Two conditions, one total." },
          { type: "p", text: "This is the heart of the system. Someone adds an approved row to the log, and every affected balance updates itself. There is no month-end ritual of \"updating the tracker\". The tracker updates continuously." },
          { type: "p", text: "Priya Nair's log shows a casual leave of 3 days (LV-014, 5 to 7 November 2026). Suppose an earlier casual leave of 2 days exists. Her Casual Taken is then 5, and her balance is `=12-5`, that is 7 days. The formula never forgets a request; people do." },
          { type: "tip", text: "Extend the SUMIFS with a third condition, 'Leave Log'!$H$2:$H$50, \"Approved\", so rejected or pending requests never reduce a balance. Approve the row in the log, and the balance follows automatically." }
        ]
      },
      {
        heading: "The Balance Check with IF",
        blocks: [
          { type: "formula", name: "Over-limit flag", syntax: "=IF(taken > allowed, \"Over limit\", \"OK\")", example: "=IF(F2>C2, \"Over limit\", \"OK\")", explanation: "Compares days taken with days allowed for that leave type and raises a flag when the balance would go negative." },
          { type: "p", text: "Use this check on the balances sheet as a standing alarm. Use the same idea before approving a request. Priya applies for 3 casual days. Taken-plus-new is 5, well within 12, so the request is fine. But an employee with 2 days left applying for 4 gets flagged before approval, not after." },
          { type: "p", text: "Sneha Iyer's Earned Leave request (LV-015, 2 to 9 December) is 8 days. Against an entitlement of 15 it is fine. But it is still Pending. With the Status condition in place, it does not touch her balance until a manager approves it." },
          { type: "warning", text: "Text dates are the number-one cause of wrong leave balances. If 05-Nov-2026 sits in a cell as text, `F2-E2+1` returns an error or nonsense. Enter both dates as real dates and let Excel's own format display them." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Calculating days as =F2-E2 without the +1.", right: "Use =F2-E2+1 so both the first and last day count.", why: "A leave from Monday to Wednesday is 3 days, not 2. Without +1, every employee quietly gains a free day per request." },
      { wrong: "Editing the balance cell by hand after every request.", right: "Let SUMIFS read the log and keep balances updated automatically.", why: "Hand-edited balances forget requests within a week. The formula-based balance is always in step with the log." },
      { wrong: "Typing dates like 5/11 into cells formatted as text.", right: "Enter real dates and format them as dd-mmm-yyyy.", why: "Text dates cannot be subtracted, so the Days formula breaks and balances go wrong without any visible error." },
      { wrong: "Counting pending or rejected requests against a balance.", right: "Add the Status = Approved condition to the SUMIFS.", why: "Employees should not lose entitlement for leave that was never approved. Status filtering keeps the log honest." }
    ],
    shortcuts: [
      { keys: "Ctrl + ;", action: "Insert today's date into the selected cell" },
      { keys: "Alt + Down", action: "Open the dropdown in a validation cell to pick a leave type" },
      { keys: "Ctrl + Shift + L", action: "Filter the log to one employee or one leave type" }
    ],
    exercises: [
      { task: "In the leave-tracker workbook, complete the Days column in the Leave Log with the inclusive formula and check it against the Solution sheet.", hint: "Days = To - From + 1. LV-014 runs from 05-Nov-2026 to 07-Nov-2026.", solution: "LV-013 is 2 days, LV-014 is 3 days, LV-015 is 8 days." },
      { task: "Build the Casual Taken column with SUMIFS on Emp ID and leave type, then compute the Casual Balance for every employee.", hint: "Sum 'Leave Log' days where the ID matches and the type is Casual; subtract from the allowed column.", solution: "Priya Nair has 5 casual days taken and a balance of 7 out of 12. Match the remaining employees against the Solution sheet." },
      { task: "Add the Status = Approved condition to your SUMIFS, then change LV-015 from Pending to Approved on the Exercise sheet and watch the balances update.", hint: "Add one more criteria pair inside SUMIFS: range H, criteria \"Approved\".", solution: "Sneha Iyer's Earned Taken rises to 8 and her Earned Balance falls to 7 out of 15." }
    ],
    quiz: [
      { q: "A leave runs from 05-Nov-2026 to 07-Nov-2026. Which formula gives the right day count?", options: ["=F2-E2", "=F2-E2+1", "=F2+E2", "=F2-E2-1"], answer: 1, explanation: "Both endpoints are included, so 5, 6 and 7 November are three days: To minus From plus 1. Without the +1 you would under-count every request by one day." },
      { q: "Which function adds up all approved Casual days for one employee from the log?", options: ["SUMIF with one condition", "SUMIFS with Emp ID and leave type conditions", "COUNTIF on the Days column", "VLOOKUP on the log"], answer: 1, explanation: "Two conditions are needed (employee and leave type), so SUMIFS is the right tool. SUMIF handles only a single condition cleanly." },
      { q: "How is a leave balance calculated?", options: ["Balance = Taken - Allowed", "Balance = Allowed + Taken", "Balance = Allowed - Taken", "Balance is typed in by HR"], answer: 2, explanation: "Balance is what you are allowed minus what you have taken. Storing it as a formula over the log means it updates itself with every new request." },
      { q: "What does =IF(F2>C2, \"Over limit\", \"OK\") do on the balances sheet?", options: ["Approves the request automatically", "Flags rows where days taken exceed days allowed", "Counts casual leaves", "Formats the balance cell"], answer: 1, explanation: "It compares taken with allowed for that row and shows a flag whenever the employee has taken more than the entitlement allows." }
    ],
    summary: [
      "Keep entitlements in one policy table: Casual 12, Sick 7, Earned 15 days for Sunrise Traders.",
      "Log every request with a running number, dates, and a Status dropdown.",
      "Days = To - From + 1, because both the first and last day count.",
      "SUMIFS pulls taken days per employee and per type from the log; balance = allowed - taken.",
      "An IF check flags over-limit requests before approval, and only Approved rows touch balances."
    ],
    downloads: [
      { file: "assets/files/level5/leave-tracker.xlsx", label: "Leave tracker workbook", note: "Exercise and Solution sheets — leave entitlements and a request log" }
    ],
    keywords: ["leave management", "leave tracker", "casual sick earned leave", "sumifs leave balance", "leave request log", "leave days formula"]
  },
  {
    id: "l5-payroll",
    title: "Salary and Payroll Calculations",
    difficulty: "Intermediate",
    duration: 30,
    version: "",
    objectives: [
      "Build a salary structure with Basic, HRA, Conveyance and Special allowance",
      "Calculate PF, Professional Tax and a simple TDS estimate",
      "Derive Gross and Net salary with correct rounding",
      "Lay out a printable payslip for one employee"
    ],
    intro: "Payroll is the spreadsheet where small mistakes hurt the most. Every mistake reaches someone's bank account. The good news: salaries follow a fixed structure. Formulas love that. Build the structure once, and every month after is a two-minute job.\n\nIn this lesson you compute October 2026 salaries for the six Sunrise Traders employees. Indian terms like HRA and PF are explained simply as we go.",
    keyIdea: "Build the salary structure with formulas once, and monthly payroll becomes a two-minute job.",
    sections: [
      {
        heading: "The Salary Structure",
        blocks: [
          { type: "table", headers: ["Component", "Rule at Sunrise Traders", "Type"], rows: [
            ["Basic", "The base amount from the employee master", "Earning"],
            ["HRA", "40% of Basic (House Rent Allowance, for living costs)", "Earning"],
            ["Conveyance", "Flat 1,600 for every employee", "Earning"],
            ["Special Allowance", "20% of Basic (the flexible top-up)", "Earning"],
            ["PF", "12% of Basic (Provident Fund, your retirement savings)", "Deduction"],
            ["Professional Tax", "Flat 200 per month (a state tax)", "Deduction"],
            ["TDS", "Income tax estimate: 5% of Gross when Gross exceeds 40,000", "Deduction"]
          ] },
          { type: "p", text: "Percent-based components are deliberate. When Basic rises after an appraisal, HRA, Special Allowance and PF follow automatically. They are formulas, not typed numbers. A payroll where someone re-types every component is a payroll waiting for a mistake." },
          { type: "p", text: "PF deserves one friendly explanation. 12% of Basic is deducted from salary. The employer adds a matching amount on top. The deducted part shows in the payslip. The matching part builds your retirement fund without reducing Net." }
        ]
      },
      {
        heading: "Building the Payroll Sheet",
        blocks: [
          { type: "p", text: "One row per employee. Pull the Basic from the employee master with VLOOKUP so the two sheets can never disagree." },
          { type: "formula", name: "VLOOKUP for Basic", syntax: "=VLOOKUP(lookup_value, table_array, col_index_num, FALSE)", example: "=VLOOKUP(A2, 'Employee Master'!$A$2:$F$7, 6, FALSE)", explanation: "Finds the Emp ID in the master and returns the Basic salary stored in column F. Now a salary revision in the master flows into payroll automatically." },
          { type: "grid", title: "Payroll sheet, October 2026", rows: [
            ["Emp ID", "Name", "Basic", "HRA", "Conv.", "Special", "Gross", "PF", "Prof Tax", "TDS", "Deductions", "Net"],
            ["EMP-001", "Arjun Mehta", "25000", "=ROUND(C2*0.4, 0)", "1600", "=ROUND(C2*0.2, 0)", "=C2+D2+E2+F2", "=ROUND(C2*0.12, 0)", "200", "=IF(G2>40000, ROUND(G2*0.05, 0), 0)", "=H2+I2+J2", "=ROUND(G2-K2, 0)"],
            ["EMP-002", "Priya Nair", "40000", "=ROUND(C3*0.4, 0)", "1600", "=ROUND(C3*0.2, 0)", "=C3+D3+E3+F3", "=ROUND(C3*0.12, 0)", "200", "=IF(G3>40000, ROUND(G3*0.05, 0), 0)", "=H3+I3+J3", "=ROUND(G3-K3, 0)"]
          ] },
          { type: "p", text: "Take Arjun Mehta: Basic 25,000, HRA 10,000 (40%), Conveyance 1,600, Special 5,000 (20%). Gross is `=C2+D2+E2+F2` = 41,600. Deductions are PF 3,000 (12% of Basic), Professional Tax 200 and TDS 2,080. TDS applies because Gross 41,600 crosses the 40,000 line. Net is 41,600 - 5,280 = 36,320." },
          { type: "table", headers: ["Employee", "Basic", "Gross", "Deductions", "Net"], rows: [
            ["Arjun Mehta", "25000", "41600", "5280", "36320"],
            ["Priya Nair", "40000", "65600", "8280", "57320"],
            ["Rahul Verma", "28000", "46400", "5880", "40520"],
            ["Sneha Iyer", "22000", "34800", "2840", "31960"],
            ["Vikram Singh", "24000", "40000", "3080", "36920"],
            ["Divya Menon", "26000", "43200", "5480", "37720"]
          ] },
          { type: "p", text: "Total Gross payroll for October is 2,71,600 and total Net paid out is 2,40,760. These two totals are exactly what the company books as salary expense and salary payment." }
        ]
      },
      {
        heading: "Gross, Net and Why ROUND Matters",
        blocks: [
          { type: "ul", items: [
            "**Gross** is everything earned: `=C2+D2+E2+F2`.",
            "**Deductions** is everything removed: `=H2+I2+J2`.",
            "**Net** is what lands in the bank: `=ROUND(G2-K2, 0)`."
          ] },
          { type: "formula", name: "ROUND", syntax: "=ROUND(number, num_digits)", example: "=ROUND(G2*0.05, 0)", explanation: "Rounds to the given number of digits. With 0 digits, 2080.5 becomes 2081. The cell now truly holds the rounded value, not just a rounded display." },
          { type: "p", text: "Here is the trap ROUND solves. Suppose TDS computes to 2,080.6. Without ROUND, the cell holds 2080.6 but displays 2,081. Every sum built on it carries the hidden 0.6. Your printed payslip total then disagrees with your bank transfer by a rupee or two. Banks do not accept \"roughly correct\". ROUND makes the stored value and the displayed value identical." },
          { type: "tip", text: "Use =ROUND(x, 2) where paise matter, such as invoices, and =ROUND(x, 0) for salaries paid in whole rupees. Do not use the Decrease Decimal button for money — it changes only what is shown, not what is stored." }
        ]
      },
      {
        heading: "A Payslip Layout",
        blocks: [
          { type: "p", text: "The payroll sheet is for the accountant. The payslip is for the employee: one page, one person, with the same numbers pulled from the payroll row by VLOOKUP." },
          { type: "grid", title: "Payslip — Arjun Mehta, October 2026", rows: [
            ["", "Sunrise Traders", "Payslip for October 2026", ""],
            ["", "", "", ""],
            ["Employee", "Arjun Mehta (EMP-001)", "Designation", "Sales Executive"],
            ["", "", "", ""],
            ["Earnings", "", "Deductions", ""],
            ["Basic", "25000", "PF", "3000"],
            ["HRA", "10000", "Professional Tax", "200"],
            ["Conveyance", "1600", "TDS", "2080"],
            ["Special Allowance", "5000", "Total Deductions", "5280"],
            ["Gross", "41600", "", ""],
            ["", "", "NET PAY", "36320"]
          ] },
          { type: "steps", items: [
            "Set the print area over the payslip block (Page Layout, Print Area, Set Print Area).",
            "Set scaling to Fit Sheet on One Page and portrait orientation.",
            "Fill the employee name once in a helper cell and let every label and VLOOKUP follow it.",
            "Save a PDF copy per employee each month so old payslips never change."
          ] }
        ]
      }
    ],
    mistakes: [
      { wrong: "Typing HRA as 10000 next to Basic 25000 instead of a formula.", right: "Use =ROUND(C2*0.4, 0) so HRA follows Basic automatically.", why: "When Basic changes after an appraisal, typed components stay frozen at old values and the structure no longer adds up." },
      { wrong: "Formatting a cell to show 0 decimals while the underlying value has paise.", right: "Wrap money formulas in ROUND so the stored value is rounded, not just displayed.", why: "Sums of half-paise values disagree with printed payslips and with bank transfers by a rupee or two." },
      { wrong: "Building payroll from names instead of Emp IDs.", right: "Look up employees by Emp ID with VLOOKUP.", why: "Two people can share a name; two employees cannot share an ID." },
      { wrong: "Deducting PF from Gross instead of from Basic.", right: "Compute PF as 12% of Basic: =ROUND(C2*0.12, 0).", why: "The rule is 12% of Basic. Percentage-of-Gross inflates PF for higher earners and under-deducts it for others." }
    ],
    shortcuts: [
      { keys: "Ctrl + 1", action: "Open Format Cells to apply number and currency formats" },
      { keys: "Ctrl + D", action: "Copy the row's formulas down to the next employee" },
      { keys: "Ctrl + P", action: "Print or export the payslip to PDF" }
    ],
    exercises: [
      { task: "In the payroll-data workbook, complete the HRA, Conveyance and Special Allowance columns for all six employees on the Exercise sheet.", hint: "HRA is 40% of Basic and Special is 20% of Basic, both wrapped in ROUND(..., 0). Conveyance is a flat 1600.", solution: "Arjun Mehta: HRA 10000, Special 5000, Gross 41600. Priya Nair: HRA 16000, Special 8000, Gross 65600. Match the rest against the Solution sheet." },
      { task: "Add the PF, Professional Tax and TDS columns, then compute Deductions and Net.", hint: "PF = ROUND(Basic*0.12, 0); TDS = IF(Gross>40000, ROUND(Gross*0.05, 0), 0); Net = ROUND(Gross-Deductions, 0).", solution: "Total Gross 2,71,600, total Deductions 30,840, total Net 2,40,760." },
      { task: "Build the payslip block for Divya Menon using VLOOKUP on her Emp ID EMP-006.", hint: "Pull Name and Designation from the master, and every money figure from the payroll row.", solution: "Divya's payslip shows Basic 26000, Gross 43200, Net 37720." }
    ],
    quiz: [
      { q: "An employee's Basic is 25,000. What is the HRA at Sunrise Traders?", options: ["2,500", "4,000", "10,000", "12,000"], answer: 2, explanation: "HRA is 40% of Basic: 25000 x 0.4 = 10,000. Because HRA is a formula on Basic, any change to Basic flows through automatically." },
      { q: "How is PF computed at Sunrise Traders?", options: ["12% of Gross salary", "12% of Basic", "200 per month flat", "5% of Net"], answer: 1, explanation: "PF is 12% of Basic, deducted from the employee's pay. The employer adds a matching contribution on top, which builds the retirement fund." },
      { q: "Which set of components adds up to Gross salary?", options: ["Basic + PF", "Basic + HRA + Conveyance + Special Allowance", "Net + Deductions", "Basic + TDS"], answer: 1, explanation: "Gross is the total of all earnings before deductions. Net = Gross - Deductions, so Gross also equals Net plus Deductions." },
      { q: "Why wrap money formulas in ROUND instead of just reducing displayed decimals?", options: ["ROUND makes the file smaller", "Displayed decimals only change what you see, not what the cell stores", "ROUND is faster than formatting", "Excel requires it for currency"], answer: 1, explanation: "The Decrease Decimal button is cosmetic. ROUND changes the stored value, so your printed total, your sum and your bank transfer all agree exactly." },
      { q: "Gross is 43,200 and crosses the 40,000 line. What is the TDS at the 5% estimate rate?", options: ["0", "216", "2,160", "4,320"], answer: 2, explanation: "5% of 43,200 is 2,160. The IF formula =IF(G2>40000, ROUND(G2*0.05, 0), 0) produces exactly this and gives 0 for employees below the threshold." }
    ],
    summary: [
      "Structure first, numbers second: HRA 40% and Special 20% are formulas on Basic, Conveyance is flat.",
      "Deductions are PF 12% of Basic, Professional Tax 200, and a simple 5% TDS above a 40,000 Gross line.",
      "Gross = all earnings; Net = Gross - Deductions; both computed, never typed.",
      "ROUND makes stored values match printed values — no hidden paise.",
      "Pull Basic from the employee master by Emp ID and print payslips as one-page blocks."
    ],
    downloads: [
      { file: "assets/files/level5/payroll-data.xlsx", label: "Payroll data workbook", note: "Exercise and Solution sheets — employee salary data for payroll calculations" }
    ],
    keywords: ["payroll excel", "salary structure", "HRA calculation", "PF deduction 12 percent", "TDS estimate", "payslip format", "gross net salary"]
  },
  {
    id: "l5-inventory",
    title: "Inventory Management",
    difficulty: "Intermediate",
    duration: 25,
    version: "",
    objectives: [
      "Split stock tracking into an item master, an inward register and an outward register",
      "Log every movement with item codes, quantities and values",
      "Compute current stock with Opening + Inward - Outward using SUMIFS",
      "Value the stock with quantity times unit cost",
      "Add reorder levels to prepare for automatic alerts"
    ],
    intro: "Ask a stores person \"how many staplers do we have?\" A paper notebook gives last week's answer. An Excel inventory gives right now's answer. Stock is calculated, not remembered.\n\nSunrise Traders stocks five products: Notebooks, Pens, Staplers, Printers and Laptops. The system has three sheets. An item master, an inward register for goods coming in, and an outward register for goods going out. All three live in the inventory-records workbook.",
    keyIdea: "Current stock is never typed anywhere — it is always calculated from three linked sheets.",
    sections: [
      {
        heading: "Three Sheets That Run Stock",
        blocks: [
          { type: "table", headers: ["Sheet", "Job", "Who writes it"], rows: [
            ["Item Master", "The catalogue: codes, names, units, opening stock, cost, reorder level", "Manager, occasionally"],
            ["Inward", "Every purchase or return that adds stock", "Stores, on every arrival"],
            ["Outward", "Every sale or issue that removes stock", "Stores, on every dispatch"]
          ] },
          { type: "p", text: "This division of labour is the design decision. Current stock is never stored anywhere — it is calculated from the three sheets. So stock can never go stale. The formula re-counts every time you add a movement." },
          { type: "tip", text: "Give every item a code (ITM-101, ITM-102) and log movements by code, not by name. Names get misspelled; codes do not. This is the same principle as Emp IDs in the employee database." }
        ]
      },
      {
        heading: "The Item Master",
        blocks: [
          { type: "grid", title: "Item Master sheet", rows: [
            ["Item Code", "Item Name", "Category", "Unit", "Opening Stock", "Unit Cost", "Reorder Level", "Current Stock", "Stock Value"],
            ["ITM-101", "Notebook", "Stationery", "Nos", "120", "45", "200", "195", "=H2*F2"],
            ["ITM-102", "Pen", "Stationery", "Nos", "300", "15", "150", "360", "=H3*F3"],
            ["ITM-103", "Stapler", "Office Tools", "Nos", "40", "130", "20", "53", "=H4*F4"],
            ["ITM-104", "Printer", "Electronics", "Nos", "8", "8500", "8", "8", "=H5*F5"],
            ["ITM-105", "Laptop", "Electronics", "Nos", "5", "42000", "2", "4", "=H6*F6"]
          ] },
          { type: "ul", items: [
            "**Opening Stock** is the count on the day you started the system — counted once, entered once.",
            "**Unit Cost** is what you pay to buy one unit, taken from the latest purchase.",
            "**Reorder Level** is the stock point at which you must order again. You will use it heavily in the next lesson."
          ] }
        ]
      },
      {
        heading: "The Inward and Outward Registers",
        blocks: [
          { type: "grid", title: "Inward register (goods received)", rows: [
            ["Date", "Item Code", "Qty", "Rate", "Value"],
            ["03-Oct-2026", "ITM-101", "100", "42", "=C2*D2"],
            ["05-Oct-2026", "ITM-102", "200", "13", "=C3*D3"],
            ["08-Oct-2026", "ITM-103", "25", "125", "=C4*D4"],
            ["12-Oct-2026", "ITM-101", "50", "44", "=C5*D5"],
            ["15-Oct-2026", "ITM-104", "3", "8200", "=C6*D6"]
          ] },
          { type: "grid", title: "Outward register (goods issued or sold)", rows: [
            ["Date", "Item Code", "Qty", "Rate", "Value"],
            ["02-Oct-2026", "ITM-101", "45", "=VLOOKUP(B2, 'Item Master'!$A$2:$F$6, 6, FALSE)", "=C2*D2"],
            ["04-Oct-2026", "ITM-102", "80", "=VLOOKUP(B3, 'Item Master'!$A$2:$F$6, 6, FALSE)", "=C3*D3"],
            ["06-Oct-2026", "ITM-104", "2", "=VLOOKUP(B4, 'Item Master'!$A$2:$F$6, 6, FALSE)", "=C4*D4"],
            ["18-Oct-2026", "ITM-105", "1", "=VLOOKUP(B8, 'Item Master'!$A$2:$F$6, 6, FALSE)", "=C8*D8"]
          ] },
          { type: "p", text: "Both registers share one shape. That keeps entry fast and formulas simple. In the outward register, VLOOKUP pulls the Rate from the item master. So goods going out are always valued at cost. That value column becomes the COGS figure in the financial reports lesson." },
          { type: "warning", text: "Log movements on the day they happen. A register filled from memory at month end is an estimate, and every downstream number — stock, reorder alerts, profit — inherits the error." }
        ]
      },
      {
        heading: "Current Stock with SUMIFS",
        blocks: [
          { type: "formula", name: "Current stock", syntax: "=Opening + SUMIFS(inward qty by item) - SUMIFS(outward qty by item)", example: "=E2+SUMIFS(Inward!$C$2:$C$100, Inward!$B$2:$B$100, A2)-SUMIFS(Outward!$C$2:$C$100, Outward!$B$2:$B$100, A2)", explanation: "Starts from Opening Stock, adds every inward quantity for this item code, subtracts every outward quantity. The result is live current stock." },
          { type: "p", text: "Read the formula in three pieces. `SUMIFS(Inward!$C$2:$C$100, Inward!$B$2:$B$100, A2)` adds inward quantities only where the item code equals this row's code. The outward SUMIFS does the same for goods going out. Everything else is the opening count." },
          { type: "table", headers: ["Item", "Opening", "Inward", "Outward", "Current"], rows: [
            ["ITM-101 Notebook", "120", "150", "75", "195"],
            ["ITM-102 Pen", "300", "200", "140", "360"],
            ["ITM-103 Stapler", "40", "25", "12", "53"],
            ["ITM-104 Printer", "8", "3", "3", "8"],
            ["ITM-105 Laptop", "5", "0", "1", "4"]
          ] },
          { type: "warning", text: "Every range pair inside SUMIFS must be the same size. If the criteria range covers rows 2-100 but the sum range covers 2-50, Excel returns an error or, worse, a wrong total. Fix both ends at 100 with dollar signs and keep them matched." }
        ]
      },
      {
        heading: "Stock Value and the Reorder Level",
        blocks: [
          { type: "formula", name: "Stock value", syntax: "=Current stock * Unit cost", example: "=H2*F2", explanation: "Multiplies how many units you hold by what each unit cost. This is money sitting on your shelves." },
          { type: "p", text: "Sunrise Traders' October stock values: Notebooks 8,775, Pens 5,400, Staplers 6,890, Printers 68,000, Laptops 1,68,000. Total stock value is 2,56,065. Note how five Laptops dominate the number even though Notebooks dominate the shelf count — value and quantity tell different stories." },
          { type: "p", text: "The Reorder Level column is now just data. In the next lesson it becomes the trigger for automatic ORDER NOW alerts and ABC-style analysis of what sells and what sits." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Typing the current stock column by hand after every movement.", right: "Calculate stock with Opening + SUMIFS(inward) - SUMIFS(outward).", why: "Typed counts are snapshots that go stale the moment the next box arrives. Formulas re-count automatically." },
      { wrong: "Logging movements by item name (\"Notebook\").", right: "Log by item code (ITM-101) and look names up from the master.", why: "SUMIFS matches exact text. Name typos split one item's history into two and corrupt the stock count." },
      { wrong: "Using a criteria range of rows 2-100 with a sum range of rows 2-50 in SUMIFS.", right: "Keep every range pair identical in size, with dollar-signed ends.", why: "Mismatched ranges cause errors or silently wrong totals — the most common SUMIFS bug." },
      { wrong: "Valuing stock at selling price.", right: "Value stock at unit cost with =Current * Unit Cost.", why: "Stock value means money invested in goods you have not sold yet. Selling price would inflate it with profit you have not earned." }
    ],
    shortcuts: [
      { keys: "Ctrl + D", action: "Copy the stock formula down the item master" },
      { keys: "Ctrl + Shift + L", action: "Filter a register to one item code" },
      { keys: "F4", action: "Toggle dollar signs on a selected range while writing the formula" }
    ],
    exercises: [
      { task: "In the inventory-records workbook, complete the Current Stock column in the Item Master with the Opening + SUMIFS - SUMIFS formula.", hint: "Inward quantities live in Inward column C, keyed by item code in column B. Mirror the same for Outward.", solution: "Notebook 195, Pen 360, Stapler 53, Printer 8, Laptop 4. Compare with the Solution sheet." },
      { task: "Fill the Rate column of the Outward register with VLOOKUP from the Item Master, then compute the Value column.", hint: "VLOOKUP(B2, 'Item Master'!$A$2:$F$6, 6, FALSE) returns the unit cost of the item code in B2.", solution: "Outward value for October totals 74,535 across all movements." },
      { task: "Complete the Stock Value column and find the total value of goods on hand.", hint: "Value = Current Stock * Unit Cost; sum the column.", solution: "Total stock value is 2,56,065, with Laptops holding 1,68,000 of it." }
    ],
    quiz: [
      { q: "How is current stock calculated in this system?", options: ["Typed in after a physical count", "Opening + Inward - Outward", "Inward - Outward", "Opening - Outward"], answer: 1, explanation: "Start from opening stock, add everything that came in, subtract everything that went out. As a formula it recounts itself with every new movement." },
      { q: "Which function totals the inward quantity for one specific item code?", options: ["COUNTIF", "SUMIFS with the item code as criteria", "VLOOKUP", "SUMIF on the whole column"], answer: 1, explanation: "SUMIFS adds quantities only where the item code matches, pulling one item's history out of a shared register. COUNTIF counts cells, not quantities." },
      { q: "The Item Master shows Current Stock 195 and Unit Cost 45 for Notebooks. What is the stock value?", options: ["8,775", "2,56,065", "8,775 pairs", "195 rupees"], answer: 0, explanation: "Stock value = current stock x unit cost = 195 x 45 = 8,775. It measures money invested in unsold goods, at cost." },
      { q: "Why do registers record item codes instead of item names?", options: ["Codes are shorter to type", "Codes are unique and exact, so SUMIFS matches reliably", "Excel cannot search names", "Codes sort faster"], answer: 1, explanation: "SUMIFS needs exact text matches. Names invite typos and variants that split one item's movements into two histories and corrupt every total." }
    ],
    summary: [
      "Split inventory into three sheets: item master, inward register, outward register.",
      "Current stock is never stored — it is Opening + SUMIFS(inward) - SUMIFS(outward).",
      "Log every movement the day it happens, keyed by item code.",
      "Stock value = current stock x unit cost, always at cost not selling price.",
      "Store reorder levels in the master now; the next lesson turns them into alerts."
    ],
    downloads: [
      { file: "assets/files/level5/inventory-records.xlsx", label: "Inventory records workbook", note: "Exercise and Solution sheets — item master, inward and outward registers" }
    ],
    keywords: ["inventory management", "stock register", "inward outward register", "current stock formula", "sumifs inventory", "stock value", "reorder level"]
  },
  {
    id: "l5-stock-tracking",
    title: "Stock Tracking and Reorder Alerts",
    difficulty: "Intermediate",
    duration: 20,
    version: "",
    objectives: [
      "Use reorder levels as trigger points for fresh orders",
      "Create a status column that flags ORDER NOW automatically",
      "Highlight below-reorder items in red with conditional formatting",
      "Identify fast and slow movers by monthly outward quantity",
      "Rank best sellers by value with RANK, LARGE and INDEX-MATCH"
    ],
    intro: "Running out of Notebooks in school season costs Sunrise Traders real sales. Ordering too early blocks money on shelves. The middle path is the reorder level. It is a stock point where you must order again. A formula watches it and never forgets.\n\nIn this lesson you turn your inventory into a tracking system. You add automatic alerts, movement analysis, and a ranked list of what actually earns money.",
    keyIdea: "A reorder level plus an IF formula shouts ORDER NOW before stock runs out.",
    sections: [
      {
        heading: "Reorder Level: The Trigger Point",
        blocks: [
          { type: "p", text: "A reorder level answers one question. At what remaining stock must the next order go out, so that new goods arrive before you run dry? It depends on how fast the item sells and how long the supplier takes to deliver." },
          { type: "p", text: "Sunrise Traders' levels: Notebooks 200, Pens 150, Staplers 20, Printers 8, Laptops 2." },
          { type: "p", text: "Set the level once per item in the Item Master. Review it a few times a year, or when a supplier changes delivery times. The daily comparison — current stock against the level — is the formula's job, not yours." }
        ]
      },
      {
        heading: "A Status Column That Decides For You",
        blocks: [
          { type: "formula", name: "Reorder status", syntax: "=IF(stock <= reorder, \"ORDER NOW\", \"OK\")", example: "=IF(H2<=G2, \"ORDER NOW\", \"OK\")", explanation: "Compares current stock (H) with the reorder level (G) in the Item Master and returns a clear instruction instead of a bare number." },
          { type: "grid", title: "Status column on the Item Master", rows: [
            ["Item Code", "Item Name", "Reorder Level", "Current Stock", "Status"],
            ["ITM-101", "Notebook", "200", "195", "=IF(H2<=G2, \"ORDER NOW\", \"OK\")"],
            ["ITM-102", "Pen", "150", "360", "=IF(H3<=G3, \"ORDER NOW\", \"OK\")"],
            ["ITM-104", "Printer", "8", "8", "=IF(H5<=G5, \"ORDER NOW\", \"OK\")"],
            ["ITM-105", "Laptop", "2", "4", "=IF(H6<=G6, \"ORDER NOW\", \"OK\")"]
          ] },
          { type: "p", text: "Notebooks at 195 against a level of 200 say ORDER NOW. Printers sitting exactly at 8 also say ORDER NOW. That is why the comparison is `<=` and not `<`. At the reorder level you are already at the trigger, so the order must go out. Laptops at 4 against a level of 2 are still OK." },
          { type: "warning", text: "Keep one standard alert phrase. If one row says ORDER NOW and another says ORDER NEXT, conditional formatting and filters catch only one of them. They match exact text. Copy the formula down; never retype it row by row." }
        ]
      },
      {
        heading: "Red Alerts with Conditional Formatting",
        blocks: [
          { type: "steps", items: [
            "Select the Status column on the Item Master, for example `I2:I6`.",
            "Go to Home, Conditional Formatting, Highlight Cells Rules, Equal To.",
            "Type ORDER NOW in the box and choose Light Red Fill with dark red text.",
            "Click OK. Every item at or below its reorder level now glows red automatically.",
            "Optionally add a whole-row rule with Use a formula: `=$I2=\"ORDER NOW\"` applied to `A2:I6`."
          ] },
          { type: "p", text: "The alert is live. The moment an outward entry pushes Notebooks below 200, the row turns red. No human checks anything. This is the real payoff of calculating stock instead of typing it." },
          { type: "tip", text: "Print or screenshot the red list every Monday morning and compare it with the purchase register from the next lesson. Items that stay red for two weeks need a chasing call to the supplier, not another alert." }
        ]
      },
      {
        heading: "Fast and Slow Movers",
        blocks: [
          { type: "p", text: "Reorder levels tell you when to buy. Movement analysis tells you what to keep buying. A fast mover deserves shelf space and prompt reorders. A slow mover deserves a smaller reorder level and fewer orders." },
          { type: "formula", name: "Monthly outward by item", syntax: "=SUMIFS(outward qty, outward item codes, item)", example: "=SUMIFS(Outward!$C$2:$C$100, Outward!$B$2:$B$100, A2)", explanation: "Adds the October outward quantity for one item code. Sort or read this column to spot fast and slow movers instantly." },
          { type: "table", headers: ["Item", "Oct outward qty", "Oct outward value", "Verdict"], rows: [
            ["Pen", "140", "2100", "Fastest by quantity"],
            ["Notebook", "75", "3375", "Fast by quantity"],
            ["Stapler", "12", "1560", "Slow"],
            ["Printer", "3", "25500", "Slow by quantity, big by value"],
            ["Laptop", "1", "42000", "Slowest by quantity, biggest by value"]
          ] },
          { type: "p", text: "Quantity and value disagree, and both matter. Pens move 140 units but earn only 2,100 because each unit costs 15. Laptops moved a single unit yet carry 42,000 of value. A store manager watching only quantities would treat Laptops as dead stock; watching value, it is the best seller." }
        ]
      },
      {
        heading: "ABC Thinking: Rank Your Best Sellers",
        blocks: [
          { type: "p", text: "Borrow a retail habit called ABC analysis. A few items earn most of the money (A items). A middle group earns some (B). A long tail earns little (C). Rank the value column to see which group each item falls in." },
          { type: "formula", name: "RANK", syntax: "=RANK(number, ref)", example: "=RANK(D2, $D$2:$D$6)", explanation: "Gives the position of this row's outward value within all five items: 1 for the biggest earner. The dollar signs fix the comparison range while the row moves." },
          { type: "formula", name: "Top seller by name", syntax: "=INDEX(names, MATCH(LARGE(values, 1), values, 0))", example: "=INDEX($B$2:$B$6, MATCH(LARGE($D$2:$D$6, 1), $D$2:$D$6, 0))", explanation: "LARGE finds the highest outward value, MATCH finds its position, INDEX returns the item name at that position. Change the 1 to a 2 to get the runner-up." },
          { type: "p", text: "For October the ranking by value is: Laptop (1), Printer (2), Notebook (3), Pen (4), Stapler (5). The top-seller formula returns Laptop." },
          { type: "p", text: "In practice you would run this on a real catalogue of hundreds of items. There, the top 20 percent of codes routinely produce 70 to 80 percent of the value. That is the A group." },
          { type: "p", text: "The action list writes itself. Keep A items always in stock, with generous reorder levels. Review B items monthly. Shrink the reorder levels of C items, before they turn into cash sitting on a shelf." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Writing =IF(H2<G2, ...) so stock exactly at the reorder level shows OK.", right: "Use <= so the level itself triggers the order.", why: "The reorder level is the trigger point, not a safety floor. Stock sitting at the level has already reached the moment to order." },
      { wrong: "Ranking items by quantity moved.", right: "Rank by outward value for purchasing decisions, and read quantity separately for shelf management.", why: "One Laptop outweighs 140 Pens in money terms. Quantity ranking hides where the cash actually moves." },
      { wrong: "Using =RANK(D2, D2:D6) without dollar signs.", right: "Fix the comparison range: =RANK(D2, $D$2:$D$6).", why: "As the formula fills down, the moving range compares each row against a different subset and produces meaningless ranks." },
      { wrong: "Changing the alert text to suit the row, like ORDER NEXT or ORDER IMMEDIATELY.", right: "Use one standard phrase everywhere and filter or format for it.", why: "Conditional formatting and filters match exact text. Mixed phrases mean some alerts stay unhighlighted and unnoticed." }
    ],
    shortcuts: [
      { keys: "Alt + H, L", action: "Open Conditional Formatting from the Home ribbon" },
      { keys: "Ctrl + Shift + L", action: "Filter the Item Master to Status = ORDER NOW" },
      { keys: "Ctrl + Shift + Down", action: "Select the full column of ranks before sorting" }
    ],
    exercises: [
      { task: "In the inventory-records workbook, add the Status column to the Item Master with the IF formula and find which items need ordering.", hint: "=IF(H2<=G2, \"ORDER NOW\", \"OK\") compares current stock with the reorder level.", solution: "Notebooks (195 vs 200) and Printers (8 vs 8) show ORDER NOW; the other three show OK." },
      { task: "Apply conditional formatting so every ORDER NOW status turns light red, then add an outward row for 1 Notebook and watch the stock and status update.", hint: "Highlight Cells Rules, Equal To, ORDER NOW.", solution: "Notebooks drop to 194 and stay red; the system re-evaluates with no manual work." },
      { task: "Build a movement summary: October outward quantity and value per item, the RANK by value, and the name of the top seller.", hint: "Two SUMIFS on the Outward register, =RANK(D2, $D$2:$D$6), then INDEX-MATCH-LARGE for the name.", solution: "By value: Laptop 42000 (rank 1), Printer 25500, Notebook 3375, Pen 2100, Stapler 1560. Top seller by value: Laptop." }
    ],
    quiz: [
      { q: "Stock is exactly equal to the reorder level. Should you order?", options: ["No, order only when stock falls below the level", "Yes, the level is the trigger, which is why the formula uses <=", "Only if the item is an A item", "Only after the physical count"], answer: 1, explanation: "The reorder level marks the point where the next order must go out so goods arrive before stock runs dry. That is why the status formula tests H2<=G2." },
      { q: "Which feature makes below-reorder rows turn red without any manual checking?", options: ["Cell styles applied by hand", "Conditional formatting driven by the status text", "Data validation", "Freeze panes"], answer: 1, explanation: "Conditional formatting watches the cells continuously. When the status shows ORDER NOW, the fill changes immediately — and reverses itself when stock is replenished." },
      { q: "In October, Pens moved 140 units and Laptops moved 1 unit. Which item earned more outward value?", options: ["Pens, because quantity wins", "Laptops, because 1 unit at 42,000 beats 140 Pens at 15", "Both are equal", "Cannot be determined"], answer: 1, explanation: "Value = quantity x unit cost. Laptops carry 42,000 against Pens' 2,100. This is why ABC analysis ranks by value, not by count." },
      { q: "What does =INDEX($B$2:$B$6, MATCH(LARGE($D$2:$D$6, 1), $D$2:$D$6, 0)) return?", options: ["The largest value itself", "The name of the item with the largest value", "The rank of the largest item", "The average value"], answer: 1, explanation: "LARGE finds the highest value, MATCH locates its position in the value column, and INDEX returns the item name from the same position. Swap 1 for 2 to get the runner-up." },
      { q: "What is the purpose of ABC-style thinking?", options: ["Sorting items alphabetically", "Grouping items by inventory value so the best earners get the most attention", "Classifying items by supplier", "Colour-coding categories"], answer: 1, explanation: "A few items usually generate most of the value. Focusing reorder attention on that A group protects revenue, while C items get leaner reorder levels." }
    ],
    summary: [
      "Reorder level is the trigger point: compare current stock against it with <=, not <.",
      "A status column =IF(H2<=G2, \"ORDER NOW\", \"OK\") converts numbers into instructions.",
      "Conditional formatting turns the alert red and stays live as stock changes.",
      "Movement analysis by quantity finds shelf movers; by value finds cash movers.",
      "RANK and INDEX-MATCH-LARGE rank best sellers for ABC thinking: feed the A items first."
    ],
    downloads: [
      { file: "assets/files/level5/inventory-records.xlsx", label: "Inventory records workbook", note: "Exercise and Solution sheets — item master, inward and outward registers" }
    ],
    keywords: ["reorder level", "stock alert", "order now formula", "conditional formatting stock", "fast slow movers", "ABC analysis", "rank best sellers"]
  },
  {
    id: "l5-purchase-management",
    title: "Purchase Management",
    difficulty: "Intermediate",
    duration: 20,
    version: "",
    objectives: [
      "Maintain a purchase order register with one row per PO",
      "Compute amounts with formulas and control Status with dropdowns",
      "Count pending POs at a glance with COUNTIF",
      "Total purchases supplier-wise with SUMIFS",
      "Track payment status from PO to payment"
    ],
    intro: "Every rupee Sunrise Traders spends starts as a purchase order. When the PO register is disciplined, three questions answer themselves instantly: what is ordered, what has arrived, and what is still unpaid.\n\nIn this lesson you build the October PO register, drive the numbers with formulas, and produce supplier-wise summaries that feed both the inventory system and the financial reports lesson.",
    keyIdea: "A disciplined PO register instantly answers what is ordered, what arrived, and what is unpaid.",
    sections: [
      {
        heading: "The Purchase Order Register",
        blocks: [
          { type: "table", headers: ["Column", "Content"], rows: [
            ["A - PO No", "Unique running number, e.g. PO-2026-014"],
            ["B - Date", "Date the PO was raised"],
            ["C - Supplier", "Who the order went to"],
            ["D - Item", "What was ordered, by item code or name"],
            ["E - Qty", "Quantity ordered"],
            ["F - Rate", "Agreed price per unit"],
            ["G - Amount", "Qty x Rate, always a formula"],
            ["H - Status", "Open, Received or Cancelled"],
            ["I - Payment", "Paid, Unpaid, Partial or NA"],
            ["J - Remarks", "Delivery dates, invoice numbers, notes"]
          ] },
          { type: "grid", title: "PO register, October 2026", rows: [
            ["PO No", "Date", "Supplier", "Item", "Qty", "Rate", "Amount", "Status", "Payment"],
            ["PO-2026-014", "02-Oct-2026", "Sharma Stationery", "ITM-101 Notebook", "100", "42", "=E2*F2", "Received", "Paid"],
            ["PO-2026-015", "05-Oct-2026", "Global Office Supplies", "ITM-104 Printer", "3", "8200", "=E3*F3", "Open", "Unpaid"],
            ["PO-2026-016", "08-Oct-2026", "Kavya Paper Mart", "ITM-102 Pen", "200", "13", "=E4*F4", "Received", "Paid"],
            ["PO-2026-017", "11-Oct-2026", "TechDesk Systems", "ITM-105 Laptop", "1", "41000", "=E5*F5", "Open", "Unpaid"],
            ["PO-2026-018", "14-Oct-2026", "Sharma Stationery", "ITM-103 Stapler", "25", "125", "=E6*F6", "Cancelled", "NA"],
            ["PO-2026-019", "16-Oct-2026", "Global Office Supplies", "ITM-104 Printer", "5", "8100", "=E7*F7", "Open", "Partial"]
          ] },
          { type: "p", text: "The PO number pattern `PO-YYYY-NNN` carries the year plus a running serial. Like Emp IDs and item codes, it is unique forever: even a cancelled PO keeps its number, because it exists in the story of the business." }
        ]
      },
      {
        heading: "Amount and Status",
        blocks: [
          { type: "formula", name: "PO amount", syntax: "=Qty * Rate", example: "=E2*F2", explanation: "Multiplies the ordered quantity by the agreed rate. Never type the amount: a typed 4200 hides a typo in either Qty or Rate forever." },
          { type: "steps", items: [
            "Select the Status cells, for example `H2:H100`.",
            "Open Data, Data Validation, Allow: List.",
            "Enter the source: `Open,Received,Cancelled` and click OK.",
            "Repeat for Payment with the list `Paid,Unpaid,Partial,NA`."
          ] },
          { type: "p", text: "Status is the column the whole office reads. Open means the order is placed but not yet delivered. Received means the goods are in the stores — and the moment you mark Received, the same quantity should appear in the inventory's inward register. Cancelled orders stay in the register with their number, but they must never be counted as purchases." },
          { type: "warning", text: "A blank Status cell is the classic PO register leak. Nobody knows if the order is pending or delivered, so it is never chased. Add a weekly habit: filter Status to Open and check each row against the supplier." }
        ]
      },
      {
        heading: "Pending POs at a Glance",
        blocks: [
          { type: "formula", name: "Pending count", syntax: "=COUNTIF(status_range, \"Open\")", example: "=COUNTIF($H$2:$H$100, \"Open\")", explanation: "Counts purchase orders still awaiting delivery. Build it once on a summary block and refresh it mentally every morning." },
          { type: "grid", title: "A small status summary", rows: [
            ["Status", "Count", "Value"],
            ["Open", "=COUNTIF($H$2:$H$100, \"Open\")", "=SUMIFS($G$2:$G$100, $H$2:$H$100, \"Open\")"],
            ["Received", "=COUNTIF($H$2:$H$100, \"Received\")", "=SUMIFS($G$2:$G$100, $H$2:$H$100, \"Received\")"],
            ["Cancelled", "=COUNTIF($H$2:$H$100, \"Cancelled\")", "=SUMIFS($G$2:$G$100, $H$2:$H$100, \"Cancelled\")"]
          ] },
          { type: "p", text: "For October: 3 POs are Open worth 1,06,100 (24,600 + 41,000 + 40,500), 2 are Received worth 6,800, and 1 is Cancelled worth 3,125. Two numbers tell the stores team exactly what to chase." }
        ]
      },
      {
        heading: "Supplier-wise Totals with SUMIFS",
        blocks: [
          { type: "formula", name: "Supplier total", syntax: "=SUMIFS(amounts, suppliers, supplier, statuses, \"Received\")", example: "=SUMIFS($G$2:$G$100, $C$2:$C$100, \"Sharma Stationery\", $H$2:$H$100, \"Received\")", explanation: "Adds only the PO amounts for one supplier whose goods were actually received. The second criteria pair keeps cancelled orders out of the money." },
          { type: "table", headers: ["Supplier", "All POs", "Received only"], rows: [
            ["Sharma Stationery", "7,325", "4,200"],
            ["Global Office Supplies", "65,100", "0"],
            ["Kavya Paper Mart", "2,600", "2,600"],
            ["TechDesk Systems", "41,000", "0"]
          ] },
          { type: "p", text: "Compare the two columns. Sharma Stationery shows 7,325 across all POs, but only 4,200 became real purchases — the 3,125 stapler PO was cancelled. Supplier talks go much better when you quote received-only figures, because that is the business you actually gave them." }
        ]
      },
      {
        heading: "Payment Status Tracking",
        blocks: [
          { type: "p", text: "A PO's life has two halves: the goods (Open, Received) and the money (Unpaid, Partial, Paid). The Payment column finishes the story." },
          { type: "formula", name: "Still to pay", syntax: "=SUMIFS(amounts, payment_status, \"Unpaid\")", example: "=SUMIFS($G$2:$G$100, $I$2:$I$100, \"Unpaid\")", explanation: "Adds the PO amounts with no payment made yet. Add a second criteria pair on Status = Received to exclude undelivered orders from the payment urgency list." },
          { type: "p", text: "For October, the fully Unpaid amount is 65,600 (the two printer POs and the laptop PO), and the Partial amount is 40,500. Combine both and you have the accounts team's payment run: who to pay, in what order, and against which PO numbers." },
          { type: "tip", text: "Write the supplier's invoice number in the Remarks column when payment is made. When a supplier disputes a payment six months later, the PO number and invoice number together settle it in one minute." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Typing the Amount column by hand.", right: "Use =E2*F2 so Amount always equals Qty x Rate.", why: "A typed amount hides any typo in quantity or rate. The formula guarantees the three columns agree." },
      { wrong: "Leaving Status blank and filling it \"later\".", right: "Set Open at entry and update to Received or Cancelled the day it happens.", why: "Blank rows vanish from every COUNTIF summary, so pending orders silently disappear from the chase list." },
      { wrong: "Counting cancelled POs in supplier totals.", right: "Add the Status = Received criteria to supplier SUMIFS.", why: "Cancelled orders inflate what you appear to buy from a supplier and mislead negotiation talks." },
      { wrong: "Deleting a cancelled PO row from the register.", right: "Keep the row with Status Cancelled and Payment NA.", why: "PO numbers must never repeat. The register is a history, not a to-do list; deleting rows breaks that promise." }
    ],
    shortcuts: [
      { keys: "Alt + Down", action: "Open the Status or Payment dropdown in a register cell" },
      { keys: "Ctrl + Shift + L", action: "Filter the register to Status = Open for the daily chase list" },
      { keys: "Ctrl + ;", action: "Stamp today's date into the Date column" }
    ],
    exercises: [
      { task: "In the purchase-records workbook, complete the Amount column with formulas and add the Status and Payment dropdowns on the Exercise sheet.", hint: "Amount = Qty x Rate. Validation lists: Open, Received, Cancelled and Paid, Unpaid, Partial, NA.", solution: "PO-2026-014 is 4,200; PO-2026-019 is 40,500. Match all six rows against the Solution sheet." },
      { task: "Build the status summary block with COUNTIF and SUMIFS for Open, Received and Cancelled.", hint: "One COUNTIF per status for the count, one SUMIFS per status for the value.", solution: "Open: 3 POs, 1,06,100. Received: 2 POs, 6,800. Cancelled: 1 PO, 3,125." },
      { task: "Compute received-only purchase totals per supplier, then compute the total Unpaid amount.", hint: "SUMIFS with two criteria pairs for supplier totals; SUMIFS on Payment = Unpaid for the payables.", solution: "Sharma 4,200, Kavya 2,600, Global 0 received, TechDesk 0 received. Unpaid total 65,600." }
    ],
    quiz: [
      { q: "How should the Amount column of a PO be filled?", options: ["Typed from the supplier's quote", "=Qty * Rate as a formula", "Copied from the invoice", "Left blank until payment"], answer: 1, explanation: "The formula ties Amount to Qty and Rate, so any correction in either automatically corrects the amount. Typed amounts hide inconsistencies." },
      { q: "Which formula counts how many POs are still pending delivery?", options: ["=SUM(H2:H100)", "=COUNTIF($H$2:$H$100, \"Open\")", "=COUNTA($H$2:$H$100)", "=COUNTIF($H$2:$H$100, \"Received\")"], answer: 1, explanation: "COUNTIF with the criteria Open counts exactly the undelivered POs. COUNTA counts all filled cells regardless of status." },
      { q: "Why add the criteria Status = Received to a supplier-wise SUMIFS?", options: ["To speed up the formula", "To exclude cancelled and still-open POs from real purchase totals", "To sort by supplier", "To count POs instead of adding"], answer: 1, explanation: "Only received goods are real purchases. Without the status filter, cancelled POs inflate the supplier total and distort negotiations." },
      { q: "What should happen to a cancelled purchase order?", options: ["Delete the row so it never confuses totals", "Keep the row with Status Cancelled", "Mark it as Received to close it", "Copy it to next month"], answer: 1, explanation: "The register is a permanent history and PO numbers are never reused. Cancelling via Status keeps the record while excluding it from purchase totals." }
    ],
    summary: [
      "One row per PO with a permanent number pattern PO-YYYY-NNN.",
      "Amount = Qty x Rate, always a formula; Status and Payment are dropdowns.",
      "COUNTIF on Status gives the daily pending list; SUMIFS gives its value.",
      "Supplier totals use SUMIFS with Status = Received to exclude cancelled orders.",
      "The Payment column completes the PO lifecycle from order to paid."
    ],
    downloads: [
      { file: "assets/files/level5/purchase-records.xlsx", label: "Purchase records workbook", note: "Exercise and Solution sheets — purchase order register" }
    ],
    keywords: ["purchase order register", "PO tracking", "pending purchase orders", "supplier wise totals", "sumifs purchase", "payment status"]
  }
);
