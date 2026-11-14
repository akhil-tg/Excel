// Excel Lab — Level 6: Advanced Excel (part 1: metadata + lessons 1-5)
window.LEVEL_META = window.LEVEL_META || {};
window.LEVEL_LESSONS = window.LEVEL_LESSONS || {};

window.LEVEL_META[6] = {
  id: 6,
  title: "Advanced Excel",
  subtitle: "Power Query, Power Pivot, DAX and automation with macros and VBA",
  description: "Go professional: nested formulas, advanced lookups, Power Query, Power Pivot, an introduction to DAX, and automation with macros and VBA.",
  difficulty: "Advanced",
  icon: "bolt"
};

(window.LEVEL_LESSONS[6] = window.LEVEL_LESSONS[6] || []).push(
  {
    id: "l6-nested-formulas",
    title: "Advanced Nested Formulas",
    difficulty: "Advanced",
    duration: 16,
    version: "",
    objectives: [
      "Combine IF with AND and OR to test several conditions in one formula",
      "Wrap VLOOKUP in IFERROR so reports show clean messages instead of errors",
      "Build readable text with TEXT, TODAY, LEFT and FIND",
      "Read any long formula from the inside out",
      "Build a complex formula in helper cells first, then combine it into one cell"
    ],
    intro: "Until now you have mostly used one function at a time. Real office work is messier: a bonus depends on the sales score AND the attendance, an invoice is overdue only if TODAY is past the due date, and a lookup should show a friendly message instead of #N/A. One lonely function cannot handle questions like that — but several functions joined together can.\n\nIn this lesson you learn the three habits professionals use to stay in control: combining functions that work well together, reading a long formula from the inside out, and building complicated formulas in small pieces first. Master these and no formula in any workbook will scare you.",
    sections: [
      {
        heading: "Why real formulas need more than one function",
        blocks: [
          { type: "p", text: "A formula is nested when one function sits inside another. The inner function calculates first, and its result is used by the outer function. Think of it like packing a box inside a box: Excel opens the outer box and finds the inner one already solved." },
          { type: "p", text: "At Sunrise Traders, almost every real question needs this. \"Should Rahul get the quarterly bonus?\" needs his score, his attendance and a rule. \"Is invoice INV-2041 overdue?\" needs today's date and the due date. You could answer with several separate cells, and sometimes you should — but often one well-built formula is cleaner, and you can copy it down a whole column in seconds." },
          { type: "table", headers: ["Combination", "What it solves", "You will use it for"], rows: [
            ["IF + AND / OR", "One decision that depends on several tests", "Bonus rules, pass/fail, approval checks"],
            ["IFERROR + VLOOKUP", "A lookup that shows a clean message instead of #N/A", "Customer names, product names, prices"],
            ["TEXT + TODAY", "A sentence that contains a date or today's date", "Report titles, overdue warnings"],
            ["LEFT + FIND", "Pulling out the part of a text code you need", "Splitting product codes, invoice prefixes"]
          ] },
          { type: "tip", text: "Nesting is not about showing off. Each nesting exists because the question itself has several parts. If you can say the rule out loud in a sentence with the word \"and\", \"or\" or \"if... then... else\", the formula will need nesting." }
        ]
      },
      {
        heading: "IF with AND and OR: the bonus rule",
        blocks: [
          { type: "p", text: "IF alone can test one condition. To test several conditions in the same IF, wrap them in AND (all conditions must be true) or OR (any one condition is enough). Here is the annual appraisal at Sunrise Traders. HR's rule: \"Excellent\" needs a score of at least 80 and attendance of at least 90%. Anything else is \"Review\"." },
          { type: "grid", title: "Appraisal sheet with a combined IF + AND formula", rows: [
            ["Employee", "Score", "Attendance", "Result"],
            ["Arjun Mehta", "86", "94%", "=IF(AND(B2>=80,C2>=90%),\"Excellent\",\"Review\")"],
            ["Priya Nair", "74", "96%", "=IF(AND(B3>=80,C3>=90%),\"Excellent\",\"Review\")"],
            ["Rahul Verma", "68", "88%", "=IF(AND(B4>=80,C4>=90%),\"Excellent\",\"Review\")"],
            ["Sneha Iyer", "91", "91%", "=IF(AND(B5>=80,C5>=90%),\"Excellent\",\"Review\")"]
          ] },
          { type: "p", text: "Arjun passes both tests, so AND returns TRUE and he is rated \"Excellent\". Rahul fails both, so he gets \"Review\". Priya is the interesting one: her score is only 74 but her attendance is 96%. Under AND she would fail. If HR decides a score of 70 or attendance of 95% is also enough, switch the inner test to OR:" },
          { type: "code", lang: "excel", text: "=IF(AND(B2>=80,C2>=90%),\"Excellent\",\n   IF(OR(B2>=70,C2>=95%),\"Good\",\"Needs training\"))" },
          { type: "p", text: "Read it as a sentence: \"If the score is 80 or more and attendance is 90% or more, say Excellent. Otherwise, if the score is 70 or more or attendance is 95% or more, say Good. In every other case, say Needs training.\" Notice how the sentence and the formula match almost word for word — that is the sign of a good nested formula." }
        ]
      },
      {
        heading: "IFERROR plus VLOOKUP: clean messages instead of errors",
        blocks: [
          { type: "p", text: "VLOOKUP returns #N/A whenever the value it looks for is not in the table. In a live report that is normal: a new product code has not been added yet, or someone typed a code with an extra space. An #N/A down a whole column looks broken to a manager, even when the formula is doing its job correctly." },
          { type: "p", text: "The professional pattern is to wrap the lookup in IFERROR. IFERROR takes two arguments: the thing to try, and what to show if that thing returns any error." },
          { type: "formula", name: "IFERROR + VLOOKUP", syntax: "=IFERROR(VLOOKUP(lookup_value, table_array, col_index_num, FALSE), \"message\")", example: "=IFERROR(VLOOKUP(A2,Products!$A$2:$D$200,2,FALSE),\"Code not found\")", explanation: "Looks up the product name for the code in A2. If the code exists, the name appears. If not, the cell shows \"Code not found\" instead of #N/A." },
          { type: "warning", text: "Do not wrap every formula in IFERROR \"just in case\". IFERROR also hides serious errors like #REF! (a deleted column) and #DIV/0! (a real mistake in your data). Use it only where \"not found\" or \"empty cell\" is a normal, expected situation — lookups are the classic case." },
          { type: "tip", text: "Make the fallback message useful. \"Code not found\" tells someone what to fix. A blank cell from =IFERROR(VLOOKUP(...),\"\") tells them nothing, and empty-looking cells get copied into reports as if they were real data." }
        ]
      },
      {
        heading: "TEXT, TODAY, LEFT and FIND: formulas that build text",
        blocks: [
          { type: "p", text: "Reports are not only numbers. Titles, statuses and warnings are text, and they often need a date or a piece of a code inside them. Four functions cover most of this work." },
          { type: "ul", items: [
            "`TODAY()` returns today's date, and it updates every day you open the file.",
            "`TEXT(value, \"format\")` turns a number or date into text that looks the way you want, such as `TEXT(C2,\"dd-mmm-yyyy\")` giving 14-Nov-2026.",
            "`FIND(\"-\", A2)` returns the position of the first dash in the text in A2.",
            "`LEFT(A2, n)` returns the first n characters of the text in A2."
          ] },
          { type: "p", text: "Here is the invoice overdue check the accounts team at Sunrise Traders runs every week. Each row has an invoice code in column A and a due date in column C. One formula builds the whole status message:" },
          { type: "code", lang: "excel", text: "=IF(TODAY()>C2,\n   \"Overdue by \" & (TODAY()-C2) & \" days\",\n   \"Due on \" & TEXT(C2,\"dd-mmm\"))" },
          { type: "p", text: "If the due date has passed, TODAY minus the due date is a number of days, and the formula builds a sentence like \"Overdue by 6 days\". If it is not yet due, TEXT formats the date so the message reads \"Due on 21-Nov\" instead of showing a raw date serial number." },
          { type: "p", text: "LEFT + FIND is the standard way to split a code. Product codes at Sunrise Traders look like NOTE-A4-48PK, PEN-BLUE-10 or STAP-NO1-METAL. To get the product family (everything before the first dash):" },
          { type: "code", lang: "excel", text: "=LEFT(A2, FIND(\"-\", A2) - 1)" },
          { type: "p", text: "FIND(\"-\", A2) returns the position of the dash — say 5 for NOTE-A4-48PK. Subtracting 1 gives 4, and LEFT(A2, 4) returns NOTE. Because FIND looks the position up instead of you counting characters, the same formula works for every code, long or short." }
        ]
      },
      {
        heading: "Read it inside out, build it in pieces",
        blocks: [
          { type: "p", text: "Long formulas look impossible until you know the trick: do not read left to right. Start at the deepest inner function and work outwards, one layer at a time. Take this real formula from the accounts sheet:" },
          { type: "code", lang: "excel", text: "=IF(TODAY()>VLOOKUP(A2,Invoices,3,FALSE),\"Overdue\",\"On time\")" },
          { type: "steps", items: [
            "Innermost layer: VLOOKUP(A2,Invoices,3,FALSE) — finds the due date of the invoice whose code is in A2.",
            "Middle layer: TODAY()>that result — asks \"is today after the due date?\" and returns TRUE or FALSE.",
            "Outer layer: IF(TRUE/FALSE, \"Overdue\", \"On time\") — turns the answer into a readable status."
          ] },
          { type: "p", text: "Three short steps and the formula is fully understood. To build such a formula, reverse the process: make one small piece per helper cell, check each result is correct, then combine. In the invoice sheet you might build it like this:" },
          { type: "grid", title: "Helper cells first, combined formula last", rows: [
            ["Invoice", "Due date", "E (due date)", "F (is it late?)", "G (final)"],
            ["INV-2041", "02-Nov-2026", "=VLOOKUP(A2,Invoices,3,FALSE)", "=TODAY()>E2", "=IF(F2,\"Overdue\",\"On time\")"],
            ["INV-2042", "28-Nov-2026", "=VLOOKUP(A3,Invoices,3,FALSE)", "=TODAY()>E3", "=IF(F3,\"Overdue\",\"On time\")"]
          ] },
          { type: "p", text: "When every helper cell shows the right answer, replace E2 inside F2, then F2 inside G2, and you get the one-cell version above. If the combined formula misbehaves later, split it back into helpers to find the broken piece. Helper cells are also kinder to whoever inherits your file — they can see your thinking." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Writing =IF(B2>=80, C2>=90%, \"Excellent\", \"Review\") and expecting both tests to be checked together.", right: "Wrap the tests in a single condition function: =IF(AND(B2>=80,C2>=90%),\"Excellent\",\"Review\").", why: "IF accepts exactly one condition. Written as above, Excel compares the two tests against each other and returns strange results. AND and OR exist precisely to combine several tests into one." },
      { wrong: "Typing a five-function nested formula from memory and trusting it works.", right: "Build each part in its own helper cell, confirm every result, then combine them into one cell.", why: "One wrong bracket or comma in a long formula is nearly impossible to spot. Helper cells turn one big problem into several tiny, checkable ones." },
      { wrong: "Wrapping every formula on the sheet in IFERROR so the report \"looks clean\".", right: "Use IFERROR only for expected situations like a missing lookup code; leave real errors visible.", why: "IFERROR cannot tell a missing product code from a broken reference. Hiding #REF! errors means a deleted column can silently show wrong totals for months." },
      { wrong: "Forgetting one closing bracket in a long nested formula and getting an error dialog.", right: "Count brackets while writing, or let the tooltip's coloured brackets guide you — every ( needs its ).", why: "Excel refuses to accept a formula with unbalanced brackets. Matching them in a 200-character formula wastes far more time than building it in helper cells in the first place." }
    ],
    shortcuts: [
      { keys: "F4", action: "While editing a reference, cycle between A2, $A$2, A$2 and $A2 to lock ranges before copying" },
      { keys: "Ctrl + `", action: "Show all formulas on the sheet instead of their results, to review nesting at a glance" },
      { keys: "F9", action: "With part of a formula selected in edit mode, calculate just that piece to see what it returns" }
    ],
    exercises: [
      { task: "In a blank sheet, set up the Sunrise Traders appraisal table: A2:A5 the four salespeople (Arjun Mehta, Priya Nair, Rahul Verma, Sneha Iyer), B2:B5 their scores (86, 74, 68, 91) and C2:C5 their attendance as percentages (94%, 96%, 88%, 91%). Build the two-level rating formula from the lesson in D2 and copy it down.", hint: "Outer IF tests AND(score>=80, attendance>=90%). The value_if_false is another whole IF using OR(score>=70, attendance>=95%).", solution: "=IF(AND(B2>=80,C2>=90%),\"Excellent\",IF(OR(B2>=70,C2>=95%),\"Good\",\"Needs training\")) — Arjun and Sneha show Excellent, Priya shows Good, Rahul shows Needs training." },
      { task: "Build the invoice status formula: A2 holds INV-2041 and C2 holds the due date 02-Nov-2026. The cell should read \"Overdue by n days\" when the due date has passed, and \"Due on dd-mmm\" when it has not.", hint: "Compare TODAY() with C2. To join text and numbers use &. TEXT(C2,\"dd-mmm\") turns the date into readable text.", solution: "=IF(TODAY()>C2,\"Overdue by \"&(TODAY()-C2)&\" days\",\"Due on \"&TEXT(C2,\"dd-mmm\"))" },
      { task: "Product codes in A2:A6 are NOTE-A4-48PK, PEN-BLUE-10, STAP-NO1-METAL, PRINT-LJ-220 and LAP-15-I5. In B2, write one formula that returns the product family before the first dash, then copy it down.", hint: "Find the position of the first dash with FIND, subtract 1, and take that many characters from the left.", solution: "=LEFT(A2,FIND(\"-\",A2)-1) — returns NOTE, PEN, STAP, PRINT and LAP." }
    ],
    quiz: [
      { q: "With B2 = 86 and C2 = 94%, what does =AND(B2>=80, C2>=90%) return?", options: ["TRUE", "FALSE", "86", "#VALUE!"], answer: 0, explanation: "AND returns TRUE only when every test inside it is true. 86>=80 is TRUE and 94%>=90% is TRUE, so the whole AND is TRUE." },
      { q: "Which formula shows \"Not found\" instead of #N/A when a VLOOKUP misses?", options: ["=IF(VLOOKUP(A2,T,2,FALSE),\"Not found\")", "=IFERROR(VLOOKUP(A2,T,2,FALSE),\"Not found\")", "=VLOOKUP(IFERROR(A2),T,2,FALSE)", "=VLOOKUP(A2,T,2,FALSE)+\"Not found\""], answer: 1, explanation: "IFERROR takes two arguments: the formula to try and the value to show if it returns an error. Wrapping the VLOOKUP is exactly its purpose." },
      { q: "What does =LEFT(\"PEN-BLUE-10\", FIND(\"-\",\"PEN-BLUE-10\")-1) return?", options: ["PEN", "PEN-", "BLUE", "#VALUE!"], answer: 0, explanation: "FIND returns 4 (the position of the first dash). Minus 1 gives 3, and LEFT takes the first 3 characters: PEN." },
      { q: "You need a long nested formula you have never built before. What is the professional way to start?", options: ["Type the whole thing in one cell and hope", "Build each part in its own helper cell, check the results, then combine", "Copy any long formula from the internet and edit it live", "Use longer function names so it is easier to read"], answer: 1, explanation: "Helper cells let you test each layer against real data before combining. If the combined formula fails later, you can split it back apart to find the broken piece." },
      { q: "Why does =TODAY() show a different value when you open the workbook next week?", options: ["It is broken and should be replaced by a typed date", "It always recalculates to the current date", "It stores the date the file was created", "It follows the computer's time zone, not the date"], answer: 1, explanation: "TODAY() is a volatile function: every recalculation (including reopening the file) makes it return the current date. That is exactly what you want for overdue checks — and why you must not use it for dates that must stay fixed." }
    ],
    summary: [
      "Nesting means one function inside another; the inner function is calculated first.",
      "AND requires every test to be true; OR needs only one; both collapse several tests into the single condition IF needs.",
      "IFERROR around VLOOKUP turns normal \"not found\" situations into readable messages — but never use it to hide structural errors like #REF!.",
      "TEXT formats numbers and dates inside sentences; TODAY brings the live current date; LEFT+FIND splits codes at a known character.",
      "Read long formulas from the inside out, and build new ones in helper cells before combining — these two habits make any formula manageable."
    ],
    downloads: [],
    keywords: ["nested formulas", "IF AND OR", "IFERROR VLOOKUP", "TEXT function", "TODAY function", "LEFT FIND", "helper cells", "reading long formulas", "invoice overdue check"]
  },
  {
    id: "l6-advanced-lookups",
    title: "Advanced Lookup Techniques",
    difficulty: "Advanced",
    duration: 18,
    version: "",
    objectives: [
      "Do a two-way lookup with INDEX and two MATCH functions",
      "Look up to the left with INDEX/MATCH when the result column sits before the search column",
      "Match on two or more criteria using a helper column or an array formula",
      "Use approximate-match lookups for tax slabs and commission tiers",
      "Recognise XMATCH, the modern 365 replacement for MATCH"
    ],
    intro: "VLOOKUP answers one question: \"find this value in the first column and return something from the right.\" Real workbooks ask harder questions. \"What is the price of Staplers in Chennai?\" needs a row and a column at the same time. \"What is Sneha Iyer's employee ID?\" needs a value from a column on the left, which VLOOKUP refuses to do. \"What commission rate applies to ₹1,75,000 of sales?\" needs the closest lower value, not an exact one.\n\nEvery one of these has a clean answer, and they all rest on one idea: separate finding a position from fetching a value. That is exactly what INDEX and MATCH do, and it is why professional workbooks use them far more than VLOOKUP.",
    sections: [
      {
        heading: "INDEX and MATCH: the pair that does everything",
        blocks: [
          { type: "p", text: "Recall the pattern from Level 5: `=INDEX(values, MATCH(what, where, 0))`. MATCH finds the position of a value in a list — for example, Staplers is the 3rd item — and INDEX returns the item at that position in another list. Neither function knows anything about the other; you are simply using the number MATCH returns as the address INDEX needs." },
          { type: "p", text: "This separation is the whole secret. Because MATCH can search any column and INDEX can fetch from any range, you can look up in any direction — left, right, up, down. VLOOKUP can only ever look right. That single limitation is why so many experienced users stopped using VLOOKUP altogether." },
          { type: "table", headers: ["Need", "VLOOKUP", "INDEX/MATCH"], rows: [
            ["Return a column left of the search column", "Not possible", "Works normally"],
            ["Insert or delete columns inside the table", "Breaks (col_index_num is a hardcoded number)", "Survives (ranges are named directly)"],
            ["Match on two criteria at once", "Needs a helper column", "Works with one array formula"],
            ["Two-way lookup (row and column)", "Not possible in one formula", "Two MATCHes inside one INDEX"]
          ] }
        ]
      },
      {
        heading: "Two-way lookup: one INDEX, two MATCHes",
        blocks: [
          { type: "p", text: "Sunrise Traders keeps one price list per city, laid out as a matrix: products down the side, cities across the top. To fetch the Chennai price for Staplers you need to find Staplers' row and Chennai's column at the same time. This is a two-way lookup, and it is one MATCH for each direction inside a single INDEX." },
          { type: "grid", title: "City price matrix (A1:E5) with the two-way lookup in G3", rows: [
            ["Product", "Delhi", "Mumbai", "Chennai", "Kochi"],
            ["Notebooks", "120", "125", "118", "122"],
            ["Pens", "35", "38", "34", "36"],
            ["Staplers", "210", "215", "205", "212"],
            ["Printers", "8500", "8600", "8400", "8550"]
          ] },
          { type: "code", lang: "excel", text: "Product name in G2, city name in H2, price wanted in I2:\n\n=INDEX(B2:E5, MATCH(G2, A2:A5, 0), MATCH(H2, B1:E1, 0))" },
          { type: "p", text: "Walk through it with Staplers and Chennai. MATCH(\"Staplers\", A2:A5, 0) returns 3, because Staplers is the third product row. MATCH(\"Chennai\", B1:E1, 0) returns 3, because Chennai is the third city column. INDEX then returns the item at row 3, column 3 of B2:E5 — 205, the Chennai stapler price. Change either dropdown cell and the price follows." },
          { type: "tip", text: "Keep the two ranges consistent in your head: the first MATCH looks down the row labels (A2:A5), the second looks across the column headers (B1:E1), and the INDEX range (B2:E5) is the inner rectangle where they cross. If you get #N/A, check which of the three ranges you typed wrong." }
        ]
      },
      {
        heading: "Lookup to the left",
        blocks: [
          { type: "p", text: "Here is the employee table where VLOOKUP fails. The employee ID is in column A, the city in column B, and the full name in column C. You know the name and want the ID — the answer sits to the left of what you are searching. VLOOKUP cannot do this without rebuilding the whole table." },
          { type: "grid", title: "Employees sheet — the answer (ID) is left of the search column (Name)", rows: [
            ["EmpID", "City", "Name"],
            ["S001", "Delhi", "Arjun Mehta"],
            ["S002", "Kochi", "Priya Nair"],
            ["S003", "Mumbai", "Rahul Verma"],
            ["S004", "Chennai", "Sneha Iyer"]
          ] },
          { type: "code", lang: "excel", text: "Name typed in E2, ID wanted in F2:\n\n=INDEX(A2:A5, MATCH(E2, C2:C5, 0))" },
          { type: "p", text: "MATCH finds \"Priya Nair\" in the Name column — position 2 — and INDEX returns the second item of the EmpID range: S002. The two ranges do not have to touch or even be next to each other; they only need to be the same height, so that position 2 means the same row in both." },
          { type: "warning", text: "The two ranges must line up exactly: if MATCH searches C2:C50, INDEX must fetch from a range that starts at row 2 and ends at row 50, such as A2:A50. A mismatched start row silently returns the value from the wrong person — the most dangerous kind of lookup bug." }
        ]
      },
      {
        heading: "Multi-criteria lookup: matching on two things at once",
        blocks: [
          { type: "p", text: "Sales data rarely has one unique key. The sheet below lists many rows per salesperson and per month, so \"Arjun Mehta\" alone is not unique, and neither is \"Jan\". Only the pair together is. The classic solution is a helper column that glues the two values into one key, and a lookup on that key." },
          { type: "grid", title: "Sales rows with a helper key in column D", rows: [
            ["Salesperson", "Month", "Amount", "Key (helper)", "Lookup"],
            ["Arjun Mehta", "Jan", "82000", "=A2&\"|\"&B2", "Arjun Mehta|Jan"],
            ["Priya Nair", "Jan", "91500", "=A3&\"|\"&B3", "Arjun Mehta|Feb"],
            ["Arjun Mehta", "Feb", "78000", "=A4&\"|\"&B4", "=INDEX(C2:C4,MATCH(E2,D2:D4,0))"]
          ] },
          { type: "code", lang: "excel", text: "Helper key in D2, copied down:   =A2&\"|\"&B2\n\nLookup formula:  =INDEX(C2:C50, MATCH(E2, D2:D50, 0))\nwhere E2 contains the combined key, e.g. \"Arjun Mehta|Feb\"" },
          { type: "p", text: "If you cannot or do not want to add a helper column, there is an array version that does the gluing in memory. It multiplies two TRUE/FALSE lists — which Excel treats as 1/0 — and finds the row where the product is 1, meaning both tests were TRUE:" },
          { type: "code", lang: "excel", text: "=INDEX(C2:C50, MATCH(1, (A2:A50=E2)*(B2:B50=F2), 0))\n\nOlder Excel (2019 and earlier): confirm with Ctrl+Shift+Enter.\nMicrosoft 365: plain Enter works." },
          { type: "p", text: "Do not worry about the array mechanics; the sentence to remember is: \"give me the row where salesperson equals E2 AND month equals F2\". Each comparison produces TRUE/FALSE down the column, TRUE times TRUE is 1, and MATCH looks for that single 1. In older Excel versions such formulas must be confirmed with Ctrl+Shift+Enter — you will see curly braces appear around the formula. In Microsoft 365, just press Enter." },
          { type: "tip", text: "The helper-column version is slower to set up but easier for colleagues to read and debug. Use the array version when the sheet is a downloaded export you cannot restructure, or when adding columns would break someone else's layout." }
        ]
      },
      {
        heading: "Approximate match: tax slabs and commission tiers",
        blocks: [
          { type: "p", text: "Sometimes the lookup value will not be in the table at all — and should not be. A commission of ₹1,75,000 of sales does not appear as a row; it falls into a tier. This is the fourth argument of VLOOKUP set to TRUE (or MATCH with 1 as its third argument): find the largest value that is still less than or equal to the lookup value." },
          { type: "table", headers: ["Sales from (₹)", "Rate", "Meaning"], rows: [
            ["0", "1%", "The first ₹49,999 earns 1%"],
            ["50000", "3%", "₹50,000 up to ₹99,999 earns 3%"],
            ["100000", "5%", "₹1,00,000 up to ₹1,99,999 earns 5%"],
            ["200000", "8%", "₹2,00,000 and above earns 8%"]
          ] },
          { type: "code", lang: "excel", text: "Tier table in E2:F5, monthly sales in B2, commission rate in C2:\n\n=VLOOKUP(B2, $E$2:$F$5, 2, TRUE)" },
          { type: "p", text: "With B2 = 175000, VLOOKUP scans down the first column for the largest value not above 175000 — that is 100000 — and returns 5%. This is how income-tax slabs, courier weight brackets and bulk discounts are all modelled in real finance sheets. The lookup column must be sorted in ascending order, or approximate match returns nonsense." },
          { type: "p", text: "One honest note for Microsoft 365 users: the functions XLOOKUP and XMATCH exist as modern replacements. XMATCH is MATCH with exact match as the default (no need for the ,0 argument), can search from the bottom up, and supports the same approximate behaviour: `=XMATCH(F2, A2:A50)` finds F2 exactly. If your whole team is on 365 you may prefer it, but INDEX/MATCH remains the skill that works on every Excel any client or employer will hand you, so learn it first." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Leaving out the FALSE in VLOOKUP on a table that happens to be sorted.", right: "Always write the fourth argument explicitly: =VLOOKUP(A2, T, 3, FALSE) for exact match.", why: "With the argument omitted, VLOOKUP quietly does an approximate match. On sorted data it can return the row above the true match and give a plausible-looking wrong answer — the worst kind of error." },
      { wrong: "Typing a multi-criteria array formula in Excel 2019 or earlier and pressing Enter.", right: "Confirm it with Ctrl+Shift+Enter, or use the helper-column version instead.", why: "In pre-365 Excel, array formulas must be entered with Ctrl+Shift+Enter. Plain Enter makes the formula compare only the first row and return #VALUE! or a wrong match." },
      { wrong: "Leaving the tier table unsorted when using approximate match.", right: "Sort the lookup column ascending: 0, 50000, 100000, 200000.", why: "Approximate match assumes ascending order and stops at the wrong tier, often returning the rate for a lower slab than the customer earned." },
      { wrong: "Making the INDEX range and the MATCH range start at different rows.", right: "Keep them aligned: INDEX(A2:A50, MATCH(...,C2:C50,0)) — both start at row 2.", why: "INDEX and MATCH talk in positions, not cell addresses. Position 5 must be the same physical row in both ranges, or you fetch another person's data with no error message to warn you." }
    ],
    shortcuts: [
      { keys: "F4", action: "Lock the tier table reference ($E$2:$F$5) while writing the formula, so it survives copying down" },
      { keys: "Ctrl + Shift + Enter", action: "Confirm array formulas in Excel 2019 and earlier (365 users can just press Enter)" },
      { keys: "F9", action: "Select (A2:A50=E2)*(B2:B50=F2) inside the formula and press F9 to see the TRUE/FALSE products while debugging" }
    ],
    exercises: [
      { task: "Build the city price matrix from the lesson (Products in A2:A5: Notebooks, Pens, Staplers, Printers; columns B to E for Delhi, Mumbai, Chennai, Kochi with any plausible ₹ prices). In G2 and H2 type a product and a city, and in I2 write the two-way lookup that returns the price.", hint: "INDEX over the price rectangle B2:E5, one MATCH down A2:A5 for the product, one MATCH across B1:E1 for the city.", solution: "=INDEX(B2:E5, MATCH(G2, A2:A5, 0), MATCH(H2, B1:E1, 0)) — with Staplers and Chennai it returns the Chennai stapler price." },
      { task: "Recreate the employee table (A: EmpID S001-S004, B: City, C: Name for Arjun Mehta, Priya Nair, Rahul Verma, Sneha Iyer). In E2 type a name and fetch their ID in F2 — the ID column is to the LEFT of the name column.", hint: "MATCH the name in C2:C5; INDEX into A2:A5. The ranges must start and end on the same rows.", solution: "=INDEX(A2:A5, MATCH(E2, C2:C5, 0)) — Priya Nair returns S002." },
      { task: "Set up the commission tier table from the lesson in E2:F5 (0→1%, 50000→3%, 100000→5%, 200000→8%). In B2 enter 175000 and write the formula that returns the correct rate. Then test B2 with 40000, 60000 and 250000.", hint: "Fourth argument TRUE means approximate: the largest tier value not greater than the sales figure.", solution: "=VLOOKUP(B2,$E$2:$F$5,2,TRUE) — returns 5% for 175000; 1% for 40000; 3% for 60000; 8% for 250000." },
      { task: "Add a helper key column to any two-column list (Salesperson and Month) and write a lookup that finds the Amount for \"Sneha Iyer|Feb\". Then rewrite it without the helper column using the array version.", hint: "Key: =A2&\"|\"&B2. Array: MATCH(1,(range=E2)*(range=F2),0) inside INDEX. In Excel 2019 or earlier confirm with Ctrl+Shift+Enter.", solution: "Helper: =INDEX(C2:C50, MATCH(E2, D2:D50, 0)) with E2 = \"Sneha Iyer|Feb\". Array: =INDEX(C2:C50, MATCH(1,(A2:A50=E2)*(B2:B50=F2),0))" }
    ],
    quiz: [
      { q: "In =INDEX(B2:E5, MATCH(G2, A2:A5, 0), MATCH(H2, B1:E1, 0)), what does the second MATCH do?", options: ["Finds the row of the product", "Finds the column of the city", "Returns the price itself", "Counts how many cities there are"], answer: 1, explanation: "The first MATCH walks down the product names to get the row; the second walks across the city headers B1:E1 to get the column. INDEX then uses both numbers as row and column." },
      { q: "The value you want is in column A and you are searching in column D. What works?", options: ["VLOOKUP with a negative column number", "INDEX(A:A, MATCH(...,D:D,0))", "VLOOKUP set to TRUE", "HLOOKUP turned sideways"], answer: 1, explanation: "VLOOKUP can never return a value from a column left of the search column. INDEX/MATCH is direction-blind, which is exactly why professionals prefer it." },
      { q: "In the array formula =MATCH(1,(A2:A50=E2)*(B2:B50=F2),0), why multiply the two comparisons?", options: ["To average the two columns", "TRUE times TRUE equals 1, which marks the row where both tests pass", "Multiplication is faster than addition in Excel", "Because MATCH cannot compare text directly"], answer: 1, explanation: "Each comparison gives a column of TRUE/FALSE, treated as 1/0. Only a row where both are TRUE gives 1×1 = 1, and MATCH finds that one position." },
      { q: "Monthly sales of ₹1,75,000 fall in the 100000 → 5% tier. Which formula returns 5%?", options: ["=VLOOKUP(175000, T, 2, FALSE)", "=VLOOKUP(175000, T, 2, TRUE)", "=VLOOKUP(175000, T, 1, TRUE)", "=MATCH(175000, T, 0)"], answer: 1, explanation: "Approximate match (TRUE) finds the largest tier value not above 175000 — that is 100000 — and returns 5%. FALSE would demand an exact 175000 row and return #N/A." },
      { q: "What is XMATCH?", options: ["An older name for MATCH, removed in 2016", "A 365 function like MATCH but exact by default, with more search options", "A function that matches two sheets automatically", "The VBA command for MATCH"], answer: 1, explanation: "XMATCH (Microsoft 365 only) defaults to exact match, so you no longer type the ,0 argument, and it can search from the bottom up. INDEX/MATCH still works everywhere and remains the safer universal habit." }
    ],
    summary: [
      "INDEX returns the value at a position; MATCH finds a position. Together they look up in any direction, which VLOOKUP cannot.",
      "Two-way lookup = one INDEX with a row MATCH and a column MATCH — perfect for product × city price matrices.",
      "Multi-criteria lookups use a glued helper key (=A2&\"|\"&B2) or the array pattern MATCH(1,(test1)*(test2),0).",
      "Approximate match (VLOOKUP TRUE / MATCH 1) drives tax slabs, commission tiers and discount brackets; the tier table must be sorted ascending.",
      "XMATCH exists in Microsoft 365 with exact match as default — worth knowing, but INDEX/MATCH works in every Excel you will meet."
    ],
    downloads: [],
    keywords: ["INDEX MATCH", "two-way lookup", "lookup left of column", "multi-criteria lookup", "array formula", "Ctrl Shift Enter", "approximate match VLOOKUP", "commission tiers", "tax slabs", "XMATCH"]
  },
  {
    id: "l6-let-lambda",
    title: "LET and LAMBDA: Modern Formula Tools",
    difficulty: "Advanced",
    duration: 18,
    version: "Microsoft 365",
    objectives: [
      "Name intermediate results inside a formula with LET for readability and speed",
      "Convert a repeated multi-step calculation into one LET formula",
      "Create your own reusable function with LAMBDA in Name Manager",
      "Call your custom function like a built-in one anywhere in the workbook",
      "Know what to use instead when colleagues are on older Excel"
    ],
    intro: "LET and LAMBDA are the two features that make formulas feel like real programming — and both need Microsoft 365 or Excel 2021 onward for LET, and Microsoft 365 for LAMBDA. If you or your team are on Excel 2016 or 2019, read this lesson anyway: the thinking is valuable, and this lesson also shows the older-Excel equivalents so nothing here blocks you.\n\nWhy do professionals care? Because reports repeat the same sub-calculation five times inside one formula, and nobody can read the result. LET lets you give a piece of a formula a name. LAMBDA goes further: it lets you invent a function that did not exist before, save it in the workbook, and call it by name like SUM or VLOOKUP. No VBA required.",
    sections: [
      {
        heading: "LET: give your intermediate results a name",
        blocks: [
          { type: "p", text: "Look at this bonus calculation from the Sunrise Traders appraisal. The sales figure B2 is used three times and the rate logic is repeated twice:" },
          { type: "code", lang: "excel", text: "=ROUND(B2*IF(B2>=100000,5%,IF(B2>=50000,3%,1%)),0)\n\n+IF(B2>=100000,2000,0)\n\n+B2*0.5%*IF(AND(B2>=100000,C2>=90%),1,0)" },
          { type: "p", text: "It works, but every reader must re-derive what B2 is and what the tiers mean. LET fixes both problems. Its syntax is a list of name–value pairs followed by the final calculation that uses those names:" },
          { type: "code", lang: "excel", text: "=LET(\n   sales, B2,\n   attend, C2,\n   rate, IF(sales>=100000, 5%, IF(sales>=50000, 3%, 1%)),\n   loyalty, IF(sales>=100000, 2000, 0),\n   attendanceBonus, sales*0.5%*IF(AND(sales>=100000,attend>=90%),1,0),\n   ROUND(sales*rate,0) + loyalty + attendanceBonus\n)" },
          { type: "p", text: "Read it top to bottom like a recipe: sales means B2, attend means C2, rate is the tier lookup, and the last line is the actual answer. Each name is defined once and used as many times as needed. If HR changes the tier thresholds next year, you edit one line instead of hunting through five copies of the same IF." },
          { type: "tip", text: "LET can also make workbooks faster. A sub-expression written out three times is calculated three times; named once in LET, it is calculated once and reused. On big sheets with heavy lookups, this is a real, measurable speed-up." }
        ]
      },
      {
        heading: "A full LET worked example: the variance message",
        blocks: [
          { type: "p", text: "Monthly targets are the bread and butter of retail reporting. The sheet below tracks each salesperson's actual sales against target, and the manager wants one column that speaks plainly: \"Above target by 8.2%\" or \"Below target by 3.5%\". Without LET, that sentence needs a long nest of repeated subtraction and division. With LET it reads like a sentence itself." },
          { type: "grid", title: "Target tracking — one LET formula filled down column E", rows: [
            ["Salesperson", "Actual (₹)", "Target (₹)", "Message (the LET formula)"],
            ["Arjun Mehta", "432000", "400000", "Above target by 8.0%"],
            ["Rahul Verma", "386000", "400000", "Below target by 3.5%"]
          ] },
          { type: "code", lang: "excel", text: "In E2, filled down:\n\n=LET(\n   actual, B2,\n   target, C2,\n   diff, actual - target,\n   pct, diff / target,\n   IF(diff>=0, \"Above target by \" & TEXT(pct,\"0.0%\"),\n                   \"Below target by \" & TEXT(-pct,\"0.0%\"))\n)" },
          { type: "p", text: "Four names, one answer. actual and target are plain cell references; diff is actual minus target; pct is diff divided by target. The final line builds the message, using -pct for the below-target case so the number reads positive (\"Below by 3.5%\", not \"Below by -3.5%\"). Notice how TEXT formats the fraction as a percentage inside the sentence — the same TEXT skill from the nested-formulas lesson." },
          { type: "steps", items: [
            "Start =LET( and add one name–value pair per idea: actual, target, diff, pct.",
            "Write the final calculation on the last line using only those names.",
            "Close the bracket and press Enter. If you get an error, check that every name except the last is followed by both a value and a comma.",
            "Copy the formula down. Because the names point at B2 and C2, they adjust to B3 and C3 automatically."
          ] },
          { type: "warning", text: "LET names cannot look like cell references. A name such as \"B2\" or \"CT1\" will be rejected or cause confusion. Use words: sales, target, rate, diff. Also, a LET name cannot be defined twice in the same formula." },
          { type: "p", text: "On Excel 2016 or 2019, LET is not available. The equivalent discipline still works: put each intermediate value in its own helper column with a clear header (Actual, Diff, Pct), and build the message from those. Same clarity, one column per name instead of one name per line." }
        ]
      },
      {
        heading: "LAMBDA: invent the function Excel is missing",
        blocks: [
          { type: "p", text: "LAMBDA lets you define a function with a name and arguments, exactly like the built-ins. The raw form looks odd at first: `=LAMBDA(x, y, x+y)(3, 4)` — it takes arguments, then a body, and you can even call it immediately with values in brackets at the end. Calling it directly like that is only for testing. The real power is saving it under a name." },
          { type: "steps", items: [
            "Test the logic in a cell first: =LAMBDA(sales, rate, ROUND(sales*rate,0))(82000, 5%) should return 4100.",
            "Select the whole LAMBDA expression (without the test values at the end) and copy it.",
            "Go to Formulas > Name Manager > New... Give it the name Commission.",
            "In \"Refers to\", paste the LAMBDA, keep the = sign, and click OK.",
            "Now use it anywhere: =Commission(B2, 5%) — exactly like a built-in function."
          ] },
          { type: "code", lang: "excel", text: "Definition saved in Name Manager as \"Commission\":\n=LAMBDA(sales, rate, ROUND(sales*rate, 0))\n\nUsage anywhere in the workbook:\n=Commission(B2, 5%)\n=Commission(120000, 8%)" },
          { type: "p", text: "Anyone opening the workbook can now call Commission, and Excel's formula tooltip shows the argument names you chose. If the commission rule changes — say a flat ₹500 minimum — you edit the definition once in Name Manager and every cell that uses it updates on recalculation. That is the same benefit a VBA function would give, but without macros, without security warnings, and it recalculates automatically like any formula." }
        ]
      },
      {
        heading: "A full LAMBDA worked example: NetAmount with GST",
        blocks: [
          { type: "p", text: "Sunrise Traders invoices follow one rule everywhere: net amount = quantity × unit price, minus any discount, plus 18% GST on the discounted subtotal. Doing this inline in every invoice sheet invites inconsistency — one file will use 18%, another 17.9%, a third will forget the discount. This is precisely the situation LAMBDA was built for." },
          { type: "code", lang: "excel", text: "Saved in Name Manager as \"NetAmount\":\n=LAMBDA(qty, price, discount,\n   LET(\n      gross, qty*price,\n      net, gross*(1-discount),\n      gst, ROUND(net*18%, 2),\n      ROUND(net+gst, 2)\n   )\n)\n\nUsage:\n=NetAmount(12, 210, 10%)   -> 2,226.96\n=NetAmount(B2, C2, D2)     -> filled down an invoice column" },
          { type: "p", text: "This is LET and LAMBDA working together: LAMBDA gives the function its name and arguments, LET organises the three inside steps. Run the numbers once so you trust the function: twelve staplers at ₹210 with a 10% discount gives a gross of 2520, a net of 2268, GST of 2268 × 18% = 408.24, and a total of 2276.24. Every invoice sheet in the company can now produce that figure identically, because the rule exists in exactly one place." },
          { type: "tip", text: "LAMBDA definitions live inside the workbook file. If you email the workbook, the functions travel with it. But a new blank workbook does not know them — keep a small \"function library\" workbook and copy definitions via Name Manager when starting a new project." }
        ]
      },
      {
        heading: "Which tool when: LET, LAMBDA, or old-school helpers",
        blocks: [
          { type: "table", headers: ["Situation", "Best tool", "Older-Excel alternative"], rows: [
            ["One formula with repeated sub-steps", "LET", "Helper columns with clear headers"],
            ["Same calculation repeated across many workbooks or sheets", "LAMBDA in Name Manager", "Copy a hidden \"calc\" sheet, or a VBA function (Level 8)"],
            ["A rule that changes often (rates, tiers)", "LAMBDA: edit the definition once", "One master helper cell that everything else references"],
            ["Colleagues must audit your formula without training", "LET with meaningful names", "Helper columns — most auditable of all"]
          ] },
          { type: "p", text: "One closing warning from the field: LET and LAMBDA formulas break with #NAME? errors the moment someone opens the file in Excel 2019. Before shipping a 365-only solution to a client or a new employer, ask what version they run. The safest professional pattern is LET/LAMBDA for your own analysis sheets, and plain helper columns for anything you hand to others." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Naming a LET variable \"B2\" or \"Total ₹\".", right: "Use simple word names: sales, target, rate, diff.", why: "Names that look like cell references are rejected or silently confusing, and special characters break the parser. Words are also easier for the next reader." },
      { wrong: "Forgetting that the last argument of LET is the actual result and stopping after the name–value pairs.", right: "Always end with the calculation that uses the names, then close the bracket.", why: "LET returns only its final argument. A LET that ends with definitions returns the last value pair, which is almost never what you meant." },
      { wrong: "Typing =Commission(B2,5%) in a fresh workbook and getting #NAME?, then concluding LAMBDA is broken.", right: "Check Name Manager in that workbook — custom functions exist per workbook, not per computer.", why: "LAMBDA definitions are stored in the workbook's Name Manager. A new file has no definitions until you add them (or copy them from your function library)." },
      { wrong: "Sending a LET/LAMBDA-heavy sheet to a colleague on Excel 2019.", right: "Confirm the recipient's Excel version, or provide a helper-column fallback.", why: "LET needs Excel 2021 or Microsoft 365 and LAMBDA needs Microsoft 365. Older versions show #NAME? for every such formula, with no hint of what went wrong." }
    ],
    shortcuts: [
      { keys: "Ctrl + F3", action: "Open Name Manager directly to add or edit LAMBDA definitions" },
      { keys: "Alt + Enter", action: "Inside a long LET formula, break it onto a new line for readability while editing" },
      { keys: "F9", action: "Select one LET name's value while editing to inspect what it holds" }
    ],
    exercises: [
      { task: "Rewrite the tier bonus =ROUND(B2*IF(B2>=100000,5%,IF(B2>=50000,3%,1%)),0) as a LET formula with names sales and rate. Test with B2 = 82000 (expect 2460) and B2 = 120000 (expect 6000).", hint: "LET(sales, B2, rate, IF(...), ROUND(sales*rate,0)) — the last line is the answer.", solution: "=LET(sales, B2, rate, IF(sales>=100000,5%,IF(sales>=50000,3%,1%)), ROUND(sales*rate,0))" },
      { task: "Build the variance message formula in a sheet: B2 actual 432000, C2 target 400000. The result must read \"Above target by 8.0%\". Change actual to 386000 and it must read \"Below target by 3.5%\".", hint: "Define diff and pct in LET. For the below case, use TEXT(-pct,\"0.0%\") so the number reads positive.", solution: "=LET(actual,B2,target,C2,diff,actual-target,pct,diff/target,IF(diff>=0,\"Above target by \"&TEXT(pct,\"0.0%\"),\"Below target by \"&TEXT(-pct,\"0.0%\")))" },
      { task: "Save a LAMBDA called Commission (sales, rate) that rounds sales×rate to 0 decimals, using Name Manager. Call it for 82000 at 3%.", hint: "Test =LAMBDA(sales,rate,ROUND(sales*rate,0))(82000,3%) in a cell first, then copy the definition without the test values into Name Manager.", solution: "Definition: =LAMBDA(sales, rate, ROUND(sales*rate, 0)) — then =Commission(82000,3%) returns 2460." },
      { task: "Create the NetAmount LAMBDA from the lesson (qty, price, discount with 18% GST) and verify =NetAmount(12, 210, 10%). Then compute an invoice line: 30 pens at ₹35 with no discount.", hint: "Inside the LAMBDA, use LET with gross, net, gst. GST is 18% of the discounted net.", solution: "=NetAmount(12, 210, 10%) returns 2276.24 (net 2268 + GST 408.24), and =NetAmount(30, 35, 0) returns 1239.00 (net 1050 + GST 189)." }
    ],
    quiz: [
      { q: "What is the main benefit of LET?", options: ["It makes formulas run on older Excel versions", "It names intermediate results once and reuses them, for clarity and speed", "It replaces VLOOKUP with a faster function", "It converts text to numbers automatically"], answer: 1, explanation: "LET defines name–value pairs and a final calculation. Each sub-step is written once, calculated once, and reads like a labelled recipe." },
      { q: "In =LET(a, B2, b, C2, a+b), what does Excel return?", options: ["#NAME? because a and b are invalid names", "The sum of B2 and C2", "The value of B2", "A dialog asking for a and b"], answer: 1, explanation: "The pairs define a = B2 and b = C2, and the final argument a+b is the result. Simple word names are perfectly valid LET names." },
      { q: "You defined a LAMBDA called Commission in Workbook A. You open a new blank Workbook B and type =Commission(100,5%). What happens?", options: ["It returns 5", "It returns #NAME? — definitions are stored per workbook", "Excel copies it from Workbook A automatically", "It asks you to define it"], answer: 1, explanation: "LAMBDA definitions live in each workbook's Name Manager. Workbook B does not know Commission until you add the definition there too." },
      { q: "Which Excel version can use LAMBDA?", options: ["Excel 2016 and later", "Excel 2019 and later", "Microsoft 365 (subscription)", "Any Excel with Power Pivot enabled"], answer: 2, explanation: "LAMBDA is a Microsoft 365 feature. LET needs 365 or Excel 2021. On Excel 2016/2019, helper columns are the honest equivalent." },
      { q: "Why would a professional combine LAMBDA with LET inside the definition?", options: ["LET makes the function run faster on every Excel version", "LET organises the internal steps with names, making the function easier to write and audit", "LAMBDA cannot do arithmetic without LET", "It is the only way to pass three arguments"], answer: 1, explanation: "LAMBDA handles the name, arguments and reuse; LET structures the body into readable named steps. They solve different problems and combine naturally." }
    ],
    summary: [
      "LET names intermediate results inside one formula: name–value pairs, then the final calculation — clearer to read and often faster to recalculate.",
      "LET needs Microsoft 365 or Excel 2021; on Excel 2016/2019, use helper columns with the same discipline.",
      "LAMBDA lets you define a custom function, save it in Name Manager, and call it like a built-in — no macros and no security warnings.",
      "LAMBDA definitions are stored per workbook, so a new file needs the definition copied in before it recognises your function.",
      "Use LET/LAMBDA for your own analysis; use helper columns for sheets you hand to users on older Excel."
    ],
    downloads: [],
    keywords: ["LET function", "LAMBDA function", "Name Manager", "custom function", "named intermediate results", "Microsoft 365 formulas", "readable formulas", "formula performance", "NetAmount GST"]
  },
  {
    id: "l6-power-query",
    title: "Power Query: Import and Clean Data Automatically",
    difficulty: "Advanced",
    duration: 25,
    version: "Excel 2016+",
    objectives: [
      "Import data from CSV files, other workbooks and whole folders with Get Data",
      "Use the Applied Steps pane to record and re-order every cleaning action",
      "Apply the everyday transforms: remove columns, change type, split, replace values, remove duplicates, trim and unpivot",
      "Load the cleaned result as a table and refresh it with one click",
      "Decide when Power Query beats manual cleaning"
    ],
    intro: "Every office has the same monthly ritual: download an export, delete useless columns, fix the dates that arrived as text, trim the spaces someone typed before every product name, remove the duplicated rows, reformat it all by hand — and repeat the entire ritual next month. Power Query, built into Excel 2016 and later at no extra cost, does the cleaning once and then repeats it perfectly forever with one click.\n\nPower Query is not a formula and not a macro. It is a recording of your cleaning steps, attached to the source file. When next month's export arrives with the same shape, you drop it in, press Refresh, and every step replays on the new data. This lesson takes you from a genuinely messy export to a clean, refreshable table.",
    sections: [
      {
        heading: "Why professionals clean with Power Query, not by hand",
        blocks: [
          { type: "p", text: "Manual cleaning has three hidden costs. First, it is slow: an hour of deleting, trimming and fixing types, every single month. Second, it is inconsistent: today you remember to remove duplicates, next month you forget one row. Third — and this is the killer — the work is thrown away. Nothing about hand-cleaning is saved; the next file starts from zero." },
          { type: "p", text: "Power Query inverts that. You clean the first file by clicking through menus, and Excel writes down every action in order. That written recipe is called a query. Next month you point the same query at the new file, press Refresh, and the whole cleanup replays in two or three seconds. You do the thinking once; the machine does the repetition." },
          { type: "ul", items: [
            "It is free and built in: on the Data tab, look for the Get & Transform group. In Excel 2016 it is called Get & Transform; in Microsoft 365 the group is simply Get Data.",
            "Users on Excel 2010 or 2013 can install the free \"Microsoft Power Query for Excel\" add-in from Microsoft's site and follow along with almost everything in this lesson.",
            "It never modifies your source file. All cleaning happens on a copy inside Excel; the original export stays untouched on disk.",
            "The result loads into a normal Excel table, so everything you already know — filters, pivots, lookups — works on it as before."
          ] }
        ]
      },
      {
        heading: "Getting data in: from files, workbooks and folders",
        blocks: [
          { type: "p", text: "Everything starts on the Data tab with Get Data. The three sources you will use weekly are files (Text/CSV), other Excel workbooks, and folders full of files." },
          { type: "steps", items: [
            "From Text/CSV: Data > Get Data > From File > From Text/CSV. Pick the file; a preview window opens showing how the columns were detected. Click Transform Data to enter the editor, or Load to accept it as is.",
            "From Workbook: Data > Get Data > From File > From Workbook. Pick the workbook, then choose the sheet or table you want from the navigator list.",
            "From Folder: Data > Get Data > From File > From Folder. Point at a folder and Power Query combines every file inside it into one table — ideal when each branch or month sends its own export.",
            "In the preview or navigator, always choose Transform Data rather than Load when cleaning is needed. That opens the Power Query Editor, where the real work happens."
          ] },
          { type: "p", text: "For this lesson, use the practice download: a deliberately messy Sunrise Traders export with extra spaces, numbers stored as text, dates as text, and duplicated rows. Open the workbook first to see the mess with your own eyes, then import it: Data > Get Data > From File > From Workbook, select the messy sheet, and click Transform Data." },
          { type: "warning", text: "When you import a CSV, Power Query guesses data types from the first rows. If the first 200 rows of an amount column are clean numbers and row 201 contains \"1 250\", the column may be typed as number and the odd row becomes an error. Always glance at the type icons in each column header before you trust the load." }
        ]
      },
      {
        heading: "The Power Query Editor and the Applied Steps pane",
        blocks: [
          { type: "p", text: "The editor looks like a stripped-down Excel: your data in the middle, a ribbon of transforms on top, and on the right side a pane called Query Settings with a list named Applied Steps. That list is the heart of the whole tool." },
          { type: "p", text: "Every click you make — removed a column, changed a type, replaced a value — is appended to Applied Steps as a named step. Click any earlier step and the data preview rewinds to how it looked at that moment; click the last step to return to the present. This gives you two superpowers manual cleaning never had:" },
          { type: "ul", items: [
            "Undo without Ctrl+Z: click the step above the mistake, delete the bad step (click the X next to it), and continue. Steps after it re-evaluate on the corrected data.",
            "A visible audit trail: a colleague can open the query, read the step names — Removed Columns, Changed Type, Removed Duplicates — and understand the entire cleanup without asking you anything."
          ] },
          { type: "tip", text: "Rename a step to make the audit trail honest: right-click it > Rename, and turn \"Replaced Value1\" into \"Trimmed product names\". Six months later, that sentence will save someone an hour." }
        ]
      },
      {
        heading: "The everyday transforms",
        blocks: [
          { type: "p", text: "You will use perhaps eight transforms for 95% of cleanup work. All of them act on the column you select first (click the column header; Ctrl+click for several). Here is the full toolkit applied to the messy practice export:" },
          { type: "table", headers: ["Transform", "How", "Fixes"], rows: [
            ["Remove columns", "Right-click column header > Remove", "Export padding like row numbers, timestamps, internal codes"],
            ["Change type", "Type icon in the header (123/ABC) > pick Date, Number...", "Dates that arrived as text, amounts stored as text"],
            ["Split column", "Transform > Split Column > By Delimiter", "\"NOTE-A4-48PK\" into product family and size"],
            ["Replace values", "Right-click column > Replace Values", "\"N/A\" into empty, \"INR \" prefixes removed"],
            ["Remove duplicates", "Right-click column > Remove Duplicates", "Rows exported twice from the source system"],
            ["Trim", "Transform > Format > Trim", "Leading/trailing spaces that break lookups: \" Pens\" ≠ \"Pens\""],
            ["Remove blank rows", "Home > Remove Rows > Remove Blank Rows", "Spacer rows from reports printed to CSV"],
            ["Unpivot columns", "Transform > Unpivot Columns", "A month-per-column report turned into a proper list"]
          ] },
          { type: "p", text: "Two of these deserve special attention. Trim: text exports love to hide a space before values, and a lookup against \" Pens\" fails with #N/A even though the eye sees a perfect match. Trim removes the padding from every selected column in one action. Unpivot: the moment a report has one column per month (Jan, Feb, Mar across the top), it is a cross-tab, and no pivot or lookup works properly on it. Select the columns that are not months, right-click the month columns, choose Unpivot, and the cross-tab becomes a tidy three-column list: item, month, value." },
          { type: "code", lang: "excel", text: "Cross-tab (before unpivot):          After unpivot:\n\nProduct   | Jan | Feb                 Product   | Month | Sales\nNotebooks | 120 | 125                 Notebooks | Jan   | 120\nPens      |  35 |  38                 Notebooks | Feb   | 125\n                                    Pens      | Jan   | 35\n                                    Pens      | Feb   | 38" },
          { type: "p", text: "The unpivoted shape is what the Data Model and Power Pivot (next lesson) expect. Cross-tabs are for humans to read; lists are for Excel to compute. Power Query converts between the two in one click — manually that conversion is copy-paste surgery." }
        ]
      },
      {
        heading: "Close & Load, and the one-click refresh",
        blocks: [
          { type: "steps", items: [
            "Finish your transforms and check the last Applied Step's preview looks right.",
            "Home tab > Close & Load > Close & Load To... Choose Table, and New worksheet (or Existing worksheet).",
            "The cleaned data lands as a normal Excel table with the query's name. Everything downstream — pivots, formulas — now works on it.",
            "Next month: save the new export in the same location (or drop it in the connected folder), then Data > Refresh All (or right-click the table > Refresh). Every applied step replays automatically.",
            "If a column arrives with unexpected values, reopen the editor (double-click the query in the Queries & Connections pane), add or adjust a step, and Close & Load again."
          ] },
          { type: "p", text: "That final click is the entire payoff of the tool. A cleanup that took an hour by hand now takes the time to drop a file in a folder and press Refresh All. And because the steps are saved in the workbook, whoever opens it next month inherits the whole process — the knowledge stops living only in your head." },
          { type: "warning", text: "When to choose Power Query over doing it by hand: the cleanup will recur (monthly exports), the file is large (tens of thousands of rows), several files need combining, or the same cleanup must be shared across a team. When not to: a one-off 20-row fix, or data that arrives already clean. Power Query has a small learning cost — spend it on jobs that repeat." },
          { type: "warning", text: "Refresh fails if the source file has been moved, renamed or deleted — the error names the path it could not find. Keep source files in a stable folder (or the same folder as the workbook for relative paths), and tell your team never to rename the export file." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Editing the loaded table on the sheet (deleting rows, retyping values) instead of editing the query.", right: "Double-click the query in Queries & Connections and fix the step in the editor, then Close & Load.", why: "Any hand-edit is wiped on the next Refresh, because the table is rebuilt from the source through the steps. If a value is wrong, fix the data or add a Replace Values step — never the loaded result." },
      { wrong: "Trusting the automatic type detection and loading straight away.", right: "Check the 123/ABC type icon on every column in the editor before loading; set dates and amounts explicitly.", why: "Dates stored as text sort wrongly and fail date arithmetic; amounts stored as text break SUM. Fixing types in the query means every future refresh is correct too." },
      { wrong: "Cleaning the original export file manually \"just this once\" and then importing it.", right: "Import the untouched original and do all cleaning inside Power Query.", why: "The query must replay on next month's original file. If the process depends on you having hand-fixed the source first, the refresh produces the old mess." },
      { wrong: "Removing duplicates on a table where some rows differ only by a hidden trailing space.", right: "Trim text columns first, then remove duplicates.", why: "\"Staplers \" and \"Staplers\" are different text values, so the duplicate survives. Trim (and, if needed, Clean for line breaks) before the duplicate check." }
    ],
    shortcuts: [
      { keys: "Ctrl + click", action: "Select multiple columns in the editor before applying one transform to all of them" },
      { keys: "Ctrl + Alt + F5", action: "Refresh the selected query only (Data > Refresh All refreshes every query)" },
      { keys: "Double-click", action: "A query in the Queries & Connections pane reopens the editor at any time" }
    ],
    exercises: [
      { task: "Open the practice file assets/files/level6/power-query-messy-data.xlsx and study the messy export sheet. List at least four problems you can see (types, spaces, duplicates, extra columns).", hint: "Look at the type icons, the alignment of text vs numbers, and whether any two rows repeat exactly.", solution: "Typical problems: dates/amounts stored as text, leading and trailing spaces in product and salesperson names, duplicated rows, and one or more useless columns (row number or timestamp) that can be removed." },
      { task: "Import the messy workbook with Data > Get Data > From File > From Workbook > Transform Data. Then: remove the useless column, trim all text columns, fix the date and amount types, and remove duplicate rows. Load the result as a table in a new worksheet.", hint: "One transform per selection, checking the Applied Steps pane after each. Trim lives under Transform > Format; duplicates under right-click on any column.", solution: "Applied Steps should end with something like: Removed Columns, Trimmed Text, Changed Type, Removed Duplicates. After Close & Load, the table sorts and SUMs correctly — proof the types are right." },
      { task: "After loading, change one value in the original messy source file (add a space before a product name and re-save). Refresh the query. What happened to your hand-edits if you made any on the loaded table?", hint: "Refresh replays the steps against the source. Anything typed onto the loaded table is not a step.", solution: "The table rebuilds from the source: hand-edits vanish, and the query's steps re-apply (the extra space you added is removed by the Trim step). This demonstrates why fixes belong in the query, not the table." },
      { task: "In the editor, build a small unpivot: select the month columns of any cross-tab-shaped block and unpivot them into a Product / Month / Sales list, then load it to a new sheet.", hint: "Select the month columns, right-click a selected header > Unpivot Columns. Rename the two new columns afterwards by double-clicking their headers.", solution: "The cross-tab becomes a long list with one row per product-month combination — the shape Power Pivot expects in the next lesson." }
    ],
    quiz: [
      { q: "What is an Applied Step in Power Query?", options: ["A formula typed into a cell", "A recorded cleaning action that replays on every refresh", "A copy of the source file", "A keyboard shortcut for the editor"], answer: 1, explanation: "Each click in the editor is saved as a step in the Applied Steps list. On refresh, all steps replay in order against the new source data." },
      { q: "Next month's export has arrived. What do you do to clean it?", options: ["Repeat all the cleaning manually", "Save it in the same location and press Data > Refresh All", "Rebuild the query from scratch", "Copy the cleaned table from last month over it"], answer: 1, explanation: "The query is attached to the source path. Drop the new file in place and refresh — every step replays automatically. That is the entire point of Power Query." },
      { q: "A lookup against the loaded table fails even though the values look identical. What is the most likely cause?", options: ["The lookup formula has too many brackets", "Hidden leading/trailing spaces in the text — fixed by a Trim step", "The workbook needs saving as .xlsm", "Excel's calculation is set to manual"], answer: 1, explanation: "Text exports often carry invisible padding, so \"Pens\" and \"Pens \" are different values. A Trim step in the query fixes it for every future refresh as well." },
      { q: "What does Unpivot do?", options: ["Deletes the selected columns", "Turns a month-per-column cross-tab into a long list of item / month / value rows", "Rotates the worksheet 90 degrees", "Splits one column into several"], answer: 1, explanation: "Unpivot converts wide cross-tab layouts into tidy lists — the shape that pivots, the Data Model and DAX all require. Split Column is the tool that divides one column into several." },
      { q: "Where can you get Power Query?", options: ["Buy it as part of Office Professional Plus", "It is built into Excel 2016 and later for free (a free add-in exists for 2010/2013)", "Only with a Microsoft 365 subscription", "Download it from the Power BI website for a fee"], answer: 1, explanation: "Excel 2016 onward includes it on the Data tab (Get & Transform / Get Data) at no cost. Excel 2010/2013 users can install Microsoft's free Power Query add-in and follow the same workflow." }
    ],
    summary: [
      "Power Query records your cleaning as a replayable query attached to a source file — free and built into Excel 2016 and later.",
      "Import with Data > Get Data (From Text/CSV, From Workbook, From Folder); always choose Transform Data when cleanup is needed.",
      "The Applied Steps pane is both an undo system and an audit trail; rename steps to document the cleanup.",
      "Core transforms: remove columns, change type, split, replace values, remove duplicates, trim, remove blank rows, and unpivot cross-tabs.",
      "Close & Load to a table, then refresh with one click each month. Fix problems in the query — never by editing the loaded table."
    ],
    downloads: [
      { file: "assets/files/level6/power-query-messy-data.xlsx", label: "Messy data for Power Query", note: "A deliberately messy export (extra spaces, wrong types, duplicated rows) to clean with Power Query" }
    ],
    keywords: ["Power Query", "Get Data", "Get & Transform", "Applied Steps", "remove duplicates", "trim spaces", "change data type", "unpivot", "Close and Load", "refresh query", "clean data"]
  },
  {
    id: "l6-power-pivot",
    title: "Power Pivot and the Data Model",
    difficulty: "Advanced",
    duration: 22,
    version: "Excel 2016+",
    objectives: [
      "Enable the free Power Pivot add-in in Excel 2016 and later",
      "Import tables into the Data Model",
      "Create one-to-many relationships between tables",
      "Explain why relationships replace VLOOKUP across large data",
      "Build a PivotTable directly from the Data Model"
    ],
    intro: "VLOOKUP is brilliant for one table. But real businesses do not have one table — they have a Sales table with product codes, a Products table with names and prices, a Regions table with cities and managers, and they keep growing. The classic fix is to VLOOKUP the product name into the sales rows, the region manager into the sales rows, the category into the sales rows... until the file has fifteen helper columns and takes a minute to open.\n\nPower Pivot, included free in Excel 2016 and later, offers a better idea: keep each table as it is, and tell Excel how the tables relate. Sales rows link to Products by ProductID; Products link to nothing else; Regions link by RegionID. Excel then treats the connected tables as one model, and a single PivotTable can combine fields from all of them — no helper columns, no VLOOKUP, millions of rows if needed.",
    sections: [
      {
        heading: "Turning Power Pivot on",
        blocks: [
          { type: "p", text: "Power Pivot ships with Excel but its tab is hidden by default in most installations. Enabling it takes one minute and you only ever do it once per computer." },
          { type: "steps", items: [
            "File > Options > Add-ins.",
            "At the bottom, in the Manage box, choose COM Add-ins and click Go.",
            "Tick Microsoft Power Pivot for Excel and click OK.",
            "A new Power Pivot tab appears on the ribbon with buttons like Manage, Measures and Add to Data Model."
          ] },
          { type: "p", text: "In current Microsoft 365 installations the Data Model is often reachable without the add-in (look for \"Add this data to the Data Model\" checkboxes), but enabling the tab is still worth it: the Power Pivot window is where you see and manage everything in the model. There is no separate licence or fee — it is part of Excel." },
          { type: "tip", text: "Not available at all? The standalone Excel 2016 \"Professional Plus\"-style editions all include it; only some basic OEM bundles omit it. If the add-in list is empty on a work machine, ask IT — it is a two-tick job for them." }
        ]
      },
      {
        heading: "Getting tables into the Data Model",
        blocks: [
          { type: "p", text: "The Data Model is a hidden storage area inside the workbook where Excel keeps tables in a compressed, fast form. Tables reach it by three roads, and you will use all three over time:" },
          { type: "ul", items: [
            "From Power Query: in Close & Load To..., choose \"Add this data to the Data Model\" (with or without also loading a table to the sheet). This is the cleanest route and connects directly to the previous lesson.",
            "From a normal sheet: Insert > PivotTable, tick \"Add this data to the Data Model\" in the dialog. Or select the table and use Power Pivot > Add to Data Model.",
            "From the Power Pivot window: Power Pivot > Manage opens the model window, where Get External Data lets you import straight from databases, feeds or files without ever touching a worksheet."
          ] },
          { type: "p", text: "For this lesson, open the practice workbook: it contains three clean tables — Sales (one row per sale: Date, ProductID, RegionID, Salesperson, Amount), Products (ProductID, Product, Category, Unit Price) and Regions (RegionID, Region, City, Manager). Add each to the Data Model: click inside the Sales table, Power Pivot > Add to Data Model; repeat for Products and Regions. The Power Pivot window now shows all three as separate tabs — you can switch between them exactly like worksheet tabs." },
          { type: "warning", text: "Give every table a real name before importing (via Table Design > Table Name). The model uses these names everywhere — in the field list, in DAX next lesson — and \"Table1\" tells nobody anything." }
        ]
      },
      {
        heading: "Relationships: telling Excel how the tables connect",
        blocks: [
          { type: "p", text: "Tables in the model are strangers until you define a relationship. The Sales table stores only ProductID — it does not know that P104 is a Laptop or that it costs ₹52,000. The Products table holds that knowledge, one row per product. Sales has many rows per product; Products has exactly one row per product. That shape is called one-to-many, and it is the standard pattern for almost every business model." },
          { type: "table", headers: ["Table", "Role in the relationship", "Key column"], rows: [
            ["Products", "The \"one\" side (each product appears once)", "ProductID"],
            ["Sales", "The \"many\" side (each product appears on many rows)", "ProductID"],
            ["Regions", "The \"one\" side for the region lookup", "RegionID"]
          ] },
          { type: "steps", items: [
            "In the Power Pivot window, open Home > View > Diagram View to see the tables as cards.",
            "Drag ProductID from the Products card and drop it onto ProductID on the Sales card. A connecting line appears — the relationship exists.",
            "Repeat by dragging RegionID from Regions onto RegionID on Sales.",
            "Click a line to see and edit the relationship (which table is which side, and active or not).",
            "Return to Data View to browse the tables again."
          ] },
          { type: "p", text: "That is the entire act of \"joining\" the data: two drags. No VLOOKUP column, no copying names across, no formulas to break. If a new sale references a ProductID that is missing from Products, the relationship simply leaves that row unmatched — it does not error, which is why keeping the dimension tables complete still matters." },
          { type: "tip", text: "Order does not matter when dragging: Excel detects which table is the one side and which is the many from the data itself. The arrow on the relationship line always points from the \"one\" table toward the \"many\" table." }
        ]
      },
      {
        heading: "Why relationships beat VLOOKUP on real data",
        blocks: [
          { type: "p", text: "The VLOOKUP-into-every-row approach has three failure modes that every analyst eventually meets. Power Pivot's relationships remove all three at once." },
          { type: "ul", items: [
            "Size: a VLOOKUP recalculates for every row of the Sales table, every time anything changes. With 50,000 sales rows and five lookups each, that is 250,000 scans per recalculation. The Data Model stores tables compressed and answers pivots from pre-built aggregations — files that crawled in formulas fly here.",
            "Integrity: with helper columns, an edited product name must be re-copied into every sales row or the report shows stale names. With a relationship, the name lives only in Products; sales rows display whatever Products currently says. Fix it once, correct everywhere.",
            "Flexibility: a helper column is baked in — category today means adding another column tomorrow for supplier, and another for price band. A model can follow any path of relationships: Sales → Products for category, Sales → Regions for manager, and pivots can use fields from every connected table at once."
          ] },
          { type: "p", text: "There is a mental shift here worth naming. VLOOKUP copies knowledge into the fact table; relationships let Excel follow pointers at calculation time. The first scales badly and breaks quietly. The second scales to millions of rows and stays correct because the truth is stored in exactly one place." }
        ]
      },
      {
        heading: "Build a PivotTable from the model",
        blocks: [
          { type: "steps", items: [
            "On a new sheet: Insert > PivotTable > From Data Model (in the Power Pivot group: PivotTable > From Data Model).",
            "In the field list, tables appear as expandable groups: Sales, Products, Regions.",
            "Drag Products > Category into Rows, Regions > Region into Columns.",
            "Drag Sales > Amount into Values. Excel sums it per category and region — combining three tables in one pivot with zero helper columns.",
            "Add Sales > Date into Filters or a slicer to slice by month. Drag Salesperson into Rows below Category to see each person's category mix."
          ] },
          { type: "grid", title: "What the model pivot gives you — fields from three tables, one sheet", rows: [
            ["Sum of Amount", "North", "South", "West", "Total"],
            ["Laptops", "215000", "168000", "142000", "525000"],
            ["Printers", "86000", "64000", "51000", "201000"],
            ["Stationery", "41000", "38000", "29000", "108000"]
          ] },
          { type: "p", text: "Notice what you did not do: no VLOOKUP of category onto sales rows, no VLOOKUP of region names, no concatenations. The pivot reads Amount from Sales, Category from Products and Region from Regions, following the two relationships you drew. In plain-Excel terms you replaced two columns of VLOOKUPs and one source-file dependency with two lines in a diagram." },
          { type: "p", text: "One habit to unlearn: the classic pivot field list shows one flat list of every column. The model field list groups columns under their tables — Sales fields under Sales, and so on. Read the grouping as the diagram: fields come from tables, tables are wired by keys. Once that clicks, the rest of Power Pivot — including DAX measures in the next lesson — is a natural continuation." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Trying to relate Sales.ProductID to Products.ProductName.", right: "Relate on the key columns both tables actually share: ProductID to ProductID.", why: "A relationship needs a column whose values match exactly on both sides. Keys (IDs) exist for that purpose; names may repeat, drift or be misspelled, and the relationship would silently fail to match." },
      { wrong: "Leaving duplicate values in the \"one\" side's key column.", right: "Products must have exactly one row per ProductID — run Remove Duplicates in Power Query before import.", why: "The one-to-many logic assumes one definition per product. Two rows for P104 makes aggregations ambiguous and pivot totals confusing." },
      { wrong: "Building relationships on text keys with invisible differences (\"P104 \" vs \"P104\").", right: "Trim keys in Power Query first, and prefer short, clean ID columns.", why: "Relationship matching is exact text comparison. A single padded space stops the join for every affected row, and unmatched rows just drop out of results with no error." },
      { wrong: "Adding tables to the model but still writing VLOOKUP helper columns in the Sales sheet.", right: "Use the model field list in the pivot and take fields straight from Products and Regions.", why: "Mixing both approaches keeps the maintenance burden you were trying to escape. The whole point is that the pivot reaches across tables by itself." }
    ],
    shortcuts: [
      { keys: "Alt + F5", action: "Refresh the Data Model pivot after the source tables change" },
      { keys: "Drag and drop", action: "In Diagram View, dragging a key column onto its partner creates the relationship" },
      { keys: "Right-click", action: "A relationship line to edit, deactivate or delete it" }
    ],
    exercises: [
      { task: "Enable the Power Pivot add-in on your machine (File > Options > Add-ins > COM Add-ins > Go) and confirm the Power Pivot tab appears on the ribbon.", hint: "If Excel is a Microsoft 365 subscription, the add-in should already be listed; just tick it. Restart Excel if the tab does not appear.", solution: "The ribbon gains a Power Pivot tab with Manage, Measures, Add to Data Model and other buttons." },
      { task: "Open assets/files/level6/power-pivot-data.xlsx and add all three tables (Sales, Products, Regions) to the Data Model. Verify in the Power Pivot window that three tabs appear.", hint: "Click inside each table, then Power Pivot > Add to Data Model. In the Manage window, each table is a tab like a worksheet.", solution: "The Power Pivot window shows Sales, Products and Regions as separate tabs with all columns preserved." },
      { task: "In Diagram View, create both relationships: Products[ProductID] to Sales[ProductID], and Regions[RegionID] to Sales[RegionID].", hint: "Home > View > Diagram View, then drag the key from the \"one\" table onto the same-named key on Sales.", solution: "Two lines connect the tables, arrows pointing from Products and Regions toward Sales (the many side)." },
      { task: "Insert a PivotTable From Data Model on a new sheet. Put Products[Category] in Rows, Regions[Region] in Columns and Sales[Amount] in Values. Then add Sales[Salesperson] under Category in Rows and slice with Regions[City].", hint: "Insert > PivotTable > From Data Model. The field list groups fields under table names — pick fields from all three tables.", solution: "One pivot shows revenue by category and region and can break down by salesperson, using all three tables with no VLOOKUP helper columns anywhere." }
    ],
    quiz: [
      { q: "What problem does the Data Model solve that plain VLOOKUP handles badly?", options: ["Formatting numbers with the right currency symbol", "Combining several related tables without helper columns, at scale", "Protecting sheets with a password", "Making pivot tables load faster by disabling refresh"], answer: 1, explanation: "Relationships let pivots reach across tables (Sales to Products to Regions) directly. No helper columns, no per-row lookups, and it scales to very large tables." },
      { q: "Sales has 40,000 rows; Products has 12 rows. Which relationship shape is this?", options: ["One-to-one", "Many-to-many", "One-to-many (Products is the one side)", "It cannot be related"], answer: 2, explanation: "Each product appears exactly once in Products (the \"one\" side) but on many Sales rows (the \"many\" side). This is the standard fact-plus-lookup pattern." },
      { q: "How do you create a relationship between two tables in the model?", options: ["Type a VLOOKUP formula into the Sales table", "Drag one table's key column onto the other's key in Diagram View", "Copy both tables into one sheet", "Save the workbook as .xlsm"], answer: 1, explanation: "Diagram View makes it a drag-and-drop: drop ProductID onto ProductID and the relationship line appears. No formulas involved." },
      { q: "A pivot built From Data Model shows fields grouped under Sales, Products and Regions. What does that grouping reflect?", options: ["The worksheet order", "The table structure of the Data Model and its relationships", "Alphabetical sorting of columns", "The order you added the tables"], answer: 1, explanation: "The model field list mirrors the model: each table is a group of its own fields, and pivots can combine fields across groups because the relationships connect the tables." },
      { q: "What must be true of the columns used to relate two tables?", options: ["They must have the same header text", "They must contain matching values on both sides — usually IDs", "They must both be date columns", "They must be on the first worksheet"], answer: 1, explanation: "Relationships match values, not headers. Keys like ProductID exist precisely so both tables can refer to the same product in exactly the same way." }
    ],
    summary: [
      "Power Pivot is free and built into Excel 2016+; enable the add-in once via File > Options > Add-ins > COM Add-ins.",
      "Tables enter the Data Model via Power Query's \"Add to the Data Model\" option, the PivotTable dialog checkbox, or the Power Pivot window.",
      "Relationships are drag-and-drop in Diagram View; the standard pattern is one-to-many from lookup tables (Products, Regions) to the fact table (Sales).",
      "Relationships replace per-row VLOOKUPs: better speed on large data, single-source-of-truth names, and pivots that combine fields from every connected table.",
      "Build pivots with Insert > PivotTable > From Data Model and take fields from all tables — then continue with DAX measures in the next lesson."
    ],
    downloads: [
      { file: "assets/files/level6/power-pivot-data.xlsx", label: "Power Pivot data model tables", note: "Three linked tables — Sales, Products and Regions — ready for the data model and DAX practice" }
    ],
    keywords: ["Power Pivot", "Data Model", "relationships", "one-to-many", "diagram view", "PivotTable from data model", "replace VLOOKUP", "add to data model", "COM add-in", "product ID relationship"]
  }
);
