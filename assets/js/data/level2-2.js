// Excel Lab — Level 2 (Formulas and Functions), part 2 of 2
// Lessons 8-14 only. Metadata and lessons 1-7 live in level2-1.js.

window.LEVEL_LESSONS = window.LEVEL_LESSONS || {};

(window.LEVEL_LESSONS[2] = window.LEVEL_LESSONS[2] || []).push(
  {
    id: "l2-sumifs-countifs",
    title: "SUMIFS, COUNTIFS and AVERAGEIFS",
    difficulty: "Intermediate",
    duration: 16,
    version: "",
    objectives: [
      "Write SUMIFS with the sum range first and criteria pairs after it",
      "Build COUNTIFS and AVERAGEIFS with two or more conditions",
      "Filter by a date range using >= and <= joined to cell references",
      "Combine region, month and product conditions in one formula"
    ],
    keyIdea: "SUMIFS and COUNTIFS handle two or more conditions: sum range first, then range-and-criteria pairs.",
    intro: "SUMIF and COUNTIF handle one condition. But managers usually ask with two or three: \"What did the **North** region sell in **March**?\" \"How many **Stapler** orders came from the **East**?\"\n\nFor those you need the plural versions: SUMIFS, COUNTIFS and AVERAGEIFS. Remember one big difference from day one. In SUMIFS the **sum range comes first**, not last. After it, conditions travel in pairs — a range, then its criteria. Get that order right and the rest follows.",
    sections: [
      {
        heading: "SUMIFS: sum range first, then pairs",
        blocks: [
          {
            type: "formula",
            name: "SUMIFS",
            syntax: "=SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)",
            example: "=SUMIFS(E2:E10, A2:A10, \"North\", B2:B10, \"Staplers\")",
            explanation: "Adds the Amount column only for rows where the Region is North AND the Product is Staplers. Every pair after the sum range works like one SUMIF."
          },
          { type: "p", text: "Notice the changed order compared with SUMIF: the range to add leads the formula. After it, each condition is a two-part pair — the column to test, then the test itself. You can stack as many pairs as the report needs." },
          {
            type: "grid",
            title: "Sunrise Traders sales record (same table as lesson 7)",
            rows: [
              ["Region", "Product", "Units", "Price (₹)", "Amount (₹)"],
              ["North", "Notebooks", "40", "45", "=C2*D2"],
              ["South", "Pens", "100", "12", "=C3*D3"],
              ["East", "Staplers", "15", "130", "=C4*D4"],
              ["North", "Printers", "2", "8500", "=C5*D5"],
              ["West", "Laptops", "1", "42000", "=C6*D6"],
              ["South", "Notebooks", "60", "45", "=C7*D7"],
              ["East", "Pens", "80", "12", "=C8*D8"],
              ["North", "Staplers", "10", "130", "=C9*D9"],
              ["West", "Notebooks", "25", "45", "=C10*D10"]
            ]
          },
          { type: "p", text: "`=SUMIFS(E2:E10, A2:A10, \"North\", B2:B10, \"Staplers\")` walks the table row by row. North **and** Staplers is true only in row 9, so the result is 1300. Change the product pair to Notebooks and it returns 1800. Only row 2 is North and Notebook." }
        ]
      },
      {
        heading: "COUNTIFS and AVERAGEIFS: same pairs, no sum range for counting",
        blocks: [
          {
            type: "formula",
            name: "COUNTIFS",
            syntax: "=COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)",
            example: "=COUNTIFS(A2:A10, \"North\")",
            explanation: "Counts rows matching all the pairs. With one pair it behaves exactly like COUNTIF; with two or more it narrows the count."
          },
          {
            type: "formula",
            name: "AVERAGEIFS",
            syntax: "=AVERAGEIFS(average_range, criteria_range1, criteria1, ...)",
            example: "=AVERAGEIFS(E2:E10, A2:A10, \"North\")",
            explanation: "Averages (not totals) the rows matching all pairs. Like SUMIFS, the range to average comes first."
          },
          {
            type: "grid",
            title: "Three plural functions side by side",
            rows: [
              ["Question", "Formula", "Result"],
              ["North orders", "=COUNTIFS(A2:A10, \"North\")", "3"],
              ["North total", "=SUMIFS(E2:E10, A2:A10, \"North\")", "20100"],
              ["North average", "=AVERAGEIFS(E2:E10, A2:A10, \"North\")", "6700"]
            ]
          },
          { type: "p", text: "The average of 20100 across 3 orders is 6700 — the typical North invoice. Averages answer \"what is a normal order like?\" Totals answer \"how much altogether?\" Managers want both. You can now produce them with the same criteria structure." },
          { type: "warning", text: "All criteria ranges must be the same size and shape as the sum range. If one pair runs `A2:A9` while another runs `A2:A10`, SUMIFS returns the #VALUE! error. Unlike SUMIF, it refuses to guess." }
        ]
      },
      {
        heading: "Date ranges: between two dates",
        blocks: [
          { type: "p", text: "\"Total for March\" really means \"total where the date is on or after 1 March and on or before 31 March\". That is two conditions on the same date column." },
          { type: "p", text: "So you need two pairs, using >= and <=. Hard-coding dates inside formulas is fragile. Instead, join the operators to cells that hold the dates." },
          {
            type: "grid",
            title: "Date-stamped sales for February and March",
            rows: [
              ["Date", "Region", "Product", "Amount (₹)"],
              ["05-Feb", "North", "Notebooks", "1800"],
              ["07-Feb", "South", "Pens", "1200"],
              ["11-Feb", "East", "Staplers", "1950"],
              ["18-Feb", "North", "Printers", "17000"],
              ["25-Feb", "West", "Laptops", "42000"],
              ["02-Mar", "South", "Notebooks", "2700"],
              ["09-Mar", "East", "Pens", "960"],
              ["16-Mar", "North", "Staplers", "1300"],
              ["24-Mar", "West", "Notebooks", "1125"]
            ]
          },
          { type: "p", text: "Suppose H2 holds the start date 01-Mar and H3 holds the end date 31-Mar. The March total is:" },
          {
            type: "code",
            lang: "excel",
            text: "=SUMIFS(D2:D10, A2:A10, \">=\"&H2, A2:A10, \"<=\"&H3)"
          },
          { type: "p", text: "The ampersand glues the operator text to the cell. \">=\"&H2 becomes \">=01-Mar-2026\" in Excel's eyes. The formula adds rows 7 to 10: 2700 + 960 + 1300 + 1125 = 6085. Change H2 to 01-Feb and the same formula returns the February total, 63950. No editing needed." },
          { type: "tip", text: "This \">=\"&cell pattern is the single most reused trick in reporting. Keep the start and end dates in labelled cells near your report so anyone can change the period without touching a formula." }
        ]
      },
      {
        heading: "Putting it together: region, month and product",
        blocks: [
          { type: "p", text: "Real reports combine it all. \"North region, March, Staplers\" needs three pairs — one for the region, two for the date window:" },
          {
            type: "code",
            lang: "excel",
            text: "=SUMIFS(D2:D10,\n        A2:A10, \"North\",\n        B2:B10, \"Staplers\",\n        A2:A10, \">=\"&H2,\n        A2:A10, \"<=\"&H3)"
          },
          { type: "p", text: "The date column appears twice — once for the start test, once for the end test. That is perfectly normal. Each pair is an independent condition, and several pairs may point at the same column. In the table above, only one row matches: the March North Stapler order (row 8, 1300). So the formula returns 1300." },
          {
            type: "ul",
            items: [
              "**Build one pair at a time.** Start with =SUMIFS(D2:D10, A2:A10, \"North\"), check the result, then add the next pair.",
              "**Keep criteria in cells** (region in one cell, dates in two others) so the report has control boxes instead of hard-coded values.",
              "**Wrap the finished formula in IFERROR** from lesson 6 if there is any chance no row matches — an empty match is not an error here, but the habit of testing first still applies."
            ]
          }
        ]
      }
    ],
    mistakes: [
      {
        wrong: "=SUMIFS(A2:A10, \"North\", E2:E10) — borrowed order from SUMIF",
        right: "=SUMIFS(E2:E10, A2:A10, \"North\")",
        why: "SUMIFS puts the sum range first, then range-criteria pairs. Swapping them makes Excel treat the region column as the range to add."
      },
      {
        wrong: "Criteria ranges of different lengths: =SUMIFS(E2:E10, A2:A9, \"North\")",
        right: "=SUMIFS(E2:E10, A2:A10, \"North\")",
        why: "SUMIFS requires all ranges to match the sum range exactly. Mismatched sizes produce #VALUE!."
      },
      {
        wrong: "Writing the operator without the ampersand: =SUMIFS(D2:D10, A2:A10, \">=H2\")",
        right: "=SUMIFS(D2:D10, A2:A10, \">=\"&H2)",
        why: "Inside the quotes, H2 is just the text H2. The & joins the operator to the cell's real value."
      },
      {
        wrong: "Using two date pairs with the operator outside: A2:A10, >=, H2",
        right: "A2:A10, \">=\"&H2, A2:A10, \"<=\"&H3",
        why: "Each pair is exactly two arguments: a range, then a text-built condition. A bare >= breaks the pair structure."
      }
    ],
    shortcuts: [],
    exercises: [
      {
        task: "Open the sumif-practice workbook, Exercise sheet. In cell H2, total the revenue where the Region is East AND the Product is Pens.",
        hint: "Sum range first, then two range-criteria pairs.",
        solution: "H2: =SUMIFS(E2:E10, A2:A10, \"East\", B2:B10, \"Pens\"), which returns 960."
      },
      {
        task: "In cell H3, count the North orders with a total above 1000, and in H4 their average value.",
        hint: "One formula counts with two pairs; the other averages with the sum range leading.",
        solution: "H3: =COUNTIFS(A2:A10, \"North\", E2:E10, \">1000\") returns 2 (17000 and 1300). H4: =AVERAGEIFS(E2:E10, A2:A10, \"North\", E2:E10, \">1000\") returns 9150."
      },
      {
        task: "The date-stamped table on the same sheet has dates in column A, regions in B and amounts in E. With the start date in H6 and the end date in H7, build the total for that period in H8.",
        hint: "Two pairs on the same date column, operators joined with &.",
        solution: "H8: =SUMIFS(E2:E10, A2:A10, \">=\"&H6, A2:A10, \"<=\"&H7). With 01-Mar and 31-Mar it returns 6085."
      },
      {
        task: "Extend exercise 3 so it also filters by the region typed in cell H9. Compare your answer with the Solution sheet for both March and February.",
        hint: "Add one more pair: region column against the cell reference.",
        solution: "=SUMIFS(E2:E10, B2:B10, H9, A2:A10, \">=\"&H6, A2:A10, \"<=\"&H7). For March and North it returns 1300."
      }
    ],
    quiz: [
      {
        q: "What is the correct argument order in SUMIFS?",
        options: ["Sum range first, then range-criteria pairs", "Criteria pairs first, sum range last", "Exactly the same order as SUMIF", "Criteria first, then count of pairs"],
        answer: 0,
        explanation: "SUMIFS leads with the range to add, followed by pairs of (criteria range, criteria). SUMIF has the sum range last — the order swap is the classic beginner error."
      },
      {
        q: "Which formula counts North orders above 1000 (regions in A, amounts in E)?",
        options: ["=COUNTIFS(A2:A10, \"North\", E2:E10, \">1000\")", "=COUNTIF(A2:A10, \"North\") + COUNTIF(E2:E10, \">1000\")", "=COUNTIFS(E2:E10, \">1000\", A2:A10, \"North\", E2:E10)", "=SUMIFS(A2:A10, \"North\", E2:E10, \">1000\")"],
        answer: 0,
        explanation: "COUNTIFS takes pairs of range and criteria, and both conditions must hold in the same row. The plus version counts the two conditions separately and over-counts."
      },
      {
        q: "Why does the date filter need two pairs, \">=\"&H2 and \"<=\"&H3?",
        options: ["\"Between\" is two separate tests on the same column, and each pair is one test", "Excel cannot compare dates", "One pair can only hold one cell", "The second pair corrects the first"],
        answer: 0,
        explanation: "A date window is an AND of two comparisons: on or after the start and on or before the end. Two pairs on the date column do that."
      },
      {
        q: "What does SUMIFS return when no row matches all the pairs?",
        options: ["0 — an empty sum is zero, not an error", "#N/A", "#VALUE!", "An empty cell"],
        answer: 0,
        explanation: "Adding no rows gives 0. Mismatched range sizes give #VALUE!, but a valid formula with no matches is simply 0."
      },
      {
        q: "In =SUMIFS(E2:E10, A2:A10, \"North\"), why is the result the same as =SUMIF(A2:A10, \"North\", E2:E10)?",
        options: ["With one pair, SUMIFS behaves like SUMIF", "They are the same function internally", "It is a coincidence of the data", "Both ignore the criteria"],
        answer: 0,
        explanation: "One pair in SUMIFS is equivalent to SUMIF. The plural form earns its keep from the second pair onward."
      }
    ],
    summary: [
      "SUMIFS leads with the sum range, then range-criteria pairs; COUNTIFS needs only the pairs.",
      "AVERAGEIFS mirrors SUMIFS but averages the matches — same leading range, same pair structure.",
      "Date windows are two pairs on the date column: \">=\"&start and \"<=\"&end, with the dates in labelled cells.",
      "All ranges must be the same size, or SUMIFS returns #VALUE!.",
      "Build one pair at a time and verify against the visible rows — complex criteria grow correctly one step at a time."
    ],
    downloads: [
      {
        file: "assets/files/level2/sumif-practice.xlsx",
        label: "Conditional aggregation workbook",
        note: "Exercise and Solution sheets for SUMIF, COUNTIF, AVERAGEIF, SUMIFS, COUNTIFS and AVERAGEIFS"
      }
    ],
    keywords: ["SUMIFS", "COUNTIFS", "AVERAGEIFS", "multiple criteria", "date range criteria", "between two dates", "region and month totals"]
  },
  {
    id: "l2-vlookup",
    title: "VLOOKUP",
    difficulty: "Intermediate",
    duration: 16,
    version: "",
    objectives: [
      "Write =VLOOKUP(lookup_value, table_array, col_index_num, FALSE) with confidence",
      "Understand why the lookup value must sit in the first column of the table",
      "Lock the table array with $ so the formula can be filled down safely",
      "Interpret #N/A and fix the common causes",
      "Avoid the two classic mistakes: unlocked ranges and wrong column numbers"
    ],
    keyIdea: "VLOOKUP finds your value in a table's first column and returns an answer from another column.",
    intro: "You have a product name; you need its price. You have an employee code; you need the department. VLOOKUP is the classic answer: find this in a list and bring back that.\n\nIt is one of the most demanded office skills, and interviewers still ask about it. VLOOKUP searches the first column of a table for your value. Then it returns whatever sits in a chosen column of the same row. VLOOKUP is strict — three small details decide whether it works.",
    sections: [
      {
        heading: "The four arguments",
        blocks: [
          {
            type: "formula",
            name: "VLOOKUP",
            syntax: "=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])",
            example: "=VLOOKUP(\"Staplers\", G2:H6, 2, FALSE)",
            explanation: "Searches the first column of G2:H6 for the text Staplers, then returns the value from column 2 of the table — the price. FALSE means exact match only."
          },
          {
            type: "ul",
            items: [
              "**lookup_value** — what you are searching for: a product name, an ID, a code.",
              "**table_array** — the table to search, selected from its first column through the last column you might return from.",
              "**col_index_num** — which column of the table holds the answer. Column 1 is the table's own left edge, not column A of the sheet.",
              "**FALSE** — demand an exact match. Almost every office task needs FALSE."
            ]
          },
          {
            type: "grid",
            title: "Price list (columns G and H)",
            rows: [
              ["Product", "Price (₹)"],
              ["Notebooks", "45"],
              ["Pens", "12"],
              ["Staplers", "130"],
              ["Printers", "8500"],
              ["Laptops", "42000"]
            ]
          },
          { type: "p", text: "`=VLOOKUP(\"Staplers\", G2:H6, 2, FALSE)` returns 130. Excel finds Staplers in the first column and walks across to column 2 — the price column. The `col_index_num` counts from the left edge of the selected table. Here column G is 1 and column H is 2." }
        ]
      },
      {
        heading: "The rule that causes the most pain",
        blocks: [
          { type: "p", text: "VLOOKUP can only look to its **right**. The value you search for must be in the first column of your table. The answer must sit in a column further right. There is no way to make plain VLOOKUP return something to the left." },
          {
            type: "ul",
            items: [
              "Works: product name in the left column of the price list, price to its right.",
              "Works: employee ID in the left column, department to its right.",
              "Fails: you have the department and want the ID, and ID sits left of department. VLOOKUP returns #N/A because it can never match against a middle column.",
              "The workaround for looking left is INDEX and MATCH (lesson 11) or XLOOKUP (lesson 10)."
            ]
          },
          { type: "tip", text: "When a VLOOKUP keeps returning #N/A for a value you can clearly see in the table, check this rule first: is the lookup value really in the FIRST column of the table_array you selected?" }
        ]
      },
      {
        heading: "Locking the table with $",
        blocks: [
          { type: "p", text: "The price list lives in one place, but you want to fill the formula down a whole invoice. Fill-down shifts every reference one row at a time — including your table. The table then slides off the data. The fix is the dollar sign, which pins a reference in place." },
          {
            type: "code",
            lang: "excel",
            text: "Unlocked (breaks when filled down):\n=VLOOKUP(B2, G2:H6, 2, FALSE)\n\nLocked (safe to fill down):\n=VLOOKUP(B2, $G$2:$H$6, 2, FALSE)"
          },
          { type: "p", text: "`$G$2:$H$6` means \"always G2:H6, no matter where the formula is copied\". The lookup value `B2` has no dollars, so it correctly becomes `B3`, `B4`, `B5` as the formula fills down. One part moves, one part stays. That is exactly what you want." },
          {
            type: "grid",
            title: "Invoice with the locked lookup",
            rows: [
              ["Item", "Qty", "Price (₹)", "Amount (₹)"],
              ["Notebooks", "40", "=VLOOKUP(A2, $G$2:$H$6, 2, FALSE)", "=B2*C2"],
              ["Pens", "100", "=VLOOKUP(A3, $G$2:$H$6, 2, FALSE)", "=B3*C3"],
              ["Staplers", "15", "=VLOOKUP(A4, $G$2:$H$6, 2, FALSE)", "=B4*C4"]
            ]
          },
          { type: "p", text: "The price column pulls 45, 12 and 130 from the price list. Type the formula once in C2 with the locked table, press Enter, then drag the fill handle down — the table reference survives every copy." },
          { type: "tip", text: "Instead of typing the dollars, select the table reference in the formula and press [[F4]] — Excel cycles G2:H6, $G$2:$H$6, G$2:H$6, $G$2:$H$6. For a fixed lookup table you want the fully locked version." }
        ]
      },
      {
        heading: "Reading and fixing #N/A",
        blocks: [
          { type: "p", text: "#N/A from VLOOKUP means exactly one thing: the lookup value was not found in the first column of the table. The usual causes, in order of how often they bite:" },
          {
            type: "ul",
            items: [
              "**Extra spaces.** \"Staplers \" with a trailing space does not equal \"Staplers\". Wrap both sides in TRIM or clean the source data (lesson 12 teaches TRIM).",
              "**Spelling differences.** \"Stapler\" is not \"Staplers\". Excel matches text character for character.",
              "**The value is in a different column.** You selected a table whose first column is not the one holding the lookup values.",
              "**The value genuinely does not exist** — a discontinued product, a new code not yet added to the list."
            ]
          },
          { type: "warning", text: "This is where IFERROR from lesson 6 earns its keep, but only after you have verified the data: =IFERROR(VLOOKUP(A2, $G$2:$H$6, 2, FALSE), \"Not in price list\"). A friendly message beats a red #N/A in a customer-facing invoice." }
        ]
      }
    ],
    mistakes: [
      {
        wrong: "=VLOOKUP(B2, G2:H6, 2, FALSE) filled down without dollars",
        right: "=VLOOKUP(B2, $G$2:$H$6, 2, FALSE)",
        why: "The unlocked table slides down with the fill: row 5's formula searches G5:H9, which is mostly empty, so rows return #N/A or wrong prices."
      },
      {
        wrong: "Counting col_index_num from column A of the sheet",
        right: "Count from the first column of the table_array you selected",
        why: "In a table starting at G, column G is column 1. Counting from A of the sheet gives a column number far too large and returns values from the wrong place."
      },
      {
        wrong: "=VLOOKUP(A2, G2:H6, 2) — leaving out the fourth argument",
        right: "=VLOOKUP(A2, $G$2:$H$6, 2, FALSE)",
        why: "Without FALSE, VLOOKUP does an approximate match that assumes sorted data. On unsorted lists it silently returns the nearest wrong value — the most dangerous mistake of all."
      },
      {
        wrong: "Selecting only the lookup column: =VLOOKUP(B2, G2:G6, 2, FALSE)",
        right: "Select the lookup column through the return column: $G$2:$H$6",
        why: "Column 2 must exist inside the selected table. A one-column table has no column 2, so the formula returns #REF!."
      }
    ],
    shortcuts: [
      { keys: "F4", action: "Cycle a selected reference through $G$2, G$2, $G2 and back — press once to fully lock a lookup table" }
    ],
    exercises: [
      {
        task: "Open the lookup-practice workbook, Exercise sheet. It has a product price list and an invoice area. In the invoice's price column, pull the price for the first item using VLOOKUP with an exact match.",
        hint: "Lookup value, locked table, column 2, FALSE.",
        solution: "=VLOOKUP(A2, $G$2:$H$6, 2, FALSE), adjusted to your sheet's actual price-list range."
      },
      {
        task: "Fill the price formula down the invoice so every item picks up its price automatically.",
        hint: "Lock the table before you fill; press [[F4]] to add the dollars.",
        solution: "The locked table $G$2:$H$6 stays put while each row's item name changes, so every row returns its correct price."
      },
      {
        task: "Type a product name that does not exist in the price list into the invoice and observe the error. Then wrap the formula in IFERROR with the message \"Not in price list\".",
        hint: "First verify the #N/A is a genuinely missing product, not a typo.",
        solution: "=IFERROR(VLOOKUP(A2, $G$2:$H$6, 2, FALSE), \"Not in price list\")"
      },
      {
        task: "On the Solution sheet, the employee table has IDs in the first column and departments to the right. Build a formula that returns the department for the ID typed into a cell.",
        hint: "Same four arguments, different table and column count.",
        solution: "=VLOOKUP(ID, $A$2:$C$6, 3, FALSE) if the departments sit in the table's third column."
      }
    ],
    quiz: [
      {
        q: "In =VLOOKUP(B2, $G$2:$H$6, 2, FALSE), what does the 2 mean?",
        options: ["Return the value from the 2nd column of the selected table", "Look in row 2", "Return 2 results", "The second argument of the formula"],
        answer: 0,
        explanation: "col_index_num counts columns inside the table_array: G is 1, H is 2. The 2 makes VLOOKUP return the price column."
      },
      {
        q: "What does the FALSE argument do?",
        options: ["Requires an exact match of the lookup value", "Stops the formula from showing errors", "Makes the search faster", "Returns zero when not found"],
        answer: 0,
        explanation: "FALSE is the exact-match switch. Without it, VLOOKUP does an approximate match that assumes sorted data and can silently return wrong values."
      },
      {
        q: "Why must the lookup value be in the first column of table_array?",
        options: ["VLOOKUP only searches the first column and can only return values to the right", "Excel sorts the table automatically", "It saves memory", "Otherwise the formula returns zero"],
        answer: 0,
        explanation: "This is VLOOKUP's built-in limitation: match on the left edge, return from the right. To look left you need XLOOKUP or INDEX/MATCH."
      },
      {
        q: "A filled-down VLOOKUP returns #N/A from row 5 onward. What is the most likely cause?",
        options: ["The table array was not locked with $, so it slid off the data", "The FALSE argument is missing", "The column number is too small", "Excel is offline"],
        answer: 0,
        explanation: "Without $G$2:$H$6 the table reference moves down with each row. Lock it with [[F4]] and fill again."
      },
      {
        q: "A VLOOKUP shows #N/A even though the value is visible in the table. Which cause should you check FIRST?",
        options: ["Extra spaces or small spelling differences in the lookup value", "The workbook is not saved", "The table has too many rows", "The column number is 2"],
        answer: 0,
        explanation: "Matching is character by character: \"Staplers \" with a trailing space is a different text from \"Staplers\". TRIM or retype the value."
      }
    ],
    summary: [
      "VLOOKUP searches the first column of a table and returns a value from a numbered column of the same row.",
      "Always end with FALSE for exact match; omitting it triggers approximate matching on unsorted data.",
      "Lock the table array with $ (press [[F4]]) before filling the formula down.",
      "The lookup value must be in the table's first column; VLOOKUP cannot look left.",
      "#N/A means not found — check spaces, spelling and column position, then wrap in IFERROR for a friendly message."
    ],
    downloads: [
      {
        file: "assets/files/level2/lookup-practice.xlsx",
        label: "Lookup practice workbook",
        note: "Exercise and Solution sheets with product and employee tables for VLOOKUP, XLOOKUP and INDEX/MATCH"
      }
    ],
    keywords: ["VLOOKUP", "exact match", "FALSE", "table array", "column index", "F4 lock", "dollar signs", "N/A error", "price lookup"]
  },
  {
    id: "l2-xlookup",
    title: "HLOOKUP and XLOOKUP",
    difficulty: "Intermediate",
    duration: 16,
    version: "Microsoft 365",
    objectives: [
      "Use HLOOKUP for horizontal tables where the labels run across a row",
      "Write =XLOOKUP(lookup, lookup_array, return_array, [if_not_found]) in full",
      "Look left and right without rearranging your tables",
      "Explain when XLOOKUP is the better choice over VLOOKUP"
    ],
    keyIdea: "XLOOKUP searches one range and returns from another — any direction, with a built-in not-found message.",
    intro: "VLOOKUP has one blind spot: it cannot look left. It also counts columns by number, which is easy to get wrong. XLOOKUP fixes both.\n\nXLOOKUP searches any column you name and returns from any other column. The catch: it needs Microsoft 365 or Excel 2021 or later. In Excel 2019 or older it simply does not exist. We start with a quick look at HLOOKUP, VLOOKUP's sideways sibling. Then the rest of the lesson is about XLOOKUP.",
    sections: [
      {
        heading: "HLOOKUP: the horizontal lookup",
        blocks: [
          {
            type: "formula",
            name: "HLOOKUP",
            syntax: "=HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])",
            example: "=HLOOKUP(\"Q2\", B1:E2, 2, FALSE)",
            explanation: "Searches the first ROW of the table for Q2 and returns the value from the 2nd row of the table. H stands for Horizontal."
          },
          {
            type: "grid",
            title: "Quarterly targets laid out sideways",
            rows: [
              ["", "Q1", "Q2", "Q3", "Q4"],
              ["Target (₹)", "120000", "150000", "140000", "180000"]
            ]
          },
          { type: "p", text: "`=HLOOKUP(\"Q2\", B1:E2, 2, FALSE)` returns 150000. Everything you learned about VLOOKUP applies — exact match with FALSE, locked table with $. But the search runs across a row, and the answer comes from a numbered row below. You will meet this layout in dashboards that show months or quarters as columns." },
          { type: "p", text: "In practice, most teams turn horizontal data vertical before looking anything up. HLOOKUP is worth knowing; XLOOKUP makes even it unnecessary, as you will see next." }
        ]
      },
      {
        heading: "XLOOKUP: name the column, find the value",
        blocks: [
          {
            type: "formula",
            name: "XLOOKUP",
            syntax: "=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])",
            example: "=XLOOKUP(\"Priya Nair\", B2:B6, A2:A6, \"Not found\")",
            explanation: "Searches the Name column B2:B6 for Priya Nair and returns the matching ID from A2:A6 — a column to the LEFT. The fourth argument is a built-in not-found message."
          },
          { type: "p", text: "Notice what changed from VLOOKUP. You name exactly two ranges: where to search and what to bring back. There is no table and no column number to count. The two ranges can point in any direction — left, right, up or down." },
          {
            type: "grid",
            title: "Employee table (columns A to C)",
            rows: [
              ["ID", "Name", "Department"],
              ["EMP001", "Arjun Mehta", "Sales"],
              ["EMP002", "Priya Nair", "Accounts"],
              ["EMP003", "Rahul Verma", "Operations"],
              ["EMP004", "Sneha Iyer", "Marketing"]
            ]
          },
          { type: "p", text: "With a plain VLOOKUP, finding the ID from a name would be impossible. Names sit to the right of IDs. XLOOKUP does not care." },
          { type: "p", text: "`=XLOOKUP(\"Priya Nair\", B2:B6, A2:A6, \"Not found\")` searches the Name column and returns EMP002 from the ID column to its left. The same function returns the department just as easily: `=XLOOKUP(\"EMP003\", A2:A6, C2:C6, \"Not found\")` gives Operations." }
        ]
      },
      {
        heading: "Exact match by default, and a built-in not-found message",
        blocks: [
          { type: "p", text: "Two quiet improvements make XLOOKUP safer than VLOOKUP out of the box." },
          {
            type: "ul",
            items: [
              "**Exact match is the default.** VLOOKUP without its fourth argument does a risky approximate match; XLOOKUP does the safe exact match unless you deliberately ask for something else.",
              "**The fourth argument is the not-found fallback.** =XLOOKUP(A2, B2:B6, C2:C6, \"Not found\") shows that text when the lookup value is missing — no separate IFERROR wrapper needed."
            ]
          },
          { type: "p", text: "Without the fourth argument, a missing lookup value still shows #N/A. XLOOKUP does not guess a fallback for you. It just makes one easy to add. In the employee table, `=XLOOKUP(\"EMP009\", A2:A6, C2:C6, \"Check the ID\")` returns the friendly text instead of an error." },
          {
            type: "grid",
            title: "A price lookup built three ways",
            rows: [
              ["Method", "Formula", "Needs"],
              ["VLOOKUP", "=VLOOKUP(A2, $G$2:$H$6, 2, FALSE)", "Locking, column counting, first-column rule"],
              ["XLOOKUP", "=XLOOKUP(A2, $G$2:$G$6, $H$2:$H$6, \"Not found\")", "Two named ranges, no counting"],
              ["XLOOKUP with fallback", "=XLOOKUP(A2, $G$2:$G$6, $H$2:$H$6, \"Not found\")", "Same formula — fallback built in"]
            ]
          },
          { type: "p", text: "The XLOOKUP version says exactly what it does: search the products, return the prices, say \"Not found\" when there is no match. Six months later, anyone reading it understands instantly. That is a real advantage when colleagues inherit your workbooks." }
        ]
      },
      {
        heading: "When to choose which lookup",
        blocks: [
          {
            type: "table",
            headers: ["Situation", "Best choice", "Why"],
            rows: [
              ["Your Excel is Microsoft 365 or 2021+", "XLOOKUP", "Looks any direction, exact by default, built-in fallback"],
              ["Your Excel is 2019 or older, right-side lookup", "VLOOKUP", "The only simple option; lock the table, use FALSE"],
              ["Your Excel is 2019 or older, left-side lookup", "INDEX and MATCH (next lesson)", "VLOOKUP cannot look left at all"],
              ["Data laid out horizontally", "XLOOKUP", "Searches rows just as happily as columns, so HLOOKUP is rarely needed"],
              ["Editing shared workbooks in mixed Excel versions", "VLOOKUP", "Opens and calculates identically everywhere"]
            ]
          },
          { type: "warning", text: "If a colleague opens your XLOOKUP workbook in Excel 2019, every XLOOKUP cell shows #NAME? because the function does not exist there. For workbooks that must travel between Excel versions, VLOOKUP (or INDEX/MATCH) remains the safe choice." }
        ]
      }
    ],
    mistakes: [
      {
        wrong: "Using XLOOKUP in Excel 2019 or 2016",
        right: "Use VLOOKUP, or INDEX and MATCH, in versions before 2021",
        why: "XLOOKUP does not exist in those versions. Every cell shows #NAME? and the workbook appears broken to the recipient."
      },
      {
        wrong: "Lookup array and return array of different lengths: =XLOOKUP(A2, B2:B6, C2:C10)",
        right: "Keep both arrays the same size: =XLOOKUP(A2, B2:B6, C2:C6)",
        why: "XLOOKUP pairs the arrays position by position. Unequal sizes return the #VALUE! error."
      },
      {
        wrong: "Adding IFERROR around every XLOOKUP out of habit",
        right: "Use the built-in fourth argument: =XLOOKUP(A2, B2:B6, C2:C6, \"Not found\")",
        why: "The fallback argument is cleaner and shows exactly which lookups have a defined not-found behaviour."
      }
    ],
    shortcuts: [],
    exercises: [
      {
        task: "Open the lookup-practice workbook, Exercise sheet. Using the employee table, return the department for the name typed in a cell, with the fallback \"Not an employee\".",
        hint: "Lookup array is the name column; return array is the department column.",
        solution: "=XLOOKUP(IDcell, B2:B6, C2:C6, \"Not an employee\") — ranges adjusted to the sheet."
      },
      {
        task: "Now do the reverse: from the ID column, return the employee name. Notice the return array sits to the LEFT of the lookup array.",
        hint: "This is impossible with VLOOKUP; XLOOKUP handles it naturally.",
        solution: "=XLOOKUP(IDcell, A2:A6, B2:B6, \"Unknown ID\") returns the name for whatever ID you type."
      },
      {
        task: "Build the same price lookup you made with VLOOKUP in the last lesson, but with XLOOKUP and a \"Not in price list\" fallback. Compare the two formulas side by side.",
        hint: "Two named ranges replace the table and the column number.",
        solution: "=XLOOKUP(A2, $G$2:$G$6, $H$2:$H$6, \"Not in price list\"). Same result as the VLOOKUP version, with no column counting."
      },
      {
        task: "On the Solution sheet, find the horizontal quarterly table and use XLOOKUP to return the target for the quarter named in a cell.",
        hint: "XLOOKUP works on a row exactly like on a column: lookup array is the row of quarter labels.",
        solution: "=XLOOKUP(Qcell, B1:E1, B2:E2, \"Unknown quarter\") returns the matching quarter's target."
      }
    ],
    quiz: [
      {
        q: "Which is the correct XLOOKUP argument order?",
        options: ["lookup_value, lookup_array, return_array, [if_not_found]", "lookup_value, return_array, lookup_array", "table_array, col_index, lookup_value", "lookup_array, lookup_value, return_array"],
        answer: 0,
        explanation: "First what you are looking for, then where to search, then what to return, and optionally the not-found text."
      },
      {
        q: "What does XLOOKUP do by default when no exact match exists?",
        options: ["Returns #N/A, or the fourth argument's text if you supplied one", "Returns zero", "Returns the nearest value above", "Refuses to calculate"],
        answer: 0,
        explanation: "Exact match is the default and a miss produces #N/A — unless you use the built-in if_not_found argument to supply friendlier text."
      },
      {
        q: "Why can XLOOKUP return a value from a column to the LEFT of the search column?",
        options: ["Because you name the lookup array and return array separately — no table or column counting", "It searches from right to left by default", "It sorts the table first", "Because FALSE is required"],
        answer: 0,
        explanation: "The two arrays are independent. XLOOKUP finds the position in one range and returns the same position from any other range."
      },
      {
        q: "What does HLOOKUP search?",
        options: ["The first row of a table, returning from a numbered row below", "The first column, like VLOOKUP", "All columns at once", "Only numbers"],
        answer: 0,
        explanation: "H is for Horizontal: match across the top row, return from the chosen row underneath. Same exact-match rules as VLOOKUP."
      },
      {
        q: "A colleague on Excel 2019 opens your workbook and all lookup cells show #NAME?. What happened?",
        options: ["The workbook uses XLOOKUP, which does not exist in Excel 2019", "The file is corrupted", "The lookup values were deleted", "Excel needs an update"],
        answer: 0,
        explanation: "XLOOKUP arrived with Microsoft 365/Excel 2021. Older versions show #NAME? for unknown functions; use VLOOKUP or INDEX/MATCH for shared files."
      }
    ],
    summary: [
      "HLOOKUP searches across the first row of a table and returns from a numbered row below it.",
      "XLOOKUP(lookup, lookup_array, return_array, [if_not_found]) names the search column and the return column directly.",
      "XLOOKUP looks left, right, up and down — the first-column rule of VLOOKUP disappears.",
      "Exact match is the default, and the fourth argument provides a built-in not-found message.",
      "XLOOKUP needs Microsoft 365 or Excel 2021+; for workbooks shared with older Excels, stick to VLOOKUP or INDEX/MATCH."
    ],
    downloads: [
      {
        file: "assets/files/level2/lookup-practice.xlsx",
        label: "Lookup practice workbook",
        note: "Exercise and Solution sheets with product and employee tables for VLOOKUP, XLOOKUP and INDEX/MATCH"
      }
    ],
    keywords: ["XLOOKUP", "HLOOKUP", "lookup array", "return array", "not found message", "Microsoft 365", "look left", "VLOOKUP comparison"]
  },
  {
    id: "l2-index-match",
    title: "INDEX and MATCH",
    difficulty: "Intermediate",
    duration: 18,
    version: "",
    objectives: [
      "Use MATCH to find the position of a value in a column",
      "Use INDEX to return the value at a position in a range",
      "Combine them as =INDEX(return_range, MATCH(lookup, lookup_range, 0))",
      "Explain why INDEX/MATCH survives column insertions that break VLOOKUP",
      "Build a lookup step by step, verifying MATCH first"
    ],
    keyIdea: "MATCH finds a position; INDEX returns the value at that position. Together they beat VLOOKUP.",
    intro: "INDEX and MATCH are two simple functions that become powerful together. MATCH answers: which position is this value in? INDEX answers: what is the value at this position?\n\nTogether they do everything VLOOKUP does. They also do two things VLOOKUP cannot: look to the left, and survive changes to the sheet. The technique works in every Excel version. That is why professionals still use it, even after XLOOKUP arrived. Build it in two steps, exactly as this lesson shows.",
    sections: [
      {
        heading: "MATCH: where is the value?",
        blocks: [
          {
            type: "formula",
            name: "MATCH",
            syntax: "=MATCH(lookup_value, lookup_array, [match_type])",
            example: "=MATCH(\"Staplers\", G2:G6, 0)",
            explanation: "Returns the POSITION of Staplers within G2:G6 — here 3 — not the value and not a cell reference. The 0 means exact match."
          },
          { type: "p", text: "The third argument, match_type, decides how MATCH behaves. The rule for daily work: **always type 0** for an exact match." },
          {
            type: "table",
            headers: ["match_type", "Behaviour", "Use it when"],
            rows: [
              ["0", "Exact match only", "Almost always — the professional default"],
              ["1", "Largest value less than or equal to the lookup; data must be sorted ascending", "Grading bands and sorted tables — rare"],
              ["-1", "Smallest value greater than or equal to the lookup; data must be sorted descending", "Very rare"],
              ["(omitted)", "Behaves like 1 — approximate!", "Never omit it by accident"]
            ]
          },
          { type: "p", text: "In the Sunrise Traders price list, the products in `G2:G6` are Notebooks, Pens, Staplers, Printers, Laptops. `=MATCH(\"Staplers\", G2:G6, 0)` returns 3 because Staplers is the third item. If the product is not there, MATCH returns #N/A — the same \"not found\" signal as the lookups." },
          { type: "warning", text: "The number MATCH returns is the position within the range you selected, not the row number of the sheet. \"Staplers\" is position 3 in `G2:G6`, even though it sits in row 4 of the sheet." }
        ]
      },
      {
        heading: "INDEX: what is at that position?",
        blocks: [
          {
            type: "formula",
            name: "INDEX",
            syntax: "=INDEX(array, row_num)",
            example: "=INDEX(H2:H6, 3)",
            explanation: "Returns the 3rd value of the range H2:H6 — the price 130. INDEX does no searching at all; it just picks the item at a position."
          },
          { type: "p", text: "On its own, INDEX is a plain \"give me the nth item\" function. `=INDEX(H2:H6, 3)` walks three items into the prices (45, 12, 130) and returns 130. Combine the two, and the position is no longer hard-coded. MATCH works it out for you." }
        ]
      },
      {
        heading: "INDEX + MATCH: the full technique, built step by step",
        blocks: [
          { type: "p", text: "The combined formula reads backwards for beginners, so build it in the order your brain works: find the position first, then fetch the value." },
          {
            type: "steps",
            items: [
              "Step 1 — write MATCH alone in a scratch cell: =MATCH(A2, $G$2:$G$6, 0). Check it returns a sensible position number, for example 3 for Staplers.",
              "Step 2 — in the real price cell, write =INDEX($H$2:$H$6, 3) using that number, and confirm it returns the right price.",
              "Step 3 — replace the hard-coded 3 with the MATCH formula: =INDEX($H$2:$H$6, MATCH(A2, $G$2:$G$6, 0)).",
              "Step 4 — delete the scratch cell and fill the final formula down the column.",
              "Step 5 — test with a product that does not exist and confirm you get #N/A, then wrap in IFERROR if the report needs it."
            ]
          },
          {
            type: "formula",
            name: "INDEX and MATCH combined",
            syntax: "=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))",
            example: "=INDEX($H$2:$H$6, MATCH(A2, $G$2:$G$6, 0))",
            explanation: "MATCH finds the row position of the product in the product column; INDEX returns the price from that same position in the price column."
          },
          { type: "p", text: "Read it from the inside out. \"The position of A2 within the products\" feeds \"the value at that position within the prices\". Both ranges are locked with $, so the formula fills down safely — the same lesson you learned with VLOOKUP." }
        ]
      },
      {
        heading: "Why INDEX/MATCH beats VLOOKUP",
        blocks: [
          {
            type: "ul",
            items: [
              "**It looks left.** The return range and the lookup range are completely independent. To get the employee ID from a name, search the Name column and return from the ID column — no rearranging tables.",
              "**It survives inserted columns.** VLOOKUP's col_index_num is a number that silently points at the wrong column when someone inserts a column inside the table. INDEX/MATCH names its ranges directly; inserting a column shifts nothing meaningful.",
              "**It is faster on very large sheets.** VLOOKUP loads the whole table for every lookup; INDEX/MATCH touches only the two narrow columns it needs.",
              "**The workbooks stay compatible.** It runs identically in every Excel version, old or new."
            ]
          },
          {
            type: "grid",
            title: "Employee table: ID from name — VLOOKUP cannot do this",
            rows: [
              ["ID", "Name", "Department"],
              ["EMP001", "Arjun Mehta", "Sales"],
              ["EMP002", "Priya Nair", "Accounts"],
              ["EMP003", "Rahul Verma", "Operations"],
              ["EMP004", "Sneha Iyer", "Marketing"],
              ["Lookup", "Rahul Verma", "=INDEX(A2:A5, MATCH(B6, B2:B5, 0))"]
            ]
          },
          { type: "p", text: "MATCH finds \"Rahul Verma\" at position 3 of the names. INDEX returns the 3rd ID, EMP003. The search column and the return column have no required order at all. That single property is why professionals keep this pair, even in the XLOOKUP era." }
        ]
      }
    ],
    mistakes: [
      {
        wrong: "=MATCH(\"Staplers\", G2:G6) without the 0",
        right: "=MATCH(\"Staplers\", G2:G6, 0)",
        why: "Omitted match_type behaves like 1 — an approximate match that assumes sorted data and can return the wrong position silently."
      },
      {
        wrong: "Mixing range sizes: =INDEX(H2:H6, MATCH(A2, G2:G10, 0))",
        right: "Keep both ranges the same length and the same start row: =INDEX($H$2:$H$6, MATCH(A2, $G$2:$G$6, 0))",
        why: "The position from MATCH must mean the same thing in both ranges. Offset or unequal ranges fetch values from the wrong rows."
      },
      {
        wrong: "Treating MATCH's result as a sheet row number",
        right: "Remember MATCH returns a position within the selected range",
        why: "Position 3 in G2:G6 is row 4 of the sheet. Writing =INDEX(H2:H6, 4) would grab the next item down."
      },
      {
        wrong: "Building =INDEX($H$2:$H$6, MATCH(...)) in one go and debugging it blind",
        right: "Write MATCH alone first, check the number, then wrap it in INDEX",
        why: "A combined formula that fails gives you one error with no clue where. The two-step build isolates the problem immediately."
      }
    ],
    shortcuts: [],
    exercises: [
      {
        task: "Open the lookup-practice workbook, Exercise sheet. In a scratch cell, write a MATCH formula that returns the position of Printers in the product column. Confirm the number before going further.",
        hint: "Exact match needs the third argument 0.",
        solution: "=MATCH(\"Printers\", G2:G6, 0) returns 4."
      },
      {
        task: "Using that position, write an INDEX formula that returns the price of Printers.",
        hint: "Same position number inside the price range.",
        solution: "=INDEX(H2:H6, 4) returns 8500."
      },
      {
        task: "Now combine the two so the price is looked up from the item name typed in a cell, with both ranges locked. Fill it down and compare with the Solution sheet.",
        hint: "Replace the hard-coded position with the MATCH formula.",
        solution: "=INDEX($H$2:$H$6, MATCH(A2, $G$2:$G$6, 0))"
      },
      {
        task: "In the employee table on the same sheet, return the ID for the name \"Sneha Iyer\" — a lookup to the LEFT, which VLOOKUP cannot do.",
        hint: "Search the Name column, return from the ID column.",
        solution: "=INDEX(A2:A5, MATCH(\"Sneha Iyer\", B2:B5, 0)) returns EMP004."
      }
    ],
    quiz: [
      {
        q: "What does =MATCH(\"Staplers\", G2:G6, 0) return?",
        options: ["3 — the position of Staplers within the range", "130 — the price", "G4 — the cell reference", "TRUE"],
        answer: 0,
        explanation: "MATCH returns a position number, not a value or a reference. Staplers is the third item in G2:G6."
      },
      {
        q: "Which match_type should you use for an exact match?",
        options: ["0", "1", "-1", "Omit it"],
        answer: 0,
        explanation: "0 is exact. Omitting the argument behaves like 1 — approximate — which assumes sorted data and can quietly return wrong positions."
      },
      {
        q: "In =INDEX(H2:H6, 3), what does the 3 do?",
        options: ["Returns the 3rd value of the range H2:H6", "Looks in row 3 of the sheet", "Counts 3 columns across", "Returns 3 results"],
        answer: 0,
        explanation: "INDEX picks the item at that position within the range. Position 3 of the prices is 130."
      },
      {
        q: "What can INDEX/MATCH do that VLOOKUP cannot?",
        options: ["Return a value from a column to the left of the lookup column", "Handle text lookups", "Return exact matches", "Work in old Excel versions"],
        answer: 0,
        explanation: "INDEX/MATCH names two independent ranges, so the return range can sit anywhere. VLOOKUP can only look right."
      },
      {
        q: "Why is INDEX/MATCH safer than VLOOKUP when someone inserts a column inside the lookup table?",
        options: ["It names its return range directly instead of using a column number", "It recalculates faster", "It ignores inserted columns", "VLOOKUP and it are identical here"],
        answer: 0,
        explanation: "A VLOOKUP's hard-coded col_index_num silently points elsewhere after an insert. INDEX/MATCH's return range moves with the column it points at."
      }
    ],
    summary: [
      "MATCH(lookup, range, 0) returns the position of a value; always include the 0 for exact match.",
      "INDEX(range, position) returns the value at that position within the range.",
      "The combined pattern is =INDEX(return_range, MATCH(lookup_value, lookup_range, 0)) — find the position, then fetch the value.",
      "INDEX/MATCH looks left, survives inserted columns, and works in every Excel version.",
      "Build in two steps — MATCH alone first, then wrap in INDEX — and lock both ranges with $ before filling down."
    ],
    downloads: [
      {
        file: "assets/files/level2/lookup-practice.xlsx",
        label: "Lookup practice workbook",
        note: "Exercise and Solution sheets with product and employee tables for VLOOKUP, XLOOKUP and INDEX/MATCH"
      }
    ],
    keywords: ["INDEX", "MATCH", "match type", "exact match position", "INDEX MATCH combined", "look left", "insert columns", "VLOOKUP alternative"]
  },
  {
    id: "l2-text-functions",
    title: "Text Functions",
    difficulty: "Intermediate",
    duration: 18,
    version: "",
    objectives: [
      "Pull parts of a code or name with LEFT, RIGHT and MID",
      "Measure and clean text with LEN and TRIM",
      "Standardise case with UPPER, LOWER and PROPER",
      "Join text with CONCAT and the & operator",
      "Use SUBSTITUTE, FIND versus SEARCH, and TEXT for formatting numbers as text"
    ],
    keyIdea: "Text functions cut, clean, join and reformat text so messy data becomes report-ready.",
    intro: "Real office data is messy. Names arrive with stray spaces. Product codes mix letters and numbers. Dates come as text. Everything is typed in a different case.\n\nText functions are the cleaning tools. They cut, join, measure and standardise text. Your reports will look like one careful clerk wrote them. They also prepare you for a fact from the date lesson: Excel treats text and numbers as different things. Half of all lookup failures are text problems in disguise.",
    sections: [
      {
        heading: "LEFT, RIGHT, MID: cutting text into parts",
        blocks: [
          {
            type: "formula",
            name: "LEFT",
            syntax: "=LEFT(text, [num_chars])",
            example: "=LEFT(\"STP-013\", 3)",
            explanation: "Returns the first characters of the text. Here the first 3 characters of STP-013, which is STP — the product family code."
          },
          {
            type: "formula",
            name: "MID",
            syntax: "=MID(text, start_num, num_chars)",
            example: "=MID(\"STP-013-BLACK\", 5, 3)",
            explanation: "Reads 3 characters starting from position 5 of the code — the item number 013. MID is how you dig a block out of the middle."
          },
          {
            type: "formula",
            name: "RIGHT",
            syntax: "=RIGHT(text, [num_chars])",
            example: "=RIGHT(\"STP-013-BLACK\", 5)",
            explanation: "Returns the last 5 characters — BLACK, the colour suffix of the product code."
          },
          { type: "p", text: "Positions count characters from the left, starting at 1. In `STP-013-BLACK`: S is 1, T is 2, P is 3, the dash is 4, 0 is 5. That is why MID starts at 5 to grab the number block. Leave the count out of LEFT or RIGHT and you get one character: `=LEFT(\"STP-013\")` gives S." },
          { type: "p", text: "Where is this used? Splitting warehouse codes into family, item and colour; taking the city code off an invoice number; pulling the year out of an order reference like `ORD-2026-0142` with `=MID(\"ORD-2026-0142\", 5, 4)`." }
        ]
      },
      {
        heading: "LEN and TRIM: measuring and cleaning",
        blocks: [
          {
            type: "formula",
            name: "TRIM",
            syntax: "=TRIM(text)",
            example: "=TRIM(\"  Priya Nair  \")",
            explanation: "Removes leading and trailing spaces and squeezes double spaces in the middle down to one. Returns the clean text: Priya Nair."
          },
          {
            type: "formula",
            name: "LEN",
            syntax: "=LEN(text)",
            example: "=LEN(\"STP-013\")",
            explanation: "Counts every character, spaces included. STP-013 is 7 characters long."
          },
          { type: "p", text: "TRIM is the single most valuable data-cleaning function in Excel. Invisible trailing spaces break VLOOKUP matches — remember the #N/A mystery from the lookup lesson. A TRIM column over imported names fixes most of them. Put `=TRIM(B2)` next to the raw import, then paste the results back as values." },
          { type: "p", text: "LEN is the diagnostic partner. `=LEN(TRIM(B2))` tells you whether a name still has odd spacing. And `=LEN(A2)-LEN(SUBSTITUTE(A2, \" \", \"\"))` counts the spaces in `A2` — a neat trick once you know SUBSTITUTE, coming up in a moment." },
          { type: "warning", text: "TRIM removes the normal space character (code 32). Data pasted from web pages sometimes holds a non-breaking space (code 160), which TRIM leaves untouched. If TRIM seems to do nothing, that is usually why. Lesson 13's CHAR family and SUBSTITUTE handle the stubborn cases." }
        ]
      },
      {
        heading: "UPPER, LOWER and PROPER: standardising case",
        blocks: [
          {
            type: "ul",
            items: [
              "`=UPPER(\"north\")` returns NORTH — handy for region codes that should shout.",
              "`=LOWER(\"NOTEBOOKS\")` returns notebooks — good for email-style identifiers.",
              "`=PROPER(\"priya nair\")` returns Priya Nair — capitalises the first letter of each word."
            ]
          },
          {
            type: "grid",
            title: "Cleaning an imported name column",
            rows: [
              ["", "Raw name (column B)", "Cleaned formula", "Result"],
              ["2", "  priya nair ", "=PROPER(TRIM(B2))", "Priya Nair"],
              ["3", "RAHUL  verma", "=PROPER(TRIM(B3))", "Rahul Verma"],
              ["4", "sneha iyer", "=PROPER(TRIM(B4))", "Sneha Iyer"]
            ]
          },
          { type: "p", text: "The nested pair `=PROPER(TRIM(B2))` is a standard cleaning formula. TRIM first fixes the spacing. PROPER then fixes the case. Note row 3: TRIM also squeezed the double space between RAHUL and verma down to one, before PROPER capitalised it. Cleaning functions chain like this all the time." },
          { type: "p", text: "One caution: PROPER capitalises every word. So `=PROPER(\"l2a-3 staplers\")` produces \"L2A-3 Staplers\" — the letter after a dash or number also gets capitalised. Check PROPER's output on codes. UPPER and LOWER are predictable; PROPER is not." }
        ]
      },
      {
        heading: "Joining text: & and CONCAT",
        blocks: [
          {
            type: "formula",
            name: "CONCAT (and the older CONCATENATE)",
            syntax: "=CONCAT(text1, [text2], ...)",
            example: "=CONCAT(B2, \" \", C2)",
            explanation: "Joins pieces of text into one. CONCAT replaces the older CONCATENATE, which still works in every Excel version for compatibility."
          },
          { type: "p", text: "The ampersand does the same job with less typing: `=B2&\" \"&C2` joins the first name, a space and the surname. The space between the quotes is the part beginners forget — without it, \"Arjun\" and \"Mehta\" become \"ArjunMehta\"." },
          {
            type: "grid",
            title: "Building labels for reports",
            rows: [
              ["", "First (B)", "Last (C)", "Full name", "Report label"],
              ["2", "Arjun", "Mehta", "=B2&\" \"&C2", "=UPPER(B2)&\" - \"&UPPER(C2)"],
              ["3", "Priya", "Nair", "=B3&\" \"&C3", "=UPPER(B3)&\" - \"&UPPER(C3)"]
            ]
          },
          { type: "p", text: "The label column produces ARJUN - MEHTA and PRIYA - NAIR, ready for a report header. Joining is also how you build lookup-friendly keys. `=A2&\"-\"&B2` glues region and product into NORTH-NOTEBOOKS, which COUNTIF can then search for." },
          { type: "p", text: "Any text can be joined with a formula result, including numbers. But notice the next section for what happens to their formatting." }
        ]
      },
      {
        heading: "SUBSTITUTE, FIND versus SEARCH, and TEXT",
        blocks: [
          {
            type: "formula",
            name: "SUBSTITUTE",
            syntax: "=SUBSTITUTE(text, old_text, new_text, [instance_num])",
            example: "=SUBSTITUTE(\"STP-013-BLACK\", \"-\", \" \")",
            explanation: "Replaces every dash with a space, giving STP 013 BLACK. The optional instance number replaces only that occurrence — the 2nd dash, for example."
          },
          { type: "p", text: "SUBSTITUTE replaces known text everywhere it appears. Use it to standardise separators, strip unwanted characters, or swap an old prefix for a new one. To \"remove all spaces\" you write `=SUBSTITUTE(A2, \" \", \"\")` — replacing spaces with nothing." },
          {
            type: "table",
            headers: ["Function", "Case sensitivity", "Wildcards", "Returns"],
            rows: [
              ["FIND", "Case-sensitive", "Not allowed", "Position of the text, or #VALUE! if absent"],
              ["SEARCH", "Case-insensitive", "Allowed", "Position of the text, or #VALUE! if absent"]
            ]
          },
          { type: "p", text: "Both return a position number you can feed into MID or LEFT. `=FIND(\"-\", \"STP-013-BLACK\")` returns 4 — the first dash sits at position 4. So `=LEFT(\"STP-013-BLACK\", FIND(\"-\", \"STP-013-BLACK\")-1)` returns STP without hard-coding the 3. Use FIND when case matters (codes). Use SEARCH when it does not (typed names)." },
          {
            type: "formula",
            name: "TEXT",
            syntax: "=TEXT(value, format_text)",
            example: "=TEXT(D2, \"dd-mmm-yyyy\")",
            explanation: "Converts a number or date into text wearing a specific format. A date of 14-Nov-2026 becomes the literal text \"14-Nov-2026\" that can be joined into a report title."
          },
          { type: "p", text: "TEXT shines when a number must appear inside a sentence. `=\"Report for \"&TEXT(TODAY(), \"dd-mmm-yyyy\")` builds a title that updates itself. The format code does the styling — `\"#,##0.00\"` for thousand-separated money." },
          { type: "p", text: "Remember the trade-off from the ROUND lesson, in reverse. TEXT returns **text**, so the result can no longer be added up. Format for display, TEXT only for labels." }
        ]
      }
    ],
    mistakes: [
      {
        wrong: "=B2&C2 to join names",
        right: "=B2&\" \"&C2",
        why: "Without the quoted space the result is ArjunMehta. Whatever sits between the joined pieces must be typed inside quotes."
      },
      {
        wrong: "Comparing or looking up imported text without TRIM",
        right: "=VLOOKUP(TRIM(A2), ...)",
        why: "Invisible trailing spaces make two visually identical texts different, and the lookup returns #N/A for a value you can plainly see."
      },
      {
        wrong: "Using PROPER on product codes: =PROPER(\"stP-013\")",
        right: "Use UPPER for codes: =UPPER(\"stP-013\") gives STP-013",
        why: "PROPER capitalises letters after dashes and numbers too, which mangles codes. Reserve PROPER for human names and addresses."
      },
      {
        wrong: "Treating =TEXT(D2, \"dd-mmm-yyyy\") as a date you can calculate with",
        right: "Keep the real date in the cell; use TEXT only for labels and titles",
        why: "TEXT returns text. A formula like D2+30 fails or gives nonsense on a TEXT result — keep the underlying value for maths."
      }
    ],
    shortcuts: [],
    exercises: [
      {
        task: "Open the text-date-practice workbook, Exercise sheet. Column B holds product codes like STP-013-BLACK. In column C extract the family code (the first 3 letters).",
        hint: "One function returns the leading characters.",
        solution: "C2: =LEFT(B2, 3), filled down."
      },
      {
        task: "In column D extract the item number (the 3 characters after the first dash) and in column E the colour (the last 5 characters).",
        hint: "One needs a start position and a count; the other counts from the end.",
        solution: "D2: =MID(B2, 5, 3). E2: =RIGHT(B2, 5)."
      },
      {
        task: "Column F holds raw customer names with stray spaces and lowercase typing. Clean them in column G with one nested formula.",
        hint: "Fix spacing first, then fix case — two functions in one formula.",
        solution: "G2: =PROPER(TRIM(F2)), filled down."
      },
      {
        task: "In column H build the report label: the cleaned name in capitals, then \" - \", then the family code from exercise 1, also in capitals. Compare with the Solution sheet.",
        hint: "Join with & and wrap each part in UPPER.",
        solution: "H2: =UPPER(G2)&\" - \"&UPPER(C2), producing PRIYA NAIR - STP."
      }
    ],
    quiz: [
      {
        q: "What does =MID(\"STP-013-BLACK\", 5, 3) return?",
        options: ["013", "STP", "LACK", "-01"],
        answer: 0,
        explanation: "MID starts at position 5 (the first 0) and reads 3 characters. Positions 1 to 4 are S, T, P and the dash."
      },
      {
        q: "What does =TRIM(\"  Priya  Nair \") return?",
        options: ["Priya Nair — spaces at the edges and doubles removed", "PriyaNair — all spaces gone", "  Priya  Nair  — unchanged", "An error"],
        answer: 0,
        explanation: "TRIM strips leading and trailing spaces and collapses internal double spaces to single ones."
      },
      {
        q: "Which function is case-insensitive and allows wildcards?",
        options: ["SEARCH", "FIND", "MID", "LEN"],
        answer: 0,
        explanation: "SEARCH ignores case and accepts wildcards. FIND is case-sensitive and treats wildcards as literal characters."
      },
      {
        q: "What does =TEXT(D2, \"dd-mmm-yyyy\") return when D2 holds the date 14-Nov-2026?",
        options: ["The text \"14-Nov-2026\", which can be joined into a title", "The number 14", "The date unchanged", "An error"],
        answer: 0,
        explanation: "TEXT formats a value as text. It is perfect for report titles, but the result can no longer be used in date arithmetic."
      },
      {
        q: "Why does =B2&C2 make \"ArjunMehta\"?",
        options: ["The ampersand joins exactly what you give it — no space was supplied", "Excel removes spaces when joining", "C2 lost its leading space", "Ampersand is the wrong joining tool"],
        answer: 0,
        explanation: "Joining is literal. To get a space between the pieces, include it in quotes: =B2&\" \"&C2."
      }
    ],
    summary: [
      "LEFT, RIGHT and MID cut text from the start, the end and a chosen middle position — positions count from 1.",
      "TRIM fixes invisible spacing problems that break lookups; LEN measures text including its spaces.",
      "UPPER, LOWER and PROPER standardise case; PROPER is for names, UPPER is for codes.",
      "Join with & or CONCAT, and remember to include the separator (\" \", \" - \") in quotes.",
      "SUBSTITUTE replaces known text, FIND/SEARCH locate positions, and TEXT formats numbers as text for labels — never for calculations."
    ],
    downloads: [
      {
        file: "assets/files/level2/text-date-practice.xlsx",
        label: "Text and date practice workbook",
        note: "Exercise and Solution sheets for text and date functions"
      }
    ],
    keywords: ["LEFT", "RIGHT", "MID", "LEN", "TRIM", "UPPER", "LOWER", "PROPER", "CONCAT", "ampersand", "SUBSTITUTE", "FIND", "SEARCH", "TEXT function", "data cleaning"]
  },
  // __APPEND__
);
