// Excel Lab — Level 1 (Excel Fundamentals) — lessons 7-12 (no metadata here)
window.LEVEL_LESSONS = window.LEVEL_LESSONS || {};

(window.LEVEL_LESSONS[1] = window.LEVEL_LESSONS[1] || []).push(
  {
    id: "l1-cell-formatting",
    title: "Cell Formatting",
    difficulty: "Beginner",
    duration: 15,
    version: "",
    objectives: [
      "Apply bold, italic, underline, fonts and font colours",
      "Use fill colour and borders to build a clean report layout",
      "Align text and wrap long content inside cells",
      "Copy formatting quickly with the Format Painter"
    ],
    intro: "Formatting is the difference between a sheet only you can read and a report your whole office can read. The numbers are the same; the presentation decides whether people trust and use your work.\n\nThis lesson covers everyday formatting from the Home tab: fonts, colours, borders, alignment and the tools that apply them fast. Your target: a sheet that looks like a neat office document, not a wall of numbers.",
    sections: [
      {
        heading: "Font basics: bold, italic, underline and size",
        blocks: [
          { type: "p", text: "All of these live on the **Home** tab, in the Font group. Select the cell or range first, then apply:" },
          { type: "ul", items: [
            "**Bold** ([[Ctrl]] + [[B]]) — for headings and totals. The workhorse of formatting.",
            "**Italic** ([[Ctrl]] + [[I]]) — for notes and captions. Use sparingly; italic numbers are harder to read.",
            "**Underline** ([[Ctrl]] + [[U]]) — rare in spreadsheets; a bottom border usually looks better.",
            "**Font size** — the dropdown with numbers. Titles 14-16, headings 11-12 bold, body 10-11.",
            "**Font family** — Calibri is the default. Stick with one or two fonts per workbook; offices keep it simple."
          ] },
          { type: "p", text: "A simple professional recipe for any report: **title** in 14 bold, **column headers** in 11 bold, **totals row** in bold, everything else regular. That is genuinely how most good office sheets are dressed." },
          { type: "p", text: "You can also change the **font colour** (the A with a coloured bar) — dark grey or navy reads better than pure black on some screens. Reserve red for negative or warning values, because people read red as \"problem\"." }
        ]
      },
      {
        heading: "Fill colour and borders",
        blocks: [
          { type: "p", text: "**Fill colour** (the paint bucket icon) colours the background of selected cells. One soft fill behind the header row is the single most effective formatting move in Excel: it separates headers from data instantly. Pick light shades — pale blue, light grey — so black text stays readable. Loud dark fills fight with the data." },
          { type: "p", text: "**Borders** (the grid icon next to the bucket) draw lines around selected cells. Click the dropdown to choose: bottom border, all borders, thick box border, and more. Why do borders matter in office reports? Because a printed or emailed sheet has no gridlines — the faint grey lines you see on screen do not print by default. Without borders, a printed table is a floating mess of numbers. Draw **All Borders** on the data block and the printout instantly becomes a proper table." },
          { type: "grid", title: "A formatted Sunrise Traders expense report", rows: [
            ["Item", "Cost"],
            ["Notebooks", "1800"],
            ["Printer ink", "2400"],
            ["Printer paper", "900"],
            ["Total", "5100"]
          ] },
          { type: "p", text: "In the finished report behind this grid, the header row has a light fill and bold text, the data block has all borders, and the Total row is bold with a thicker top border. Take a moment to imagine that — or better, build it in the exercise." },
          { type: "tip", text: "A classic accounting trick: apply only a **bottom border** under each number in a totals column, and a **double top border** above the final total. It looks tidy and it is exactly how accountants present figures." }
        ]
      },
      {
        heading: "Alignment and wrap text",
        blocks: [
          { type: "p", text: "The Alignment group controls where content sits inside the cell. **Horizontal** alignment: left, centre, right. **Vertical** alignment: top, middle, bottom — useful when rows are tall. Remember from the data entry lesson that text defaults left and numbers right; you can override for headers, but **never** centre-align a column of numbers in financial data — the digits stop lining up, and the column becomes hard to compare and audit." },
          { type: "p", text: "**Wrap Text** (Home > Alignment group) fixes the classic long-text problem: a long sentence in a narrow cell spills over neighbours or gets cut off. Turn on Wrap Text and the cell grows taller, showing every word across multiple lines. Use it for column headers like \"Amount Pending Approval\" and for description columns. Widen the row height back down by double-clicking the row border if it overshoots." },
          { type: "p", text: "Related buttons beside it: **Increase Indent** nudges text inward (nice for sub-items), and **Orientation** rotates text — handy for very narrow date columns on printed reports, but rare in daily work." }
        ]
      },
      {
        heading: "Merge & Center — useful but use it sparingly",
        blocks: [
          { type: "p", text: "Home > Alignment > **Merge & Center** combines several cells into one wide cell and centres the text. The classic use: a single title row across your report. Select `A1:E1`, click Merge & Center, type \"Sunrise Traders — Weekly Expense Report\", and you have a proper title bar." },
          { type: "p", text: "So why \"sparingly\"? Because merged cells cause three real problems:" },
          { type: "ul", items: [
            "**Sorting and filtering break** — Excel cannot sort a range that contains merged cells without a complaint.",
            "**Selection behaves oddly** — clicking a merged cell selects the whole merge; copying parts of the block misbehaves.",
            "**Formulas stumble** — some functions and copy-paste operations choke on merged areas."
          ] },
          { type: "p", text: "The safer habit used by experienced users: **Center Across Selection**. Open the Format Cells dialog ([[Ctrl]] + [[1]]), Alignment tab, set Horizontal to \"Center Across Selection\". The title looks centred over several columns, but each cell stays separate, so sorting and formulas keep working. For titles it feels identical; for data, avoid merging entirely." },
          { type: "warning", text: "Never merge cells that hold data used by formulas. Merge is for presentation rows like titles — not for data." }
        ]
      },
      {
        heading: "Format Painter: copy the look, not the data",
        blocks: [
          { type: "p", text: "You made a header row look perfect — fill, bold, borders. Another sheet needs the same look. Do not rebuild it by hand. Use the **Format Painter**:" },
          { type: "steps", items: [
            "Select the already-formatted cells.",
            "Click **Format Painter** (the paintbrush icon, Home tab, left side). The cursor gains a brush.",
            "Drag over the destination cells. The formatting copies; the data does not."
          ] },
          { type: "p", text: "Single-click the painter for a one-shot copy. **Double-click** it to lock the brush on, so you can paint several separate ranges one after another; press [[Esc]] or click the painter again to stop. There is also a keyboard route: select the source, press [[Ctrl]] + [[C]], select the destination, then use Paste Special > Formats — the same result through the clipboard." }
        ]
      },
      {
        heading: "Cell styles: one-click consistency",
        blocks: [
          { type: "p", text: "On the Home tab, the **Cell Styles** gallery offers ready-made looks: Good, Bad and Neutral (green, red, yellow fills), heading levels, and themed accents. One click applies a complete combination of font, fill and border. Clicking a style again replaces it — formatting stays easily changeable, which is the point." },
          { type: "p", text: "Their real value in offices is **consistency**. If every workbook from your team marks problem cells with the Bad style and headers with Heading 2, anyone opening any file instantly understands the language. Styles can also be redefined once (right-click a style > Modify) and every styled cell in the workbook updates together." },
          { type: "tip", text: "Use the Good/Bad/Neutral styles during data checking: mark verified rows green and doubtful rows yellow. In a shared file, colleagues can see the checking status at a glance." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Adding a currency symbol or the word \"approx\" into cells to make the sheet communicate.", right: "Use number formats and comments — formatting is presentation, data stays pure.", why: "Cells decorated by editing their contents can no longer be added, sorted or matched reliably." },
      { wrong: "Centre-aligning every column, including the numbers.", right: "Keep numbers right-aligned; centre only headings and short labels.", why: "Right-aligned digits line up vertically, so a column of amounts can be compared and totalled by eye at a glance." },
      { wrong: "Merging cells across data ranges to \"make it look tidy\".", right: "Merge only title rows, or use Center Across Selection instead.", why: "Merged cells break sorting, filtering and some copy-paste operations, which costs far more time than the look is worth." },
      { wrong: "Rebuilding the same header formatting on every sheet by hand.", right: "Format one header well, then copy the look with Format Painter.", why: "The painter takes two seconds and guarantees the headers all match exactly." }
    ],
    shortcuts: [
      { keys: "Ctrl + B", action: "Bold the selection" },
      { keys: "Ctrl + I", action: "Italicise the selection" },
      { keys: "Ctrl + U", action: "Underline the selection" },
      { keys: "Ctrl + 1", action: "Open the Format Cells dialog" }
    ],
    exercises: [
      { task: "In the Beginner practice workbook's Exercise sheet, format the header row: bold text and a light fill colour of your choice. Then apply All Borders to the whole data block.", hint: "Select the header cells first, use [[Ctrl]] + [[B]], then the paint bucket. Borders: the grid icon dropdown > All Borders.", solution: "The header stands out with bold on a light fill, and every cell of the data block has a border so the table will look correct when printed." },
      { task: "A description column in the Exercise sheet holds long sentences that are cut off. Apply Wrap Text to that column and auto-fit the row heights.", hint: "Home > Alignment group > Wrap Text. Double-click a row border to auto-fit height.", solution: "With Wrap Text on, each cell shows its full sentence across multiple lines and no text spills into neighbouring columns." },
      { task: "Add a title row above the data: \"Sunrise Traders - Weekly Expense Report\", using Center Across Selection (not Merge) over the data columns. Then copy the finished header formatting to the Solution-style blank area using the Format Painter.", hint: "Type the title in the first column's cell, then [[Ctrl]] + [[1]] > Alignment > Horizontal: Center Across Selection. Double-click Format Painter to paint several areas.", solution: "The title appears centred over the table while the cells stay unmerged, and the Format Painter transfers the header look exactly." },
      { task: "Use the Good, Bad and Neutral cell styles to mark three sample rows in the Exercise sheet: one green (verified), one red (problem), one yellow (check later).", hint: "Select the row, Home > Cell Styles > pick the style.", solution: "Each row takes its colour from one click, demonstrating how styles give a shared sheet a common visual language." }
    ],
    quiz: [
      { q: "Why do borders matter so much in office reports?", options: ["Gridlines on screen do not print, so printed sheets need real borders", "Borders make formulas faster", "Excel deletes data without borders", "Borders are required by Excel to save"], answer: 0, explanation: "The faint grey gridlines are screen-only. A printed or PDF report needs drawn borders to look like a table." },
      { q: "What does Wrap Text do?", options: ["Shows long content on multiple lines inside the cell by growing the row", "Removes extra spaces from text", "Sends the cell content to the next sheet", "Turns text into numbers"], answer: 0, explanation: "Wrap Text breaks long content into multiple lines within the cell, so nothing is cut off or spills over." },
      { q: "Why should merged cells be used sparingly in data areas?", options: ["They break sorting, filtering and some copy-paste operations", "They use too much memory", "They delete hidden values", "They turn off formatting"], answer: 0, explanation: "Merged cells confuse many range operations. Titles are fine; data ranges should stay unmerged." },
      { q: "What does the Format Painter copy?", options: ["Both formatting and contents", "Contents only", "Formulas without formatting", "Formatting only — not the contents"], answer: 3, explanation: "Format Painter transfers the look (font, fill, borders, number format) but never the data." },
      { q: "Which alignment keeps a column of amounts easiest to compare by eye?", options: ["Right-aligned", "Centre-aligned", "Left-aligned", "It does not matter"], answer: 0, explanation: "Right-aligned digits line up vertically, so thousands and decimals sit in the same visual position in every row." }
    ],
    summary: [
      "Standard recipe: bold headers, light fill on the header row, bold totals, one or two fonts only.",
      "Borders are essential for anything printed, because screen gridlines do not print.",
      "Wrap Text keeps long content visible; keep numbers right-aligned and never centre data columns.",
      "Merge & Center only for titles — prefer Center Across Selection so sorting keeps working.",
      "Format Painter (double-click for multiple targets) and Cell Styles give fast, consistent formatting."
    ],
    downloads: [
      { file: "assets/files/level1/beginner-practice.xlsx", label: "Beginner practice workbook", note: "Exercise and Solution sheets covering data entry, formatting, autofill and copy/paste tasks — the Exercise sheet contains the tables to format in this lesson" }
    ],
    keywords: ["cell formatting", "bold italic underline", "font color fill color", "borders in excel", "wrap text", "merge and center", "format painter", "cell styles", "center across selection", "print gridlines"]
  },
  {
    id: "l1-number-formatting",
    title: "Number, Date, Currency and Percent Formatting",
    difficulty: "Beginner",
    duration: 16,
    version: "",
    objectives: [
      "Open and use the Format Cells dialog with Ctrl + 1",
      "Choose between General, Number, Currency and Accounting formats",
      "Apply Percent, Date, Time and Comma styles correctly",
      "Understand that formatting changes appearance, never the stored value"
    ],
    intro: "The number 12500 can legitimately appear as 12,500.00, ₹12,500, 12500.00 and 1.25E+04 — same value, different clothes. Number formatting is how you choose the clothes without touching the value.\n\nThis lesson walks through the Format Cells dialog and the everyday formats office work needs. You will also learn the distinction professionals rely on: the display changes, the stored number never does.",
    sections: [
      {
        heading: "The Format Cells dialog: your control room",
        blocks: [
          { type: "p", text: "Select cells and press **[[Ctrl]] + [[1]]**. The **Format Cells** dialog opens with tabs: Number, Alignment, Font, Border, Fill and Protection. (Alignment, Font, Border and Fill overlap with the Home tab buttons; the **Number** tab is what we are here for.)" },
          { type: "p", text: "On the Number tab, the **Category** list runs down the left: General, Number, Currency, Accounting, Date, Time, Percentage, Custom and more. Pick a category, adjust its options (decimal places, symbol), and the **Sample** box at the top previews exactly how the active cell will look. Click OK to apply." },
          { type: "p", text: "The Home tab also offers quick dropdowns: the Number Format dropdown (it says \"General\" by default) and the style buttons for Currency, Percent and Comma. The dialog is the full toolbox; the buttons are the fast path for the common cases." }
        ]
      },
      {
        heading: "General vs Number",
        blocks: [
          { type: "p", text: "**General** is the default format: Excel shows what you typed, trimming anything unnecessary. Type 45, you see 45. Type 45.5, you see 45.5. General is fine for counts and IDs, but it has no discipline: one column can hold 45, 45.5, 45.5033 and 45000.0001 side by side, and nothing lines up meaningfully." },
          { type: "p", text: "The **Number** category brings order. Choose it and set **decimal places** — usually 2 for money — and tick **Use 1000 Separator**. Now every number in the column shows exactly two decimals and thousands separators: 1,250.00 next to 45.00 next to 12,500.00. The digits line up in tidy columns, which is precisely what an accounts clerk needs to check figures against a paper bill." },
          { type: "grid", title: "Same values — General vs Number format", rows: [
            ["Raw (General)", "Number, 2 decimals"],
            ["12500", "12,500.00"],
            ["45.5", "45.50"],
            ["0.5", "0.50"]
          ] },
          { type: "p", text: "Nothing under the hood changed — the stored values are still 12500, 45.5 and 0.5. Only the display gained discipline." }
        ]
      },
      {
        heading: "Currency vs Accounting",
        blocks: [
          { type: "p", text: "Both formats show a currency symbol, but they behave differently:" },
          { type: "table", headers: ["Format", "Symbol position", "Look and use"], rows: [
            ["Currency", "Right next to the first digit, e.g. 12,500", "Compact; fine for dashboards and personal sheets"],
            ["Accounting", "Pinned to the left edge of the cell", "The office standard for financial reports; symbols align in a column, zeros show as dashes"]
          ] },
          { type: "p", text: "Pick **Accounting** for anything an accounts department will read. In a column of Accounting-formatted cells, every rupee symbol sits in a perfect vertical line at the left edge regardless of amount size, and a value of zero displays as a dash ( - ) instead of 0.00 — a convention accountants read instantly as \"nothing\". Choose the symbol in the dialog: ₹ or $ or none." },
          { type: "p", text: "The Home tab's **Accounting Number Format** button applies your default currency in one click; its dropdown lets you pick other symbols. The **Currency** button does the compact version. Same data, different presentation styles." }
        ]
      },
      {
        heading: "Percent and Comma styles",
        blocks: [
          { type: "p", text: "The **Percent** format multiplies the display by 100 and adds a % sign. A cell holding 0.18 shows 18%; 0.075 shows 7.50%. This is the one formatting trap for beginners, so read carefully: the stored value does not change. 0.18 stays 0.18 underneath; only the display says 18%. If you type **18** into a percent-formatted cell, Excel shows **1800%**, because 18 × 100% = 1800%. When entering percentages directly, either type `18%` (Excel stores 0.18) or type 0.18 and format." },
          { type: "p", text: "Percent formatting earns its keep in real work: VAT/GST at 18%, a savings rate of 15%, a discount of 7.5%, growth of 12% over last month. All of these are stored as 0.18, 0.15, 0.075 and 0.12 and displayed with the % sign." },
          { type: "p", text: "The **Comma Style** button (Home tab, the comma icon) is a one-click version of the Number format with separator and two decimals, no currency symbol. Quick and useful for large plain numbers: 1,250,000.00." },
          { type: "tip", text: "Remember the direction: percent formatting scales the display up by 100. Store the true fraction (0.15), display the readable percent (15%)." }
        ]
      },
      {
        heading: "Date and Time formats",
        blocks: [
          { type: "p", text: "Excel stores every date as a **serial number** — the count of days since 1 January 1900. Today's date is some number in the forties of thousands. That is why dates right-align (they are numbers) and why you can subtract one date from another to get days between. Formatting decides how the serial appears:" },
          { type: "table", headers: ["Format chosen", "Cell displays"], rows: [
            ["Short Date", "14-Mar-2026 (or 14/03/2026 depending on region)"],
            ["Long Date", "Saturday, March 14, 2026"],
            ["Custom dd-mmm-yy", "14-Mar-26"]
          ] },
          { type: "p", text: "Open Format Cells > Date and choose the style your office uses. The **Time** category works the same way: a time is a fraction of a day, so 12:00 noon is 0.5. A cell showing 14-Mar-2026 12:00 holds the day serial plus 0.5. You rarely need those internals — but they explain why dates sort correctly and subtract cleanly, and why \"8 hours\" is often written as 8:00 in timesheets." },
          { type: "warning", text: "If a date shows as a plain number like 46100, the cell lost its date format — reapply a Date format. If a date sits left-aligned, it is text, not a real date; it will not sort or calculate until re-entered in a format Excel recognises." }
        ]
      },
      {
        heading: "Formatting never changes the value — and custom formats prove it",
        blocks: [
          { type: "p", text: "This is the lesson's core principle, worth stating plainly: **number formatting is cosmetic**. Format a cell as currency, and `=B2*2` still uses the same stored number. Format 0.5 as percent, and it still contributes 0.5 — half — to any formula. The formula bar always shows the stored value; the cell shows the formatted costume." },
          { type: "p", text: "Why does this matter in practice? Because you can format fearlessly. Apply Accounting to a whole expense column and nothing breaks. Change from Currency to Comma style and the totals stay identical. Compare this with typing symbols into the cells, which permanently corrupts the data. Formatting is reversible and safe; editing values is not." },
          { type: "grid", title: "One stored value, many costumes", rows: [
            ["Stored value", "Formatted display"],
            ["0.15", "15%"],
            ["0.15", "0.15"],
            ["0.15", "15.00%"]
          ] },
          { type: "p", text: "All three display rows hold exactly 0.15. Any formula using them gets 0.15." },
          { type: "p", text: "The **Custom** category at the bottom of the Number tab proves the point hands-on: it lets you write a format code that changes only the display. Two codes cover most beginner needs:" },
          { type: "table", headers: ["Code", "Displays 45.5 as", "Displays 12500 as", "Meaning"], rows: [
            ["0.00", "45.50", "12500.00", "Always two decimals, no separator"],
            ["#,##0", "46", "12,500", "No decimals, thousands separator (45.5 shows 46 — rounded for display)"],
            ["#,##0.00", "45.50", "12,500.00", "Separator plus two decimals"]
          ] },
          { type: "p", text: "Read the codes simply: `0` is a forced digit, `#` is an optional digit, and the comma inserts the thousands separator. `#,##0.00` says \"group thousands, always show two decimals\". You can build on these later, but even these three make any number column consistent." },
          { type: "tip", text: "Custom formats affect display only, like everything on the Number tab. 45.5 formatted with `#,##0` still stores 45.5 — the formula bar will prove it. Change any custom format back to General or Accounting at any time; the stored value never moved." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Typing 18 into a cell and formatting it as percent, expecting 18%.", right: "Either type 18% directly, or store 0.18 and format as percent.", why: "Percent format multiplies the display by 100. The number 18 displays as 1800%, a classic silent error." },
      { wrong: "Typing currency symbols or dashes into cells so they \"look formatted\".", right: "Type pure values and apply Currency or Accounting formatting.", why: "Typed symbols turn values into text; totals then skip them. Formatting achieves the look while keeping the maths intact." },
      { wrong: "Assuming a Percent-formatted cell holding 15% contains the number 15.", right: "Click the cell: the formula bar shows 0.15.", why: "Formatting never changes the stored value. Formulas reading the cell receive 0.15 — knowing this prevents half of all percent mistakes." },
      { wrong: "Leaving mixed formats in one column: some cells General, some Currency.", right: "Select the whole column and apply one consistent format.", why: "Mixed formats misalign columns and suggest inconsistent data. Uniform formatting is a mark of a reliable sheet." }
    ],
    shortcuts: [
      { keys: "Ctrl + 1", action: "Open the Format Cells dialog" },
      { keys: "Ctrl + Shift + 1", action: "Apply Number format with two decimals (comma style variant)" },
      { keys: "Ctrl + Shift + 4", action: "Apply Currency format" },
      { keys: "Ctrl + Shift + 5", action: "Apply Percent format" }
    ],
    exercises: [
      { task: "In the Beginner practice workbook's Exercise sheet, format the Price column with the Accounting format (two decimals, rupee symbol). Then format the Rate column as Percent and confirm in the formula bar that a cell showing 18% stores 0.18.", hint: "Select the column, press [[Ctrl]] + [[1]], Number tab. Percent is the fifth category down; the formula bar always shows the stored value.", solution: "Accounting aligns every rupee symbol at the left edge and shows zeros as dashes. The percent cell's formula bar shows 0.18 — the stored value never changed." },
      { task: "The Exercise sheet has dates typed correctly but displaying as plain numbers like 46100. Reapply a Date format (Short Date) to that column.", hint: "Select the column, [[Ctrl]] + [[1]], choose Date > Short Date.", solution: "The serial numbers once again display as recognizable dates such as 14-Mar-2026; the underlying values were always dates." },
      { task: "Apply a custom format `#,##0.00` to the Total column in the Exercise sheet, then change your mind and switch it to Comma style. Confirm the totals themselves never changed.", hint: "Format Cells > Custom, type the code in the Type box. Compare formula bar values before and after.", solution: "Both formats display separators and two decimals; switching between them changes nothing in the stored numbers, proving formatting is cosmetic." },
      { task: "In the count column of the Exercise sheet, set the Number format with 0 decimal places and use 1000 separator on. Enter 12500 and 45 and check how they display.", hint: "[[Ctrl]] + [[1]] > Number > set Decimals to 0 and tick the separator box.", solution: "12500 displays as 12,500 and 45 as 45 — clean, consistent, and still pure numbers underneath." }
    ],
    quiz: [
      { q: "What does number formatting change?", options: ["Only the appearance — the stored value stays the same", "The stored value", "The cell's address", "The formulas in other cells"], answer: 0, explanation: "Formatting is cosmetic. The formula bar always shows the true stored value, unchanged by any format." },
      { q: "A cell is formatted as Percent. You type 18 and press Enter. What displays?", options: ["1800%", "18%", "0.18", "An error"], answer: 0, explanation: "Percent display multiplies by 100, so stored 18 shows as 1800%. Type 18% or store 0.18 instead." },
      { q: "Which format pins the currency symbol to the left edge and shows zeros as dashes?", options: ["Currency", "General", "Accounting", "Comma"], answer: 2, explanation: "Accounting aligns symbols in a column at the left and displays zero as a dash — the standard for financial reports." },
      { q: "A cell shows 46100 but you know it should be a date. What happened and what do you do?", options: ["The date format was lost — reapply a Date format", "The date was deleted — retype it", "The cell is corrupt — start a new sheet", "Dates cannot show as numbers"], answer: 0, explanation: "Dates are stored as serial numbers. When the Date format is stripped, the serial shows. Reapply the format — the value was never wrong." },
      { q: "What does the custom format #,##0.00 do to the value 45.5?", options: ["Displays 45.50 while storing 45.5", "Changes the value to 45.50", "Rounds the stored value to 46", "Converts it to text"], answer: 0, explanation: "The code forces two displayed decimals with a thousands separator. The stored value remains 45.5." }
    ],
    summary: [
      "Ctrl + 1 opens Format Cells; the Number tab is the control room for how values display.",
      "General is default and undisciplined; Number format adds fixed decimals and thousands separators.",
      "Accounting (symbols aligned left, zeros as dashes) is the financial-report standard; Currency is the compact variant.",
      "Percent displays values times 100 — store 0.15 to show 15%.",
      "Formatting never changes the stored value; custom codes like #,##0.00 give precise display control."
    ],
    downloads: [
      { file: "assets/files/level1/beginner-practice.xlsx", label: "Beginner practice workbook", note: "Exercise and Solution sheets covering data entry, formatting, autofill and copy/paste tasks — the Exercise sheet has price, rate, date and total columns to format in this lesson" }
    ],
    keywords: ["number formatting", "format cells ctrl 1", "currency vs accounting", "percent format", "date format excel", "comma style", "custom number format", "0.00", "#,##0", "stored value vs display"]
  },
  {
    id: "l1-basic-calculations",
    title: "Basic Calculations and Your First Formulas",
    difficulty: "Beginner",
    duration: 18,
    version: "",
    objectives: [
      "Write formulas that start with = and use +, -, *, / correctly",
      "Predict results using Excel's order of operations (PEMDAS)",
      "Calculate a total with =B2*C2 and a sum with =SUM(...)",
      "Edit formulas and display all formulas on a sheet"
    ],
    intro: "This is the lesson where Excel stops being a typewriter and starts being a calculator. Everything from here on — functions, references, reports — builds on one small rule: a formula begins with the equals sign.\n\nYou will write your first real formulas, understand why Excel sometimes gives a \"surprising\" answer (order of operations), and meet `SUM`, the most used function in the world of work.",
    sections: [
      {
        heading: "Every formula starts with =",
        blocks: [
          { type: "p", text: "Type a calculation like `2+3` into a cell and Excel shows you the **text** \"2+3\", left-aligned. Type `=2+3` and Excel shows **5**. That single equals sign is the switch: it tells Excel \"do not store this as text — calculate it\"." },
          { type: "p", text: "After the equals sign you can combine numbers, cell addresses, and operators. Type `=A2+B2` and Excel fetches whatever is in those cells **right now** and adds them. This is the crucial leap: formulas reference cells, not fixed numbers. Change `A2` and the formula's result updates instantly. Your sheet becomes a live machine, not a snapshot." },
          { type: "grid", title: "Text vs formula in a cell", rows: [
            ["", "A", "B"],
            ["1", "2+3", "5"],
            ["2", "stored as text", "=2+3"]
          ] },
          { type: "p", text: "In the grid, `A1` holds the characters 2+3 (a left-aligned label), while `B2` holds `=2+3` and therefore shows 5. The formula bar proves which is which — click each cell and look up." }
        ]
      },
      {
        heading: "The arithmetic operators",
        blocks: [
          { type: "table", headers: ["Operator", "Does", "Example", "Result"], rows: [
            ["+", "Add", "=10+4", "14"],
            ["-", "Subtract", "=10-4", "6"],
            ["*", "Multiply", "=10*4", "40"],
            ["/", "Divide", "=10/4", "2.5"],
            ["^", "Raise to a power", "=10^2", "100"],
            ["%", "Percent of the preceding number", "=200*10%", "20"]
          ] },
          { type: "p", text: "Note the symbols Excel expects: `*` for multiply (not x), `/` for divide, `^` for powers. The percent shortcut `=200*10%` works because 10% means 0.1 — you learned in the last lesson that percent is just display of a fraction." },
          { type: "p", text: "Spaces inside formulas are optional: `=A2+B2` and `=A2 + B2` behave identically. Use whatever keeps the formula readable to you." }
        ]
      },
      {
        heading: "Order of operations: why =2+3*4 is not 20",
        blocks: [
          { type: "p", text: "Excel does not calculate strictly left to right. It follows the standard mathematical **order of operations**, remembered as **PEMDAS**:" },
          { type: "ol", items: [
            "**P**arentheses ( ) first",
            "**E**xponents ^ (powers)",
            "**M**ultiply * and **D**ivide / (left to right, equal rank)",
            "**A**dd + and **S**ubtract - (left to right, equal rank)"
          ] },
          { type: "p", text: "So `=2+3*4` calculates the multiplication first: 3*4 = 12, then 2+12 = **14** — not 20. And `=(2+3)*4` forces the addition first with parentheses: 5*4 = **20**. Parentheses are your steering wheel: they cost nothing and remove all doubt." },
          { type: "table", headers: ["Formula", "Excel computes", "Result"], rows: [
            ["=2+3*4", "2 + (3*4)", "14"],
            ["=(2+3)*4", "(2+3) first, then *4", "20"],
            ["=10-4-3", "(10-4)-3, left to right", "3"],
            ["=2^3*2", "(2^3) first, then *2", "16"],
            ["=200*10%", "200 * 0.1", "20"]
          ] },
          { type: "p", text: "Real-office example: a bill of 1250 with 18% tax. The total is `=1250*1.18` (multiply once by 1.18) or `=1250+1250*0.18` (amount plus the tax part). Both give 1475 — and in the second one, the multiplication happens before the addition, exactly as PEMDAS promises. If you wrote `=1250+1250*0.18` hoping for a different grouping, add parentheses: `=1250+(1250*0.18)` makes the intent explicit." },
          { type: "warning", text: "A missing parenthesis is the most common silent formula bug for beginners. When mixing + - with * /, add parentheses even when you believe they are unnecessary — they cost nothing and make the meaning obvious to the next reader." }
        ]
      },
      {
        heading: "Your first real formula: total = price * qty",
        blocks: [
          { type: "p", text: "Time to do actual office work. A sales list for Sunrise Traders has prices and quantities. The Total column should be price times quantity — expressed with cell addresses so it updates automatically:" },
          { type: "grid", title: "Building =B2*C2 in D2", rows: [
            ["", "B", "C", "D"],
            ["1", "Price", "Qty", "Total"],
            ["2", "45", "10", "=B2*C2"],
            ["3", "60", "5", "=B3*C3"],
            ["4", "12500", "1", "=B4*C4"]
          ] },
          { type: "steps", items: [
            "Click `D2` (the first empty Total cell).",
            "Type **=**, then click cell `B2` — Excel writes the address for you.",
            "Type **\***, then click `C2`.",
            "Press **[[Enter]]**. D2 shows 450, the product of 45 and 10.",
            "Click `D2` again: the formula bar shows `=B2*C2`."
          ] },
          { type: "p", text: "Notice the click-to-reference trick: instead of typing B2, you click the cell while the formula is open. Excel inserts the address. Beginners who type addresses by hand make more typos (like `=B2*C3`), which produces plausible-looking wrong numbers. Clicking the cells makes that impossible." },
          { type: "p", text: "Now the payoff: change `C2` from 10 to 20 and watch `D2` jump to 900 with no action from you. The formula is permanently wired to those cells." }
        ]
      },
      {
        heading: "=SUM: adding ranges the professional way",
        blocks: [
          { type: "p", text: "To add several cells you *could* write `=D2+D3+D4`. It works — until the list grows to 30 rows and you are writing 30 additions. The **SUM** function does the job:" },
          { type: "formula", name: "SUM", syntax: "=SUM(number1, [number2], ...)", example: "=SUM(D2:D4)", explanation: "Adds all the numbers in the range D2:D4. The colon range notation from the selection lesson does the counting for you." },
          { type: "p", text: "Type `=SUM(D2:D4)` under the totals and Excel returns 13550. If a new row is inserted inside the range later, SUM stretches to include it automatically — a quiet advantage over hand-written additions, which never grow." },
          { type: "grid", title: "SUM under a column of totals", rows: [
            ["", "D"],
            ["4", "12500"],
            ["5", "=SUM(D2:D4)"],
            ["", ""]
          ] },
          { type: "p", text: "Compare the two approaches honestly: `=D2+D3` is fine for two cells you will never extend; `=SUM(D2:D100)` is the office answer for anything longer, and it ignores text accidentally mixed into the range instead of producing errors. When in doubt, choose SUM." },
          { type: "tip", text: "The AutoSum button (Home tab, the sigma symbol, or [[Alt]] + [[=]]) writes =SUM() for you: select the empty cell under a column of numbers, press AutoSum, check the proposed range, press Enter." }
        ]
      },
      {
        heading: "Editing formulas and showing them all",
        blocks: [
          { type: "p", text: "Formulas edit like any other content: click the cell, look at the formula bar to see the real instruction, then edit there — or press **F2** to edit inside the cell. While editing, Excel colour-highlights the referenced cells on screen, so you can visually confirm the formula is touching the right cells before pressing Enter." },
          { type: "p", text: "To see **every** formula on the sheet at once, press **[[Ctrl]] + [[`]]** (the backquote key, left of 1 on most keyboards). The grid switches from results to formulas: every cell shows its instruction. This is the professional's X-ray view for checking a sheet before sending it. Press the same shortcut again to return to normal view." },
          { type: "grid", title: "Ctrl + ` view: formulas instead of results", rows: [
            ["", "D"],
            ["1", "Total"],
            ["2", "=B2*C2"],
            ["3", "=B3*C3"],
            ["4", "=SUM(D2:D3)"]
          ] },
          { type: "p", text: "If a formula shows an error code like `#NAME?` (often from a misspelled function name like `=SUN(...)`) or `#DIV/0!` (dividing by zero or an empty cell), do not panic — read the code, check the formula bar, fix the reference or spelling, and the result returns." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Typing a calculation without the equals sign and wondering why it shows as text.", right: "Start with = — every formula, no exceptions.", why: "Without =, Excel stores the characters as text: left-aligned, never calculated, ignored by SUM." },
      { wrong: "Writing =2+3*4 and expecting 20.", right: "Use parentheses when you need a specific order: =(2+3)*4.", why: "Excel follows PEMDAS — multiplication before addition. Unparenthesised mixed operations give correct-but-surprising results." },
      { wrong: "Building long chains like =D2+D3+D4+D5+D6 for a running list.", right: "Use =SUM(D2:D6).", why: "SUM survives inserted rows and range growth; hand-built chains must be re-edited every time the list changes." },
      { wrong: "Typing cell addresses by hand into formulas and hitting the wrong one (like =B2*C3).", right: "Click the cells while the formula is open.", why: "Clicking inserts the exact address and highlights it on screen, eliminating reference typos that produce plausible wrong totals." }
    ],
    shortcuts: [
      { keys: "F2", action: "Edit the formula in the active cell" },
      { keys: "Ctrl + `", action: "Show all formulas on the sheet (press again to revert)" },
      { keys: "Alt + =", action: "AutoSum the adjacent column or row" },
      { keys: "Esc", action: "Cancel the formula being typed" }
    ],
    exercises: [
      { task: "In the formula-basics workbook's Exercise sheet, write formulas in the Total column so each row shows price times quantity. Use click-to-reference, not typing the addresses.", hint: "Click the Total cell, type =, click the Price cell, type *, click the Qty cell, Enter.", solution: "Each Total cell holds a formula like =B2*C2. Editing any price or quantity updates that row's total instantly." },
      { task: "Add a Grand Total cell under the Total column using =SUM. Then insert a new expense row inside the range and confirm the SUM still covers everything.", hint: "Type =SUM( then select the range with the mouse, close the parenthesis, Enter. Insert rows by right-clicking a row number > Insert.", solution: "SUM written over a range automatically includes inserted rows inside it — the grand total stays correct without editing the formula." },
      { task: "Predict first, then verify: without typing, decide what =10-4-3, =2^3*2 and =(10+5)*2 return. Enter them in the practice area and check.", hint: "Apply PEMDAS: equal-rank operations run left to right; ^ beats *; parentheses win over everything.", solution: "3, 16 and 30. If any guess differed, press Ctrl + ` to view the formulas and re-walk the order of operations." },
      { task: "One row in the Exercise sheet has a typo: its formula multiplies the wrong quantity cell. Press Ctrl + ` to expose all formulas, find the odd one, and fix it.", hint: "Look for the row whose formula does not follow the =B*pattern of its neighbours. Fix with F2 or in the formula bar.", solution: "In formula view the inconsistent reference is obvious against the pattern. Correcting it restores the right total — a habit you will use on every real sheet you inherit." }
    ],
    quiz: [
      { q: "What must every Excel formula begin with?", options: ["A cell address", "The SUM function", "A plus sign", "The = sign"], answer: 3, explanation: "The equals sign switches Excel from storing text to calculating. =2+3 shows 5; 2+3 shows the text 2+3." },
      { q: "What does =2+3*4 return?", options: ["14", "20", "24", "An error"], answer: 0, explanation: "PEMDAS: multiplication first (3*4=12), then addition (2+12=14). Parentheses — =(2+3)*4 — would give 20." },
      { q: "Why prefer =SUM(D2:D100) over =D2+D3+...+D100?", options: ["SUM stretches over ranges and survives inserted rows", "SUM is faster to display", "Both are identical in every way", "SUM converts text to numbers"], answer: 0, explanation: "A range-based SUM automatically includes rows inserted inside it, while hand-built addition chains must be edited by hand each time." },
      { q: "You click a Total cell and the formula bar shows =B2*C3, but row 2's quantity is in C2. What is the problem?", options: ["The formula references the wrong row", "Nothing — C3 is fine", "The formula needs = instead of *", "The file must be re-saved"], answer: 0, explanation: "A reference typo points the formula at the wrong cell. Click-to-reference while building formulas prevents exactly this." },
      { q: "Which shortcut displays every formula on the sheet instead of results?", options: ["Ctrl + `", "Ctrl + F", "F9", "Alt + F"], answer: 0, explanation: "Ctrl + ` (backquote) toggles the X-ray view showing all formulas — the standard way to audit a sheet before sharing." }
    ],
    summary: [
      "Every formula starts with =; formulas reference cells, so results update when data changes.",
      "Operators: + - * / ^ %; Excel follows PEMDAS, so use parentheses to control order.",
      "Total = price * qty is written =B2*C2 — build formulas by clicking cells, not typing addresses.",
      "=SUM(range) is the professional way to add a column; AutoSum (Alt + =) writes it for you.",
      "Edit formulas via the formula bar or F2; press Ctrl + ` to expose all formulas for checking."
    ],
    downloads: [
      { file: "assets/files/level1/formula-basics.xlsx", label: "Formula basics workbook", note: "Exercise and Solution sheets with basic calculations, references and the weekly budget project — the Exercise sheet contains the price/qty table used in this lesson" }
    ],
    keywords: ["first formula excel", "equals sign", "arithmetic operators", "order of operations", "PEMDAS", "SUM function", "autosum", "show formulas", "price times quantity", "basic calculations"]
  },
  {
    id: "l1-cell-references",
    title: "Relative and Absolute References",
    difficulty: "Beginner",
    duration: 18,
    version: "",
    objectives: [
      "Explain how relative references change when copied",
      "Lock a cell with absolute references using $",
      "Toggle reference types with F4 while building a formula",
      "Choose the right reference type for rates, tax and conversion cells"
    ],
    intro: "You already know formulas adjust when copied — `=B2*C2` becomes `=B3*C3` in the next row. That automatic shifting is brilliant most of the time and wrong some of the time. This lesson teaches you the other half: how to nail part of a formula to one cell so it **never** moves.\n\nThis is the skill that separates people who copy a tax rate formula to 500 rows and get 500 wrong answers, from people who get it right the first time.",
    sections: [
      {
        heading: "Relative references: shifting with the copy",
        blocks: [
          { type: "p", text: "A normal reference like `B2` is **relative**. It does not mean \"cell B2\" so much as \"the cell two columns left of me, same row\" — direction and distance, not a fixed address. When you copy the formula elsewhere, the relationship travels with it:" },
          { type: "grid", title: "What each copied formula becomes", rows: [
            ["Copied to", "Formula shown", "It really computes"],
            ["D2 (original)", "=B2*C2", "B2 times C2"],
            ["D3 (one row down)", "=B3*C3", "B3 times C3"],
            ["E2 (one col right)", "=C2*D2", "C2 times D2"]
          ] },
          { type: "p", text: "Copy down, row numbers shift. Copy right, column letters shift. This is exactly what you want for row-by-row calculations like price times quantity — the previous lesson's Totals column worked because of relative references." }
        ]
      },
      {
        heading: "Where relative references break: the fixed rate cell",
        blocks: [
          { type: "p", text: "Now a realistic problem. Sunrise Traders adds 18% tax to each sale, and the rate sits in one cell, say `G1`, so it can be updated in one place. You write the first row's formula as `=F2*G1` (amount times rate) and copy it down. Disaster:" },
          { type: "grid", title: "Copying =F2*G1 down — watch the rate cell drift", rows: [
            ["Row", "Formula after copy", "Problem"],
            ["2", "=F2*G1", "Correct — 9500 times G1"],
            ["3", "=F3*G2", "Multiplies by G2 — wrong cell!"],
            ["4", "=F4*G3", "Multiplies by G3 — wrong cell!"]
          ] },
          { type: "p", text: "The `F2` part shifting to `F3`, `F4` is right — each row must use its own amount. But `G1` shifting too means rows 3 and 4 multiply by empty or wrong cells, and the results are garbage. You need the amount reference to stay relative and the rate reference to stay **absolute** — nailed to `G1` no matter where the formula lands." }
        ]
      },
      {
        heading: "Absolute references: the $ lock",
        blocks: [
          { type: "p", text: "Put **$** signs in a reference and it stops shifting. `$G$1` reads as \"column G, row 1 — permanently\". Copy the formula anywhere on the sheet and `$G$1` still points at G1. The $ works like a padlock on both the column letter and the row number." },
          { type: "grid", title: "The fixed formula: =F2*$G$1", rows: [
            ["Row", "Formula after copy", "Result"],
            ["2", "=F2*$G$1", "Row 2 amount times the rate"],
            ["3", "=F3*$G$1", "Row 3 amount times the rate"],
            ["4", "=F4*$G$1", "Row 4 amount times the rate"]
          ] },
          { type: "p", text: "Now every row multiplies its own amount by the single tax rate cell. Better still, the real power move: update `G1` from 18% to 21% and **every row recalculates instantly**. One cell, one edit, 500 rows corrected. That is why professionals never type tax rates into formulas directly — they put the rate in one cell and lock it." },
          { type: "p", text: "The same pattern handles any fixed factor: a USD to INR conversion rate, a standard discount, a store's service charge. Formula shape is always the same: `=amount * $rate$` — relative for the amount, absolute for the rate." }
        ]
      },
      {
        heading: "F4: the one-key toggle",
        blocks: [
          { type: "p", text: "Typing $ signs by hand is slow and error-prone. Excel gives you a key: while typing or editing a formula, click on a reference (or just after typing it) and press **F4**. The reference cycles through its four states with each press:" },
          { type: "table", headers: ["Press F4", "Reference becomes", "Meaning"], rows: [
            ["1st", "$G$1", "Fully absolute — column and row locked"],
            ["2nd", "G$1", "Row locked, column relative"],
            ["3rd", "$G1", "Column locked, row relative"],
            ["4th", "G1", "Back to fully relative"]
          ] },
          { type: "steps", items: [
            "Click the cell with the formula and press F2 to edit (or start typing a new formula).",
            "Click on the reference you want to lock, e.g. G1.",
            "Press **F4** once — it becomes `$G$1`.",
            "Press Enter. Copy the formula down — the lock holds."
          ] },
          { type: "p", text: "The two-in-between states (`G$1`, `$G1`) are called **mixed references**. They lock one dimension only. A classic use: a multiplication table, where the top row of factors must lock the row (`B$1`) and the left column must lock the column (`$A2`). You will meet mixed references again in later levels; for now, know they exist and that F4 walks through them." },
          { type: "tip", text: "Build the habit now: whenever you reference a rate, a total, or any single fixed cell inside a formula you plan to copy, press F4 immediately after entering that reference. It becomes reflex, and reflexes like this are what make Excel work feel fast." }
        ]
      },
      {
        heading: "A conversion-rate example you will actually use",
        blocks: [
          { type: "p", text: "A store receives invoices in USD and reports in rupees. The exchange rate — say 83.5 — lives in `B1`, formatted clearly with a label. The conversion column needs `=C4*$B$1` and nothing else:" },
          { type: "grid", title: "Invoices converted with one locked rate", rows: [
            ["", "B", "C", "D"],
            ["1", "Rate: 83.5", "", ""],
            ["2", "Invoice", "USD", "INR"],
            ["3", "INV-101", "250", "=C3*$B$1"],
            ["4", "INV-102", "480", "=C4*$B$1"]
          ] },
          { type: "p", text: "Copy `D3` down as far as the invoices go; every row uses its own USD amount and the single locked rate. When the rate changes next month, edit `B1` once and the whole column updates. No find-and-replace, no retyping, no formula errors." },
          { type: "p", text: "Notice also the label in `B1` — \"Rate: 83.5\" in our sketch would actually be split across two cells in practice: `A1` holding the text \"Rate\" and `B1` holding the pure number 83.5, so the formula can reference the number cleanly. Keep labels in their own cells, values in theirs." }
        ]
      },
      {
        heading: "Quick decision guide: do I need $ here?",
        blocks: [
          { type: "table", headers: ["Situation", "Reference type", "Why"], rows: [
            ["Total = this row's price times this row's qty", "Relative (`B2*C2`)", "Each row must use its own cells"],
            ["Every row multiplied by one tax-rate cell", "Absolute (`$G$1`)", "All rows share one fixed cell"],
            ["Every row times one conversion rate", "Absolute (`$B$1`)", "Same reason — one source of truth"],
            ["Summing a fixed grand total repeatedly", "Absolute (`$D$20`)", "The total lives in one place"],
            ["Copying along a row that must always read one column", "Mixed (`$A2`)", "Lock the column, let the row move"]
          ] },
          { type: "p", text: "The test in one sentence: **will this reference need to move when I copy, or must it stay?** Move → relative. Stay → absolute ($ on both parts). When unsure, build the formula, copy it one or two cells, and check the results with your own eyes before trusting it at scale." },
          { type: "warning", text: "Test copies before mass-filling. Write the formula, copy it to the next two rows, and check the numbers by hand. If row 2 was right and rows 3-4 are wrong, you have a reference problem — fix it with F4 before copying to 500 rows." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Copying =F2*G1 down a column and trusting every result.", right: "Lock the fixed cell: =F2*$G$1.", why: "The relative G1 drifts to G2, G3... on each copy, multiplying by wrong or empty cells. The results look plausible but are wrong." },
      { wrong: "Typing the tax rate 18% directly inside every formula, like =F2*1.18.", right: "Put the rate in one labelled cell and reference it absolutely: =F2*$G$1.", why: "A rate change then needs one edit instead of editing every formula — and mixed old rates scattered in formulas are an audit nightmare." },
      { wrong: "Locking everything by habit: =SUM($D$2:$D$20) in a formula you copy down.", right: "Lock only what must stay fixed; leave row-wise references relative.", why: "Over-locking makes copies keep pointing at the same rows, so every copied row returns the identical total." },
      { wrong: "Copying a formula to 500 rows and spot-checking nothing.", right: "Copy to two or three rows first and verify the results manually.", why: "Reference errors repeat identically on every row. Catching them at row 3 takes seconds; catching them at row 500 takes an audit." }
    ],
    shortcuts: [
      { keys: "F4 (while editing a formula)", action: "Cycle a reference through $G$1, G$1, $G1, G1" },
      { keys: "F2", action: "Edit the formula so you can change references" },
      { keys: "Ctrl + `", action: "View all formulas to check references at a glance" }
    ],
    exercises: [
      { task: "In the formula-basics workbook's Exercise sheet, find the tax rate cell (18% as 0.18). Write a formula in the Tax column that multiplies each row's amount by the rate, and copy it down the whole column. Check three rows against hand calculation.", hint: "Type =, click the amount cell, type *, click the rate cell and press F4 once, Enter. Then copy down.", solution: "Each cell reads like =C4*$G$1 — the amount reference shifts with the row while the rate stays locked on one cell." },
      { task: "Change the rate cell from 0.18 to 0.21 and watch the Tax column. Then change it back. Confirm you never touched any formula.", hint: "Only edit the single rate cell.", solution: "Every row recalculates the moment the rate cell changes — demonstrating why the rate lives in one absolute-referenced cell." },
      { task: "The Exercise sheet has a USD amount column and one conversion rate cell. Build the INR column with a locked-rate formula and fill it down for all invoice rows.", hint: "Same pattern: =USD_cell*$rate_cell with F4 on the rate.", solution: "The INR column converts every row using the single rate; editing the rate reconverts the entire column instantly." },
      { task: "Practice F4: in any empty cell, type =, click any cell, then press F4 four times, watching the reference cycle. Finish with the fully relative form and press Esc.", hint: "Press F2 first if editing an existing cell; in a new formula just click the reference.", solution: "The reference cycles $G$1 > G$1 > $G1 > G1, then wraps. Esc cancels without entering anything — pure muscle-memory practice." }
    ],
    quiz: [
      { q: "You copy =B2*C2 from D2 down to D5. What does D5 contain?", options: ["=B5*C5", "=B2*C2", "=B5*C2", "=B2*C5"], answer: 0, explanation: "Relative references shift with the copy: each dimension moves by the same offset, so both row numbers become 5." },
      { q: "What does $G$1 mean in a formula?", options: ["Column G and row 1 are both locked — the reference never shifts when copied", "The cell is protected from editing", "G1 contains currency", "The formula is an error"], answer: 0, explanation: "The $ signs lock the column letter and the row number. Copies of the formula keep pointing at G1." },
      { q: "Which key toggles a reference between relative and absolute while you edit a formula?", options: ["F2", "F4", "F9", "$"], answer: 1, explanation: "F4 cycles through $G$1, G$1, $G1 and G1 with each press while the cursor sits on that reference." },
      { q: "Your formula =F3*G2 (copied down from =F2*G1) gives a wrong result. What is the fix?", options: ["Lock the rate reference: =F3*$G$1", "Delete and retype the formula", "Change G2 to G1 by hand in every row", "Use SUM instead"], answer: 0, explanation: "The rate cell must not drift with the copy. $G$1 keeps every row's formula anchored to the single rate cell." },
      { q: "Every row of a column must multiply by a conversion rate stored in B1. What should the formula look like before copying down?", options: ["=C4*$B$1", "=C4*B1", "=$C$4*$B$1", "=C4+$B$1"], answer: 0, explanation: "The row's own amount stays relative (C4 shifts per row) while the rate is locked ($B$1). Locking both, or neither, gives wrong copies." }
    ],
    summary: [
      "Relative references shift with copies — perfect for row-by-row maths like =B2*C2.",
      "Fixed cells (tax rate, conversion rate) must be locked with $: $G$1 never moves when copied.",
      "F4 cycles a reference through absolute, mixed and relative states while you edit.",
      "The professional pattern: data in one labelled rate cell, formulas like =amount*$rate$.",
      "Test copies on two or three rows and verify before filling hundreds."
    ],
    downloads: [
      { file: "assets/files/level1/formula-basics.xlsx", label: "Formula basics workbook", note: "Exercise and Solution sheets with basic calculations, references and the weekly budget project — includes the tax-rate and conversion-rate columns used in this lesson" }
    ],
    keywords: ["relative reference", "absolute reference", "dollar sign excel", "f4 toggle reference", "mixed reference", "tax rate formula", "conversion rate excel", "locked cell reference", "copy formulas"]
  },
  {
    id: "l1-keyboard-shortcuts",
    title: "Essential Keyboard Shortcuts",
    difficulty: "Beginner",
    duration: 14,
    version: "",
    objectives: [
      "Explain why keyboard speed matters in office work",
      "Use the core navigation, selection and editing shortcuts without thinking",
      "Apply formatting and file shortcuts in daily work",
      "Enter today's date and time instantly"
    ],
    intro: "Watch an experienced Excel user work: hands barely leave the keyboard. Data flies in, ranges select themselves, formats snap on. That speed is not talent — it is two dozen shortcuts used until they are automatic.\n\nThis lesson collects the shortcuts that matter most in real office jobs, grouped by purpose, with honest advice on which to learn first. You already met many of them in earlier lessons; now they live in one reference table.",
    sections: [
      {
        heading: "Why shortcuts matter in office jobs",
        blocks: [
          { type: "p", text: "Consider one honest calculation. A data entry clerk copies and pastes perhaps 150 times a day. Reaching for the mouse each time costs about 4 seconds (move hand, aim, click, move back); [[Ctrl]] + [[C]] then [[Ctrl]] + [[V]] costs about 1. That is roughly **7 minutes saved every single day** on pasting alone — and the same arithmetic repeats for saving, editing, navigating and formatting. Together, a shortcut-fluent user saves 30-60 minutes a day over a mouse-only colleague." },
          { type: "p", text: "The second benefit is less measurable but real: **flow**. Switching to the mouse breaks your reading rhythm; you lose your place in a long list. Keyboard workers keep eyes on the data and make fewer entry mistakes. In job interviews and probation periods, supervisors notice within days who works with both hands on the keyboard." },
          { type: "quote", text: "You do not need 200 shortcuts. You need about 25, used until you stop thinking about them." }
        ]
      },
      {
        heading: "Navigation: moving through the sheet",
        blocks: [
          { type: "table", headers: ["Shortcut", "Action"], rows: [
            ["Arrow keys", "Move one cell in any direction"],
            ["Ctrl + Arrow", "Jump to the edge of the data block in that direction"],
            ["Ctrl + Home", "Jump to cell A1"],
            ["Ctrl + End", "Jump to the last used cell of the sheet"],
            ["Page Down / Page Up", "Move one screen down / up"],
            ["Ctrl + Page Down / Up", "Switch to the next / previous worksheet tab"]
          ] },
          { type: "p", text: "The navigation pair to internalise first is **[[Ctrl]] + [[Arrow]]** and **[[Ctrl]] + [[Home]]**: dive to the bottom of a 3,000-row list, and return home. **Ctrl + Page Up/Down** flips between sheets — the fastest way to hop from a data sheet to a summary sheet and back." }
        ]
      },
      {
        heading: "Selection: grabbing ranges fast",
        blocks: [
          { type: "table", headers: ["Shortcut", "Action"], rows: [
            ["Shift + Arrow", "Extend the selection one cell at a time"],
            ["Ctrl + Shift + Arrow", "Extend the selection to the end of the data block"],
            ["Ctrl + A", "Select the entire sheet (press once more for the data region)"],
            ["Shift + Space", "Select the entire current row"],
            ["Ctrl + Space", "Select the entire current column"]
          ] },
          { type: "p", text: "**Ctrl + Shift + Arrow** is the selection workhorse: click the first cell of a column, press it, and the whole column of data selects instantly — ready for formatting, copying or a SUM. **Shift + Space** and **Ctrl + Space** select whole rows and columns, which is exactly what you need before inserting, hiding or formatting them." }
        ]
      },
      {
        heading: "Editing: enter, fix, move data",
        blocks: [
          { type: "table", headers: ["Shortcut", "Action"], rows: [
            ["F2", "Edit the active cell in place"],
            ["Ctrl + C / Ctrl + X / Ctrl + V", "Copy / cut / paste"],
            ["Ctrl + Alt + V", "Paste Special (values, formats, transpose)"],
            ["Ctrl + Z", "Undo the last action"],
            ["Ctrl + Y", "Redo"],
            ["Delete", "Clear the contents of selected cells"],
            ["Ctrl + ;", "Insert today's date into the active cell"],
            ["Ctrl + Shift + ;", "Insert the current time into the active cell"]
          ] },
          { type: "p", text: "Two gems deserve special attention. **[[Ctrl]] + [[;]]** stamps **today's date** into the cell as a real date value — invaluable for logging received payments, marking \"checked on\" columns and daily reports. **[[Ctrl]] + [[Shift]] + [[;]]** stamps the current time the same way. Both produce true date/time values that sort and calculate, unlike typed text." },
          { type: "tip", text: "A daily-report combination: select the log's date cell, press Ctrl + ; then Tab into the next column — the date arrives in under a second, correctly formatted, every day, no typing errors possible." }
        ]
      },
      {
        heading: "Formatting and files",
        blocks: [
          { type: "table", headers: ["Shortcut", "Action"], rows: [
            ["Ctrl + B / I / U", "Bold / italic / underline"],
            ["Ctrl + 1", "Open the Format Cells dialog"],
            ["Alt + H", "Open the Home tab (then follow KeyTips for any command)"],
            ["Ctrl + S", "Save the workbook"],
            ["Ctrl + N", "New blank workbook"],
            ["Ctrl + P", "Open Print"],
            ["Ctrl + F", "Find (search the sheet)"],
            ["Ctrl + `", "Show all formulas"]
          ] },
          { type: "p", text: "**[[Alt]] + [[H]]** deserves a note: it opens the Home tab and reveals KeyTips — little letters on every button — so you can press, say, then 1 for fill colour options, entirely from the keyboard. Every Ribbon command has such a keyboard path; [[Alt]] is your entry point when a button has no direct shortcut." },
          { type: "p", text: "And the two that protect your livelihood: **[[Ctrl]] + [[S]]** (save constantly) and **[[Ctrl]] + [[Z]]** (undo mistakes instantly). Between those two, almost no accident in Excel is fatal." }
        ]
      },
      {
        heading: "How to actually learn them",
        blocks: [
          { type: "p", text: "Reading a shortcut table teaches nothing; fingers teach. A method that works in one week:" },
          { type: "steps", items: [
            "Pick **three** shortcuts — for example Ctrl + C/V, Ctrl + Z and Ctrl + Home.",
            "Ban the mouse for those actions only. Every time you would normally click, use the keys — even if slower at first.",
            "Once the three feel automatic (about two days of real work), add three more.",
            "Repeat. In two weeks you carry 20+ shortcuts without effort."
          ] },
          { type: "p", text: "The mistake to avoid: trying to memorise the whole table in one sitting. Nobody retains that. Shortcuts are motor skills, like typing — they arrive through repetition during real work, not through reading." },
          { type: "tip", text: "Print the shortcut table from this lesson (or keep it on a second screen) for your first practice week. Cross out each shortcut the day it starts feeling automatic — the growing list of crossed-out keys is motivating." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Trying to learn 50 shortcuts in one evening.", right: "Adopt three at a time and use them in real work until automatic.", why: "Shortcuts are motor skills built by repetition, not facts memorised by reading. Three at a time sticks; fifty at once evaporates." },
      { wrong: "Typing today's date by hand into daily log columns.", right: "Press Ctrl + ; to stamp the real current date.", why: "The stamp is instant, always correct, and enters a true date value that sorts and calculates — typed text dates often fail both tests." },
      { wrong: "Reaching for the mouse to save, undo and paste all day.", right: "Use Ctrl + S, Ctrl + Z, Ctrl + C/V until they need no thought.", why: "These are the highest-frequency actions in Excel; automating them alone saves real minutes every day." }
    ],
    shortcuts: [
      { keys: "Ctrl + S", action: "Save the workbook" },
      { keys: "Ctrl + Z / Ctrl + Y", action: "Undo / redo" },
      { keys: "Ctrl + C / X / V", action: "Copy / cut / paste" },
      { keys: "Ctrl + Arrow keys", action: "Jump to the data block edge" },
      { keys: "Ctrl + Shift + Arrow keys", action: "Select to the data block edge" },
      { keys: "Ctrl + Home", action: "Jump to A1" },
      { keys: "F2", action: "Edit the active cell" },
      { keys: "Ctrl + 1", action: "Format Cells dialog" }
    ],
    exercises: [
      { task: "In the Beginner practice workbook's Exercise sheet, select the first cell of the price column and extend the selection to the bottom using only the keyboard. Then undo the selection with an arrow key press and return to A1.", hint: "Ctrl + Shift + Down selects to the data end; pressing an arrow key collapses the selection; Ctrl + Home returns home.", solution: "The column selects in one press and the active cell returns to A1 — the core navigation-selection cycle used dozens of times daily." },
      { task: "In an empty log area of the Exercise sheet, enter today's date with Ctrl + ; in the first cell and the current time with Ctrl + Shift + ; in the next. Then check both are right-aligned (true date/time values).", hint: "Press the shortcut and Enter — no typing. Right alignment confirms Excel stored a real date or time.", solution: "The stamps appear instantly and right-align, proving they are real date and time values, not text. Try sorting a column of them — they order correctly." },
      { task: "Format the header row of the Exercise sheet using only the keyboard: bold with Ctrl + B, then open Format Cells with Ctrl + 1 and apply a border from the Border tab.", hint: "Select the header with keyboard selection first. In the dialog, use Alt keys or Tab to reach the Border tab.", solution: "The header becomes bold with borders without touching the mouse — demonstrating that full formatting is reachable from the keyboard." },
      { task: "Practice the sheet-flip: move to the second sheet with Ctrl + Page Down, enter something there, and flip back with Ctrl + Page Up.", hint: "The Practice workbook has Exercise and Solution sheets to hop between.", solution: "Sheet switching takes two keystrokes in each direction — far faster than aiming at the small tabs with the mouse." }
    ],
    quiz: [
      { q: "Which shortcut inserts today's date as a real date value?", options: ["Ctrl + ;", "Ctrl + D", "Alt + D", "F5"], answer: 0, explanation: "Ctrl + ; stamps the current date into the active cell as a true date that sorts and calculates — faster and more reliable than typing." },
      { q: "Which pair of shortcuts selects a whole column of data from its first cell?", options: ["Ctrl + Shift + Down, then work across as needed", "Ctrl + C then Ctrl + V", "Alt + H then A", "F2 then Enter"], answer: 0, explanation: "Ctrl + Shift + Arrow extends the selection to the end of the data block — the fastest way to grab a full column." },
      { q: "What does Ctrl + Page Up / Page Down do?", options: ["Switches between worksheet tabs", "Scrolls the current sheet", "Undoes page setup changes", "Prints the workbook"], answer: 0, explanation: "Ctrl + Page Up/Down flips to the previous or next worksheet — the keyboard way to hop between sheets." },
      { q: "What is the recommended way to build lasting shortcut skill?", options: ["Adopt about three at a time and use them in real work until automatic", "Memorise a full table in one evening", "Only use shortcuts on Fridays", "Learn them after mastering the mouse fully"], answer: 0, explanation: "Shortcuts are motor skills. Small groups adopted during real work become permanent; mass memorisation does not." },
      { q: "Which shortcut opens the Format Cells dialog?", options: ["Ctrl + F1", "Alt + 1", "Ctrl + 1", "F1"], answer: 2, explanation: "Ctrl + 1 opens Format Cells — the fastest route to number formats, borders, alignment and more." }
    ],
    summary: [
      "Around 25 well-chosen shortcuts save 30-60 minutes a day and keep your eyes on the data.",
      "Navigation core: Ctrl + Arrow to dive, Ctrl + Home to return, Ctrl + Page Up/Down between sheets.",
      "Selection core: Ctrl + Shift + Arrow grabs a full data range instantly.",
      "Ctrl + ; and Ctrl + Shift + ; stamp true date and time values in one press.",
      "Adopt three shortcuts at a time in real work — motor memory beats memorisation."
    ],
    downloads: [],
    keywords: ["excel keyboard shortcuts", "ctrl shortcuts list", "excel navigation keys", "ctrl semicolon date", "selection shortcuts", "excel speed office work", "alt h ribbon", "f2 edit", "essential excel shortcuts"]
  },
  {
    id: "l1-mini-project",
    title: "Mini Project: Build a Weekly Budget",
    difficulty: "Beginner",
    duration: 20,
    version: "",
    objectives: [
      "Set up a clean weekly income and expense budget from scratch",
      "Format the sheet professionally with currency and borders",
      "Use SUM for totals and subtraction for the balance",
      "Apply an absolute reference to a savings-rate cell"
    ],
    intro: "Time to put every Level 1 skill into one real piece of work. You will build a complete weekly budget for a small office: data entry, currency formatting, SUM totals, a balance calculation, and an absolute-referenced savings rate.\n\nNothing here is new — that is the point. A project stitches separate skills into one confident workflow. Work in the formula-basics workbook (Exercise sheet) or in your own blank workbook, and take your time. Aim for a sheet you would be happy to hand to a manager.",
    sections: [
      {
        heading: "The plan: what you are building",
        blocks: [
          { type: "p", text: "Sunrise Traders wants a simple weekly money picture: what came in, what went out, what remains, and how much of the remainder should be moved to savings (the office follows a fixed savings rule of 20%). The finished sheet looks like this:" },
          { type: "grid", title: "The finished weekly budget (the look you are aiming for)", rows: [
            ["", "A", "B"],
            ["1", "Sunrise Traders - Weekly Budget (Week 46)", ""],
            ["2", "Savings rate", "20%"],
            ["3", "", ""],
            ["4", "Income", ""],
            ["5", "Cash sales", "18400"],
            ["6", "Card payments received", "9250"],
            ["7", "Total income", "=SUM(B5:B6)"],
            ["8", "", ""],
            ["9", "Expenses", ""],
            ["10", "Notebooks restock", "3600"],
            ["11", "Printer paper", "1200"],
            ["12", "Electricity", "2400"],
            ["13", "Courier", "450"],
            ["14", "Total expenses", "=SUM(B10:B13)"],
            ["15", "", ""],
            ["16", "Balance (income - expenses)", "=B7-B14"],
            ["17", "To savings (balance * rate)", "=B16*$B$2"]
          ] },
          { type: "p", text: "Every element maps to a lesson: data entry (rows of items), number formatting (currency), `SUM` (totals), subtraction (balance), and an absolute reference (the savings rate in `B2`). Build it section by section, checking against this picture as you go." }
        ]
      },
      {
        heading: "Step 1: set up the workbook and enter the data",
        blocks: [
          { type: "steps", items: [
            "Open the formula-basics workbook and go to the **Exercise** sheet (or create a blank workbook with [[Ctrl]] + [[N]]).",
            "In `A1`, type the title: Sunrise Traders - Weekly Budget (Week 46).",
            "In `A2`, type Savings rate; in `B2`, type **20%** (typing the % sign stores 0.20 — exactly what you want).",
            "In `A4`, type Income. Below it, enter the two income rows: `A5` Cash sales with `B5` = 18400; `A6` Card payments received with `B6` = 9250.",
            "In `A9`, type Expenses. Then rows 10-13: Notebooks restock 3600; Printer paper 1200; Electricity 2400; Courier 450.",
            "Leave `A7`, `A14`, `A16`, `A17` empty for now — formulas go there in Step 3."
          ] },
          { type: "p", text: "Data-entry reminders from earlier lessons: press [[Enter]] to move down a column; check that all amounts sit **right-aligned** (pure numbers); if any entry sits left, something text-like slipped in — fix it with F2 rather than retyping the cell. Notice the deliberate blank rows 3, 8 and 15: they visually separate the three blocks, a simple and very effective layout habit." },
          { type: "warning", text: "Type amounts as pure numbers — 18400, not 18,400 or Rs 18400. Formatting comes next; typed symbols would turn values into text and break every later formula." }
        ]
      },
      {
        heading: "Step 2: format the sheet",
        blocks: [
          { type: "steps", items: [
            "Title: select `A1`, make it **bold**, size 12-14.",
            "Section headers: bold `A4` (Income) and `A9` (Expenses). Give both a light fill colour from the paint bucket dropdown.",
            "Amounts: select `B5:B6` and `B10:B13`, press [[Ctrl]] + [[1]], choose **Accounting** with 0 decimals (or 2 if you prefer paise visible). The rupee symbols align at the left edge.",
            "Select `B2` and format it as **Percent** — it already holds 0.20, so it will display 20%.",
            "Borders: select the whole used block (`A1:B17`), Home > Borders dropdown > **All Borders**."
          ] },
          { type: "p", text: "Also auto-fit: double-click the border between the A and B column headers so the long labels like \"Card payments received\" display fully. If a cell ever shows ##### , widen the column — never retype the number." },
          { type: "tip", text: "Resist over-decorating. Bold headers, one light fill, Accounting numbers, borders — stop there. Professional sheets are quiet; the numbers are the stars." }
        ]
      },
      {
        heading: "Step 3: the formulas — SUM, balance and the savings calculation",
        blocks: [
          { type: "steps", items: [
            "In `A7`, type Total income (bold). In `B7`, type **=SUM(B5:B6)** and press Enter.",
            "In `A14`, type Total expenses (bold). In `B14`, type **=SUM(B10:B13)**.",
            "In `A16`, type Balance (income - expenses). In `B16`, type **=B7-B14** and press Enter.",
            "In `A17`, type To savings (balance * rate). In `B17`, type **=B16*$B$2** — and here is the key moment: after clicking `B2`, press **F4** once to lock it as `$B$2`."
          ] },
          { type: "p", text: "Walk through what each formula does. `=SUM(B5:B6)` adds the income rows: 18400 + 9250 = **27650**. `=SUM(B10:B13)` adds the four expenses: 3600 + 1200 + 2400 + 450 = **7650**. The balance `=B7-B14` gives 27650 - 7650 = **20000** — this is a PEMDAS-free formula, a single subtraction, so no parentheses are needed. Finally `=B16*$B$2` multiplies the balance by the savings rate: 20000 × 0.20 = **4000**." },
          { type: "p", text: "Why `$B$2` and not `B2`? Today one formula sits in `B17`, but budgets grow: next week you may add rows below, or copy the savings line into several weekly blocks on one sheet. A relative `B2` would drift to `B3`, `B4`... and multiply by empty cells. The locked `$B$2` is immune to all of that. This is exactly the decision guide from the references lesson: **one fixed cell feeding many formulas → lock it.**" },
          { type: "grid", title: "Formula check — what each cell should hold", rows: [
            ["Cell", "Formula", "Shows"],
            ["B7", "=SUM(B5:B6)", "27,650"],
            ["B14", "=SUM(B10:B13)", "7,650"],
            ["B16", "=B7-B14", "20,000"],
            ["B17", "=B16*$B$2", "4,000"]
          ] },
          { type: "tip", text: "The click-to-reference habit pays off here: build each formula by clicking the cells instead of typing addresses. Press F4 the instant you reference B2 — one press, fully locked, done." }
        ]
      },
      {
        heading: "Step 4: final checks and polish",
        blocks: [
          { type: "p", text: "Before calling it done, run the checks a careful colleague would run:" },
          { type: "ul", items: [
            "Press **[[Ctrl]] + [[`]]** to view all formulas. Confirm `B7` and `B14` are SUMs over the right ranges, `B16` subtracts the two totals, and `B17` locks `$B$2`. Press the shortcut again to return.",
            "Test the live wiring: change Cash sales from 18400 to 20000. The balance should jump to 21600 and savings to 4320 — with **no formula edits**. Change it back.",
            "Test the rate: change `B2` to 25% (that is, type 25% or 0.25). Savings should become 5000. Change it back to 20%.",
            "Check alignment: every amount right-aligned, labels left-aligned, no ##### anywhere."
          ] },
          { type: "p", text: "If any check fails, use [[Ctrl]] + [[Z]] to step back and fix the cause — a text-typed number (left-aligned), a missing F4 lock, or a SUM range missing a row. Fixing is part of the job; even experienced users test and adjust." },
          { type: "p", text: "Finish by pressing [[Ctrl]] + [[S]]. If you built this in your own blank workbook, name it \"Sunrise Traders - Weekly Budget.xlsx\". You now hold a reusable template: next week, overwrite the income and expense values and the whole budget recalculates itself — totals, balance and savings — with zero formula work." }
        ]
      },
      {
        heading: "Extend it: exercises on the same sheet",
        blocks: [
          { type: "p", text: "The four exercises below all build on the budget you just finished. They replay Level 1 skills in new combinations — that repetition is what turns knowledge into skill." },
          { type: "p", text: "When you want to compare your work, the formula-basics workbook also contains a **Solution** sheet showing a completed version of the budget. Compare honestly: check formulas with [[Ctrl]] + [[`]], not just the displayed numbers." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Typing amounts as \"Rs 18,400\" while entering the data.", right: "Enter pure numbers (18400) and apply Accounting format in Step 2.", why: "Typed symbols and commas make text; every SUM then silently skips the cell and the whole budget undercounts." },
      { wrong: "Writing =B16*B2 for the savings line, planning to \"fix it later if needed\".", right: "Lock immediately: =B16*$B$2 with F4.", why: "Unlocked rate references drift the moment the formula is copied or rows are inserted — the classic budget-breaking error." },
      { wrong: "Building the grand total as =B5+B6+B10+B11+B12+B13 in one cell.", right: "Two separate SUMs (income and expenses), then one subtraction for the balance.", why: "Separate totals are readable, auditable and reusable — a single mixed chain hides its logic and breaks when any block changes." },
      { wrong: "Testing nothing after building, and shipping the first result.", right: "Change an input value, watch the outputs move, then change it back.", why: "A quick wiring test proves the formulas reference the right cells. Shipping an unverified budget is how wrong numbers reach managers." }
    ],
    shortcuts: [
      { keys: "Ctrl + S", action: "Save the finished budget" },
      { keys: "Ctrl + 1", action: "Format the amounts as Accounting" },
      { keys: "F4", action: "Lock the savings-rate reference while building the formula" },
      { keys: "Ctrl + `", action: "Audit all formulas before finishing" },
      { keys: "Ctrl + Z", action: "Undo any wrong step during the build" }
    ],
    exercises: [
      { task: "Add a second week to the same sheet: rows for Week 47 with new income and expense values, repeating the block structure (income rows, SUM total, expense rows, SUM total). Give Week 47's balance its own =total-total formula.", hint: "Copy the Week 46 block with Ctrl + C / Ctrl + V below, then overwrite the values — the SUM ranges adjust to the new rows automatically.", solution: "Copying the block brings formulas and formatting along; relative SUM ranges re-point to the pasted rows. Only the values need retyping." },
      { task: "Add a \"Week 46 + Week 47 combined\" line: a grand income (SUM of both weeks' income totals), a grand expense total, and a combined balance. Verify the numbers by hand.", hint: "SUM can add non-adjacent totals: =SUM(B7,B21) style — or add the two total cells directly.", solution: "Grand income = the two income totals added; same for expenses; balance = grand income minus grand expenses. Hand-checking confirms the wiring." },
      { task: "Raise the savings rate to 30% and confirm every savings figure updates. Then think: which single cell did you edit, and how many formulas felt it?", hint: "Only B2 (or your Week 47 equivalent rate cell) changes.", solution: "One edit to the locked rate cell recalculated every dependent savings formula — the exact payoff of absolute references." },
      { task: "Compare your Exercise sheet against the workbook's Solution sheet using Ctrl + ` on both. List any formula that differs and decide which version is better and why.", hint: "Toggle formula view with Ctrl + ` on each sheet in turn. Differences in ranges or locks matter more than cosmetic differences.", solution: "A careful comparison typically surfaces range or lock differences. Any version where the totals, balance and locked rate all verify by hand is correct — understanding why beats matching exactly." }
    ],
    quiz: [
      { q: "Which formula correctly computes the balance if B7 holds total income and B14 holds total expenses?", options: ["=B7-B14", "=SUM(B7:B14)", "=B14-B7", "=B7+B14"], answer: 0, explanation: "The balance is income minus expenses: =B7-B14. A single subtraction needs no parentheses and no SUM." },
      { q: "Why must the savings formula reference $B$2 instead of B2?", options: ["B2 contains text that breaks formulas", "The $ makes the result positive", "The rate cell must stay fixed if the formula is copied or rows are added", "Excel requires $ in all formulas"], answer: 2, explanation: "A relative B2 would drift to other cells on copy or row insertion. $B$2 anchors every savings formula to the one rate cell." },
      { q: "You typed 20% into B2 and formatted B17 as =B16*$B$2. What does Excel actually multiply by?", options: ["0.20 — percent entries store the fraction", "The text \"20%\"", "20", "It depends on the formatting"], answer: 0, explanation: "Typing 20% stores 0.20. Formatting displays it as 20%, but the stored value used by formulas is the fraction." },
      { q: "After finishing the budget, which check best proves the formulas are wired correctly?", options: ["Change an input value and confirm the totals, balance and savings update, then change it back", "Look at the numbers and see if they seem reasonable", "Print the sheet and read it", "Save the file twice"], answer: 0, explanation: "A live wiring test — edit an input, watch every dependent output move — proves the references. Reasonableness of numbers proves nothing about references." },
      { q: "What makes this sheet a reusable template for next week?", options: ["Formulas reference cells, so new values recalculate everything automatically", "It is saved as xlsx", "The title mentions Week 46", "It has borders"], answer: 0, explanation: "Because totals, balance and savings are formulas over cell ranges, replacing the input values recalculates the entire budget — no formula rework needed." }
    ],
    summary: [
      "The budget stitches Level 1 together: data entry, formatting, SUM, subtraction and $-locked references.",
      "Structure first: title, one rate cell, separated Income and Expenses blocks with their own SUM totals.",
      "Balance = income total - expense total; savings = balance * locked rate cell.",
      "One rate edit recalculates every savings figure — the payoff of absolute references.",
      "Always finish with checks: Ctrl + ` formula view, a live wiring test, and Ctrl + S."
    ],
    downloads: [
      { file: "assets/files/level1/formula-basics.xlsx", label: "Formula basics workbook", note: "Exercise and Solution sheets with basic calculations, references and the weekly budget project — build your budget on the Exercise sheet and compare with the Solution sheet" }
    ],
    keywords: ["weekly budget excel", "excel mini project", "sum totals budget", "balance formula", "savings rate absolute reference", "beginner excel project", "office budget template", "formula basics practice"]
  }
);
