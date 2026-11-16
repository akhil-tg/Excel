// level4-1.js — Level 4: Data Analysis and Reporting (metadata + lessons 1-4)
window.LEVEL_META = window.LEVEL_META || {};
window.LEVEL_LESSONS = window.LEVEL_LESSONS || {};

window.LEVEL_META[4] = {
  id: 4,
  title: "Data Analysis and Reporting",
  subtitle: "Turn raw data into charts, PivotTables and professional dashboards",
  description: "Analyze data and communicate it: charts, PivotTables, slicers, KPI reporting, monthly reports and interactive dashboards.",
  difficulty: "Intermediate",
  icon: "chart"
};

(window.LEVEL_LESSONS[4] = window.LEVEL_LESSONS[4] || []).push(
  {
    id: "l4-charts",
    title: "Charts and Graphs",
    difficulty: "Intermediate",
    duration: 18,
    version: "",
    objectives: [
      "Create a chart from your data and move or resize it",
      "Choose the right chart type for each business question",
      "Edit chart titles, axis labels, legends and data labels",
      "Format charts quickly with styles and colors",
      "Add sparklines to show trends inside a single cell"
    ],
    intro: "A table of numbers hides the story. A chart shows the story in seconds. Your manager sees at a glance which region is winning and whether sales are growing. That is why every review meeting starts with charts.\n\nIn this lesson you will use the sales data of Sunrise Traders, an office supplies company. The workbook `sales-data.xlsx` has 250 rows on the Sales Data sheet. The columns are Date, Region, Salesperson, Product, Units and Revenue. The data is a Table named `SalesTable`. You already know `SUMIFS` from Level 3. So we will summarise first, then chart the summary.",
    keyIdea: "Charts turn 250 rows of data into one picture your manager can read in seconds.",
    sections: [
      {
        heading: "Why a picture beats a table",
        blocks: [
          { type: "p", text: "Excel charts answer one question at a glance. Before you insert a chart, ask: what question am I answering? The question decides the chart type, not the other way round." },
          { type: "ul", items: [
            "**Which region sold more?** You are comparing categories, so a column chart is right.",
            "**Is revenue growing month by month?** That is a trend over time, so a line chart is right.",
            "**How much of the total comes from each product?** That is a share of a whole, so a pie chart is right.",
            "**Which salesperson hit target?** You are ranking a list, so a sorted column chart or a bar chart is right."
          ] },
          { type: "p", text: "If you start from the question, you will rarely pick the wrong chart. If you start from whatever chart looks pretty, your report will confuse people even when the data is correct." },
          { type: "p", text: "Charts appear everywhere in office life. The weekly team meeting opens with a trend line. The monthly review deck leads with a revenue-by-region column chart. A dashboard keeps small charts under the KPI numbers." },
          { type: "p", text: "Charts are not decoration. They are the fastest way for a busy reader to judge the business. A reader decides in about three seconds. Your job is to make those three seconds count." }
        ]
      },
      {
        heading: "Creating your first chart",
        blocks: [
          { type: "p", text: "Excel cannot make sense of 250 raw rows on its own. First build a small summary table, then chart it. On the Sales Data sheet, type `Region` in `H1` and `Revenue` in `I1`. Type the four region names in `H2:H5` and this formula in `I2`:" },
          { type: "formula", name: "SUMIFS summary for the chart",
            syntax: "=SUMIFS(sum_range, criteria_range, criteria)",
            example: "=SUMIFS(SalesTable[Revenue], SalesTable[Region], H2)",
            explanation: "Adds the revenue of every row whose Region matches the name in `H2`. Fill it down to `I5` to get one total per region." },
          { type: "grid", title: "The summary table that feeds the chart",
            rows: [
              ["", "Region", "Revenue"],
              ["", "North", "₹6,45,000"],
              ["", "South", "₹5,90,000"],
              ["", "East", "₹6,05,000"],
              ["", "West", "₹6,40,000"]
            ] },
          { type: "steps", items: [
            "Select `H1:I5`, including the headers.",
            "On the Insert tab, in the Charts group, click the Column chart button and choose Clustered Column.",
            "Excel places the chart on the same sheet, next to your data.",
            "Click the chart once and drag it where you want it, for example below the summary table.",
            "Drag the corner handle to resize it. Drag the border to move it.",
            "Click outside the chart to deselect it."
          ] },
          { type: "p", text: "When you select a chart, two new tabs appear: Chart Design and Format. Use Chart Design to change the chart type or pick a style. Use Format only for colors and shapes." },
          { type: "p", text: "Here is an important fact: the chart is linked to the cells, not a copy of them. Change a figure in the summary table and the chart redraws itself." },
          { type: "p", text: "This is why the chart must point at a small summary range like `H1:I5`, not at 250 transaction rows. Small summary, clean chart, always current." },
          { type: "tip", text: "Select the summary range and press [[Alt]] + [[F1]] to drop a chart straight onto the sheet. Press [[F11]] instead to place the chart on its own full-page chart sheet. That is handy when the chart is the main output of a report." },
          { type: "p", text: "Handle charts with care. Click the chart border to move the whole object. Click inside on a column or line to select that series for formatting. Resize only from the corner handles; they keep the proportions." },
          { type: "p", text: "Side handles stretch the chart and distort the text inside. When a report holds several charts, line them up with the cell grid. Hold [[Alt]] while dragging to snap to the grid." }
        ]
      },
      {
        heading: "Choosing the right chart type",
        blocks: [
          { type: "table", headers: ["Chart type", "Use it for", "Sunrise Traders example"], rows: [
            ["Column", "Comparing a few categories", "Revenue by region"],
            ["Line", "Trends over time", "Revenue by month, January to December"],
            ["Pie", "Share of a total", "Revenue share by product"],
            ["Bar", "Categories with long names", "Revenue by salesperson full name"],
            ["Combo", "Two measures with different scales", "Revenue as columns, units sold as a line"]
          ] },
          { type: "p", text: "**Column charts** are the safe default for comparisons. One column per region, tallest column = best region. Readers understand them without explanation, which matters in a meeting." },
          { type: "p", text: "Two habits make column charts better at once. First, sort the categories by value before charting. The columns then fall from tallest to shortest, and the ranking is obvious without reading labels." },
          { type: "p", text: "Second, keep the value axis starting at zero. A cut-off axis makes a 5% difference look like 50%. That reads as exaggeration, not insight." },
          { type: "p", text: "**Line charts** are for time. Put dates or month names on the horizontal axis and the measure on the vertical axis. The line tells the story: rising is growth, falling is trouble." },
          { type: "p", text: "Never use a line chart for unrelated categories such as products. The connecting lines imply an order that does not exist." },
          { type: "p", text: "**Pie charts** show how a total is divided. One pie, one measure, one period. Keep it to 5 or 6 slices at most. With 12 products, the slices become unreadable slivers." },
          { type: "p", text: "With more categories, use a column chart instead. A pie of Sunrise Traders revenue by product works well, because there are only five products." },
          { type: "p", text: "**Bar charts** are column charts turned sideways. Use them when category names are long, such as `Priya Nair` or `Rahul Verma`, because horizontal bars leave room for the full label." },
          { type: "p", text: "**Combo charts** show two different measures together. A classic Sunrise Traders view: revenue as columns and units sold as a line, because rupees and units are on completely different scales." },
          { type: "steps", items: [
            "Build a region summary with Revenue in one column and Units in the next, using `SUMIFS` twice.",
            "Select the summary and insert a Clustered Column chart.",
            "With the chart selected, on the Chart Design tab click Change Chart Type and choose Combo.",
            "Set Revenue to Clustered Column and Units to Line.",
            "Tick the Secondary Axis box next to Units so the line gets its own scale on the right.",
            "Click OK and check that both series are readable."
          ] },
          { type: "warning", text: "Avoid 3-D chart types. The tilted perspective distorts the columns, makes values harder to compare, and looks dated in a professional report. Flat 2-D charts are the standard in business reporting." }
        ]
      },
      {
        heading: "Chart elements and formatting",
        blocks: [
          { type: "table", headers: ["Element", "What it does", "When to keep it"], rows: [
            ["Chart title", "States the question the chart answers", "Always, and make it specific: \"Revenue by region, 2025\""],
            ["Axis titles", "Explain what each axis measures", "When units are not obvious, for example ₹ or units"],
            ["Legend", "Identifies the series by color", "Only when there are two or more series"],
            ["Data labels", "Show the exact value on each point", "When exact numbers matter, for example KPI charts"],
            ["Gridlines", "Help the eye read values", "Light horizontal gridlines only; delete vertical ones"]
          ] },
          { type: "steps", items: [
            "Click the chart to select it.",
            "Click the plus button (+) that appears at the top-right corner of the chart.",
            "Tick Chart Title and type a clear title, such as Revenue by region, 2025.",
            "Tick Data Labels if managers need exact figures.",
            "Untick Legend when there is only one series; a lone legend is clutter."
          ] },
          { type: "p", text: "For colors and styles, use the Chart Styles gallery on the Chart Design tab. It is faster than formatting each part by hand. Need finer control? Select an element and press [[Ctrl]] + [[1]] to open the Format pane. There you can set fill, border and number formats." },
          { type: "tip", text: "You can highlight a single column. Click the series once to select all columns. Click again on one column to select only it. Then use Format > Shape Fill. Highlighting the weakest region is a simple way to focus a discussion." },
          { type: "p", text: "Two small habits make charts look professional. Keep one color per series across all charts in a report. Delete anything that does not help the reader, such as heavy borders and extra legends." },
          { type: "p", text: "Watch the axis scale when you compare charts. Two revenue charts can look completely different even with the same data. One may run from 0 to 10 lakh, the other from 4 to 6 lakh." },
          { type: "p", text: "Excel usually picks the scale automatically. So glance at the axis numbers before you present. A fair comparison needs comparable axes." }
        ]
      },
      {
        heading: "Sparklines: tiny charts in a cell",
        blocks: [
          { type: "p", text: "A sparkline is a tiny chart inside a single cell. Use it to show each item's trend in a row-per-item summary, without inserting twelve full charts." },
          { type: "steps", items: [
            "Put the twelve monthly revenue figures in a row or column, for example `C2:N2` for one region.",
            "Click the cell where you want the sparkline, for example `P2`.",
            "On the Insert tab, in the Sparklines group, click Line.",
            "In the Data Range box type `C2:N2` and click OK.",
            "Fill the sparkline down for the other regions."
          ] },
          { type: "p", text: "Sparklines update when the numbers change. They print with the sheet. Select one and the Sparkline tab lets you mark high and low points. Use them in KPI tables and month summaries. You will meet them again in the KPI reporting lesson." },
          { type: "p", text: "Excel offers three sparkline types from the same Insert menu. Line is the default and best for trends. Column shows each period as a tiny bar. It is good for spotting one unusually strong month." },
          { type: "p", text: "Win/Loss shows only up or down markers, for sequences of wins and losses. For revenue and sales reporting, line sparklines cover almost every need." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Pie chart with twelve slices, one per product.", right: "Column chart of revenue by product, sorted from largest to smallest.", why: "Beyond five or six slices, a pie becomes unreadable and the small slices carry no message. Columns handle any number of categories." },
      { wrong: "Charting the 250 raw rows of SalesTable directly.", right: "Summarise first with `SUMIFS` or a PivotTable, then chart the summary.", why: "A chart of raw transactions shows noise, not insight. Charts answer questions, and the summary table is where the question gets answered." },
      { wrong: "Line chart with one line per product across random categories.", right: "Line chart only when the axis is time, column chart for unrelated categories.", why: "Lines imply a continuous trend. Connecting Pens to Staplers to Printers suggests a progression that does not exist." }
    ],
    shortcuts: [
      { keys: "Alt + F1", action: "Create a chart from the selected data on the current sheet" },
      { keys: "F11", action: "Create a chart on its own chart sheet" },
      { keys: "Ctrl + 1", action: "Open the Format pane for the selected chart element" }
    ],
    exercises: [
      { task: "In sales-data.xlsx, build a region summary in `H1:I5` with `SUMIFS` and create a clustered column chart of it. Title the chart \"Revenue by region, 2025\".",
        hint: "Use `=SUMIFS(SalesTable[Revenue], SalesTable[Region], H2)` in `I2` and fill down to `I5`.",
        solution: "Select `H1:I5`, Insert > Charts > Clustered Column. Click the title and type Revenue by region, 2025. Your columns should match the Solution idea: North ₹6,45,000, South ₹5,90,000, East ₹6,05,000, West ₹6,40,000." },
      { task: "Add a Month column to SalesTable with `=MONTH([@Date])`, build a 12-row month summary with `SUMIFS`, and create a line chart of revenue by month.",
        hint: "Type the header Month in `G1` next to the table; Excel extends the table. Then use `=SUMIFS(SalesTable[Revenue], SalesTable[Month], 1)` for month 1 and change the number per row.",
        solution: "Select the month summary and insert a Line chart. The line should rise and fall across the year. Compare your shape with the Solution sheet pattern." },
      { task: "Create a combo chart: revenue by region as columns and units sold by region as a line on the secondary axis.",
        hint: "Build a second summary column with `=SUMIFS(SalesTable[Units], SalesTable[Region], H2)`, then use Chart Design > Change Chart Type > Combo.",
        solution: "Revenue stays on the primary axis and Units moves to the secondary axis on the right. Both series should be clearly readable; if the line hugs the floor, the secondary axis is not ticked." }
    ],
    quiz: [
      { q: "You want to show whether revenue is rising or falling from January to December. Which chart type fits best?",
        options: ["Pie chart", "Line chart", "Bar chart", "Doughnut chart"],
        answer: 1,
        explanation: "A trend over time calls for a line chart. Pie and doughnut charts show shares of a total, and a bar chart compares categories without emphasising time." },
      { q: "How many slices should a pie chart have at most?",
        options: ["2 or 3", "5 or 6", "10 to 12", "As many as there are categories"],
        answer: 1,
        explanation: "Past five or six slices the small wedges become unreadable. With more categories, switch to a column or bar chart sorted from largest to smallest." },
      { q: "In a combo chart of revenue (₹) and units sold, why put units on the secondary axis?",
        options: ["Because secondary axes are always on the right side", "Because ₹ and units have very different scales", "Because Excel does not allow two series on one axis", "To make the chart taller"],
        answer: 1,
        explanation: "Revenue is in lakhs of rupees while units are in hundreds, so one shared axis would flatten one of the series. The secondary axis gives each measure its own scale." },
      { q: "What does pressing F11 do when data is selected?",
        options: ["Creates a chart on its own chart sheet", "Opens Format Cells", "Inserts a PivotTable", "Saves the workbook"],
        answer: 0,
        explanation: "F11 builds a chart from the selection and places it on a dedicated full-page chart sheet. Alt + F1 embeds the chart on the current sheet instead." }
    ],
    summary: [
      "Start from the question, then pick the chart type: column to compare, line for time, pie for share, bar for long labels.",
      "Summarise first with formulas such as `SUMIFS`, then chart the small summary table, never the 250 raw rows.",
      "Set a specific chart title, keep only useful elements, and prefer flat 2-D charts over 3-D.",
      "Use secondary axes in combo charts when two measures have different scales.",
      "Sparklines give a per-row trend in a single cell and update with the data."
    ],
    downloads: [
      { file: "assets/files/level4/sales-data.xlsx",
        label: "Sales dataset for charts and pivots",
        note: "250 rows of realistic sales data (date, region, salesperson, product, units, revenue) — Exercise sheet; pivot/chart builds are described in the lessons" }
    ],
    keywords: ["charts", "column chart", "line chart", "pie chart", "bar chart", "combo chart", "secondary axis", "sparklines", "data labels", "chart formatting"]
  },
  {
    id: "l4-pivot-tables",
    title: "PivotTables",
    difficulty: "Intermediate",
    duration: 20,
    version: "",
    objectives: [
      "Explain what a PivotTable is and why analysts use it",
      "Create a PivotTable from an Excel Table",
      "Place fields in the Rows, Columns, Values and Filters areas",
      "Switch a value field between Sum, Count and Average",
      "Refresh a PivotTable after the source data changes"
    ],
    intro: "A PivotTable answers many questions without a single formula. Drag Region into rows and you get revenue by region. Drag Product there instead and you get revenue by product. Add months and you get both together. Analysts love them because a summary of twenty `SUMIFS` formulas takes twenty seconds of dragging.\n\nThere is a second reason: safety. A PivotTable never touches your raw data. It reads the source and writes its own summary area. Sorting, filtering and rearranging the pivot leave the 250 rows of Sunrise Traders sales data exactly as they were. In this lesson you will build your first pivots from the `SalesTable` in `sales-data.xlsx`.",
    keyIdea: "A PivotTable is a summary table you build by dragging field names — Excel does all the counting and adding.",
    sections: [
      {
        heading: "What a PivotTable is and why analysts use it",
        blocks: [
          { type: "p", text: "A PivotTable is an automatic summary of a dataset. You choose which fields to group by and which numbers to add up. Excel builds the whole table for you." },
          { type: "ul", items: [
            "**Speed**: one revenue-by-region-and-product summary would take many `SUMIFS` formulas; a pivot takes seconds.",
            "**Flexibility**: change your mind about the breakdown and just drag a different field in.",
            "**Safety**: the pivot is a separate summary; your raw rows are never modified.",
            "**Freshness**: one Refresh command pulls in any new or changed source rows."
          ] },
          { type: "p", text: "The name comes from pivoting, or rotating, the layout. Move a field from rows to columns and the summary reorganises instantly." },
          { type: "p", text: "Here is a typical office use. Every month-end you get the transactions export. The review meeting needs revenue by region, by product and by salesperson. With a pivot, all three views live on one sheet." },
          { type: "p", text: "Compare that with the formula approach from Level 3. A `SUMIFS` summary is excellent when the layout is fixed and final. But every new breakdown is a new block of formulas, and every change of mind is an hour of editing." },
          { type: "p", text: "A pivot answers the \"what if we looked at it by...?\" questions in real time. That is exactly how analysis happens in a meeting. Many analysts use both: pivots to explore, `SUMIFS` for the fixed report layout published every month." }
        ]
      },
      {
        heading: "Creating a PivotTable from a Table",
        blocks: [
          { type: "p", text: "Your data should be an Excel Table first, as `SalesTable` already is. Tables grow automatically when you add rows. So the pivot picks up new data on refresh, without any re-selection." },
          { type: "steps", items: [
            "Open sales-data.xlsx and click any cell inside the Sales Data table.",
            "On the Insert tab, click PivotTable.",
            "Excel pre-fills Table/Range with `SalesTable`. Keep it.",
            "Choose New Worksheet so the pivot gets its own clean sheet.",
            "Click OK. Excel creates a blank pivot placeholder on the left and the PivotTable Fields pane on the right."
          ] },
          { type: "p", text: "The Fields pane lists every column of the table: Date, Region, Salesperson, Product, Units and Revenue. Below it are four drop areas: Filters, Columns, Rows and Values. You build everything by dragging fields between them." },
          { type: "tip", text: "If the Fields pane disappears, click anywhere inside the pivot and it comes back. Renaming the new sheet to something like \"PT Revenue\" right away keeps big workbooks easy to navigate." },
          { type: "p", text: "A new pivot reads \"Row Labels\" and \"Sum of Revenue\". That is Excel's wording, not yours. Click the cell that says Row Labels and type \"Region\"." },
          { type: "p", text: "Rename value headers the same way through Value Field Settings. Use names a manager would write, such as \"Revenue (₹)\". Renamed headers survive refreshes and make the pivot readable to anyone." }
        ]
      },
      {
        heading: "The four areas: Rows, Columns, Values, Filters",
        blocks: [
          { type: "table", headers: ["Area", "What goes there", "Example with Sunrise Traders"], rows: [
            ["Rows", "Fields you want listed downwards", "Region, or Region then Product"],
            ["Columns", "Fields you want spread sideways", "Product, to see region by product"],
            ["Values", "Numeric fields to summarise", "Revenue (sum), Units (sum or average)"],
            ["Filters", "Fields applied to the whole pivot", "Salesperson, to view one person at a time"]
          ] },
          { type: "p", text: "Drag **Region** to Rows and **Revenue** to Values. The pivot instantly shows one row per region with the sum of revenue, plus a Grand Total row." },
          { type: "grid", title: "The pivot after Region to Rows and Revenue to Values",
            rows: [
              ["Region", "Sum of Revenue"],
              ["North", "₹6,45,000"],
              ["South", "₹5,90,000"],
              ["East", "₹6,05,000"],
              ["West", "₹6,40,000"],
              ["Grand Total", "₹24,80,000"]
            ] },
          { type: "p", text: "Now drag **Product** to Columns as well. You get a cross-tab: regions downwards, products sideways. Each cell answers \"how much revenue did the North make from Printers?\". Finally drag **Salesperson** to Filters. A filter dropdown appears above the pivot, so you can view one person's numbers." },
          { type: "p", text: "A useful habit: put the field with fewer items in Columns and the one with more items in Rows. Five products across the top stays readable. Twelve months across the top is often too wide." },
          { type: "p", text: "You can stack more than one field in an area. Drop Salesperson under Region in Rows. The pivot lists the four salespeople inside each region, collapsible with the minus buttons." },
          { type: "p", text: "Two fields in Values work the same way: Sum of Revenue next to Sum of Units in one pivot. Excel adds automatic subtotals for outer row fields. Start with one field per area while learning, then add a second only when the question needs it." }
        ]
      },
      {
        heading: "Sum, Count and Average: changing the calculation",
        blocks: [
          { type: "p", text: "By default Excel sums numeric fields in Values. But if the field contains text or blank cells, the pivot quietly switches to Count. Many beginners are surprised when the header reads \"Count of Revenue\" instead of \"Sum of Revenue\"." },
          { type: "p", text: "Count merely counts rows. It does not add anything." },
          { type: "steps", items: [
            "Right-click any number inside the Values area.",
            "Choose Value Field Settings.",
            "In the Summarize by list choose Sum, Count, Average, Max or Min.",
            "Type a clear Custom Name, for example \"Average Units\", so the pivot header means something.",
            "Click OK."
          ] },
          { type: "p", text: "When is each calculation right? Use **Sum** for money and units: total revenue, total units. Use **Count** when the question is \"how many transactions?\". Use **Average** for typical values, such as average order size per region." },
          { type: "p", text: "For example, put Units in Values twice. Set one copy to Sum of Units and the other to Average Units. Now you can see volume and typical order size side by side." },
          { type: "p", text: "A worked example makes the difference concrete. Suppose the North sold 3,800 units across 95 transactions. Sum of Units says 3,800. Count of Units says 95, the number of rows. Average of Units says 40, the typical order size." },
          { type: "p", text: "Three different numbers from the same column, each answering a different question. So choosing the calculation is a business decision, not an Excel setting. Ask what the reader needs to know before you right-click." },
          { type: "p", text: "Number formatting belongs in the same dialog. In Value Field Settings, click Number Format and choose Currency with the ₹ symbol and no decimals. This format lives inside the pivot field. So it survives refreshes, moves and layout changes." },
          { type: "warning", text: "Do not format pivot numbers by selecting the cells and applying a cell format. When you refresh or move the pivot, those hand-applied formats land on the wrong cells. Always use Value Field Settings > Number Format." }
        ]
      },
      {
        heading: "Refresh, and your data stays untouched",
        blocks: [
          { type: "p", text: "A PivotTable is a snapshot of the source at the moment it was built. If you add rows to `SalesTable` or edit revenue figures, the pivot does not change until you refresh it." },
          { type: "p", text: "Click inside the pivot and press PivotTable Analyze > Refresh. Or right-click the pivot and choose Refresh. The shortcut is [[Alt]] + [[F5]]." },
          { type: "p", text: "Because `SalesTable` is a Table, new rows are included automatically. The next refresh pulls them into the pivot. So the monthly update cycle is simple: paste the new month's rows, refresh the pivots, done." },
          { type: "p", text: "Sorting, filtering, collapsing groups, even deleting the pivot sheet: none of it affects the source rows. The pivot only reads. One thing you cannot do is type into the values area. Excel blocks it, because pivot cells are calculated, not stored." },
          { type: "tip", text: "If a pivot looks wrong after a data edit, refresh first. Nine out of ten \"my pivot is broken\" moments are just a stale snapshot waiting for a refresh." },
          { type: "p", text: "When a workbook holds several pivots, refresh them all at once with Data > Refresh All. This matters in monthly reporting. Five pivots fed by one table must show the same version of the truth." },
          { type: "p", text: "If one pivot is refreshed and another is not, the two views disagree. In a meeting, that is far more damaging than being a day late." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Selecting the pivot cells and applying currency formatting by hand.", right: "Set the number format inside Value Field Settings > Number Format.", why: "Hand-applied cell formats break and land on wrong cells when the pivot is refreshed or rearranged. Field-level formatting travels with the field." },
      { wrong: "Wondering why the pivot shows \"Count of Revenue\" instead of a sum.", right: "Open Value Field Settings and change Summarize by to Sum.", why: "Excel switches to Count automatically when the value column contains text or blanks. Counting rupees is not the same as adding them." },
      { wrong: "Keeping the source as a plain range instead of a Table, so new rows never reach the pivot.", right: "Format the source as a Table with [[Ctrl]] + [[T]] before creating the pivot.", why: "A pivot built on a fixed range keeps the old address forever. A Table grows, so Refresh always includes the new rows." },
      { wrong: "Typing a corrected number directly into a pivot value cell.", right: "Fix the number in the source data, then refresh the pivot.", why: "Pivot values are calculated by Excel and cannot be edited. The source data is the only place where numbers should ever change." }
    ],
    shortcuts: [
      { keys: "Alt + F5", action: "Refresh the selected PivotTable" },
      { keys: "Ctrl + T", action: "Format the selected range as a Table before pivoting" }
    ],
    exercises: [
      { task: "In sales-data.xlsx, create a PivotTable on a new sheet that shows Sum of Revenue by Region, formatted as ₹ with no decimals.",
        hint: "Insert > PivotTable from `SalesTable`. Drag Region to Rows, Revenue to Values, then use Value Field Settings > Number Format.",
        solution: "You should see North ₹6,45,000, South ₹5,90,000, East ₹6,05,000, West ₹6,40,000 and a Grand Total of ₹24,80,000." },
      { task: "Add Product to Columns so you get revenue by region and product. Which product earns the most in the North?",
        hint: "Keep Region in Rows and drag Product into the Columns area.",
        solution: "Laptops lead the North, as they do in every region: the laptops column holds the largest values across the board." },
      { task: "Put Salesperson in Filters and view only Arjun Mehta. Then change Revenue to Count to see how many transactions he handled.",
        hint: "Drag Salesperson to the Filters area, use its dropdown, and right-click a value cell > Value Field Settings to switch Sum to Count.",
        solution: "With Arjun Mehta selected, Count of Revenue equals his number of sales rows in the dataset. Switch back to Sum when finished." },
      { task: "Show Sum of Units and Average Units side by side by region.",
        hint: "Drag Units into Values twice. Change the second copy to Average via Value Field Settings and rename it \"Average Units\".",
        solution: "Sum of Units shows total volume per region; Average Units shows the typical order size. Format the average with one decimal place so it reads cleanly." }
    ],
    quiz: [
      { q: "You want products spread across the top of the pivot and regions down the side. Where do you put each field?",
        options: ["Region in Columns, Product in Rows", "Region in Rows, Product in Columns", "Both in Rows", "Both in Values"],
        answer: 1,
        explanation: "Fields in the Rows area list downwards, fields in the Columns area spread sideways. Region goes to Rows, Product goes to Columns." },
      { q: "Your pivot shows \"Count of Revenue\" although you expected a sum. What is the most likely reason?",
        options: ["The Revenue column contains text or blank cells", "The pivot is not refreshed", "Count is the Excel default for all fields", "You used the wrong area"],
        answer: 0,
        explanation: "Excel automatically uses Count when a value field contains text or blanks, because text cannot be summed. Clean the source or change Summarize by to Sum in Value Field Settings." },
      { q: "You just added 20 new rows to the source table. What must you do for the PivotTable to include them?",
        options: ["Nothing, pivots update instantly", "Recreate the pivot from scratch", "Refresh the pivot", "Save and reopen the file"],
        answer: 2,
        explanation: "A pivot is a snapshot until you refresh it. Click in the pivot and press Refresh (or [[Alt]] + [[F5]]). Because the source is a Table, the new rows are picked up automatically." },
      { q: "What happens to the raw sales data when you sort, filter or rearrange a PivotTable?",
        options: ["It is sorted and filtered too", "It is unchanged; the pivot only reads it", "It is deleted", "It is copied to a new sheet"],
        answer: 1,
        explanation: "The pivot writes its summary in its own area and never modifies the source. That separation is exactly why analysts prefer pivots for exploration." }
    ],
    summary: [
      "A PivotTable is an automatic, drag-and-drop summary that never touches your raw data.",
      "Build pivots from an Excel Table so Refresh always includes new rows.",
      "Rows list downwards, Columns spread sideways, Values hold the numbers, Filters apply to the whole pivot.",
      "Change Sum, Count and Average in Value Field Settings, and set number formats there too.",
      "Refresh ([[Alt]] + [[F5]]) after any change to the source data."
    ],
    downloads: [
      { file: "assets/files/level4/sales-data.xlsx",
        label: "Sales dataset for charts and pivots",
        note: "250 rows of realistic sales data (date, region, salesperson, product, units, revenue) — Exercise sheet; pivot/chart builds are described in the lessons" }
    ],
    keywords: ["PivotTable", "pivot table", "Rows area", "Columns area", "Values area", "Filters area", "Value Field Settings", "Sum vs Count", "refresh pivot", "number formatting in pivot"]
  },
  {
    id: "l4-pivot-grouping",
    title: "Grouping and Calculations in PivotTables",
    difficulty: "Intermediate",
    duration: 18,
    version: "",
    objectives: [
      "Group dates by month, quarter and year in a PivotTable",
      "Group numeric fields into ranges",
      "Use Show Values As for percentages and running totals",
      "Create a calculated field from existing fields",
      "Sort, filter and drill down into pivot values"
    ],
    intro: "A PivotTable with 250 dates down the side is useless. Grouped by month, it becomes a one-page story of the year. Grouping, Show Values As and calculated fields turn a basic pivot into real analysis. Not just \"how much?\" but \"what share?\" and \"how has it accumulated?\".\n\nThis lesson continues with the same `sales-data.xlsx` workbook and the `SalesTable` you pivoted in the previous lesson. Every feature here builds on a pivot you already know how to create.",
    keyIdea: "Grouping, Show Values As and calculated fields turn a plain PivotTable into real analysis.",
    sections: [
      {
        heading: "Grouping dates by month, quarter and year",
        blocks: [
          { type: "p", text: "Put **Date** in Rows and the pivot lists every individual date, one row per transaction day. Right-click any date cell and choose Group. A dialog opens, offering options from Seconds up to Years. Select Months, Quarters and Years together and click OK." },
          { type: "p", text: "The pivot now shows years, quarters and months in a hierarchy. Click the minus buttons to collapse to quarters, or expand to drill into a month. This grouped date field is the backbone of monthly reporting. 250 rows become 12 meaningful lines." },
          { type: "warning", text: "In newer Excel versions, dates are sometimes grouped automatically the moment you drag them to Rows. That is fine; use the hierarchy as it is, or right-click > Group to change the levels. If grouping fails with an error, the dates are probably stored as text. Clean the column first so every value is a real date." },
          { type: "tip", text: "Group by Months only and different years merge together; all Januaries become one row. To keep January 2025 separate from January 2026, group by both Months and Years." },
          { type: "p", text: "Read a grouped pivot the way a manager would. With Date grouped by Months and Revenue in Values, each row is a month. The eye finds the strongest and weakest months at once. With Region in Columns, you can see whether a weak month was weak everywhere or only in one region." },
          { type: "p", text: "Grouping does not just shrink the table. It decides what level of story the table can tell. Before a review meeting, group to the level the meeting discusses: months for a monthly review, quarters for a board summary, days only when investigating a specific week." }
        ]
      },
      {
        heading: "Grouping numbers into ranges",
        blocks: [
          { type: "p", text: "The same trick works on numbers. Suppose you want to know whether big orders or small orders drive the business. Put **Units** in Rows and Revenue in Values. Then right-click a units value and choose Group." },
          { type: "ul", items: [
            "**Starting at**: the smallest value of the first range, for example 0.",
            "**Ending at**: the largest value of the last range, for example 100.",
            "**By**: the width of each range, for example 10, which creates 0-9, 10-19, 20-29 and so on."
          ] },
          { type: "p", text: "The pivot now answers a different question: how much revenue comes from orders of 0-9 units versus 10-19 units? If small frequent orders bring in a large share, that changes how Sunrise Traders thinks about delivery costs." },
          { type: "p", text: "Choose the range width to match the question. Grouping Units by 10 gives about ten ranges, which suits exploring order sizes. Grouping by 50 gives a coarse picture for a board pack slide." },
          { type: "p", text: "You can regroup at any time. Right-click a grouped value, choose Group again and change the settings. Or click Ungroup to return to individual values. Nothing in the source data is affected." }
        ]
      },
      {
        heading: "Show Values As: percentages and running totals",
        blocks: [
          { type: "p", text: "Sometimes the interesting number is not the value itself but its share or its accumulation. Show Values As re-displays any value field without changing the underlying sum." },
          { type: "table", headers: ["Show Values As option", "What the cell shows", "Sunrise Traders use"], rows: [
            ["% of Grand Total", "Each cell as a share of the whole pivot", "Which region contributes most of total revenue?"],
            ["% of Column Total", "Each cell as a share of its column", "Within the North, how do products split?"],
            ["Running Total In", "A cumulative sum across the field", "How is revenue accumulating month by month?"]
          ] },
          { type: "steps", items: [
            "Build a pivot with Revenue by Product (Product in Rows, Revenue in Values).",
            "Right-click a value cell and choose Show Values As > % of Grand Total.",
            "The same numbers now appear as percentages of ₹24,80,000.",
            "To keep both views, drag Revenue into Values a second time and set the copy to % of Grand Total.",
            "For a monthly view, group dates by month, add Revenue, and set Show Values As > Running Total In > Date (Months)."
          ] },
          { type: "grid", title: "Revenue by region with Show Values As columns",
            rows: [
              ["Region", "Revenue", "% of Grand Total", "Running Total"],
              ["North", "₹6,45,000", "26.0%", "₹6,45,000"],
              ["South", "₹5,90,000", "23.8%", "₹12,35,000"],
              ["East", "₹6,05,000", "24.4%", "₹18,40,000"],
              ["West", "₹6,40,000", "25.8%", "₹24,80,000"]
            ] },
          { type: "p", text: "The running total in a month pivot is effectively your year-to-date line, month after month. Managers read it as: by June we had already banked half the year." },
          { type: "p", text: "Pick the Show Values As option by asking: what should the 100% be? % of Grand Total answers \"how big is this piece of everything?\". % of Column Total answers \"how is this column split?\". % of Row Total answers the same question for a row." },
          { type: "p", text: "There is also Difference From. It shows each month against a chosen base month instead of a share. The values underneath never change; only the lens does. Keep the plain sum in one column and a Show Values As copy beside it, so the reader can always check the raw number." }
        ]
      },
      {
        heading: "Calculated fields: new numbers from existing fields",
        blocks: [
          { type: "p", text: "A calculated field is a formula you add to the pivot itself. It is built from the source fields and appears in the Values area like any other field. Use it when a ratio should come from the grouped totals, not row by row." },
          { type: "steps", items: [
            "Click inside the pivot.",
            "On the PivotTable Analyze tab, choose Fields, Items & Sets > Calculated Field.",
            "Type a name, for example \"Avg Price\".",
            "In the Formula box, keep the `=`, remove the default `0`, and enter `=Revenue/Units`.",
            "Click Add, then OK. The new field appears in the Values area.",
            "Open its Value Field Settings and format it as ₹ with two decimals."
          ] },
          { type: "formula", name: "Calculated field for average selling price",
            syntax: "=Revenue/Units",
            example: "=Revenue/Units",
            explanation: "For each region the pivot computes total revenue divided by total units, giving the blended average price per unit sold in that region." },
          { type: "p", text: "Understand what the field does: it works on the sums. On a region row, `=Revenue/Units` gives total revenue over total units. That is the correct blended average price." },
          { type: "p", text: "But a formula like `=Revenue/250` hard-codes a row count and breaks when data changes. So build calculated fields only from other fields." },
          { type: "warning", text: "Calculated fields cannot use cell references, named ranges or worksheet functions that need ranges. They combine pivot fields with basic arithmetic only. For anything more complex, compute a helper column in the source Table instead." },
          { type: "p", text: "Calculated fields are managed from the same dialog. Open Fields, Items & Sets > Calculated Field, choose the field's name in the dropdown, and modify or delete it." },
          { type: "p", text: "A common refinement for Sunrise Traders is a field named \"Rev per Order\". A plain field-to-field formula such as `=Revenue/Units` is the reliable pattern. Calculated fields always compute from the grouped sums of the fields you name. If a calculation needs to respect individual rows, for example a per-transaction discount, do it in a source Table column instead. Each row then carries its own result before the pivot sums it." }
        ]
      },
      {
        heading: "Sort, filter and drill down",
        blocks: [
          { type: "p", text: "**Sorting by value**: right-click any value cell and choose Sort > Largest to Smallest. The rows reorder by revenue, so the best region jumps to the top. Sorting by label (A to Z) is in the same menu." },
          { type: "p", text: "**Value filters**: click the Row Labels dropdown > Value Filters > Top 10. Change 10 to 3 and you see the top three products by revenue. Label filters do the same for names, for example Region begins with N. Slicers, in the next lesson, are simply button-based filters." },
          { type: "p", text: "**Drill down**: double-click any value cell, for example the East region total. Excel inserts a new sheet with exactly the source rows behind that number: every East transaction in the ₹6,05,000. This is the fastest way to answer \"show me the rows behind that figure\"." },
          { type: "warning", text: "The drill-down sheet is a copy, not a live view. It does not update when the source changes. Deleting rows on it does nothing to the pivot or the source. Treat it as a temporary extract and delete the sheet when done." },
          { type: "p", text: "Two housekeeping details keep pivots tidy. First, a \"(blank)\" row means some source rows have an empty cell in that field. Clean the source, refresh, and the blank row disappears." },
          { type: "p", text: "Second, after filtering a pivot, clear the filter before building the next view. Otherwise you will analyse a subset while believing you see everything. Only the filter dropdowns and icons reveal that a subset is on screen." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Trying to group a date column and getting an error.", right: "Check the column for text values and convert them to real dates first, then group.", why: "Excel can only group genuine dates or numbers. A single text entry such as \"30/06/2025 \" with a stray space makes the whole column text." },
      { wrong: "Confusing % of Column Total with % of Grand Total.", right: "Decide first what the 100% should be: the column, the row, or the whole pivot.", why: "With Product in Columns, % of Column Total shows each region's product mix; % of Grand Total shows each cell's share of the year. They answer different questions." },
      { wrong: "Writing a calculated field like `=Revenue/250` with a hard-coded row count.", right: "Build calculated fields only from other pivot fields, such as `=Revenue/Units`.", why: "A hard-coded number stops being true the moment rows are added or refreshed. Field-to-field formulas stay correct automatically." },
      { wrong: "Editing the drill-down sheet and expecting the pivot or source to change.", right: "Fix data in the source Table, refresh, then drill down again if needed.", why: "Drill-down output is a one-time copy of the rows behind a value. It is a reading aid, not a data entry screen." }
    ],
    shortcuts: [],
    exercises: [
      { task: "Create a pivot with Date in Rows and Sum of Revenue in Values, then group the dates by Months and Years. How many month rows do you get?",
        hint: "Right-click any date cell in the pivot and choose Group. Tick Months and Years.",
        solution: "The 250 daily rows collapse into 12 month rows under the year 2025. Collapse the year to see the 12-month summary at a glance." },
      { task: "Show Revenue by Product twice: once as the plain sum and once as % of Grand Total.",
        hint: "Drag Revenue into Values twice. Right-click the second copy > Show Values As > % of Grand Total.",
        solution: "Laptops contribute about 33% of total revenue (₹8,20,000 of ₹24,80,000), followed by Printers. The percentages should add to 100%." },
      { task: "Apply a Value Filter to show only the top 3 salespeople by Sum of Revenue.",
        hint: "Put Salesperson in Rows, then use the Row Labels dropdown > Value Filters > Top 10 and change 10 to 3.",
        solution: "Three of the four names remain, listed with their revenue. Sort Largest to Smallest afterwards to see the ranking clearly." },
      { task: "Create a calculated field named \"Avg Price\" with the formula `=Revenue/Units`, placed next to Revenue by region, formatted as ₹ with two decimals.",
        hint: "PivotTable Analyze > Fields, Items & Sets > Calculated Field. Type the name, then `=Revenue/Units` in the Formula box.",
        solution: "Each region row now shows total revenue and the blended average price per unit. Compare the Solution sheet: prices should be in a similar range across regions." }
    ],
    quiz: [
      { q: "How do you group a date field in a PivotTable by month?",
        options: ["Insert > Group from the ribbon's Data tab only", "Right-click any date cell in the pivot and choose Group", "You cannot group dates in a pivot", "Add a helper column with MONTH in the pivot area"],
        answer: 1,
        explanation: "Right-click any date cell in the Rows area and choose Group, then select Months, Quarters or Years. Excel builds the hierarchy for you." },
      { q: "Which Show Values As option gives you a cumulative year-to-date figure across months?",
        options: ["% of Grand Total", "Difference From", "Running Total In", "Index"],
        answer: 2,
        explanation: "Running Total In adds each month on top of the previous ones, which is exactly a year-to-date accumulation. % of Grand Total shows shares instead." },
      { q: "What does a calculated field in a PivotTable do?",
        options: ["Edits the source data with a formula", "Creates a new value field computed from existing pivot fields", "Groups rows automatically", "Formats numbers as currency"],
        answer: 1,
        explanation: "A calculated field is a new Values-area field defined by a formula over other fields, such as `=Revenue/Units`. It never modifies the source." },
      { q: "What happens when you double-click a value cell in a PivotTable?",
        options: ["The cell becomes editable", "Excel opens the Value Field Settings dialog", "Excel creates a new sheet with the source rows behind that value", "The pivot refreshes"],
        answer: 2,
        explanation: "Double-clicking drills down: Excel inserts a sheet listing exactly the source rows that sum to that value. It is a copy, so treat it as a temporary extract." }
    ],
    summary: [
      "Right-click a date or number in a pivot and choose Group to build month, quarter, year or range summaries.",
      "Show Values As re-expresses the same sums as percentages or running totals without extra formulas.",
      "Calculated fields such as `=Revenue/Units` add ratio analysis inside the pivot, built only from pivot fields.",
      "Sort by values and apply Top N value filters to surface winners quickly.",
      "Double-click any value to drill down to the exact source rows behind it, remembering the output is a copy."
    ],
    downloads: [
      { file: "assets/files/level4/sales-data.xlsx",
        label: "Sales dataset for charts and pivots",
        note: "250 rows of realistic sales data (date, region, salesperson, product, units, revenue) — Exercise sheet; pivot/chart builds are described in the lessons" }
    ],
    keywords: ["pivot grouping", "group dates", "group months", "group numbers", "Show Values As", "% of Grand Total", "running total", "calculated field", "value filters", "drill down"]
  },
  {
    id: "l4-pivot-charts-slicers",
    title: "Pivot Charts, Slicers and Timelines",
    difficulty: "Intermediate",
    duration: 20,
    version: "Excel 2013+",
    objectives: [
      "Insert a PivotChart directly from a PivotTable",
      "Filter pivots with slicer buttons",
      "Use a timeline to filter by date periods",
      "Connect one slicer to several PivotTables with Report Connections",
      "Assemble a simple interactive report page"
    ],
    intro: "Your pivots answer questions, but the reader still has to open filter dropdowns and click tiny menus. Slicers turn filtering into big, friendly buttons that anyone can use. Timelines do the same for dates. PivotCharts add the visual layer on top. Together they turn a static pivot into an interactive report.\n\nThis lesson uses `sales-data.xlsx` again. Everything builds on the pivoting and grouping skills from the last two lessons. Timelines need Excel 2013 or later; slicers work from Excel 2010.",
    keyIdea: "Slicers turn pivot filtering into big buttons, and PivotCharts keep the picture in sync with the numbers.",
    sections: [
      {
        heading: "From pivot to PivotChart",
        blocks: [
          { type: "p", text: "A PivotChart is a chart bound to a PivotTable. Change the pivot and the chart changes with it. Filter the pivot and the chart filters too. You never build the chart series by hand." },
          { type: "steps", items: [
            "Build or open a pivot, for example Sum of Revenue with Date grouped by month in Rows and Region in Columns.",
            "Click inside the pivot.",
            "On the PivotTable Analyze tab, click PivotChart.",
            "Choose Clustered Column and click OK.",
            "The chart appears beside the pivot, already titled \"Sum of Revenue\". Edit the title to say something meaningful, such as \"Monthly revenue by region, 2025\"."
          ] },
          { type: "p", text: "The pivot and the chart are two views of the same fields. Drag a new field into Columns and the chart re-plots instantly. Filter the pivot and the chart shrinks to match." },
          { type: "p", text: "This is the key difference from the ordinary charts of the first lesson. A PivotChart cannot drift out of sync with its numbers." },
          { type: "p", text: "A few PivotChart quirks are worth knowing before a meeting. The chart shows small field buttons on the plot area. They are useful for exploring but cluttered for presentation. Right-click one and choose Hide All Field Buttons on Chart before showing it." },
          { type: "p", text: "Collapsed pivot groups collapse in the chart too, so a date hierarchy collapsed to quarters shows a shorter, cleaner chart. If you move or resize the pivot, the chart follows its data, not its position. So reposition both deliberately when you lay out a report page." },
          { type: "p", text: "When should you use an ordinary chart instead? Use one when the summary is fixed and formula-driven, such as a `SUMIFS` region summary. An ordinary chart is simpler and has no pivot buttons to hide. PivotCharts are best when the reader will filter and explore." },
          { type: "p", text: "One warning: not every chart type works as a PivotChart. Column, bar, line, pie and area work fine. Scatter and bubble are not offered, because a pivot's category structure does not fit them. For those cases, chart a `SUMIFS` summary instead." }
        ]
      },
      {
        heading: "Slicers: filtering with buttons",
        blocks: [
          { type: "p", text: "A slicer is a floating panel of filter buttons for one field. Instead of hunting through a dropdown menu, you click a button. In a review meeting, this is the difference between fumbling with Excel and steering the discussion." },
          { type: "p", text: "Slicers also fix a big problem with dropdowns: you cannot see the filter state. A dropdown filter hides inside a small icon, and a reader can easily miss that the pivot is filtered at all." },
          { type: "p", text: "A slicer shows its selection as pressed buttons in plain sight. Everyone in the room can see what is on screen. That transparency, more than the speed, made slicers the standard in shared reports." },
          { type: "steps", items: [
            "Click inside the pivot.",
            "On the PivotTable Analyze tab (or Insert tab), click Slicer.",
            "Tick Region and Product, then click OK.",
            "Two slicer panels appear. Click North in the Region slicer: the pivot shows only the North, and the Product slicer grays out products the North never sold.",
            "Hold [[Ctrl]] and click to select several buttons at once, for example North and East.",
            "Click the Clear Filter icon at the top-right corner of a slicer to reset it."
          ] },
          { type: "p", text: "Slicers show their state clearly: selected buttons are highlighted, and unavailable combinations are grayed out. That visible state is worth a lot when someone else takes over your file." },
          { type: "tip", text: "Format a slicer to fit your report. Select it and use the Slicer tab to pick a style. Set Columns to 2 or 3 so tall lists become compact grids. Resize and position slicers beside the pivot, and turn off the sheet gridlines." },
          { type: "p", text: "Give each slicer a clear caption when the field name is not self-explanatory. With the slicer selected, the Slicer tab has a Caption box. Renaming \"Salesperson\" to \"Filter by person\" takes two seconds." },
          { type: "p", text: "Practise multi-select until it is reflex. A plain click replaces the selection. [[Ctrl]] + click adds to it. Clicking the selected button again deselects it. After exploring, the clear-filter icon at the top-right corner is the fastest reset." }
        ]
      },
      {
        heading: "Timelines: slicers for dates",
        blocks: [
          { type: "p", text: "A timeline is a slicer designed for date fields. It shows a horizontal strip of years and quarters, and one drag selects the period to display." },
          { type: "steps", items: [
            "Click inside a pivot whose Rows contain the Date field.",
            "On the PivotTable Analyze tab, click Insert Timeline.",
            "Select Date and click OK.",
            "In the timeline strip, switch the level dropdown between Years, Quarters and Months.",
            "Drag the window across the strip to select, say, April to June, and watch the pivot and PivotChart update."
          ] },
          { type: "p", text: "The timeline only appears for pivots that contain a real date field. If Insert Timeline is grayed out, your dates are text, or the Date field is not in the pivot yet. Fix that first. Timelines need Excel 2013 or later." },
          { type: "p", text: "Timelines and slicers work together naturally. In a sales review, set the period on the timeline first, then click through regions on the slicer as the discussion moves." },
          { type: "p", text: "Both controls stay visible on the sheet, so the person driving Excel can answer any question in two clicks. For example: \"what was the South doing in the second quarter?\". That responsiveness is what makes a pivot-based report feel interactive." },
          { type: "warning", text: "Before you send any interactive file to a manager, click Clear Filter on every slicer and timeline. A file that opens showing only \"West, Staplers, March\" looks broken to the next reader, even though it is just a leftover filter." },
          { type: "p", text: "One detail makes timelines feel professional: match the default level to your reporting rhythm. If your meetings discuss months, set the level dropdown to Months before saving. A quarterly default forces everyone to fumble at the start of every meeting. The choice is saved with the file, so setting it once is enough." }
        ]
      },
      {
        heading: "One slicer, many pivots: Report Connections",
        blocks: [
          { type: "p", text: "A serious report page usually has more than one pivot: revenue by month, revenue by product, and a KPI table. You do not want three separate slicers for Region. Connect one slicer to all of them. Then a single click filters the entire page." },
          { type: "steps", items: [
            "Place two or three pivots on the same sheet, built from the same source table `SalesTable`.",
            "Insert one slicer for Region.",
            "Right-click the slicer and choose Report Connections (in older versions: PivotTable Connections).",
            "Tick every pivot that should obey the slicer and click OK.",
            "Click a region button: all connected pivots filter together."
          ] },
          { type: "p", text: "Connections only work between pivots that share the same source data. That is one more reason to keep one clean Table as your single source of truth. A slicer connected to pivots from different tables cannot sync them." },
          { type: "p", text: "Prepare an interactive report for handover the way you would prepare a projector. Refresh every pivot so the numbers are current. Clear all slicers and timelines. Check that each grand total matches the expected full-year figure." },
          { type: "p", text: "Save, then reopen the file once to see exactly what the next person will see: an unfiltered, up-to-date report page. This two-minute ritual is the difference between a file that impresses and a file that raises questions." },
          { type: "grid", title: "A simple interactive report layout on one sheet",
            rows: [
              ["SUNRISE TRADERS — SALES REPORT", "", "", "", ""],
              ["[Slicer: Region]", "[Slicer: Product]", "[Timeline: Date]", "", ""],
              ["Pivot 1: Revenue by month", "", "Pivot 2: Revenue by product", "", ""],
              ["PivotChart: monthly trend", "", "Pivot 3: KPI summary", "", ""]
            ] },
          { type: "p", text: "That one-page layout is a real deliverable: title on top, one control strip for slicers and timeline, and every table and chart obeying them. You will refine this pattern into a full dashboard later in this level." },
          { type: "p", text: "Build the page in this order and it falls into place. Pivots first, so you know their size. PivotCharts second, aligned beside them. Slicers and timeline last, sized to fill the leftover strip." },
          { type: "p", text: "Placing slicers first always ends with moving everything twice. Leave one empty row between blocks as breathing room. And keep the pivot sheets in the workbook, because a reader who wants detail will drill down from the pivot." }
        ]
      }
    ],
    mistakes: [
      { wrong: "Sending the file with slicers still filtered from your last exploration.", right: "Clear all slicer and timeline filters before saving or exporting.", why: "The next reader sees a partial view and assumes the report is wrong or incomplete. Clear filters make the default view show everything." },
      { wrong: "Building three pivots from three different copied ranges and wondering why one slicer cannot control them all.", right: "Keep one source Table and build every pivot from it, then use Report Connections.", why: "Slicers can only drive pivots that share the same source. Duplicated ranges also risk drifting out of sync." },
      { wrong: "Trying to create a PivotChart of type scatter or bubble.", right: "Use column, bar, line, pie or area PivotCharts, or build a scatter from a `SUMIFS` summary instead.", why: "PivotCharts plot pivot categories, and the pivot category structure does not support scatter or bubble types." },
      { wrong: "Leaving the pivot chart title as \"Sum of Revenue\".", right: "Type a specific title such as \"Monthly revenue by region, 2025\".", why: "A generic title forces the reader to reconstruct the context. The title is where your chart states its question." }
    ],
    shortcuts: [
      { keys: "Alt + F5", action: "Refresh the selected PivotTable after source changes" }
    ],
    exercises: [
      { task: "In sales-data.xlsx, insert a PivotChart from a pivot of monthly revenue (dates grouped by month). Use a line chart.",
        hint: "Click inside the pivot, then PivotTable Analyze > PivotChart > Line.",
        solution: "The chart should show one line rising and falling across the twelve months of 2025, updating automatically whenever you change the pivot layout." },
      { task: "Add slicers for Region and Product to your pivot and practice: show only East, then East and North together, then clear the filter.",
        hint: "PivotTable Analyze > Slicer. Multi-select with [[Ctrl]] + click; clear with the funnel-with-x icon on the slicer.",
        solution: "When East only is selected, the pivot total equals East's ₹6,05,000. When North and East are both selected, the total is their sum, ₹12,50,000." },
      { task: "Insert a timeline for the Date field and restrict the pivot to the second quarter only.",
        hint: "PivotTable Analyze > Insert Timeline. Set the level dropdown to Quarters and select Q2.",
        solution: "The pivot now sums only April, May and June transactions. Clear the timeline filter when finished." },
      { task: "Build two pivots from `SalesTable` on one sheet (revenue by month, revenue by product) and connect a single Region slicer to both with Report Connections.",
        hint: "Right-click the slicer > Report Connections and tick both pivots.",
        solution: "Clicking West should now filter both pivots at once. Verify both show only West's figures, then clear the filter before closing." }
    ],
    quiz: [
      { q: "What is the main difference between a PivotChart and the ordinary charts from the first lesson?",
        options: ["PivotCharts are always 3-D", "A PivotChart is bound to a PivotTable and filters with it", "PivotCharts cannot have titles", "Ordinary charts cannot show columns"],
        answer: 1,
        explanation: "A PivotChart reads its data from a PivotTable, so changing or filtering the pivot instantly updates the chart. They can never drift out of sync." },
      { q: "What does a slicer do?",
        options: ["It slices the chart into sections", "It filters a pivot with clickable buttons", "It splits data into multiple sheets", "It sorts values"],
        answer: 1,
        explanation: "A slicer is a panel of filter buttons for one field. Clicking buttons filters the connected pivot, and multi-select is possible with [[Ctrl]] + click." },
      { q: "A timeline can only be inserted for which kind of pivot field?",
        options: ["Any text field", "Numeric fields", "Date fields", "Fields in the Values area"],
        answer: 2,
        explanation: "Timelines are date slicers: they show years, quarters and months on a draggable strip. They need Excel 2013 or later." },
      { q: "How do you make one Region slicer filter three different PivotTables on the same sheet?",
        options: ["It happens automatically for all pivots", "Right-click the slicer > Report Connections and tick the pivots", "Copy the slicer three times", "Group the three pivots with Ctrl + G"],
        answer: 1,
        explanation: "Report Connections (PivotTable Connections in some versions) links the slicer to every pivot that shares the same source data." }
    ],
    summary: [
      "A PivotChart is created straight from a pivot with PivotTable Analyze > PivotChart and stays in sync with it.",
      "Slicers replace filter menus with buttons; use [[Ctrl]] + click for multi-select and the clear icon to reset.",
      "Timelines filter pivots by year, quarter or month and require Excel 2013 or later.",
      "Report Connections makes one slicer control several pivots built from the same source table.",
      "Always clear slicer filters before sending an interactive file to someone else."
    ],
    downloads: [
      { file: "assets/files/level4/sales-data.xlsx",
        label: "Sales dataset for charts and pivots",
        note: "250 rows of realistic sales data (date, region, salesperson, product, units, revenue) — Exercise sheet; pivot/chart builds are described in the lessons" }
    ],
    keywords: ["PivotChart", "pivot chart", "slicers", "slicer", "timeline", "Report Connections", "interactive report", "filter buttons", "clear filters", "pivot dashboard"]
  }
);
