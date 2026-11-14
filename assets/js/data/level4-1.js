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
    intro: "A table of 250 rows tells you everything and shows you nothing. A chart shows the one thing your manager needs to see: which region is winning, whether sales are growing, which product carries the business. That is why every monthly review meeting starts with charts, not with raw data.\n\nIn this lesson you will use the sales dataset of Sunrise Traders, a retail and wholesale office supplies company. The workbook `sales-data.xlsx` has 250 rows on the Sales Data sheet, with columns Date, Region, Salesperson, Product, Units and Revenue, formatted as a Table named `SalesTable`. You already know how to build summaries with `SUMIFS` from Level 3, so here you will summarise first and chart second, which is the habit of every good analyst.",
    sections: [
      {
        heading: "Why a picture beats a table",
        blocks: [
          { type: "p", text: "Excel charts exist for one reason: to answer a question at a glance. Before you insert any chart, ask yourself what question you are answering. The question decides the chart type, not the other way round." },
          { type: "ul", items: [
            "**Which region sold more?** You are comparing categories, so a column chart is right.",
            "**Is revenue growing month by month?** That is a trend over time, so a line chart is right.",
            "**How much of the total comes from each product?** That is a share of a whole, so a pie chart is right.",
            "**Which salesperson hit target?** You are ranking a list, so a sorted column chart or a bar chart is right."
          ] },
          { type: "p", text: "If you start from the question, you will rarely pick the wrong chart. If you start from whatever chart looks pretty, your report will confuse people even when the data is correct." },
          { type: "p", text: "Think about where charts actually appear in office life. The weekly team meeting opens with a trend line. The monthly review deck leads with a revenue-by-region column chart. A performance dashboard keeps a row of small charts under the KPI numbers. In all three places, the chart is not decoration: it is the fastest way for a busy reader to form a judgement, and the reader forms it in about three seconds. Your job is to make those three seconds count." }
        ]
      },
      {
        heading: "Creating your first chart",
        blocks: [
          { type: "p", text: "Excel cannot chart 250 raw rows into something meaningful by itself. You first build a small summary table, then chart that summary. On the Sales Data sheet, type `Region` in `H1` and `Revenue` in `I1`, the four region names in `H2:H5`, and this formula in `I2`:" },
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
          { type: "p", text: "When a chart is selected, two extra tabs appear on the ribbon: Chart Design and Format. Use Chart Design to change the chart type, swap rows and columns, or pick a style. Use Format only for colors and shapes." },
          { type: "p", text: "One more thing worth understanding early: the chart is linked to the cells, not a copy of them. Change a figure in the summary table, or refresh the source data behind a `SUMIFS`, and the chart redraws itself. This is why the chart must point at a small, stable summary range like `H1:I5` and not at 250 transaction rows. Small summary, clean chart, always current." },
          { type: "tip", text: "Select the summary range and press [[Alt]] + [[F1]] to drop a chart of it straight onto the sheet. Press [[F11]] instead to place the chart on its own full-page chart sheet, which is handy when the chart is the main output of a report." },
          { type: "p", text: "Handle charts with care once they exist. Click the chart border to move the whole object; click inside on a column or line to select that series for formatting. Resize only from the corner handles, which keep the proportions, and avoid stretching from the side handles, which distorts text inside the chart. When a report holds several charts, move and size them against the cell grid so edges line up; hold [[Alt]] while dragging to snap to it." }
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
          { type: "p", text: "Two column-chart habits raise quality immediately. Sort the categories by value before charting, so the columns fall from tallest to shortest and the ranking is obvious without reading a single label. And keep the value axis starting at zero: a column chart with a truncated axis makes a 5% difference look like a 50% difference, which reads as exaggeration rather than insight." },
          { type: "p", text: "**Line charts** are for time. Put dates or month names on the horizontal axis and the measure on the vertical axis. The slope of the line tells the story: rising is growth, falling is trouble. Never use a line chart for unrelated categories such as products, because the connecting lines imply an order that does not exist." },
          { type: "p", text: "**Pie charts** show how a total is divided. One pie, one measure, one period. Keep it to 5 or 6 slices maximum; with 12 products the slices become unreadable slivers. If you have more categories, use a column chart instead. A pie of Sunrise Traders revenue by product with five products works well." },
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
          { type: "p", text: "For colors and styles, use the Chart Styles gallery on the Chart Design tab instead of formatting each part by hand. If you need finer control, press [[Ctrl]] + [[1]] with an element selected to open the Format pane, where you can set fill, border and number formats for that element." },
          { type: "tip", text: "Change the color of a single column to highlight it: click the series once to select all columns, click again on one column to select only it, then use Format > Shape Fill. Highlighting the weakest region in a different color is a simple way to focus a discussion." },
          { type: "p", text: "Two small habits make charts look professional: keep one consistent color per series across all charts in a report, and delete anything that does not help read the chart, such as heavy borders, vertical gridlines and unnecessary legends." },
          { type: "p", text: "Watch the axis scale when you compare charts. Two revenue charts side by side, one scaled from 0 to 10 lakh and the other from 4 to 6 lakh, will look completely different even if they show the same data. Excel usually chooses the scale automatically, so glance at the axis numbers before you present: a fair comparison needs comparable axes." }
        ]
      },
      {
        heading: "Sparklines: tiny charts in a cell",
        blocks: [
          { type: "p", text: "A sparkline is a miniature chart that lives inside a single cell. It is perfect for a row-per-item summary where you want to see each item's trend without inserting twelve full charts." },
          { type: "steps", items: [
            "Put the twelve monthly revenue figures in a row or column, for example `C2:N2` for one region.",
            "Click the cell where you want the sparkline, for example `P2`.",
            "On the Insert tab, in the Sparklines group, click Line.",
            "In the Data Range box type `C2:N2` and click OK.",
            "Fill the sparkline down for the other regions."
          ] },
          { type: "p", text: "Sparklines update when the numbers change, print with the sheet, and can be formatted with markers for high and low points from the Sparkline tab that appears when they are selected. Use them in KPI tables and month summaries; you will meet them again in the KPI reporting lesson." },
          { type: "p", text: "Excel offers three sparkline types from the same Insert menu. Line is the default and best for trends. Column shows each period as a tiny bar, good for spotting a single unusually strong month. Win/Loss shows only up or down markers, designed for sequences of wins and losses. For revenue and sales reporting, line sparklines cover almost every need." }
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
    intro: "Ask ten questions of your data and a PivotTable answers them all without a single formula. Drag Region into rows and you get revenue by region. Drag Product there instead and you get revenue by product. Add months and you get revenue by month and product. Analysts love PivotTables because a summary that takes twenty `SUMIFS` formulas takes twenty seconds of dragging.\n\nThe second reason is safety. A PivotTable never touches your raw data. It reads the source and writes its own summary area, so sorting, filtering and rearranging the pivot leave the 250 rows of Sunrise Traders sales data exactly as they were. In this lesson you will build your first pivots from the `SalesTable` in `sales-data.xlsx`.",
    sections: [
      {
        heading: "What a PivotTable is and why analysts use it",
        blocks: [
          { type: "p", text: "A PivotTable is an automatic summary of a dataset. You choose which fields to group by and which numbers to add up, and Excel builds the cross-tab for you. The name comes from the fact that you can pivot, or rotate, the layout: move a field from rows to columns and the whole summary reorganises instantly." },
          { type: "ul", items: [
            "**Speed**: one revenue-by-region-and-product summary would take many `SUMIFS` formulas; a pivot takes seconds.",
            "**Flexibility**: change your mind about the breakdown and just drag a different field in.",
            "**Safety**: the pivot is a separate summary; your raw rows are never modified.",
            "**Freshness**: one Refresh command pulls in any new or changed source rows."
          ] },
          { type: "p", text: "Typical office use: every month-end you get the transactions export, and the review meeting needs revenue by region, by product and by salesperson. With a pivot, all three views live on one sheet and update from the same source." },
          { type: "p", text: "Compare that with the formula approach from Level 3. A `SUMIFS` summary is excellent when the layout is fixed and final, but every new breakdown is a new block of formulas, and every change of mind is an hour of editing. A pivot answers the \"what if we looked at it by...?\" questions in real time, which is exactly how analysis happens in a meeting. Many analysts use both: pivots to explore, `SUMIFS` for the fixed report layout that gets published every month." }
        ]
      },
      {
        heading: "Creating a PivotTable from a Table",
        blocks: [
          { type: "p", text: "Your data should be an Excel Table first, as `SalesTable` already is. Tables grow automatically when you add rows, which means the pivot picks up new data on refresh without any re-selection." },
          { type: "steps", items: [
            "Open sales-data.xlsx and click any cell inside the Sales Data table.",
            "On the Insert tab, click PivotTable.",
            "Excel pre-fills Table/Range with `SalesTable`. Keep it.",
            "Choose New Worksheet so the pivot gets its own clean sheet.",
            "Click OK. Excel creates a blank pivot placeholder on the left and the PivotTable Fields pane on the right."
          ] },
          { type: "p", text: "The Fields pane lists every column of the table: Date, Region, Salesperson, Product, Units and Revenue. Below it are four drop areas: Filters, Columns, Rows and Values. Everything you build happens by dragging fields between this list and those areas." },
          { type: "tip", text: "If the Fields pane disappears, click anywhere inside the pivot and it comes back. Renaming the new sheet to something like \"PT Revenue\" right away keeps big workbooks easy to navigate." },
          { type: "p", text: "When a pivot first takes shape it reads \"Row Labels\" and \"Sum of Revenue\", which is Excel's wording, not yours. Click the cell that says Row Labels and type \"Region\". Rename value headers the same way through Value Field Settings, using names a manager would write, such as \"Revenue (₹)\". Renamed headers survive refreshes and make the pivot readable to someone who has never seen it before." }
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
          { type: "p", text: "Now drag **Product** to Columns as well. You get a cross-tab: regions downwards, products sideways, each cell answering \"how much revenue did the North make from Printers?\". Finally drag **Salesperson** to Filters and a filter dropdown appears above the pivot, letting you view one person's numbers without changing the layout." },
          { type: "p", text: "A useful habit: put the field with the fewest items in Columns and the field with more items in Rows. Five products across the top stays readable; twelve months across the top is often too wide." },
          { type: "p", text: "You can stack more than one field in an area. Drop Salesperson under Region in Rows and the pivot lists the four salespeople inside each region, collapsible with the minus buttons beside each region name. Two fields in Values work the same way: Sum of Revenue next to Sum of Units in one pivot. Excel adds automatic subtotals for outer row fields, so each region shows its own total above its salespeople. Start with one field per area while learning, then add the second only when the question needs it." }
        ]
      },
      {
        heading: "Sum, Count and Average: changing the calculation",
        blocks: [
          { type: "p", text: "By default Excel sums numeric fields in Values. But if the field contains text or blank cells, the pivot silently switches to Count, which surprises many beginners: the value header suddenly reads \"Count of Revenue\" instead of \"Sum of Revenue\". Count merely counts rows; it does not add anything." },
          { type: "steps", items: [
            "Right-click any number inside the Values area.",
            "Choose Value Field Settings.",
            "In the Summarize by list choose Sum, Count, Average, Max or Min.",
            "Type a clear Custom Name, for example \"Average Units\", so the pivot header means something.",
            "Click OK."
          ] },
          { type: "p", text: "When is each calculation right? Use **Sum** for money and units: total revenue, total units. Use **Count** when the question is \"how many transactions?\". Use **Average** for typical values: average order size per region. For example, put Units in Values twice, once as Sum of Units and once, via Value Field Settings, as Average Units, and you can see both volume and typical order size side by side." },
          { type: "p", text: "A worked example makes the difference concrete. Suppose the North sold 3,800 units across 95 transactions. Sum of Units says 3,800, Count of Units says 95 (the number of rows with a units value), and Average of Units says 40, the typical order size. Three different numbers from the same column, each answering a different question. Choosing the calculation is therefore a business decision, not an Excel setting: ask what the reader needs to learn before you right-click." },
          { type: "p", text: "Number formatting belongs in the same dialog. In Value Field Settings, click the Number Format button and choose Currency with the ₹ symbol and no decimals. This formatting lives inside the pivot field, so it survives refreshes, moves and layout changes." },
          { type: "warning", text: "Do not format pivot numbers by selecting the cells and applying a cell format. The moment you refresh, move or restructure the pivot, those hand-applied formats land on the wrong cells. Always set number formats through Value Field Settings > Number Format." }
        ]
      },
      {
        heading: "Refresh, and your data stays untouched",
        blocks: [
          { type: "p", text: "A PivotTable is a snapshot of the source at the moment it was built. If you add rows to `SalesTable`, edit revenue figures or correct a region name, the pivot does not change until you refresh it. Click inside the pivot and press PivotTable Analyze > Refresh, or right-click the pivot and choose Refresh. The shortcut is [[Alt]] + [[F5]]." },
          { type: "p", text: "Because `SalesTable` is a Table, any rows you append below it are included in the table automatically, and the next refresh pulls them into the pivot. That is the full update cycle for a monthly report: paste the new month's rows into the source table, refresh the pivots, done." },
          { type: "p", text: "Sorting the pivot, filtering it, collapsing and expanding groups, even deleting the pivot sheet entirely: none of it affects the source rows. The pivot only reads. The one thing you cannot do is type directly into the values area; Excel blocks it, because pivot cells are calculated, not stored." },
          { type: "tip", text: "If a pivot looks wrong after a data edit, refresh first before debugging. Nine out of ten \"my pivot is broken\" moments are simply a stale snapshot waiting for a refresh." },
          { type: "p", text: "When a workbook holds several pivots, refresh them all at once with Data > Refresh All on the ribbon's Data tab. This matters in monthly reporting, where five pivots fed by one table must show the same version of the truth. If one pivot is refreshed and another is not, the two views disagree, and that discrepancy is far more damaging in a meeting than being a day late." }
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
    intro: "A PivotTable with 250 dates down the side is useless. Grouped by month it becomes a one-page story of the year. Grouping, Show Values As and calculated fields turn a basic pivot into real analysis: not just \"how much?\" but \"what share?\", \"how has it accumulated?\" and \"what is the average price?\".\n\nThis lesson continues with the same `sales-data.xlsx` workbook and the `SalesTable` you pivoted in the previous lesson. Every feature here is built on top of a pivot you already know how to create.",
    sections: [
      {
        heading: "Grouping dates by month, quarter and year",
        blocks: [
          { type: "p", text: "Put **Date** in Rows and the pivot lists every individual date, one row per transaction day. Right-click any date cell and choose Group. Excel opens a grouping dialog offering Seconds up to Years; select Months, Quarters and Years together and click OK." },
          { type: "p", text: "The pivot now shows years, quarters and months in a hierarchy. Click the collapse buttons (the minus signs) to see only quarters, or expand to drill into a month's days. This grouped date field is the backbone of monthly reporting: 250 rows become 12 meaningful lines." },
          { type: "warning", text: "In newer Excel versions, date fields are sometimes grouped automatically into Years, Quarters and Months the moment you drag them to Rows. That is fine; use the resulting hierarchy as it is, or right-click > Group to change the levels. If grouping fails with an error, the dates are probably stored as text; clean the column first so every value is a real date." },
          { type: "tip", text: "Group by Months only and different years will merge together (all Januaries become one row). If you want January 2025 separate from January 2026, group by both Months and Years." },
          { type: "p", text: "Read a grouped pivot the way a manager would. With Date grouped by Months and Revenue in Values, each row is a month and the eye immediately finds the strongest and weakest months. With Region in Columns, each month splits into four regions, and you can see whether a weak month was weak everywhere or only in one region. Grouping does not just shrink the table; it decides what level of story the table can tell. Before every review meeting, group to the level the meeting discusses: months for a monthly review, quarters for a board summary, days only when investigating a specific week." }
        ]
      },
      {
        heading: "Grouping numbers into ranges",
        blocks: [
          { type: "p", text: "The same trick works on numeric fields. Suppose you want to know whether big orders or small orders drive the business. Put **Units** in Rows and Revenue in Values, then right-click a units value and choose Group." },
          { type: "ul", items: [
            "**Starting at**: the smallest value of the first range, for example 0.",
            "**Ending at**: the largest value of the last range, for example 100.",
            "**By**: the width of each range, for example 10, which creates 0-9, 10-19, 20-29 and so on."
          ] },
          { type: "p", text: "The pivot now answers a different kind of question: how much revenue comes from orders of 0-9 units versus 10-19 units? If a large share of revenue comes from small frequent orders, that changes how Sunrise Traders should think about delivery costs." },
          { type: "p", text: "Choose the range width to match the business question. Grouping Units by 10 gives ten or so ranges, which suits an exploratory look at order sizes. Grouping by 50 gives a coarse picture for a slide in a board pack. You can regroup at any time: right-click a grouped value, choose Group again, change the settings, or click Ungroup to return to individual values. Nothing in the source data is affected by any of it." }
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
          { type: "p", text: "Pick the Show Values As option by asking what the 100% should be. % of Grand Total answers \"how big is this piece of everything?\". % of Column Total answers \"how is this column split?\". % of Row Total answers the mirror question for a row. There is also Difference From, which shows each month against a chosen base month instead of a share. The values underneath never change; only the lens does. Keeping the plain sum in one column and a Show Values As copy beside it is the safest layout, because the reader can always check the raw number." }
        ]
      },
      {
        heading: "Calculated fields: new numbers from existing fields",
        blocks: [
          { type: "p", text: "A calculated field is a formula you add to the pivot itself, built from the source fields. It appears in the Values area like any other field. Use it when a ratio or difference should be computed from the grouped totals, not row by row." },
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
          { type: "p", text: "Understand what the field does: it works on the sums. `=Revenue/Units` on a region row gives total revenue over total units, which is the correct blended average. But a formula like `=Revenue/250` hard-codes a row count and breaks when data changes, so build calculated fields only from other fields." },
          { type: "warning", text: "Calculated fields cannot use cell references, named ranges or worksheet functions that need ranges. They combine pivot fields with basic arithmetic only. For anything more complex, compute a helper column in the source Table instead." },
          { type: "p", text: "Calculated fields are managed from the same dialog. Open Fields, Items & Sets > Calculated Field, choose the field's name in the dropdown, and modify or delete it. A common refinement for Sunrise Traders is a field named \"Rev per Order\" with `=Revenue/Count of Revenue`-style logic; note that a plain field-to-field division such as `=Revenue/Units` is the reliable pattern, because calculated fields always compute from the grouped sums of the fields you name. If a calculation needs to respect individual rows, for example a discount applied per transaction, do it in the source Table column instead, so each row carries its own result before the pivot sums it." }
        ]
      },
      {
        heading: "Sort, filter and drill down",
        blocks: [
          { type: "p", text: "**Sorting by value**: right-click any value cell and choose Sort > Largest to Smallest. The rows reorder by revenue, so the best region or product jumps to the top. Sorting by the label (A to Z) is available in the same menu." },
          { type: "p", text: "**Value filters**: click the Row Labels dropdown > Value Filters > Top 10. Change 10 to 3 and you instantly see the top three products by revenue. Label filters do the same for names, for example Region begins with N. Slicers, which you will meet in the next lesson, are simply a button-based way to filter." },
          { type: "p", text: "**Drill down**: double-click any value cell, for example the East region total. Excel inserts a new sheet containing exactly the source rows behind that number: every East transaction that makes up the ₹6,05,000. This is the fastest way in any review meeting to answer \"show me the rows behind that figure\"." },
          { type: "warning", text: "The drill-down sheet is a copy, not a live view. It does not update when the source changes, and deleting rows on it does nothing to the pivot or the source. Treat it as a temporary extract and delete the sheet when you are done." },
          { type: "p", text: "Two housekeeping details keep pivots tidy. First, if the pivot shows a \"(blank)\" row, some source rows have an empty cell in that field; clean the source, refresh, and the blank row disappears. Second, after filtering a pivot for one view, remember to clear the filter before building the next view, or you will analyse a subset while believing you see everything. Filtering is invisible in the numbers themselves; only the filter dropdowns and icons reveal that a subset is on screen." }
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
    intro: "So far your pivots answer questions, but the person reading them still has to open filter dropdowns and click tiny menus. Slicers turn filtering into big friendly buttons that anyone can use in a meeting, and timelines do the same for dates. PivotCharts add the visual layer on top. Together they turn a static pivot into an interactive report.\n\nThis lesson uses `sales-data.xlsx` again. Everything builds on the pivoting and grouping skills from the last two lessons. Note that timelines require Excel 2013 or later; slicers work from Excel 2010.",
    sections: [
      {
        heading: "From pivot to PivotChart",
        blocks: [
          { type: "p", text: "A PivotChart is a chart bound to a PivotTable. Change the pivot, and the chart changes with it; filter the pivot, and the chart filters too. You never build the chart series by hand." },
          { type: "steps", items: [
            "Build or open a pivot, for example Sum of Revenue with Date grouped by month in Rows and Region in Columns.",
            "Click inside the pivot.",
            "On the PivotTable Analyze tab, click PivotChart.",
            "Choose Clustered Column and click OK.",
            "The chart appears beside the pivot, already titled \"Sum of Revenue\". Edit the title to say something meaningful, such as \"Monthly revenue by region, 2025\"."
          ] },
          { type: "p", text: "The pivot and the chart are two views of the same fields. Drag a new field into Columns on the pivot and the chart re-plot instantly. Filter the pivot and the chart shrinks to match. This is the key difference from the ordinary charts of the first lesson: a PivotChart cannot drift out of sync with its numbers." },
          { type: "p", text: "A few PivotChart quirks are worth knowing before a meeting. The chart shows small field buttons for filtering directly on the plot area; useful for exploring, but cluttered for presentation, so right-click a button and choose Hide All Field Buttons on Chart before showing it. Collapsed pivot groups collapse in the chart too, so a date hierarchy collapsed to quarters shows a shorter, cleaner chart. And if you later move or resize the pivot, the chart follows its data, not its position: reposition both deliberately when you lay out a report page." },
          { type: "p", text: "When should you use an ordinary chart instead? When the summary behind the chart is fixed, small and formula-driven, such as a `SUMIFS` region summary, an ordinary chart is simpler and has no pivot buttons to hide. PivotCharts earn their keep when the reader is meant to filter and explore; ordinary charts earn theirs when the view is final." },
          { type: "p", text: "One warning: not every chart type is available as a PivotChart. Column, bar, line, pie and area work fine, but scatter and bubble charts are not offered, because a pivot's category structure does not fit them. For those rare cases, chart a `SUMIFS` summary instead." }
        ]
      },
      {
        heading: "Slicers: filtering with buttons",
        blocks: [
          { type: "p", text: "A slicer is a floating panel of filter buttons for one field. Instead of opening the Row Labels dropdown and hunting through a menu, you click a button. In a review meeting this is the difference between fumbling with Excel and steering the discussion." },
          { type: "p", text: "Slicers also remove the biggest usability gap of filter dropdowns: visibility of state. A dropdown filter hides inside a small icon, and a reader can easily miss that a pivot is filtered at all. A slicer shows its selection as pressed buttons in plain sight, which keeps everyone in the room honestly informed about what is on screen. That transparency, more than the speed, is why slicers became the standard control in shared reports." },
          { type: "steps", items: [
            "Click inside the pivot.",
            "On the PivotTable Analyze tab (or Insert tab), click Slicer.",
            "Tick Region and Product, then click OK.",
            "Two slicer panels appear. Click North in the Region slicer: the pivot shows only the North, and the Product slicer grays out products the North never sold.",
            "Hold [[Ctrl]] and click to select several buttons at once, for example North and East.",
            "Click the Clear Filter icon at the top-right corner of a slicer to reset it."
          ] },
          { type: "p", text: "Slicers show their state clearly: selected buttons are highlighted, and unavailable combinations are grayed out. That visible state is worth a lot when someone else takes over your file." },
          { type: "tip", text: "Format a slicer to fit your report: select it and use the Slicer tab to pick a style and set Columns to 2 or 3 so tall lists become compact grids. Resize and position slicers beside the pivot, and turn off the sheet gridlines so they sit on a clean background." },
          { type: "p", text: "Give each slicer a clear caption when the field name is not self-explanatory: with the slicer selected, the Slicer tab has a Caption box, and renaming \"Salesperson\" to \"Filter by person\" takes two seconds. Multi-select behaviour is worth practising until it is reflex: plain clicks replace the selection, [[Ctrl]] + click adds to it, and clicking the selected button again deselects it. After any exploration session, the clear-filter icon at the slicer's top-right corner is the fastest reset." }
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
          { type: "p", text: "The timeline only appears for pivots that contain a real date field. If Insert Timeline is grayed out, your dates are text or the Date field is not in the pivot yet; fix that first. Timelines need Excel 2013 or later." },
          { type: "p", text: "Timelines and slicers work together naturally. In a sales review, a typical flow is: set the period on the timeline first, then click through regions on the slicer while the discussion moves. Because both controls stay visible on the sheet, the person driving Excel can respond to any question in the room, such as \"what was the South doing in the second quarter?\", in two clicks and without leaving the report page. That responsiveness is what makes a pivot-based report feel interactive rather than static." },
          { type: "warning", text: "Before you send any interactive file to a manager, click the Clear Filter button on every slicer and timeline. A file that opens showing only \"West, Staplers, March\" looks broken to the next reader, even though it is just a leftover filter." },
          { type: "p", text: "One detail makes timelines feel professional: match the timeline's default level to the reporting rhythm. If your meetings discuss months, set the level dropdown to Months before saving; a quarterly default in a monthly meeting forces everyone to fumble with the dropdown at the start of every session. The level choice is saved with the file, so setting it once per report is enough." }
        ]
      },
      {
        heading: "One slicer, many pivots: Report Connections",
        blocks: [
          { type: "p", text: "A serious report page usually has more than one pivot: revenue by month, revenue by product, and a KPI table. You do not want three separate slicers for Region. Connect one slicer to all of them and a single click filters the entire page." },
          { type: "steps", items: [
            "Place two or three pivots on the same sheet, built from the same source table `SalesTable`.",
            "Insert one slicer for Region.",
            "Right-click the slicer and choose Report Connections (in older versions: PivotTable Connections).",
            "Tick every pivot that should obey the slicer and click OK.",
            "Click a region button: all connected pivots filter together."
          ] },
          { type: "p", text: "Connections only work between pivots that share the same source data, which is another reason to keep one clean Table as your single source of truth. A slicer connected to pivots from different tables cannot sync them." },
          { type: "p", text: "Prepare an interactive report for handover the way you would prepare a projector. Refresh every pivot so the numbers are current, clear all slicers and timelines, and check that each pivot's grand total matches the expected full-year figure. Save, then reopen the file once to see exactly what the next person will see: an unfiltered, up-to-date report page with buttons ready to use. This two-minute ritual is the difference between a file that impresses and a file that raises questions before the meeting starts." },
          { type: "grid", title: "A simple interactive report layout on one sheet",
            rows: [
              ["SUNRISE TRADERS — SALES REPORT", "", "", "", ""],
              ["[Slicer: Region]", "[Slicer: Product]", "[Timeline: Date]", "", ""],
              ["Pivot 1: Revenue by month", "", "Pivot 2: Revenue by product", "", ""],
              ["PivotChart: monthly trend", "", "Pivot 3: KPI summary", "", ""]
            ] },
          { type: "p", text: "That one-page layout is a real deliverable: title on top, slicers and timeline in one control strip, and every table and chart obeying them. You will refine this pattern into a full dashboard in the dashboard design lesson later in this level." },
          { type: "p", text: "Build the page in this order and it falls into place: pivots first, so you know their size; PivotCharts second, aligned beside them; slicers and timeline last, sized to fill the leftover strip. Trying to place slicers first always ends with moving everything twice. Leave one empty row between blocks as breathing room, and keep the pivot sheets themselves available in the workbook, because a reader who wants the detail will drill down from the pivot anyway." }
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
