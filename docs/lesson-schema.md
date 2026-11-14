# Excel Lab — Lesson Data Schema

Lessons are stored as JavaScript objects in `assets/js/data/level<N>-<part>.js`
(each level is split into two part files). This document is the contract.
Follow it exactly so the site renderer works.

## File structure

Each level has two files: `level<N>-1.js` and `level<N>-2.js`.
The **-1 file** defines the level metadata and the first half of the lessons.
The **-2 file** pushes only the remaining lessons (no metadata).

```js
// level1-1.js
window.LEVEL_META = window.LEVEL_META || {};
window.LEVEL_LESSONS = window.LEVEL_LESSONS || {};

window.LEVEL_META[1] = {
  id: 1,
  title: "…",             // short level name
  subtitle: "…",          // one short sentence
  description: "…",       // 1-2 sentences for the level card
  difficulty: "Beginner" | "Intermediate" | "Advanced",
  icon: "grid"            // "grid" | "fx" | "database" | "chart" | "briefcase" | "bolt" | "folder"
};

(window.LEVEL_LESSONS[1] = window.LEVEL_LESSONS[1] || []).push(
  { LESSON },
  { LESSON }
);

// level1-2.js
window.LEVEL_LESSONS = window.LEVEL_LESSONS || {};
(window.LEVEL_LESSONS[1] = window.LEVEL_LESSONS[1] || []).push(
  { LESSON },
  { LESSON }
);
```

Lessons appear on the site in the order they are pushed. Never redefine `LEVEL_META`
in a `-2` file. Do not define anything else globally.

## Lesson object

```js
{
  id: "l1-what-is-excel",            // unique, lowercase, starts with l<N>-
  title: "What Is Excel and Why Learn It?",
  difficulty: "Beginner",            // Beginner | Intermediate | Advanced
  duration: 12,                      // minutes (number)
  version: "",                       // "" | "Microsoft 365" | "Excel 2019+" | "Excel 2016+"
                                     // "" means works in every modern Excel (2010 and later)
  objectives: ["…"],                 // 3-5 short "After this lesson you can…" items (no leading verb bullet)
  intro: "…\n\n…",                   // 2 short paragraphs separated by \n\n
  sections: [ SECTION, … ],          // 3-6 sections
  mistakes: [ { wrong: "…", right: "…", why: "…" } ],   // 2-4
  shortcuts: [ { keys: "Ctrl + Z", action: "…" } ],     // 0-8, only ones actually taught
  exercises: [ { task: "…", hint: "…", solution: "…" } ], // 2-4, realistic, tied to lesson data files
  quiz: [ QUIZ ],                    // 3-5 questions
  summary: ["…"],                    // 3-5 bullets
  downloads: [ { file: "assets/files/level1/beginner-practice.xlsx",
                 label: "Beginner practice workbook",
                 note: "Exercise and Solution sheets included" } ],
  keywords: ["…"]                    // 5-10 search terms
}
```

## Section object

```js
{
  heading: "…",
  blocks: [ BLOCK, … ]
}
```

## Block types (use only these)

```js
{ type: "p", text: "…paragraph text…" }

{ type: "ul", items: ["…", "…"] }

{ type: "ol", items: ["…", "…"] }

{ type: "steps", items: ["step 1 text", "step 2 text"] }   // numbered how-to steps

{ type: "table", headers: ["Name", "Meaning"], rows: [["Ribbon", "Toolbar at the top"]] }

{ type: "grid", title: "How the sheet looks after the formula",
  rows: [ ["", "Price", "Qty"],        // each inner array = one row; first cell is column A
          ["Apple", "20", "2"],
          ["=B2*C2", "20", "2"] ] }    // optional; renders a mock Excel grid with A/B/C letters and row numbers

{ type: "formula", name: "SUMIF",
  syntax: "=SUMIF(range, criteria, [sum_range])",
  example: "=SUMIF(C2:C50, \"North\", F2:F50)",
  explanation: "Adds only the rows where …" }

{ type: "tip", text: "…" }        // green callout, "Tip"

{ type: "warning", text: "…" }    // amber callout, "Watch out"

{ type: "code", lang: "vba", text: "Sub SaveData()\n  ActiveWorkbook.Save\nEnd Sub" }
                                  // plain code block; use \n for new lines; also use for long formulas when needed

{ type: "quote", text: "…" }      // short emphasised statement, use sparingly
```

## Inline text formatting (inside any string)

- `**bold**` → bold
- `` `formula` `` or `A1` → code style. Always wrap every cell reference, range, and formula in backticks.
- `[[Ctrl]] + [[S]]` or `[[Enter]]` → keyboard key chips
- **Never** use raw HTML, markdown lists inside strings, or template literals (no backtick strings).
- Use plain double-quoted JS strings. Escape inner double quotes as `\"`.

## Quiz object

```js
{ q: "Which formula adds cells B2 to B10?",
  options: ["=SUM(B2:B10)", "=TOTAL(B2:B10)", "=ADD(B2:B10)", "=COUNT(B2:B10)"],
  answer: 0,                 // zero-based index of the single correct option
  explanation: "SUM adds numbers in a range. COUNT only counts cells that contain numbers." }
```

Rules: exactly 4 options, one correct, plausible distractors, explanations that teach.

## Downloads

Reference ONLY files from the manifest below. Path style: `assets/files/levelN/<file>.xlsx`.
Workbooks contain an **Exercise** sheet (student works here) and a **Solution** sheet unless the
lesson says otherwise. Write the `note` so students know what is inside.

## Curriculum (fixed lesson lists — do not add, remove, or rename)

See the prompt given to you for your level's exact list.

## Quality rules

1. Audience: absolute beginners growing to professional level. Simple, natural English.
   Short sentences. Explain **how**, **why**, **when and where it is useful** in real office work.
2. Every formula must be syntactically correct Excel. Test it mentally cell by cell.
   Wrap formulas in backticks inside text.
3. Examples must use realistic business data (employees, sales, invoices, inventory), not "xxx" or "abc".
4. Lessons must build on earlier lessons; do not assume anything taught later in the course.
5. Mark version-specific features honestly (`version` field + mention it in the intro).
6. Exercises must reference the actual workbook for that lesson and match its data.
7. Do not copy marketing text; write original teaching content.
8. No emojis. No references to screenshots we do not have; describe what students will see in words.
