// Extracts a compact per-level digest from the lesson data files:
// downloads, exercises, and grid blocks (exact sheet data taught in lessons).
"use strict";
const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const files = [];
for (let n = 1; n <= 7; n++) for (const p of [1, 2]) files.push(`level${n}-${p}.js`);

global.window = {};
for (const f of files) {
  const code = fs.readFileSync(path.join(ROOT, "assets", "js", "data", f), "utf8");
  // eslint-disable-next-line no-eval
  eval(code);
}

const digest = {};
for (const lvl of Object.keys(window.LEVEL_LESSONS).map(Number).sort((a, b) => a - b)) {
  digest[lvl] = (window.LEVEL_LESSONS[lvl] || []).map((l) => {
    const grids = [];
    for (const s of l.sections || []) {
      for (const b of s.blocks || []) {
        if (b.type === "grid") grids.push({ heading: s.heading, title: b.title, rows: b.rows });
      }
    }
    return {
      id: l.id,
      title: l.title,
      downloads: l.downloads || [],
      exercises: l.exercises || [],
      grids,
    };
  });
}

const out = path.join(__dirname, "digest.json");
fs.writeFileSync(out, JSON.stringify(digest, null, 1));
const bytes = fs.statSync(out).size;
console.log("wrote", out, (bytes / 1024).toFixed(1) + " KB");
for (const lvl of Object.keys(digest)) {
  console.log(`level ${lvl}: ${digest[lvl].length} lessons, ${digest[lvl].filter(l => l.downloads.length).length} with downloads`);
}
