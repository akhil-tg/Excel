/* ============================================================
   Excel Lab — lesson renderer (lesson.html?id=<lesson-id>)
   ============================================================ */
"use strict";

/* ---------- block renderers ---------- */
function rGrid(b) {
  var rows = b.rows || [];
  if (!rows.length) return "";
  var nCols = Math.max.apply(null, rows.map(function (r) { return r.length; }));
  var letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("").slice(0, nCols);
  var h = '<div class="lesson-grid-table"><div class="xlsheet">';
  if (b.title) h += '<div class="xlsheet-title">' + icon("xls") + escHtml(b.title) + "</div>";
  h += "<table><thead><tr><th class='rowhead' aria-hidden='true'></th>";
  letters.forEach(function (c) { h += "<th>" + c + "</th>"; });
  h += "</tr></thead><tbody>";
  rows.forEach(function (r, ri) {
    h += "<tr><th class='rowhead'>" + (ri + 1) + "</th>";
    for (var c = 0; c < nCols; c++) {
      var cell = r[c] !== undefined && r[c] !== null && r[c] !== "" ? String(r[c]) : "";
      var cls = "cell";
      if (cell.charAt(0) === "=") cls += " formula";
      else if (/^-?[\d,]+(\.\d+)?%?$/.test(cell.replace(/\s/g, ""))) cls += " num";
      h += "<td class='" + cls + "'>" + escHtml(cell) + "</td>";
    }
    h += "</tr>";
  });
  h += "</tbody></table></div></div>";
  return h;
}

function rFormula(b) {
  var h = '<div class="formula-box"><div class="fx-head"><span class="fx-badge">' + escHtml(b.name) + '</span><span class="tag">Formula</span></div>';
  h += '<div class="fx-row"><span class="fx-lbl">Syntax</span><div class="fx-code">' + escHtml(b.syntax) + "</div></div>";
  if (b.example) h += '<div class="fx-row"><span class="fx-lbl">Example</span><div class="fx-code">' + escHtml(b.example) + "</div></div>";
  if (b.explanation) h += '<div class="fx-row"><span class="fx-lbl">What it does</span><p class="fx-desc">' + md(b.explanation) + "</p></div>";
  return h + "</div>";
}

function rTable(b) {
  var h = '<div style="overflow-x:auto"><table class="data-table"><thead><tr>';
  (b.headers || []).forEach(function (hd) { h += "<th>" + md(hd) + "</th>"; });
  h += "</tr></thead><tbody>";
  (b.rows || []).forEach(function (r) {
    h += "<tr>" + r.map(function (c) { return "<td>" + md(String(c)) + "</td>"; }).join("") + "</tr>";
  });
  h += "</tbody></table>";
  if (b.caption) h += "<caption>" + md(b.caption) + "</caption>";
  return h + "</div>";
}

function rCallout(kind, text) {
  var ic = kind === "tip" ? "bulb" : kind === "warning" ? "alert" : "info";
  return '<div class="callout callout-' + kind + '">' + icon(ic) + "<div>" + mdp(text) + "</div></div>";
}

function rBlock(b) {
  switch (b.type) {
    case "p": return "<p>" + md(b.text) + "</p>";
    case "ul": return "<ul>" + b.items.map(function (i) { return "<li>" + md(i) + "</li>"; }).join("") + "</ul>";
    case "ol": return "<ol>" + b.items.map(function (i) { return "<li>" + md(i) + "</li>"; }).join("") + "</ol>";
    case "steps": return '<ol class="steps">' + b.items.map(function (i) { return "<li><div>" + md(i) + "</div></li>"; }).join("") + "</ol>";
    case "table": return rTable(b);
    case "grid": return rGrid(b);
    case "formula": return rFormula(b);
    case "code": return "<pre><code>" + escHtml(b.text) + "</code></pre>";
    case "tip": return rCallout("tip", b.text);
    case "warning": return rCallout("warning", b.text);
    case "quote": return rCallout("info", b.text);
    default: return "";
  }
}

/* ---------- quiz ---------- */
var quizState = null;

function renderQuiz(lesson) {
  if (!lesson.quiz || !lesson.quiz.length) return "";
  quizState = { id: lesson.id, items: lesson.quiz, answered: 0, correct: 0, done: false };
  var h = '<section class="quiz card" id="quiz" aria-label="Quiz">'
    + '<div class="quiz-head"><h2>' + icon("target") + "Knowledge check</h2>"
    + '<span class="quiz-score" id="quizScore">0 / ' + lesson.quiz.length + " correct</span></div>";
  lesson.quiz.forEach(function (q, qi) {
    h += '<div class="quiz-q" data-q="' + qi + '">'
      + '<div class="q-num">Question ' + (qi + 1) + " of " + lesson.quiz.length + "</div>"
      + '<div class="q-text">' + md(q.q) + "</div>"
      + '<div class="quiz-opts">';
    q.options.forEach(function (opt, oi) {
      h += '<button type="button" class="quiz-opt" data-qi="' + qi + '" data-oi="' + oi + '">'
        + '<span class="opt-letter">' + String.fromCharCode(65 + oi) + "</span><span>" + md(opt) + "</span></button>";
    });
    h += "</div></div>";
  });
  h += '<div class="quiz-foot" style="margin-top:18px;display:flex;gap:10px;align-items:center;flex-wrap:wrap">'
    + '<button type="button" class="btn btn-outline btn-sm" id="quizReset">' + icon("refresh") + "Reset quiz</button>"
    + '<span id="quizResult" style="font-weight:700;font-size:.9rem;color:var(--text-3)"></span></div>';
  h += "</section>";

  setTimeout(function () {
    document.querySelectorAll(".quiz-opt").forEach(function (btn) {
      btn.addEventListener("click", function () { answerQuiz(btn); });
    });
    var rst = document.getElementById("quizReset");
    if (rst) rst.addEventListener("click", function () { renderQuizInto(lesson); });
  }, 0);
  return h;
}

function answerQuiz(btn) {
  var qi = +btn.dataset.qi, oi = +btn.dataset.oi;
  if (quizState.done || btn.classList.contains("answered")) return;
  var q = quizState.items[qi];
  var wrap = document.querySelector('.quiz-q[data-q="' + qi + '"]');
  var opts = wrap.querySelectorAll(".quiz-opt");
  opts.forEach(function (b) { b.disabled = true; b.classList.add("answered"); });
  opts[q.answer].classList.add("correct");
  var ok = oi === q.answer;
  if (!ok) btn.classList.add("wrong");
  if (ok) quizState.correct++;
  quizState.answered++;
  var ex = wrap.querySelector(".quiz-explain");
  if (!ex) {
    ex = document.createElement("div");
    ex.className = "quiz-explain";
    wrap.querySelector(".quiz-opts").after(ex);
  }
  ex.innerHTML = (ok ? "<strong>Correct.</strong> " : "<strong>Not quite.</strong> The correct answer is <strong>" + String.fromCharCode(65 + q.answer) + "</strong>. ") + md(q.explanation);
  var sc = document.getElementById("quizScore");
  if (sc) sc.textContent = quizState.correct + " / " + quizState.items.length + " correct";
  if (quizState.answered === quizState.items.length) {
    quizState.done = true;
    quizSave(quizState.id, quizState.correct, quizState.items.length);
    var res = document.getElementById("quizResult");
    var pct = Math.round(quizState.correct / quizState.items.length * 100);
    if (res) {
      res.textContent = "You scored " + quizState.correct + "/" + quizState.items.length + " (" + pct + "%) — " + (pct >= 80 ? "excellent, move on!" : pct >= 50 ? "good — review the parts you missed." : "worth re-reading this lesson.");
      res.style.color = pct >= 80 ? "var(--green-600)" : "var(--amber-600)";
    }
    document.dispatchEvent(new CustomEvent("el:progress"));
  }
}

function renderQuizInto(lesson) {
  var old = document.querySelector(".quiz");
  if (!old) return;
  old.outerHTML = renderQuiz(lesson);
}

/* ---------- main render ---------- */
function renderLessonPage() {
  var id = qParam("id");
  var lesson = findLesson(id);
  var root = document.getElementById("lessonRoot");
  if (!lesson) {
    root.innerHTML = '<div class="empty-state card" style="padding:60px 30px"><h2>Lesson not found</h2><p>The lesson you are looking for does not exist.</p><a class="btn btn-primary" href="lessons.html">' + icon("book") + "Browse all lessons</a></div>";
    document.title = "Lesson not found — Excel Lab";
    return;
  }
  var lv = levelOf(lesson.id);
  var level = (window.LEVEL_META || {})[lv] || { title: "Course" };
  var nb = lessonNeighbors(lesson.id);
  document.title = lesson.title + " — Excel Lab";

  /* sidebar */
  var levelLessons = (window.LEVEL_LESSONS[lv] || []);
  var doneCount = levelLessons.filter(function (l) { return isDone(l.id); }).length;
  var sb = '<aside class="lesson-sidebar" aria-label="Lesson navigation"><div class="card sidebar-card">';
  sb += '<h3>' + icon(level.icon || "book") + " " + escHtml(level.title) + "</h3>";
  sb += '<ul class="sidebar-list">';
  levelLessons.forEach(function (l) {
    var cls = [];
    if (isDone(l.id)) cls.push("done");
    if (l.id === lesson.id) cls.push("current");
    sb += '<li class="' + cls.join(" ") + '"><a href="lesson.html?id=' + l.id + '">'
      + '<span class="dot" aria-hidden="true">' + (isDone(l.id) ? icon("check") : "") + "</span><span>" + escHtml(l.title) + "</span></a></li>";
  });
  sb += "</ul>";
  sb += '<div class="sidebar-progress"><span>' + doneCount + " of " + levelLessons.length + " complete</span>"
    + '<span class="mini-track" style="flex:1"><i style="width:' + Math.round(doneCount / levelLessons.length * 100) + '%"></i></span></div>';
  sb += "</div></aside>";

  /* head */
  var badges = '<span class="badge badge-level">Level ' + lv + " · " + escHtml(level.title) + "</span>" + diffBadge(lesson.difficulty);
  if (lesson.version) badges += '<span class="badge badge-version">' + escHtml(lesson.version) + " required</span>";
  if (isDone(lesson.id)) badges += '<span class="badge badge-done">' + icon("check") + "Completed</span>";
  var meta = '<span>' + icon("clock") + "≈ " + durationMin(lesson) + " min</span>"
    + '<span>' + icon("book") + "Lesson " + (nb.index + 1) + " of " + nb.total + "</span>";

  var h = sb + '<main class="lesson-main" id="main">';
  h += '<nav class="lesson-crumb" aria-label="Breadcrumb"><a href="index.html">Home</a>' + icon("chevron-right")
    + '<a href="lessons.html#level-' + lv + '">Level ' + lv + ": " + escHtml(level.title) + "</a>" + icon("chevron-right")
    + "<span>" + escHtml(lesson.title) + "</span></nav>";

  h += '<header class="lesson-head card"><div class="badges">' + badges + "</div><h1>" + escHtml(lesson.title) + "</h1>"
    + '<div class="meta-row">' + meta + "</div></header>";

  if (lesson.version) {
    h += '<div class="callout callout-info version-note">' + icon("info") + "<div><p><strong>Excel version note.</strong> This lesson covers <strong>"
      + escHtml(lesson.version) + "</strong> features. If you use an older version, follow along with the alternatives the lesson suggests.</p></div></div>";
  }

  /* objectives */
  if (lesson.objectives && lesson.objectives.length) {
    h += '<div class="objectives card"><h2>' + icon("target") + "In this lesson you will learn to</h2><ul>"
      + lesson.objectives.map(function (o) { return "<li>" + md(o) + "</li>"; }).join("") + "</ul></div>";
  }

  /* toc */
  if (lesson.sections && lesson.sections.length > 2) {
    h += '<nav class="toc card" aria-label="On this page"><h2>On this page</h2><ol>';
    lesson.sections.forEach(function (s, i) {
      h += '<li><a href="#sec-' + i + '">' + escHtml(s.heading) + "</a></li>";
    });
    h += "</ol></nav>";
  }

  /* intro */
  h += '<div class="prose">';
  if (lesson.intro) h += mdp(lesson.intro);

  /* sections */
  (lesson.sections || []).forEach(function (s, i) {
    h += '<section class="block-section" id="sec-' + i + '"><h2 class="sec">' + escHtml(s.heading) + "</h2>"
      + (s.blocks || []).map(rBlock).join("") + "</section>";
  });

  /* mistakes */
  if (lesson.mistakes && lesson.mistakes.length) {
    h += '<section class="block-section" id="mistakes"><h2 class="sec">Common mistakes to avoid</h2><div class="mistake-list">';
    lesson.mistakes.forEach(function (m) {
      h += '<div class="mistake"><div class="m-row m-wrong"><span class="m-lbl">Common mistake</span><div>' + md(m.wrong) + "</div></div>"
        + '<div class="m-row m-right"><span class="m-lbl">Do this instead</span><div>' + md(m.right) + "</div></div>"
        + '<p class="m-why"><strong>Why it matters:</strong> ' + md(m.why) + "</p></div>";
    });
    h += "</div></section>";
  }

  /* shortcuts */
  if (lesson.shortcuts && lesson.shortcuts.length) {
    h += '<section class="block-section" id="shortcuts"><h2 class="sec">Keyboard shortcuts in this lesson</h2>'
      + rTable({ headers: ["Shortcut", "What it does"], rows: lesson.shortcuts.map(function (s) { return ["[[" + s.keys + "]]", s.action]; }) })
      + "</section>";
  }

  /* exercises */
  if (lesson.exercises && lesson.exercises.length) {
    h += '<section class="block-section" id="exercises"><h2 class="sec">Practice exercises</h2><div class="exercise-list">';
    lesson.exercises.forEach(function (e) {
      h += '<div class="exercise"><div class="ex-head"></div><div class="ex-body"><p>' + md(e.task) + "</p>";
      if (e.hint) h += "<p><strong>Hint:</strong> " + md(e.hint) + "</p>";
      if (e.solution) h += "<details><summary>Show solution</summary><div class='solution'>" + mdp(e.solution) + "</div></details>";
      h += "</div></div>";
    });
    h += "</div></section>";
  }

  /* quiz */
  h += renderQuiz(lesson);

  /* summary */
  if (lesson.summary && lesson.summary.length) {
    h += '<section class="block-section" id="summary"><h2 class="sec">Lesson summary</h2><div class="callout callout-tip">'
      + icon("check") + "<div><ul style='margin:0;padding-left:18px'>"
      + lesson.summary.map(function (s) { return "<li>" + md(s) + "</li>"; }).join("") + "</ul></div></div></section>";
  }

  /* downloads */
  if (lesson.downloads && lesson.downloads.length) {
    h += '<section class="block-section" id="downloads"><h2 class="sec">Practice files</h2><p>Download the real Excel workbook for this lesson, open it, and practise while you read.</p><div class="download-list">';
    lesson.downloads.forEach(function (d) {
      h += '<a class="download-item" href="' + escHtml(d.file) + '" download>'
        + '<span class="dl-icon">' + icon("download") + "</span>"
        + '<span class="dl-body"><span class="dl-label">' + escHtml(d.label) + "</span>"
        + (d.note ? '<span class="dl-note">' + escHtml(d.note) + "</span>" : "") + "</span>"
        + '<span class="dl-cta">Download ' + escHtml(d.file.split(".").pop().toUpperCase()) + "</span></a>";
    });
    h += "</div></section>";
  }

  /* complete + prev/next */
  h += "</div>"; // /prose
  h += '<div class="lesson-footer">'
    + '<div class="complete-row card"><div><strong style="display:block">Finished this lesson?</strong>'
    + '<span class="msg" id="completeMsg"></span></div>'
    + '<button type="button" class="btn ' + (isDone(lesson.id) ? "btn-outline" : "btn-primary") + '" id="completeBtn"></button></div>';
  h += '<div class="prev-next">';
  h += nb.prev ? pnCard(nb.prev, "prev") : "<span></span>";
  h += nb.next ? pnCard(nb.next, "next") : "<span></span>";
  h += "</div></div></main>";

  root.innerHTML = h;
  wireLessonControls(lesson);
}

function pnCard(l, dir) {
  return '<a class="pn-card card ' + dir + '" href="lesson.html?id=' + l.id + '">'
    + icon(dir === "prev" ? "arrow-left" : "arrow-right")
    + '<span class="pn-body"><span class="pn-lbl">' + (dir === "prev" ? "Previous lesson" : "Next lesson") + "</span>"
    + '<span class="pn-title">' + escHtml(l.title) + "</span></span></a>";
}

function wireLessonControls(lesson) {
  var btn = document.getElementById("completeBtn");
  var msg = document.getElementById("completeMsg");
  function refresh() {
    var done = isDone(lesson.id);
    btn.className = "btn " + (done ? "btn-outline" : "btn-primary");
    btn.innerHTML = done ? icon("refresh") + " Mark as not completed" : icon("check") + " Mark lesson as complete";
    msg.textContent = done ? "Completed. Great work — keep the streak going!" : "Marking a lesson complete updates your progress on the dashboard.";
  }
  btn.addEventListener("click", function () {
    setDone(lesson.id, !isDone(lesson.id));
  });
  refresh();

  document.addEventListener("el:progress", function () {
    refresh();
    var l = findLesson(lesson.id);
    var lv = levelOf(lesson.id);
    var levelLessons = (window.LEVEL_LESSONS[lv] || []);
    var doneCount = levelLessons.filter(function (x) { return isDone(x.id); }).length;
    var sp = document.querySelector(".sidebar-progress");
    if (sp) sp.innerHTML = "<span>" + doneCount + " of " + levelLessons.length + " complete</span>"
      + '<span class="mini-track" style="flex:1"><i style="width:' + Math.round(doneCount / levelLessons.length * 100) + '%"></i></span>';
    levelLessons.forEach(function (x) {
      var li = document.querySelector('.sidebar-list a[href="lesson.html?id=' + x.id + '"]');
      if (li) {
        li.parentElement.classList.toggle("done", isDone(x.id));
        var dot = li.querySelector(".dot");
        if (dot) dot.innerHTML = isDone(x.id) ? icon("check") : "";
      }
    });
    var hd = document.querySelector(".lesson-head .badges");
    if (hd && isDone(lesson.id) && !hd.querySelector(".badge-done")) {
      hd.insertAdjacentHTML("beforeend", '<span class="badge badge-done">' + icon("check") + "Completed</span>");
    }
  }, { once: false });
}

document.addEventListener("DOMContentLoaded", renderLessonPage);
