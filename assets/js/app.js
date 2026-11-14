/* ============================================================
   Excel Lab — shared app utilities
   ============================================================ */
"use strict";

/* ---------- SVG icon sprite (lucide-style, stroke) ---------- */
(function injectSprite() {
  var S = '<svg xmlns="http://www.w3.org/2000/svg" style="display:none" aria-hidden="true">'
    + '<symbol id="i-grid" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></symbol>'
    + '<symbol id="i-sigma" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 6H7l5 6-5 6h11"/></symbol>'
    + '<symbol id="i-database" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M3 5v14a9 3 0 0 0 18 0V5"/><path d="M3 12a9 3 0 0 0 18 0"/></symbol>'
    + '<symbol id="i-chart" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v18h18"/><path d="M18 17V9"/><path d="M13 17V5"/><path d="M8 17v-3"/></symbol>'
    + '<symbol id="i-briefcase" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="7" width="20" height="14" rx="2"/><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/></symbol>'
    + '<symbol id="i-zap" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M13 2 3 14h9l-1 8 10-12h-9l1-8z"/></symbol>'
    + '<symbol id="i-folder" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 6a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2z"/></symbol>'
    + '<symbol id="i-sun" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M6.34 17.66l-1.41 1.41M19.07 4.93l-1.41 1.41"/></symbol>'
    + '<symbol id="i-moon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z"/></symbol>'
    + '<symbol id="i-menu" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M4 6h16M4 12h16M4 18h16"/></symbol>'
    + '<symbol id="i-search" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/></symbol>'
    + '<symbol id="i-check" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></symbol>'
    + '<symbol id="i-chevron-left" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m15 18-6-6 6-6"/></symbol>'
    + '<symbol id="i-chevron-right" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m9 18 6-6-6-6"/></symbol>'
    + '<symbol id="i-chevron-down" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></symbol>'
    + '<symbol id="i-clock" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/></symbol>'
    + '<symbol id="i-download" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m7 10 5 5 5-5"/><path d="M12 15V3"/></symbol>'
    + '<symbol id="i-info" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="10"/><path d="M12 16v-4"/><path d="M12 8h.01"/></symbol>'
    + '<symbol id="i-alert" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m21.73 18-8-14a2 2 0 0 0-3.46 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z"/><path d="M12 9v4"/><path d="M12 17h.01"/></symbol>'
    + '<symbol id="i-bulb" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 14c.2-1 .7-1.7 1.5-2.5 1-.9 1.5-2.2 1.5-3.5A6 6 0 0 0 6 8c0 1.3.5 2.6 1.5 3.5.8.8 1.3 1.5 1.5 2.5"/><path d="M9 18h6"/><path d="M10 22h4"/></symbol>'
    + '<symbol id="i-target" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/></symbol>'
    + '<symbol id="i-book" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/></symbol>'
    + '<symbol id="i-play" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="6 3 20 12 6 21 6 3"/></symbol>'
    + '<symbol id="i-arrow-right" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14"/><path d="m12 5 7 7-7 7"/></symbol>'
    + '<symbol id="i-arrow-left" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 12H5"/><path d="m12 19-7-7 7-7"/></symbol>'
    + '<symbol id="i-star" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m12 2 2.9 6.3 6.6.9-4.8 4.7 1.2 6.6L12 17.4 6.1 20.5l1.2-6.6L2.5 9.2l6.6-.9z"/></symbol>'
    + '<symbol id="i-flag" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 15s1-1 4-1 5 2 8 2 4-1 4-1V3s-1 1-4 1-5-2-8-2-4 1-4 1z"/><path d="M4 22v-7"/></symbol>'
    + '<symbol id="i-grad" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 10 12 5 2 10l10 5 10-5z"/><path d="M6 12v5c3 3 9 3 12 0v-5"/></symbol>'
    + '<symbol id="i-users" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></symbol>'
    + '<symbol id="i-refresh" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 12a9 9 0 0 1 15-6.7L21 8"/><path d="M21 3v5h-5"/><path d="M21 12a9 9 0 0 1-15 6.7L3 16"/><path d="M3 21v-5h5"/></symbol>'
    + '<symbol id="i-filter" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 3H2l8 9.5V19l4 2v-8.5z"/></symbol>'
    + '<symbol id="i-help" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="10"/><path d="M9.1 9a3 3 0 0 1 5.8 1c0 2-3 3-3 3"/><path d="M12 17h.01"/></symbol>'
    + '<symbol id="i-award" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="6"/><path d="M15.5 12.9 17 22l-5-3-5 3 1.5-9.1"/></symbol>'
    + '<symbol id="i-xls" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="m9.5 12 5 5"/><path d="m14.5 12-5 5"/></symbol>'
    + '</svg>';
  document.addEventListener("DOMContentLoaded", function () {
    if (!document.getElementById("el-icons")) {
      var div = document.createElement("div");
      div.id = "el-icons";
      div.style.display = "none";
      div.innerHTML = S;
      document.body.prepend(div);
    }
  });
})();

function icon(name, cls) {
  return '<svg class="icn' + (cls ? " " + cls : "") + '" aria-hidden="true"><use href="#i-' + name + '"></use></svg>';
}

/* ---------- Inline text formatter ---------- */
function escHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
function md(s) {
  var t = escHtml(s);
  t = t.replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>");
  t = t.replace(/`([^`]+)`/g, "<code>$1</code>");
  t = t.replace(/\[\[([^\]]+)\]\]/g, "<kbd>$1</kbd>");
  return t;
}
function mdp(s) { // multi-paragraph text
  return String(s).split(/\n\n+/).map(function (p) { return "<p>" + md(p) + "</p>"; }).join("");
}

/* ---------- Data access ---------- */
function allLessons() {
  var out = [];
  Object.keys(window.LEVEL_LESSONS || {}).map(Number).sort(function (a, b) { return a - b; }).forEach(function (lv) {
    (window.LEVEL_LESSONS[lv] || []).forEach(function (l) { out.push(l); });
  });
  return out;
}
function findLesson(id) {
  return allLessons().filter(function (l) { return l.id === id; })[0] || null;
}
function lessonNeighbors(id) {
  var all = allLessons();
  var i = all.findIndex(function (l) { return l.id === id; });
  return { prev: i > 0 ? all[i - 1] : null, next: i >= 0 && i < all.length - 1 ? all[i + 1] : null, index: i, total: all.length };
}
function levelOf(lessonId) {
  return Number(String(lessonId).charAt(1)) || 0;
}
function diffBadge(diff) {
  return '<span class="badge badge-' + diff.toLowerCase() + '">' + escHtml(diff) + "</span>";
}
function durationMin(l) { return l.duration || 15; }

/* ---------- Storage ---------- */
var Store = {
  get: function (k, fb) {
    try { var v = localStorage.getItem("ellab." + k); return v === null ? fb : JSON.parse(v); }
    catch (e) { return fb; }
  },
  set: function (k, v) {
    try { localStorage.setItem("ellab." + k, JSON.stringify(v)); } catch (e) { /* private mode */ }
  }
};
function getCompleted() {
  return Store.get("completed", []);
}
function isDone(id) { return getCompleted().indexOf(id) !== -1; }
function setDone(id, done) {
  var a = getCompleted();
  var i = a.indexOf(id);
  if (done && i === -1) { a.push(id); Store.set("completed", a); }
  if (!done && i !== -1) { a.splice(i, 1); Store.set("completed", a); }
  document.dispatchEvent(new CustomEvent("el:progress"));
}
function quizSave(lessonId, score, total) {
  var q = Store.get("quiz", {});
  q[lessonId] = { score: score, total: total, at: Date.now() };
  Store.set("quiz", q);
}
function quizGet(lessonId) {
  return (Store.get("quiz", {}) || {})[lessonId] || null;
}

/* ---------- Header (theme, nav, mobile menu, search) ---------- */
function initHeader() {
  var btn = document.getElementById("themeToggle");
  if (btn) {
    btn.addEventListener("click", function () {
      var dark = document.documentElement.classList.toggle("dark");
      Store.set("theme", dark ? "dark" : "light");
    });
  }
  var mt = document.getElementById("menuToggle");
  var header = document.getElementById("siteHeader");
  if (mt && header) {
    mt.addEventListener("click", function () {
      var nav = header.querySelector(".main-nav");
      if (nav) nav.classList.toggle("open");
      mt.setAttribute("aria-expanded", nav.classList.contains("open") ? "true" : "false");
    });
  }
  var form = document.getElementById("headerSearchForm");
  if (form) {
    form.addEventListener("submit", function (e) {
      e.preventDefault();
      var q = form.querySelector("input").value.trim();
      window.location.href = "lessons.html" + (q ? "?q=" + encodeURIComponent(q) : "");
    });
  }
  // mark current nav link
  var page = (location.pathname.split("/").pop() || "index.html");
  document.querySelectorAll(".main-nav a").forEach(function (a) {
    var href = a.getAttribute("href");
    if (href === page) a.setAttribute("aria-current", "page");
  });
}

/* ---------- Query params ---------- */
function qParam(name) {
  var m = new RegExp("[?&]" + name + "=([^&]*)").exec(location.search);
  return m ? decodeURIComponent(m[1].replace(/\+/g, " ")) : "";
}

document.addEventListener("DOMContentLoaded", initHeader);
