// Löst den Pivot-Nachbau aus dem Pivot-Repo heraus und schreibt ihn als
// eigenständige Seite pivot.html (wird auf der Übungsseite per iframe eingebunden).
//
// Aufruf: node skripte/pivot-uebernehmen.js
//
// Quelle: ../pivot-tabelle-prototyp/index.html (lokales Pivot-Repo, veröffentlichter
// Stand origin/main). Übernommen werden Menüband, Bearbeitungsleiste, Arbeitsblatt,
// Feldliste, Prüfen/Zurücksetzen, Datenquelle und die komplette Pivot-Logik.
// Entfernt werden alles Kursbezogene: Übungskatalog, Stufen/Übersicht, Fortschritt,
// Aufgabentext/Tipps (zeigt die Übungsseite selbst) und das Tracking in den
// Kurs-Bereich „pivot“ – Ereignisse gehen stattdessen per postMessage an die Übungsseite.
//
// Jede Ersetzung prüft ihren Ankertext. Ändert sich das Pivot-Repo so, dass ein
// Anker fehlt, bricht das Skript mit einer Meldung ab, statt still Falsches zu erzeugen.
"use strict";
const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");

const repo = path.join(__dirname, "..");
const pivotRepo = path.join(repo, "..", "pivot-tabelle-prototyp");

execSync("git fetch --quiet origin", { cwd: pivotRepo });
const commit = execSync("git rev-parse --short origin/main", { cwd: pivotRepo }).toString().trim();
let s = execSync("git show origin/main:index.html", { cwd: pivotRepo, maxBuffer: 64 * 1024 * 1024 })
  .toString("utf8")
  .replace(/\r\n/g, "\n");

function ersetze(beschreibung, suche, ersatz) {
  const treffer = typeof suche === "string" ? s.split(suche).length - 1 : (s.match(new RegExp(suche.source, suche.flags.includes("g") ? suche.flags : suche.flags + "g")) || []).length;
  if (treffer !== 1) throw new Error("Anker für „" + beschreibung + "“ " + treffer + "× gefunden (erwartet genau 1×)");
  s = s.replace(suche, typeof ersatz === "string" ? () => ersatz : ersatz);
}

function ausschneiden(beschreibung, von, bis, ersatz) {
  const a = s.indexOf(von);
  const b = s.indexOf(bis, a + 1);
  if (a === -1 || b === -1 || s.indexOf(von, a + 1) !== -1) throw new Error("Anker für „" + beschreibung + "“ nicht eindeutig gefunden");
  s = s.slice(0, a) + (ersatz || "") + s.slice(b);
}

/* ---------------- Kopfbereich ---------------- */

ersetze("Titel", /<title>[^<]*<\/title>/, '<title>Excel.Flo – Pivot-Aufgabe</title>\n<meta name="robots" content="noindex">');

// Hintergrundbild (als data-URL eingebettet) entfernen: Die Übungsseite hat ihren
// eigenen Hintergrund, und die Datei wird dadurch deutlich kleiner.
ersetze("Hintergrundbild", /    background-image: url\("data:image\/webp;base64,[^"]+"\);\n/, "");

ersetze(
  "CSS für die Einbettung",
  "</style>",
  [
    "  /* ---------- Eingebettet in die kostenlose Übungsseite (skripte/pivot-uebernehmen.js) ---------- */",
    "  html, body { background: transparent !important; overflow: hidden; }",
    "  main { max-width: none; padding: 0 0 4px; }",
    "</style>",
  ].join("\n")
);

/* ---------------- Markup ---------------- */

ausschneiden(
  "Kopfzeile + Übersicht",
  '<header class="site-header">',
  '  <section class="card excel-window">',
  "<main>\n  <div id=\"exercise-workspace\">\n"
);

/* ---------------- Tracking → Meldungen an die Übungsseite ---------------- */

ausschneiden(
  "Tracking",
  "  // ---------- Anonyme Nutzungs-Ereignisse ----------",
  "  // ---------- Datengenerierung",
  [
    "  // ---------- Meldungen an die Übungsseite (kostenlose Übungen) ----------",
    "  // Kein eigenes Tracking: Die einbettende Seite entscheidet, was sie davon erfasst.",
    "  var trackStats = { openedAt: Date.now(), attempts: 0 };",
    "  function melde(typ, daten) {",
    "    if (window.parent === window) return;",
    "    try {",
    '      window.parent.postMessage({ quelle: "excelflo-pivot", typ: typ, daten: daten || {} }, location.origin);',
    "    } catch (e) {}",
    "  }",
    "  function trackEvent(event, detail) {",
    '    melde("ereignis", { event: event, detail: detail || {} });',
    "  }",
    "",
    "",
  ].join("\n")
);

/* ---------------- Beispieldaten → Bauplan aus daten/aufgaben.json ---------------- */

ausschneiden(
  "Beispieltabellen",
  "  // ---------- Beispieltabellen je Übung ----------",
  "  var AGGS_NUMERIC",
  [
    "  // ---------- Beispieltabelle der Aufgabe (Bauplan aus daten/aufgaben.json) ----------",
    "  var DATASETS = {",
    "    aufgabe: function () { return window.ExcelFloPivotDatensatz.erzeugeDatensatz(AUFGABE.datensatz); }",
    "  };",
    "",
    "",
  ].join("\n")
);

/* ---------------- Übungskatalog, Fortschritt, Stufen → eine Aufgabe ---------------- */

ausschneiden(
  "Übungskatalog bis Zustand",
  "  // Jede Übung hat eine \"solution\"",
  "  // ---------- Zustand ----------",
  [
    "  // ---------- Die eine Aufgabe dieser Seite ----------",
    "  // Aufgabentext, Tipps und Lösung zeigt die Übungsseite selbst; hier nur, was",
    "  // die Pivot-Logik braucht (Prüfregel in solution, siehe checkSolution()).",
    "  var currentExercise = {",
    "    id: AUFGABE.id,",
    '    dataset: "aufgabe",',
    '    category: "",',
    "    title: AUFGABE.title,",
    '    intro: "",',
    "    hints: [],",
    "    solution: AUFGABE.loesungPruefung",
    "  };",
    "",
    "",
  ].join("\n")
);

ersetze(
  "Aufgabentext rendern",
  /  function renderExerciseText\(\) \{\n[\s\S]*?\n  \}\n/,
  "  function renderExerciseText() {\n    // Aufgabentext steht auf der Übungsseite.\n  }\n"
);

ersetze(
  "Übersicht öffnen/schließen",
  /    document\.getElementById\("overview-view"\)\.hidden = true;\n    document\.getElementById\("exercise-workspace"\)\.hidden = false;\n    document\.getElementById\("back-to-overview"\)\.hidden = false;\n/,
  ""
);

ausschneiden("Übersicht, Stufen, Übungsauswahl", "  function showOverview() {", "  // ---------- Rendering: Feldliste");

/* ---------------- Ereignisse ---------------- */

ersetze(
  "Prüfen: Ergebnis melden",
  "    var result = checkSolution();\n",
  '    var result = checkSolution();\n    melde("pruefung", { richtig: !!result.correct });\n'
);

ersetze(
  "Prüfen: Kurs-Fortschritt entfernen",
  "      markCompleted(currentExercise.id);\n      renderLevelTabs();\n      renderResetButton();\n      renderExercisePicker();\n",
  ""
);

ersetze(
  "Tipps-Schalter entfernen",
  /  document\.getElementById\("hints-toggle"\)\.addEventListener\("click", function \(\) \{\n[\s\S]*?\n  \}\);\n\n/,
  ""
);

ersetze(
  "Zurück-zur-Übersicht entfernen",
  /  document\.getElementById\("back-to-overview"\)\.addEventListener\("click", function \(e\) \{\n[\s\S]*?\n  \}\);\n\n/,
  ""
);

ersetze(
  "Start",
  /  loadDataset\(currentExercise\.dataset \|\| "sales"\);\n  renderLevelTabs\(\);\n[\s\S]*?  renderAll\(\);\n\}\)\(\);\n/,
  '  openExercise(currentExercise);\n  melde("bereit", {});\n};\n'
);

ersetze("IIFE → Startfunktion", "<script>\n(function () {\n", "<script>\nwindow.starteExcelFloPivot = function (AUFGABE) {\n");

ersetze(
  "Skripte laden",
  "</script>\n</body>",
  [
    "</script>",
    "<script>",
    "  // Aufgabe aus derselben Datendatei wie die Übungsseite laden, dann starten.",
    '  fetch("daten/aufgaben.json", { cache: "no-cache" })',
    "    .then(function (r) { return r.json(); })",
    "    .then(function (daten) {",
    '      var aufgabe = daten.aufgaben.filter(function (a) { return a.typ === "pivot"; })[0];',
    "      window.starteExcelFloPivot(aufgabe);",
    "    });",
    "</script>",
    "</body>",
  ].join("\n")
);

ersetze("Datensatz-Skript einbinden", "<script>\nwindow.starteExcelFloPivot", '<script src="assets/pivot/datensatz.js?v=1"></script>\n<script>\nwindow.starteExcelFloPivot');

/* ---------------- Kontrolle: nichts Kursbezogenes mehr übrig ---------------- */

// Geprüft wird Code und Markup, nicht Kommentare/CSS-Reste (ungenutzte CSS-Regeln
// der Übersicht stören nicht): Tracking-Ziel, Kurs-Fortschritt, Katalog, Element-IDs.
const verboten = [
  /supabase\.co/, /TRACK_KEY/, /PROGRESS_KEY/, /localStorage/, /var EXERCISES\b/, /\bEXERCISES\./, /(^|[^_])\bLEVELS\b/,
  /renderLevelTabs\(/, /renderExercisePicker\(/, /renderResetButton\(/, /exercisesForLevel\(/, /markCompleted\(/,
  /id="(overview-view|back-to-overview|level-tabs|exercise-picker|task-title|hints-toggle)"/,
  /getElementById\("(overview-view|back-to-overview|level-tabs|exercise-picker|task-title|hints-toggle|hints-list|task-intro)"\)/,
  /function build(Sales|Vertrieb|Produkte|Kunden)Dataset/,
];
const reste = [];
s.split("\n").forEach((zeile, i) => {
  verboten.forEach((v) => { if (v.test(zeile)) reste.push("  Zeile " + (i + 1) + " [" + v.source + "]: " + zeile.trim().slice(0, 140)); });
});
if (reste.length) throw new Error("Noch Kursbezogenes enthalten:\n" + reste.join("\n"));

const kopf = "<!-- Aus github.com/FlorianBeil/Excel-Aufgaben-Pivot-Tabellen- (Commit " + commit + ") erzeugt mit skripte/pivot-uebernehmen.js – nicht von Hand ändern. -->\n";
fs.writeFileSync(path.join(repo, "pivot.html"), s.replace("<!DOCTYPE html>\n", "<!DOCTYPE html>\n" + kopf));
console.log("pivot.html geschrieben (Pivot-Repo Commit " + commit + ", " + Math.round(s.length / 1024) + " KB)");
