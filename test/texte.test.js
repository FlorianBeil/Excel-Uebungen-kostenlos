// Aufruf: node test/texte.test.js
// Titel, Beschreibung, Überschrift und Einleitung stehen fest in index.html (sofort
// sichtbar, auch für Suchmaschinen) UND in daten/aufgaben.json. Dieser Test stellt
// sicher, dass beide übereinstimmen.
"use strict";
const assert = require("assert");
const fs = require("fs");
const path = require("path");

const html = fs.readFileSync(path.join(__dirname, "../index.html"), "utf8");
const t = JSON.parse(fs.readFileSync(path.join(__dirname, "../daten/aufgaben.json"), "utf8")).texte;

function ausHtml(muster, name) {
  const m = html.match(muster);
  assert.ok(m, name + " nicht in index.html gefunden");
  return m[1].replace(/&amp;/g, "&").replace(/&quot;/g, '"');
}

assert.strictEqual(ausHtml(/<title>([^<]*)<\/title>/, "title"), t.seitenTitel, "Seitentitel");
assert.strictEqual(ausHtml(/<meta name="description" content="([^"]*)">/, "description"), t.metaBeschreibung, "Meta-Beschreibung");
assert.strictEqual(ausHtml(/<h1 id="seiten-titel">([^<]*)<\/h1>/, "h1"), t.ueberschrift, "Überschrift");
assert.strictEqual(ausHtml(/<p class="frei-start__text">([^<]*)<\/p>/, "Einleitung"), t.einleitung, "Einleitung");
assert.strictEqual(ausHtml(/<p id="mobil-hinweis">([^<]*)<\/p>/, "Mobil-Hinweis"), t.mobilHinweis, "Mobil-Hinweis");
console.log("ok  Texte in index.html stimmen mit daten/aufgaben.json überein");
