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

// Suchmaschinen und Link-Vorschau
const k = JSON.parse(fs.readFileSync(path.join(__dirname, "../daten/aufgaben.json"), "utf8")).konfiguration;
const meta = (attr, name) => ausHtml(new RegExp('<meta ' + attr + '="' + name + '" content="([^"]*)">'), name);
assert.ok(/^https:\/\/.+\/$/.test(k.seitenUrl), "konfiguration.seitenUrl muss eine https-Adresse mit / am Ende sein");
assert.strictEqual(ausHtml(/<link rel="canonical" href="([^"]*)">/, "canonical"), k.seitenUrl, "canonical");
assert.strictEqual(meta("property", "og:url"), k.seitenUrl, "og:url");
assert.strictEqual(meta("property", "og:title"), t.ogTitel, "og:title");
assert.strictEqual(meta("name", "twitter:title"), t.ogTitel, "twitter:title");
assert.strictEqual(meta("property", "og:description"), t.ogBeschreibung, "og:description");
assert.strictEqual(meta("name", "twitter:description"), t.ogBeschreibung, "twitter:description");
assert.strictEqual(meta("property", "og:image"), k.seitenUrl + "assets/og-bild.jpg", "og:image");
assert.strictEqual(meta("name", "twitter:image"), k.seitenUrl + "assets/og-bild.jpg", "twitter:image");
assert.strictEqual(meta("property", "og:image:alt"), t.ogBildAlt, "og:image:alt");
assert.strictEqual(meta("name", "twitter:image:alt"), t.ogBildAlt, "twitter:image:alt");
assert.ok(fs.existsSync(path.join(__dirname, "../assets/og-bild.jpg")), "assets/og-bild.jpg fehlt");
// Google kürzt nach Pixelbreite (~600px); bis ~70 Zeichen passt ein Titel mit schmalen Buchstaben meist noch
assert.ok(t.seitenTitel.length <= 70, "Seitentitel zu lang für Google (" + t.seitenTitel.length + " Zeichen)");
assert.ok(t.metaBeschreibung.length <= 160, "Meta-Beschreibung zu lang (" + t.metaBeschreibung.length + " Zeichen)");
console.log("ok  Texte in index.html stimmen mit daten/aufgaben.json überein");
