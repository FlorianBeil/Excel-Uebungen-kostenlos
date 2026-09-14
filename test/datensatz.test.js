// Aufruf: node test/datensatz.test.js
// Vergleicht die aus dem Bauplan erzeugten Pivot-Daten mit buildSalesDataset()
// aus dem Pivot-Repo (liegt lokal daneben: ../pivot-tabelle-prototyp/index.html).
"use strict";
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const { erzeugeDatensatz } = require("../assets/pivot/datensatz.js");

const daten = JSON.parse(fs.readFileSync(path.join(__dirname, "../daten/aufgaben.json"), "utf8"));
const pivotAufgabe = daten.aufgaben.find((a) => a.typ === "pivot");

const html = fs.readFileSync(path.join(__dirname, "../../pivot-tabelle-prototyp/index.html"), "utf8").replace(/\r\n/g, "\n");
function funktionText(name) {
  const start = html.indexOf("  function " + name + "(");
  assert.ok(start !== -1, name + " nicht im Pivot-Repo gefunden");
  const ende = html.indexOf("\n  }\n", start);
  return html.slice(start, ende + 4);
}
// eslint-disable-next-line no-new-func
const original = new Function(funktionText("mulberry32") + funktionText("buildSalesDataset") + "return buildSalesDataset();")();
const neu = erzeugeDatensatz(pivotAufgabe.datensatz);

assert.deepStrictEqual(neu.columns, original.columns, "Spalten");
assert.deepStrictEqual(neu.fields, original.fields, "Felder/Icons");
assert.strictEqual(neu.rows.length, original.rows.length, "Zeilenanzahl");
neu.rows.forEach((r, i) => {
  const o = original.rows[i];
  assert.strictEqual(r.Datum.getTime(), o.Datum.getTime(), "Datum Zeile " + i);
  assert.deepStrictEqual(Object.assign({}, r, { Datum: 0 }), Object.assign({}, o, { Datum: 0 }), "Zeile " + i);
});

const summen = {};
neu.rows.forEach((r) => { summen[r.Region] = (summen[r.Region] || 0) + r.Umsatz; });
console.log("ok  " + neu.rows.length + " Zeilen identisch mit der Kurs-Übung");
console.log("    Umsatz je Region:", summen);
