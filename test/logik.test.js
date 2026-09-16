// Aufruf: node test/logik.test.js
"use strict";
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const L = require("../assets/seite/logik.js");

const daten = JSON.parse(fs.readFileSync(path.join(__dirname, "../daten/aufgaben.json"), "utf8"));
const ids = daten.aufgaben.map((a) => a.id);
let tests = 0;
function test(name, fn) {
  fn();
  tests++;
  console.log("ok  " + name);
}

function speicher(anfang) {
  const werte = Object.assign({}, anfang);
  return { getItem: (k) => (k in werte ? werte[k] : null), setItem: (k, v) => { werte[k] = v; }, werte };
}

test("Datendatei ist gültig", () => {
  assert.deepStrictEqual(L.pruefeDaten(daten), []);
  assert.strictEqual(daten.aufgaben.length, 5);
  assert.deepStrictEqual(daten.aufgaben.map((a) => a.typ), ["formel", "formel", "formel", "formel", "pivot"]);
});

test("pruefeDaten findet Fehler", () => {
  const kaputt = JSON.parse(JSON.stringify(daten));
  kaputt.aufgaben[1].id = kaputt.aufgaben[0].id;
  kaputt.konfiguration.hinweisNachAufgabe = 5;
  delete kaputt.konfiguration.datenschutzUrl;
  const p = L.pruefeDaten(kaputt);
  assert.ok(p.some((x) => x.includes("doppelt")));
  assert.ok(p.some((x) => x.includes("hinweisNachAufgabe")));
  assert.ok(p.some((x) => x.includes("datenschutzUrl")));
});

test("neuer Besucher: Aufgabe 1 von 5, kein Angebot, kein Abschluss", () => {
  const s = L.neuerStand();
  assert.strictEqual(L.aktuelleAufgabe(daten, s), 0);
  assert.strictEqual(L.fortschrittText(daten, s), "Aufgabe 1 von 5");
  assert.strictEqual(L.hinweisSichtbar(daten, s), false);
  assert.strictEqual(L.abschluss(daten, s), null);
});

test("Lösung erst nach einem Fehlversuch", () => {
  let s = L.neuerStand();
  assert.strictEqual(L.loesungErlaubt(s, ids[0]), false);
  assert.strictEqual(L.loesungAnzeigen(s, ids[0]), s, "ohne Fehlversuch keine Änderung");
  s = L.pruefungErgebnis(s, ids[0], false);
  assert.strictEqual(L.loesungErlaubt(s, ids[0]), true);
  s = L.loesungAnzeigen(s, ids[0]);
  assert.strictEqual(L.aufgabeStatus(s, ids[0]).loesungAngezeigt, true);
  assert.strictEqual(L.istBearbeitet(s, ids[0]), true);
  assert.strictEqual(L.fortschrittText(daten, s), "Aufgabe 2 von 5");
});

test("Stand wird nie verändert", () => {
  const s = L.neuerStand();
  const kopie = JSON.stringify(s);
  L.pruefungErgebnis(s, ids[0], false);
  L.pruefungErgebnis(s, ids[0], true);
  assert.strictEqual(JSON.stringify(s), kopie);
});

test("richtig nach Lösung zählt als mitLoesung, erneutes Prüfen ändert nichts", () => {
  let s = L.pruefungErgebnis(L.neuerStand(), ids[0], false);
  s = L.loesungAnzeigen(s, ids[0]);
  s = L.pruefungErgebnis(s, ids[0], true);
  assert.deepStrictEqual(L.aufgabeStatus(s, ids[0]), { fehlversuche: 1, geloest: true, loesungAngezeigt: true, mitLoesung: true });
  assert.strictEqual(L.pruefungErgebnis(s, ids[0], false), s);
});

test("Angebot erscheint nach Aufgabe 4, nicht vorher", () => {
  let s = L.neuerStand();
  for (let i = 0; i < 3; i++) s = L.pruefungErgebnis(s, ids[i], true);
  assert.strictEqual(L.hinweisSichtbar(daten, s), false);
  s = L.pruefungErgebnis(s, ids[3], true);
  assert.strictEqual(L.hinweisSichtbar(daten, s), true);
  assert.strictEqual(L.fortschrittText(daten, s), "Aufgabe 5 von 5");
  assert.strictEqual(L.abschluss(daten, s), null);
});

test("Abschluss: alle selbst gelöst", () => {
  let s = L.neuerStand();
  ids.forEach((id) => { s = L.pruefungErgebnis(s, id, true); });
  const a = L.abschluss(daten, s);
  assert.strictEqual(a.ergebnis, "Du hast 5 von 5 Aufgaben selbst gelöst.");
  assert.strictEqual(a.zusatz, null);
  assert.strictEqual(L.fortschrittText(daten, s), "Aufgabe 5 von 5");
});

test("Abschluss: zwei Lösungen angesehen (eine davon danach gelöst)", () => {
  let s = L.neuerStand();
  ids.forEach((id) => { s = L.pruefungErgebnis(s, id, id === ids[1] || id === ids[4] ? false : true); });
  s = L.loesungAnzeigen(s, ids[1]);
  s = L.loesungAnzeigen(s, ids[4]);
  s = L.pruefungErgebnis(s, ids[4], true);
  const a = L.abschluss(daten, s);
  assert.strictEqual(a.selbst, 3);
  assert.strictEqual(a.ergebnis, "Du hast 3 von 5 Aufgaben selbst gelöst.");
  assert.strictEqual(a.zusatz, "Bei 2 hast du dir die Lösung angesehen.");
});

test("Speichern und Laden (Wiederkehrer)", () => {
  const st = speicher();
  let s = L.pruefungErgebnis(L.neuerStand(), ids[0], true);
  s = L.pruefungErgebnis(s, ids[1], false);
  L.speichereStand(st, s);
  const geladen = L.ladeStand(st, daten);
  assert.deepStrictEqual(geladen, s);
  assert.strictEqual(L.fortschrittText(daten, geladen), "Aufgabe 2 von 5");
});

test("Laden: kaputter, alter oder fremder Stand ergibt neuen Stand", () => {
  assert.deepStrictEqual(L.ladeStand(speicher({ [L.STORAGE_KEY]: "{kaputt" }), daten), L.neuerStand());
  assert.deepStrictEqual(L.ladeStand(speicher({ [L.STORAGE_KEY]: '{"version":0,"aufgaben":{}}' }), daten), L.neuerStand());
  assert.deepStrictEqual(L.ladeStand(null, daten), L.neuerStand());
  const fremd = { version: 1, aufgaben: { "gibt-es-nicht": { geloest: true }, [ids[0]]: { geloest: "ja", fehlversuche: -3 } } };
  const g = L.ladeStand(speicher({ [L.STORAGE_KEY]: JSON.stringify(fremd) }), daten);
  assert.deepStrictEqual(g.aufgaben, { [ids[0]]: { fehlversuche: 0, geloest: false, loesungAngezeigt: false, mitLoesung: false } });
});

test("Speichern bei blockiertem localStorage wirft nicht", () => {
  L.speichereStand({ setItem() { throw new Error("blockiert"); } }, L.neuerStand());
});

test("Klick-Tipp-Formular erst mit echter Adresse und Feldnamen verbunden", () => {
  assert.strictEqual(L.formularVerbunden(daten.konfiguration), false, "Platzhalter in daten/aufgaben.json");
  const k = (werte) => ({ klicktipp: Object.assign({ action: "https://app.klick-tipp.com/formular", feldVorname: "fields[fieldFirstName]", feldEmail: "email", versteckteFelder: {} }, werte) });
  assert.strictEqual(L.formularVerbunden(k({})), true);
  assert.strictEqual(L.formularVerbunden(k({ action: "http://unsicher.example/" })), false, "nur https");
  assert.strictEqual(L.formularVerbunden(k({ feldEmail: "PLATZHALTER-FELDNAME-EMAIL" })), false);
  assert.strictEqual(L.formularVerbunden({}), false);
});

test("Gerätetyp: Touch oder schmal = mobil, sonst desktop", () => {
  assert.strictEqual(L.geraetTyp({ breite: 1366, grobZeiger: false }), "desktop");
  assert.strictEqual(L.geraetTyp({ breite: 768, grobZeiger: false }), "desktop");
  assert.strictEqual(L.geraetTyp({ breite: 767, grobZeiger: false }), "mobil");
  assert.strictEqual(L.geraetTyp({ breite: 1024, grobZeiger: true }), "mobil");
});

test("Kampagne: utm-Parameter und fremde Herkunfts-Domain, keine volle Adresse", () => {
  const k = L.kampagne("?utm_source=instagram&utm_medium=social&utm_campaign=story-sept", "https://l.instagram.com/?u=https%3A%2F%2Fx&e=geheim", "florianbeil.github.io");
  assert.deepStrictEqual(k, { utm_source: "instagram", utm_medium: "social", utm_campaign: "story-sept", herkunft: "l.instagram.com" });
  assert.deepStrictEqual(L.kampagne("", "https://florianbeil.github.io/Excel-Aufgaben/", "florianbeil.github.io"), { utm_source: null, utm_medium: null, utm_campaign: null, herkunft: null });
  assert.strictEqual(L.kampagne("", "kein-url", "x").herkunft, null);
  assert.strictEqual(L.kampagne("?utm_source=" + "a".repeat(300), "", "x").utm_source.length, 100);
});

console.log("\n" + tests + " Tests bestanden");
