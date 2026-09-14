/* Excel.Flo – kostenlose Übungen: Beispieldaten der Pivot-Aufgabe
 *
 * Erzeugt die Quelltabelle aus dem Bauplan in daten/aufgaben.json
 * (aufgabe.datensatz) statt 730 Zeilen mitzuliefern. Seeded Zufall, daher bei
 * jedem Aufruf identische Daten. Reihenfolge der Zufallsaufrufe (Region, Land,
 * Kunde, Umsatz je Tag) entspricht exakt buildSalesDataset() im Pivot-Repo –
 * so sind die Zahlen dieselben wie in der Kurs-Übung (geprüft in
 * test/datensatz.test.js). Kein DOM, mit node testbar.
 */

(function (root) {
  "use strict";

  function mulberry32(seed) {
    return function () {
      seed |= 0; seed = (seed + 0x6D2B79F5) | 0;
      var t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  function utcAusIso(iso) {
    var teile = iso.split("-").map(Number);
    return Date.UTC(teile[0], teile[1] - 1, teile[2]);
  }

  function erzeugeDatensatz(def) {
    var rand = mulberry32(def.seed);
    var regionen = Object.keys(def.regionen);
    var start = utcAusIso(def.vonDatum);
    var tage = (utcAusIso(def.bisDatum) - start) / 86400000 + 1; // bisDatum inklusive
    var rows = [];
    for (var i = 0; i < tage; i++) {
      var datum = new Date(start);
      datum.setUTCDate(datum.getUTCDate() + i);
      var region = regionen[Math.floor(rand() * regionen.length)];
      var r = def.regionen[region];
      var land = r.laender[Math.floor(rand() * r.laender.length)];
      var kunde = r.kunden[Math.floor(rand() * r.kunden.length)];
      var umsatz = Math.round(r.basisUmsatz + (rand() - 0.4) * 700);
      if (umsatz < 50) umsatz = 50;
      rows.push({ Datum: datum, Region: region, Land: land, Kunde: kunde, Umsatz: umsatz });
    }
    return { fields: def.felder, columns: Object.keys(def.felder), rows: rows };
  }

  var api = { erzeugeDatensatz: erzeugeDatensatz };

  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.ExcelFloPivotDatensatz = api;
})(typeof window !== "undefined" ? window : this);
