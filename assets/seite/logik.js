/* Excel.Flo – kostenlose Übungen: Ablauf-Logik ohne Bildschirmaufbau
 *
 * Arbeitet nur mit den Inhalten aus daten/aufgaben.json und einem einfachen
 * Speicherstand-Objekt – kein DOM, kein fetch, dadurch mit node testbar
 * (test/logik.test.js). Der Speicherstand wird nie verändert, jede Aktion gibt
 * einen neuen Stand zurück.
 *
 * Pro Aufgabe: fehlversuche, geloest, loesungAngezeigt, mitLoesung
 * (mitLoesung = erst nach dem Ansehen der Lösung richtig gelöst).
 * „Bearbeitet“ ist eine Aufgabe, wenn sie gelöst ODER die Lösung angesehen wurde –
 * danach geht es mit der nächsten Aufgabe weiter.
 */

(function (root) {
  "use strict";

  const STORAGE_KEY = "excelflo_kostenlos_v1";
  const STAND_VERSION = 1;
  const TYPEN = ["formel", "pivot"];

  /* ---------------- Speicherstand ---------------- */

  function neuerStand() {
    return { version: STAND_VERSION, aufgaben: {} };
  }

  function leererAufgabenStatus() {
    return { fehlversuche: 0, geloest: false, loesungAngezeigt: false, mitLoesung: false };
  }

  function aufgabeStatus(stand, id) {
    return Object.assign(leererAufgabenStatus(), stand.aufgaben[id]);
  }

  function mitStatus(stand, id, aenderung) {
    const aufgaben = Object.assign({}, stand.aufgaben);
    aufgaben[id] = Object.assign(aufgabeStatus(stand, id), aenderung);
    return Object.assign({}, stand, { aufgaben });
  }

  // storage: localStorage oder ein Objekt mit getItem/setItem (Tests).
  // Unlesbare, veraltete oder fremde Stände ergeben einen neuen Stand – ein
  // Wiederkehrer soll nie eine kaputte Seite sehen, schlimmstenfalls beginnt er neu.
  function ladeStand(storage, daten) {
    try {
      const roh = storage && storage.getItem(STORAGE_KEY);
      if (!roh) return neuerStand();
      const gelesen = JSON.parse(roh);
      if (!gelesen || gelesen.version !== STAND_VERSION || typeof gelesen.aufgaben !== "object") return neuerStand();
      const stand = neuerStand();
      daten.aufgaben.forEach((a) => {
        const s = gelesen.aufgaben[a.id];
        if (!s || typeof s !== "object") return;
        stand.aufgaben[a.id] = {
          fehlversuche: Number.isInteger(s.fehlversuche) && s.fehlversuche > 0 ? s.fehlversuche : 0,
          geloest: s.geloest === true,
          loesungAngezeigt: s.loesungAngezeigt === true,
          mitLoesung: s.mitLoesung === true,
        };
      });
      return stand;
    } catch (e) {
      return neuerStand();
    }
  }

  function speichereStand(storage, stand) {
    try {
      storage.setItem(STORAGE_KEY, JSON.stringify(stand));
    } catch (e) {
      // localStorage blockiert (privates Fenster o. Ä.) – Fortschritt gilt dann nur für diesen Besuch
    }
  }

  /* ---------------- Aktionen ---------------- */

  function pruefungErgebnis(stand, id, richtig) {
    const s = aufgabeStatus(stand, id);
    if (s.geloest) return stand; // erneutes Prüfen nach dem Lösen ändert nichts mehr
    if (richtig) return mitStatus(stand, id, { geloest: true, mitLoesung: s.loesungAngezeigt });
    return mitStatus(stand, id, { fehlversuche: s.fehlversuche + 1 });
  }

  // „Lösung anzeigen“ ist erst nach einem Fehlversuch erlaubt.
  function loesungErlaubt(stand, id) {
    return aufgabeStatus(stand, id).fehlversuche >= 1;
  }

  function loesungAnzeigen(stand, id) {
    if (!loesungErlaubt(stand, id)) return stand;
    const s = aufgabeStatus(stand, id);
    if (s.loesungAngezeigt || s.geloest) {
      return s.loesungAngezeigt ? stand : mitStatus(stand, id, { loesungAngezeigt: true });
    }
    return mitStatus(stand, id, { loesungAngezeigt: true });
  }

  /* ---------------- Abgeleitete Werte ---------------- */

  function istBearbeitet(stand, id) {
    const s = aufgabeStatus(stand, id);
    return s.geloest || s.loesungAngezeigt;
  }

  function anzahlBearbeitet(daten, stand) {
    return daten.aufgaben.filter((a) => istBearbeitet(stand, a.id)).length;
  }

  // Index der ersten nicht bearbeiteten Aufgabe, null wenn alle bearbeitet sind.
  function aktuelleAufgabe(daten, stand) {
    const idx = daten.aufgaben.findIndex((a) => !istBearbeitet(stand, a.id));
    return idx === -1 ? null : idx;
  }

  function fortschrittText(daten, stand) {
    const idx = aktuelleAufgabe(daten, stand);
    const gesamt = daten.aufgaben.length;
    return platzhalter(daten.texte.fortschritt, { nr: idx === null ? gesamt : idx + 1, gesamt });
  }

  // Angebotsblock erscheint, sobald die ersten N Aufgaben (Standard 4) bearbeitet sind.
  function angebotSichtbar(daten, stand) {
    const n = daten.konfiguration.angebotNachAufgabe;
    return daten.aufgaben.slice(0, n).every((a) => istBearbeitet(stand, a.id));
  }

  // null, solange nicht alle Aufgaben bearbeitet sind.
  function abschluss(daten, stand) {
    if (aktuelleAufgabe(daten, stand) !== null) return null;
    const selbst = daten.aufgaben.filter((a) => {
      const s = aufgabeStatus(stand, a.id);
      return s.geloest && !s.mitLoesung;
    }).length;
    const mitLoesung = daten.aufgaben.length - selbst;
    const t = daten.texte;
    return {
      selbst,
      mitLoesung,
      titel: t.abschlussTitel,
      ergebnis: platzhalter(t.abschlussErgebnis, { selbst }),
      zusatz: mitLoesung > 0 ? platzhalter(t.abschlussMitLoesung, { mitLoesung }) : null,
    };
  }

  /* ---------------- Hilfen ---------------- */

  function platzhalter(text, werte) {
    return String(text).replace(/\{(\w+)\}/g, (voll, name) => (name in werte ? String(werte[name]) : voll));
  }

  // mailto ohne Empfänger: das Mailprogramm öffnet sich, der Besucher trägt seine
  // eigene Adresse ein. Die Seite erfährt keine Adresse.
  function mailLink(daten, seitenUrl) {
    const t = daten.texte;
    return (
      "mailto:?subject=" + encodeURIComponent(t.mailBetreff) +
      "&body=" + encodeURIComponent(t.mailText + "\n\n" + seitenUrl)
    );
  }

  /* ---------------- Tracking-Hilfen ---------------- */

  // „mobil“ für Touch-Geräte ohne Maus oder schmale Fenster, sonst „desktop“.
  // Tablets zählen bewusst als mobil: Formeln tippen ist dort genauso mühsam.
  function geraetTyp(umgebung) {
    if (umgebung.grobZeiger) return "mobil";
    return umgebung.breite < 768 ? "mobil" : "desktop";
  }

  // Herkunft eines Besuchs: utm-Parameter aus der Adresse (z. B. ?utm_source=instagram)
  // und die Domain der verweisenden Seite – nie die volle Adresse, keine Personendaten.
  function kampagne(suche, referrer, eigenerHost) {
    const params = new URLSearchParams(suche || "");
    const kurz = (wert) => (wert ? String(wert).slice(0, 100) : null);
    let herkunft = null;
    try {
      const host = referrer ? new URL(referrer).host : "";
      if (host && host !== eigenerHost) herkunft = host.slice(0, 100);
    } catch (e) {
      // ungültiger Referrer – ignorieren
    }
    return {
      utm_source: kurz(params.get("utm_source")),
      utm_medium: kurz(params.get("utm_medium")),
      utm_campaign: kurz(params.get("utm_campaign")),
      herkunft,
    };
  }

  // Liefert eine Liste von Problemen (leer = alles in Ordnung).
  function pruefeDaten(daten) {
    const probleme = [];
    if (!daten || !Array.isArray(daten.aufgaben) || !daten.aufgaben.length) return ["Keine Aufgaben vorhanden"];
    if (!daten.konfiguration) probleme.push("konfiguration fehlt");
    if (!daten.texte) probleme.push("texte fehlen");
    const ids = new Set();
    daten.aufgaben.forEach((a, i) => {
      const wo = "Aufgabe " + (i + 1);
      if (!a.id) probleme.push(wo + ": id fehlt");
      else if (ids.has(a.id)) probleme.push(wo + ": id doppelt (" + a.id + ")");
      ids.add(a.id);
      if (!TYPEN.includes(a.typ)) probleme.push(wo + ": unbekannter typ " + a.typ);
      if (!a.title || !a.task || !a.solution || !a.explanation) probleme.push(wo + ": title/task/solution/explanation unvollständig");
      if (a.typ === "formel" && (!a.grid || !Object.values(a.grid.cells).some((c) => c.type === "input"))) {
        probleme.push(wo + ": Formel-Aufgabe ohne Eingabezelle");
      }
      if (a.typ === "pivot" && (!a.datensatz || !a.loesungPruefung)) probleme.push(wo + ": Pivot-Aufgabe ohne datensatz/loesungPruefung");
    });
    if (daten.konfiguration) {
      ["webinarUrl", "kursUrl", "impressumUrl", "datenschutzUrl", "seitenUrl"].forEach((k) => {
        if (!daten.konfiguration[k]) probleme.push("konfiguration." + k + " fehlt");
      });
      const n = daten.konfiguration.angebotNachAufgabe;
      if (!Number.isInteger(n) || n < 1 || n >= daten.aufgaben.length) {
        probleme.push("konfiguration.angebotNachAufgabe muss zwischen 1 und " + (daten.aufgaben.length - 1) + " liegen");
      }
    }
    return probleme;
  }

  const api = {
    STORAGE_KEY,
    neuerStand,
    ladeStand,
    speichereStand,
    aufgabeStatus,
    pruefungErgebnis,
    loesungErlaubt,
    loesungAnzeigen,
    istBearbeitet,
    anzahlBearbeitet,
    aktuelleAufgabe,
    fortschrittText,
    angebotSichtbar,
    abschluss,
    platzhalter,
    mailLink,
    geraetTyp,
    kampagne,
    pruefeDaten,
  };

  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.ExcelFloKostenlos = api;
})(typeof window !== "undefined" ? window : this);
