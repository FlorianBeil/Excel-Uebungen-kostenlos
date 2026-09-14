/* Excel.Flo – kostenlose Übungen: Bildschirmaufbau
 *
 * Baut die fünf Aufgaben untereinander auf, dazwischen (nach Aufgabe 4) den
 * Angebotsblock, darunter den Abschluss. Alle Entscheidungen (was ist gelöst,
 * wann erscheint was) trifft logik.js – hier nur Anzeige und Bedienung.
 *
 * Formel-Aufgaben nutzen das Tabellenblatt aus assets/geteilt/engine.js
 * (window.ExcelFlo). Die Pivot-Aufgabe läuft in pivot.html per iframe und meldet
 * Prüfungen per postMessage.
 *
 * Bewusst kein automatisches Scrollen: Nach dem Lösen erscheint ein Link zur
 * nächsten Aufgabe, die Seite selbst bewegt sich nicht.
 */

(function () {
  "use strict";

  const L = window.ExcelFloKostenlos;
  const E = window.ExcelFlo;
  const DATEN_URL = "daten/aufgaben.json";
  const reduzierteBewegung = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;

  let speicher = null;
  try {
    speicher = window.localStorage;
  } catch (e) {
    // blockiert – Fortschritt gilt dann nur für diesen Besuch
  }

  let daten = null;
  let stand = null;
  const eintraege = []; // pro Aufgabe: { a, nr, karte, status, feedback, loesungBtn, loesungBox }
  let angebotEl = null;
  let abschlussEl = null;

  /* ---------------- Hilfen ---------------- */

  function h(tag, attrs, kinder) {
    const node = document.createElement(tag);
    Object.entries(attrs || {}).forEach(([k, v]) => {
      if (v === null || v === undefined || v === false) return;
      if (k === "class") node.className = v;
      else if (k === "text") node.textContent = v;
      else if (k === "hidden") node.hidden = true;
      else node.setAttribute(k, v === true ? "" : v);
    });
    (kinder || []).forEach((kind) => {
      if (kind) node.appendChild(kind);
    });
    return node;
  }

  function istPlatzhalter(url) {
    return !url || /^PLATZHALTER/.test(url);
  }

  function seitenUrl() {
    const k = daten.konfiguration;
    return istPlatzhalter(k.seitenUrl) ? location.origin + location.pathname : k.seitenUrl;
  }

  function speichern() {
    L.speichereStand(speicher, stand);
  }

  function springeZu(ziel) {
    if (!ziel) return;
    ziel.scrollIntoView({ behavior: reduzierteBewegung ? "auto" : "smooth", block: "start" });
    ziel.focus({ preventScroll: true });
  }

  /* ---------------- Start ---------------- */

  function start() {
    const root = document.getElementById("aufgaben");
    fetch(DATEN_URL, { cache: "no-cache" })
      .then((res) => {
        if (!res.ok) throw new Error("HTTP " + res.status);
        return res.json();
      })
      .then((d) => {
        daten = d;
        const probleme = L.pruefeDaten(d);
        if (probleme.length) console.warn("daten/aufgaben.json:", probleme);
        stand = L.ladeStand(speicher, d);
        aufbauen(root);
        aktualisieren();
      })
      .catch((err) => {
        console.error(err);
        root.textContent = "";
        root.appendChild(h("p", { class: "frei-laden", text: "Die Aufgaben konnten nicht geladen werden. Bitte lade die Seite neu." }));
      });
  }

  function aufbauen(root) {
    const t = daten.texte;
    const k = daten.konfiguration;

    document.getElementById("seiten-titel").textContent = t.ueberschrift;
    document.querySelector(".frei-start__text").textContent = t.einleitung;

    document.getElementById("mobil-hinweis").textContent = t.mobilHinweis;
    const mail = document.getElementById("mobil-mail");
    mail.textContent = t.mobilMailLink;
    mail.href = L.mailLink(daten, seitenUrl());
    const webinar = document.getElementById("mobil-webinar");
    webinar.textContent = t.mobilWebinarLink;
    webinar.href = k.webinarUrl;

    [["link-impressum", k.impressumUrl, t.footerImpressum], ["link-datenschutz", k.datenschutzUrl, t.footerDatenschutz]].forEach(([id, url, text]) => {
      const link = document.getElementById(id);
      link.textContent = text;
      link.href = url;
      link.hidden = istPlatzhalter(url);
    });

    root.textContent = "";
    daten.aufgaben.forEach((a, i) => {
      root.appendChild(karteBauen(a, i));
      if (i + 1 === k.angebotNachAufgabe) {
        angebotEl = angebotBauen();
        root.appendChild(angebotEl);
      }
    });
    abschlussEl = abschlussBauen();
    root.appendChild(abschlussEl);

    // Wiederkehrer: bereits gelöste Aufgaben zeigen ihre Erklärung wieder
    eintraege.forEach((e) => {
      if (L.aufgabeStatus(stand, e.a.id).geloest) feedbackErfolg(e, true);
    });
  }

  /* ---------------- Aufgaben-Karte ---------------- */

  function karteBauen(a, i) {
    const nr = i + 1;
    const gesamt = daten.aufgaben.length;
    const titelId = "aufgabe-" + nr + "-titel";
    const loesungId = "aufgabe-" + nr + "-loesung";

    const status = h("span", { class: "frei-status", hidden: true });
    const karte = h("section", { class: "frei-aufgabe", id: "aufgabe-" + nr, "aria-labelledby": titelId, tabindex: "-1" }, [
      h("header", { class: "frei-aufgabe__kopf" }, [
        h("p", { class: "frei-aufgabe__nr", text: L.platzhalter(daten.texte.fortschritt, { nr, gesamt }) }),
        status,
        h("h2", { id: titelId, text: a.title }),
      ]),
      h("div", { class: "exercise-task" }, [
        a.task.intro ? h("p", { class: "exercise-task__intro", text: a.task.intro }) : null,
        a.task.steps && a.task.steps.length
          ? h("ol", { class: "exercise-task__steps" }, a.task.steps.map((s) => h("li", { text: s })))
          : null,
      ]),
    ]);

    const eintrag = {
      a,
      nr,
      karte,
      status,
      feedback: h("div", { class: "exercise-feedback frei-feedback", role: "status" }),
      loesungBtn: h("button", { type: "button", class: "btn btn--secondary", text: "Lösung anzeigen", "aria-controls": loesungId, hidden: true }),
      loesungBox: h("div", { class: "frei-loesung", id: loesungId, tabindex: "-1", hidden: true }, [
        h("p", { class: "frei-loesung__titel", text: "Lösung" }),
        h("p", { class: a.typ === "formel" ? "frei-loesung__formel" : "frei-loesung__text", text: a.solution }),
        h("p", { class: "frei-loesung__erklaerung", text: a.explanation }),
      ]),
    };
    eintraege.push(eintrag);

    if (a.typ === "pivot") pivotTeil(eintrag);
    else formelTeil(eintrag);

    eintrag.loesungBtn.addEventListener("click", () => {
      stand = L.loesungAnzeigen(stand, a.id);
      speichern();
      aktualisieren();
      eintrag.loesungBox.focus({ preventScroll: true });
    });

    karte.appendChild(eintrag.feedback);
    karte.appendChild(eintrag.loesungBox);
    if (a.hints && a.hints.length) {
      karte.appendChild(
        h("details", { class: "exercise-hints frei-tipps" }, [
          h("summary", { text: "Tipps anzeigen" }),
          h("ol", {}, a.hints.map((tipp) => h("li", { text: tipp }))),
        ])
      );
    }
    return karte;
  }

  function formelTeil(eintrag) {
    const sheet = E.createSheet(eintrag.a.grid);
    const pruefen = h("button", { type: "button", class: "btn btn--primary", text: "Prüfen" });
    const zuruecksetzen = h("button", { type: "button", class: "btn btn--secondary", text: "Zurücksetzen" });

    pruefen.addEventListener("click", () => formelPruefen(eintrag, sheet));
    zuruecksetzen.addEventListener("click", () => {
      sheet.reset();
      Object.values(sheet.inputEntries).forEach((e) => e.td.classList.remove("is-correct", "is-wrong"));
      feedbackLeeren(eintrag);
    });

    eintrag.karte.appendChild(sheet.node);
    eintrag.karte.appendChild(h("div", { class: "exercise-actions" }, [pruefen, zuruecksetzen, eintrag.loesungBtn]));
  }

  function formelPruefen(eintrag, sheet) {
    const refs = Object.keys(sheet.inputEntries);
    let beantwortet = 0;
    let richtig = 0;
    refs.forEach((ref) => {
      const e = sheet.inputEntries[ref];
      const ergebnis = E.checkCell(e.raw, e.answer, sheet.getCellValue);
      e.td.classList.remove("is-correct", "is-wrong");
      if (ergebnis === true) {
        e.td.classList.add("is-correct");
        beantwortet++;
        richtig++;
      } else if (ergebnis === false) {
        e.td.classList.add("is-wrong");
        beantwortet++;
      }
    });

    if (beantwortet === 0) {
      // zählt nicht als Fehlversuch – es wurde ja noch nichts eingegeben
      const meldung = "Trag zuerst eine Formel in die gelb markierte Zelle ein.";
      feedbackFehler(eintrag, meldung);
      E.showErrorPopup(sheet.node, meldung);
      return;
    }

    const ok = richtig === refs.length;
    const meldung = refs.length === 1 ? "Das Ergebnis stimmt noch nicht." : richtig + " von " + refs.length + " Feldern stimmen.";
    pruefungVerarbeiten(eintrag, ok);
    if (ok) {
      E.showSuccessPopup(sheet.node);
      feedbackErfolg(eintrag, false);
    } else {
      E.showErrorPopup(sheet.node, meldung);
      feedbackFehler(eintrag, meldung + (L.loesungErlaubt(stand, eintrag.a.id) ? " Schau dir die Tipps an – oder lass dir die Lösung anzeigen." : ""));
    }
  }

  function pivotTeil(eintrag) {
    const iframe = h("iframe", {
      class: "frei-pivot",
      src: "pivot.html",
      title: "Pivot-Tabelle für Aufgabe " + eintrag.nr,
      loading: "lazy",
    });

    // Höhe folgt dem Inhalt (gleiche Herkunft, daher direkt messbar) – keine zweite Scrollleiste.
    iframe.addEventListener("load", () => {
      try {
        const doc = iframe.contentDocument;
        const anpassen = () => {
          iframe.style.height = Math.ceil(doc.body.getBoundingClientRect().height) + "px";
        };
        anpassen();
        new ResizeObserver(anpassen).observe(doc.body);
      } catch (e) {
        // Messen nicht möglich – feste Höhe aus dem CSS bleibt
      }
    });

    window.addEventListener("message", (ev) => {
      if (ev.origin !== location.origin || ev.source !== iframe.contentWindow) return;
      const m = ev.data;
      if (!m || m.quelle !== "excelflo-pivot" || m.typ !== "pruefung") return;
      const ok = !!(m.daten && m.daten.richtig);
      pruefungVerarbeiten(eintrag, ok);
      // Rückmeldung zum Fehler zeigt der Pivot-Nachbau selbst; hier nur Erfolg + Weiter
      if (ok) feedbackErfolg(eintrag, false);
      else feedbackLeeren(eintrag);
    });

    eintrag.karte.appendChild(iframe);
    eintrag.karte.appendChild(h("div", { class: "exercise-actions frei-pivot__aktionen" }, [eintrag.loesungBtn]));
  }

  function pruefungVerarbeiten(eintrag, ok) {
    stand = L.pruefungErgebnis(stand, eintrag.a.id, ok);
    speichern();
    aktualisieren();
  }

  /* ---------------- Rückmeldungen ---------------- */

  function feedbackLeeren(eintrag) {
    eintrag.feedback.className = "exercise-feedback frei-feedback";
    eintrag.feedback.textContent = "";
  }

  function feedbackFehler(eintrag, meldung) {
    eintrag.feedback.className = "exercise-feedback frei-feedback is-error";
    eintrag.feedback.textContent = "";
    eintrag.feedback.appendChild(h("p", { text: meldung }));
  }

  function feedbackErfolg(eintrag, wiederhergestellt) {
    const f = eintrag.feedback;
    f.className = "exercise-feedback frei-feedback is-success";
    f.textContent = "";
    f.appendChild(h("p", { text: wiederhergestellt ? "Diese Aufgabe hast du schon gelöst." : "Richtig!" }));
    f.appendChild(h("p", { class: "exercise-feedback__explanation", text: eintrag.a.explanation }));

    const letzte = eintrag.nr === daten.aufgaben.length;
    const ziel = letzte ? abschlussEl : document.getElementById("aufgabe-" + (eintrag.nr + 1));
    const link = h("a", {
      class: "frei-weiter",
      href: letzte ? "#abschluss" : "#aufgabe-" + (eintrag.nr + 1),
      text: letzte ? "Zu deinem Ergebnis ↓" : "Weiter zu Aufgabe " + (eintrag.nr + 1) + " ↓",
    });
    link.addEventListener("click", (ev) => {
      ev.preventDefault();
      springeZu(ziel);
    });
    f.appendChild(h("p", {}, [link]));
  }

  /* ---------------- Angebot und Abschluss ---------------- */

  function webinarButton() {
    return h("a", { class: "btn frei-cta", href: daten.konfiguration.webinarUrl, text: daten.texte.webinarButton });
  }

  function angebotBauen() {
    const t = daten.texte;
    return h("section", { class: "frei-angebot", id: "angebot", "aria-labelledby": "angebot-titel", hidden: true }, [
      h("h2", { id: "angebot-titel", text: t.angebotTitel }),
      h("p", { text: t.angebotText1 }),
      h("p", { text: t.angebotText2 }),
      h("p", { class: "frei-angebot__aktion" }, [webinarButton()]),
      h("p", { class: "frei-angebot__kurs" }, [h("a", { class: "frei-textlink", href: daten.konfiguration.kursUrl, text: t.kursLink })]),
    ]);
  }

  function abschlussBauen() {
    return h("section", { class: "frei-abschluss", id: "abschluss", "aria-labelledby": "abschluss-titel", tabindex: "-1", hidden: true }, [
      h("h2", { id: "abschluss-titel" }),
      h("p", { class: "frei-abschluss__ergebnis" }),
      h("p", { class: "frei-abschluss__zusatz" }),
      h("p", { class: "frei-abschluss__aktion" }, [webinarButton()]),
    ]);
  }

  /* ---------------- Anzeige an den Stand anpassen ---------------- */

  function aktualisieren() {
    document.getElementById("fortschritt").textContent = L.fortschrittText(daten, stand);

    eintraege.forEach((e) => {
      const s = L.aufgabeStatus(stand, e.a.id);
      e.status.hidden = !(s.geloest || s.loesungAngezeigt);
      e.status.textContent = s.geloest ? "✓ Gelöst" : "Lösung angesehen";
      e.status.classList.toggle("is-loesung", !s.geloest);
      e.loesungBtn.hidden = !L.loesungErlaubt(stand, e.a.id) || s.loesungAngezeigt;
      e.loesungBox.hidden = !s.loesungAngezeigt;
    });

    angebotEl.hidden = !L.angebotSichtbar(daten, stand);

    const a = L.abschluss(daten, stand);
    abschlussEl.hidden = !a;
    if (a) {
      abschlussEl.querySelector("h2").textContent = a.titel;
      abschlussEl.querySelector(".frei-abschluss__ergebnis").textContent = a.ergebnis;
      const zusatz = abschlussEl.querySelector(".frei-abschluss__zusatz");
      zusatz.textContent = a.zusatz || "";
      zusatz.hidden = !a.zusatz;
    }
  }

  start();
})();
