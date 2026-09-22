/* Excel.Flo – kostenlose Übungen: Bildschirmaufbau
 *
 * Baut die fünf Aufgaben untereinander auf, dazwischen (nach Aufgabe 4) den
 * Hinweis auf das Übungsportal Light, darunter den Abschluss mit dem Klick-Tipp-Formular
 * (Vorname + E-Mail → Double-Opt-in → Mail mit dem Link zum Light-Portal). Dasselbe
 * Formular gibt es aufklappbar im Hinweis für Smartphones. Webinar und Kurs werden auf
 * dieser Seite bewusst nicht beworben. Alle Entscheidungen (was ist gelöst,
 * wann erscheint was) trifft logik.js – hier nur Anzeige und Bedienung.
 *
 * Formel-Aufgaben nutzen das Tabellenblatt aus assets/geteilt/engine.js
 * (window.ExcelFlo). Die Pivot-Aufgabe läuft in pivot.html per iframe und meldet
 * Prüfungen per postMessage.
 *
 * Bewusst kein automatisches Scrollen: Nach dem Lösen erscheint ein Link zur
 * nächsten Aufgabe, die Seite selbst bewegt sich nicht.
 *
 * Tracking (assets/geteilt/tracking.js → Supabase-Tabelle events, portal = 'kostenlos',
 * Auswertungen in supabase/kostenlos.sql). Jedes Ereignis trägt detail.geraet
 * („desktop“/„mobil“), Aufgaben-Ereignisse zusätzlich exercise_id und detail.nr:
 *   page_view        Seite aufgerufen       breite, wiederkehrer, bearbeitet, utm_*, herkunft
 *   exercise_view    Aufgabe ins Bild gescrollt (einmal pro Seitenaufruf)
 *   exercise_start   erste Bedienung der Aufgabe (Tabelle/Pivot angetippt oder Taste, Prüfen)
 *   check            Prüfen                 richtig, leer, versuch
 *   exercise_solved  Aufgabe gelöst         mit_loesung
 *   solution_show    Lösung angezeigt
 *   hints_open       Tipps aufgeklappt
 *   teaser_view      Hinweis auf das Light-Portal nach Aufgabe 4 im Bild
 *   form_view        Formular gesehen       ort: hinweis | abschluss | mobil (mobil = aufgeklappt)
 *   form_submit      Formular abgesendet    ort: hinweis | abschluss | mobil (vor der Weiterleitung zu Klick-Tipp)
 */

(function () {
  "use strict";

  const L = window.ExcelFloKostenlos;
  const E = window.ExcelFlo;
  const DATEN_URL = "daten/aufgaben.json";
  const TRACKING_PORTAL = "kostenlos";
  const reduzierteBewegung = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;

  let speicher = null;
  try {
    speicher = window.localStorage;
  } catch (e) {
    // blockiert – Fortschritt gilt dann nur für diesen Besuch
  }

  const geraet = L.geraetTyp({
    breite: window.innerWidth,
    grobZeiger: !!(window.matchMedia && window.matchMedia("(hover: none) and (pointer: coarse)").matches),
  });

  let daten = null;
  let stand = null;
  const eintraege = []; // pro Aufgabe: { a, nr, karte, status, feedback, loesungBtn, loesungBox, gestartet, versuche }
  let hinweisEl = null;
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

  function speichern() {
    L.speichereStand(speicher, stand);
  }

  function springeZu(ziel) {
    if (!ziel) return;
    ziel.scrollIntoView({ behavior: reduzierteBewegung ? "auto" : "smooth", block: "start" });
    ziel.focus({ preventScroll: true });
  }

  /* ---------------- Tracking ---------------- */

  // window.ExcelFloTracking erst beim Senden nachschlagen – fehlt es (Blocker, Ladefehler),
  // passiert einfach nichts. Die Übungen funktionieren immer ohne Tracking.
  function track(event, detail, eintrag) {
    const t = window.ExcelFloTracking;
    if (!t) return;
    const d = Object.assign({ geraet }, eintrag ? { nr: eintrag.nr } : {}, detail || {});
    t.track(TRACKING_PORTAL, eintrag ? eintrag.a.id : null, event, d);
  }

  function aufgabeGestartet(eintrag) {
    if (eintrag.gestartet) return;
    eintrag.gestartet = true;
    track("exercise_start", null, eintrag);
  }

  // Zählt als „gesehen“, sobald ein nennenswerter Teil sichtbar ist – bei sehr hohen
  // Blöcken (Pivot auf dem Smartphone) reicht eine halbe Bildschirmhöhe.
  function beiSichtbarkeit(element, callback) {
    if (!("IntersectionObserver" in window)) return;
    const beobachter = new IntersectionObserver(
      (eintraegeIO) => {
        eintraegeIO.forEach((io) => {
          if (!io.isIntersecting) return;
          const noetig = Math.min(io.boundingClientRect.height * 0.4, window.innerHeight * 0.5);
          if (io.intersectionRect.height >= noetig) {
            beobachter.disconnect();
            callback();
          }
        });
      },
      { threshold: [0, 0.1, 0.25, 0.4, 0.6, 1] }
    );
    beobachter.observe(element);
  }

  /* ---------------- Start ---------------- */

  function start() {
    const root = document.getElementById("aufgaben");
    const herkunft = L.kampagne(location.search, document.referrer, location.host);
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
        const bearbeitet = L.anzahlBearbeitet(d, stand);
        track("page_view", Object.assign({ breite: window.innerWidth, wiederkehrer: Object.keys(stand.aufgaben).length > 0, bearbeitet }, herkunft));
        aufbauen(root);
        aktualisieren();
      })
      .catch((err) => {
        console.error(err);
        track("page_view", Object.assign({ breite: window.innerWidth, ladefehler: true }, herkunft));
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
    mobilFormularBauen();

    [["link-impressum", k.impressumUrl, t.footerImpressum], ["link-datenschutz", k.datenschutzUrl, t.footerDatenschutz]].forEach(([id, url, text]) => {
      const link = document.getElementById(id);
      link.textContent = text;
      link.href = url;
      link.hidden = istPlatzhalter(url);
    });

    root.textContent = "";
    daten.aufgaben.forEach((a, i) => {
      root.appendChild(karteBauen(a, i));
      if (i + 1 === k.hinweisNachAufgabe) {
        hinweisEl = hinweisBauen();
        root.appendChild(hinweisEl);
      }
    });
    abschlussEl = abschlussBauen();
    root.appendChild(abschlussEl);

    eintraege.forEach((e) => {
      // Wiederkehrer: bereits gelöste Aufgaben zeigen ihre Erklärung wieder
      if (L.aufgabeStatus(stand, e.a.id).geloest) feedbackErfolg(e, true);
      beiSichtbarkeit(e.karte, () => track("exercise_view", null, e));
    });
    schutzSkriptLaden();
    beiSichtbarkeit(hinweisEl, () => track("teaser_view"));
    beiSichtbarkeit(hinweisEl.querySelector(".frei-formular"), () => track("form_view", { ort: "hinweis" }));
    beiSichtbarkeit(abschlussEl.querySelector(".frei-formular"), () => track("form_view", { ort: "abschluss" }));
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
      gestartet: false,
      versuche: 0,
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
      track("solution_show", null, eintrag);
      aktualisieren();
      eintrag.loesungBox.focus({ preventScroll: true });
    });

    karte.appendChild(eintrag.feedback);
    karte.appendChild(eintrag.loesungBox);
    if (a.hints && a.hints.length) {
      const tipps = h("details", { class: "exercise-hints frei-tipps" }, [
        h("summary", { text: "Tipps anzeigen" }),
        h("ol", {}, a.hints.map((tipp) => h("li", { text: tipp }))),
      ]);
      tipps.addEventListener("toggle", () => {
        if (tipps.open) track("hints_open", null, eintrag);
      });
      karte.appendChild(tipps);
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

    // Erste echte Bedienung der Tabelle = Aufgabe gestartet (Tab-Durchlaufen zählt nicht:
    // dessen keydown landet beim vorherigen Element)
    sheet.node.addEventListener("pointerdown", () => aufgabeGestartet(eintrag));
    sheet.node.addEventListener("keydown", () => aufgabeGestartet(eintrag));

    sheet.node.setAttribute("role", "group");
    sheet.node.setAttribute(
      "aria-label",
      "Tabelle zu Aufgabe " + eintrag.nr + ". Zellen mit den Pfeiltasten wählen, zum Bearbeiten tippen oder Enter drücken."
    );
    eintrag.karte.appendChild(sheet.node);
    eintrag.karte.appendChild(h("div", { class: "exercise-actions" }, [pruefen, zuruecksetzen, eintrag.loesungBtn]));
  }

  function formelPruefen(eintrag, sheet) {
    aufgabeGestartet(eintrag);
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
      track("check", { richtig: false, leer: true }, eintrag);
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

    iframe.addEventListener("load", () => {
      try {
        const doc = iframe.contentDocument;
        // Höhe folgt dem Inhalt (gleiche Herkunft, daher direkt messbar) – keine zweite Scrollleiste.
        const anpassen = () => {
          iframe.style.height = Math.ceil(doc.body.getBoundingClientRect().height) + "px";
        };
        anpassen();
        new ResizeObserver(anpassen).observe(doc.body);
        // Erste Bedienung im Pivot-Nachbau = Aufgabe gestartet
        doc.addEventListener("pointerdown", () => aufgabeGestartet(eintrag));
        doc.addEventListener("keydown", () => aufgabeGestartet(eintrag));
      } catch (e) {
        // Messen nicht möglich – feste Höhe aus dem CSS bleibt
      }
    });

    window.addEventListener("message", (ev) => {
      if (ev.origin !== location.origin || ev.source !== iframe.contentWindow) return;
      const m = ev.data;
      if (!m || m.quelle !== "excelflo-pivot" || m.typ !== "pruefung") return;
      aufgabeGestartet(eintrag);
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
    const vorher = L.aufgabeStatus(stand, eintrag.a.id);
    stand = L.pruefungErgebnis(stand, eintrag.a.id, ok);
    speichern();
    eintrag.versuche++;
    track("check", { richtig: ok, versuch: eintrag.versuche }, eintrag);
    const nachher = L.aufgabeStatus(stand, eintrag.a.id);
    if (!vorher.geloest && nachher.geloest) track("exercise_solved", { mit_loesung: nachher.mitLoesung }, eintrag);
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
    const link = h("a", { class: "frei-weiter", href: letzte ? "#abschluss" : "#aufgabe-" + (eintrag.nr + 1) }, [
      document.createTextNode(letzte ? "Zu deinem Ergebnis " : "Weiter zu Aufgabe " + (eintrag.nr + 1) + " "),
      h("span", { "aria-hidden": "true", text: "↓" }),
    ]);
    link.addEventListener("click", (ev) => {
      ev.preventDefault();
      springeZu(ziel);
    });
    f.appendChild(h("p", {}, [link]));
  }

  /* ---------------- Light-Portal: Hinweis, Formular, Abschluss ---------------- */

  // Nach Aufgabe 4: derselbe Anmeldeblock wie im Abschluss, damit niemand bis ganz
  // nach unten scrollen muss. Wer durchzieht, findet ihn unten noch einmal.
  function hinweisBauen() {
    const t = daten.texte;
    return h("section", { class: "frei-hinweis", id: "hinweis", "aria-labelledby": "hinweis-titel", hidden: true }, [
      h("h2", { id: "hinweis-titel", text: t.hinweisTitel }),
      h("p", { text: t.hinweisText }),
      formularBauen("hinweis"),
    ]);
  }

  // Klick-Tipp-Anmeldeformular. Absenden = normaler POST an Klick-Tipp (kein fetch: Klick-Tipp
  // leitet danach selbst auf die Bestätigungsseite weiter). Die Seite speichert keine Eingaben.
  // Solange die Klick-Tipp-Werte Platzhalter sind, wird nichts gesendet (Vorschau).
  function formularBauen(ort) {
    const t = daten.texte;
    const k = daten.konfiguration;
    const kt = k.klicktipp || {};
    const verbunden = L.formularVerbunden(k);
    const id = (name) => "formular-" + ort + "-" + name;

    const meldung = h("p", { class: "frei-formular__meldung", role: "status", hidden: true });
    const datenschutz = istPlatzhalter(k.datenschutzUrl)
      ? document.createTextNode(t.formularDatenschutzLink)
      : h("a", { href: k.datenschutzUrl, target: "_blank", rel: "noopener", text: t.formularDatenschutzLink });

    const form = h("form", { class: "frei-formular", method: "post", action: verbunden ? kt.action : null, "accept-charset": "UTF-8" }, [
      h("div", { class: "frei-formular__felder" }, [
        h("div", { class: "frei-formular__feld" }, [
          h("label", { for: id("vorname"), text: t.formularVorname }),
          h("input", { id: id("vorname"), name: verbunden ? kt.feldVorname : "vorname", type: "text", autocomplete: "given-name", required: true, maxlength: "80" }),
        ]),
        h("div", { class: "frei-formular__feld" }, [
          h("label", { for: id("email"), text: t.formularEmail }),
          h("input", { id: id("email"), name: verbunden ? kt.feldEmail : "email", type: "email", autocomplete: "email", inputmode: "email", autocapitalize: "off", spellcheck: "false", required: true, maxlength: "200" }),
        ]),
        h("button", { type: "submit", class: "btn frei-cta frei-formular__button", text: t.formularButton }),
      ]),
      h("p", { class: "frei-formular__rechtstext" }, [document.createTextNode(t.formularRechtstext + " "), datenschutz, document.createTextNode(".")]),
      meldung,
    ]);

    if (verbunden) {
      Object.entries(kt.versteckteFelder || {}).forEach(([name, wert]) => form.appendChild(h("input", { type: "hidden", name, value: String(wert) })));
    }

    let gesendet = false;
    form.addEventListener("submit", (ev) => {
      if (!verbunden) {
        ev.preventDefault();
        meldung.textContent = t.formularNichtVerbunden;
        meldung.hidden = false;
        return;
      }
      if (gesendet) {
        ev.preventDefault(); // Doppelklick: nur einmal absenden
        return;
      }
      gesendet = true;
      // tracking.js sendet mit keepalive – das Ereignis kommt auch an, wenn die Seite gleich wechselt
      track("form_submit", { ort });
    });
    return form;
  }

  // Spamschutz von Klick-Tipp (aus dem Einbettungscode). Wird bewusst erst geladen,
  // nachdem beide Formulare im Seiteninhalt stehen – sonst findet das Skript sie nicht.
  // Fehlt es (Blocker, offline), bleiben die Formulare bedienbar.
  function schutzSkriptLaden() {
    const url = (daten.konfiguration.klicktipp || {}).schutzSkript;
    if (!url || istPlatzhalter(url) || !L.formularVerbunden(daten.konfiguration)) return;
    document.body.appendChild(h("script", { src: url, async: true }));
  }

  function mobilFormularBauen() {
    const t = daten.texte;
    const aufklapper = h("details", { class: "frei-mobil__formular" }, [
      h("summary", { text: t.mobilFormularOeffnen }),
      h("p", { text: t.mobilFormularText }),
      formularBauen("mobil"),
    ]);
    let gezaehlt = false;
    aufklapper.addEventListener("toggle", () => {
      if (!aufklapper.open || gezaehlt) return;
      gezaehlt = true;
      track("form_view", { ort: "mobil" });
    });
    document.getElementById("mobil-formular").appendChild(aufklapper);
  }

  function abschlussBauen() {
    const t = daten.texte;
    return h("section", { class: "frei-abschluss", id: "abschluss", "aria-labelledby": "abschluss-titel", tabindex: "-1", hidden: true }, [
      h("h2", { id: "abschluss-titel" }),
      h("p", { class: "frei-abschluss__ergebnis" }),
      h("p", { class: "frei-abschluss__zusatz" }),
      h("div", { class: "frei-abschluss__freischalten" }, [
        h("h3", { text: t.formularTitel }),
        h("p", { class: "frei-abschluss__formulartext", text: t.formularText }),
        formularBauen("abschluss"),
      ]),
    ]);
  }

  /* ---------------- Anzeige an den Stand anpassen ---------------- */

  function aktualisieren() {
    document.getElementById("fortschritt").textContent = L.fortschrittText(daten, stand);

    eintraege.forEach((e) => {
      const s = L.aufgabeStatus(stand, e.a.id);
      e.status.hidden = !(s.geloest || s.loesungAngezeigt);
      e.status.textContent = "";
      if (s.geloest) e.status.appendChild(h("span", { "aria-hidden": "true", text: "✓ " }));
      e.status.appendChild(document.createTextNode(s.geloest ? "Gelöst" : "Lösung angesehen"));
      e.status.classList.toggle("is-loesung", !s.geloest);
      e.loesungBtn.hidden = !L.loesungErlaubt(stand, e.a.id) || s.loesungAngezeigt;
      e.loesungBox.hidden = !s.loesungAngezeigt;
    });

    hinweisEl.hidden = !L.hinweisSichtbar(daten, stand);

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
