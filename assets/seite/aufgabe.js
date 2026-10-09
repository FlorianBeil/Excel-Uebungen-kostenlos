/* Excel.Flo – kostenlose Übungen: Inhalt EINER Aufgabe (Kopfzeile, Text, Tabelle bzw.
 * Pivot-Nachbau, Prüfen, Rückmeldung, Lösung, Tipps, Weiter-Bereich) zum Einsetzen in die
 * Startseite (uebersicht.js klappt die Aufgaben dort auf und zu).
 *
 * Regeln wie bisher (logik.js): Die Lösung gibt es erst nach einem Fehlversuch; gelöst oder
 * Lösung angesehen = bearbeitet. Was nach dem Bearbeiten passiert (Weiter-Knopf, Freischalt-
 * Karte), entscheidet die Startseite über ctx.bearbeitet().
 *
 * Am Handy gibt es unter der Tabelle ein Formelfeld („Deine Formel in B5“) und die Tabelle
 * passt auf die Bildschirmbreite – wie in den Bonus-Übungen.
 *
 * ExcelFloAufgabe.bauen(a, nr, ctx) → { node, weiter, beimOeffnen() }
 *   ctx.stand() / ctx.setzeStand(s)  gespeicherter Stand (logik.js)
 *   ctx.bearbeitet(a)                nach dem ersten Bearbeiten (gelöst oder Lösung angesehen)
 */

(function () {
  "use strict";

  const F = window.ExcelFloFrei;
  const { L, el } = F;
  const E = window.ExcelFlo;

  // Erste Funktion der Formel, die zählt statt rechnet → Ergebnis ohne Zahlenformat (wie engine.js)
  const OHNE_FORMAT = new Set(["ANZAHL", "ZÄHLENWENN", "ZÄHLENWENNS", "DATEDIF", "VERGLEICH", "LÄNGE", "FINDEN"]);
  // Touch-Bedienung oder wirklich schmal (Breite 0 = Tab lädt unsichtbar im Hintergrund)
  const istMobil = () => window.matchMedia("(pointer: coarse)").matches || (window.innerWidth > 0 && window.innerWidth < 600);

  function bauen(a, nr, ctx) {
    const eintrag = { a, nr, gestartet: false, versuche: 0 };
    const schritte = a.task.steps || [];

    // Bausteine einzeln, damit sie am PC in zwei Spalten stehen können (seite.css):
    // links Kopf, Buttons, Rückmeldung/Tipps – rechts die Tabelle. Am Handy untereinander.
    const bereich = el("div", { class: "bonus-arbeitsbereich frei-teil-tabelle" });
    const aktionen = el("div", { class: "bonus-aktionen frei-teil-aktionen" });
    const feedback = el("div", { class: "bonus-feedback", role: "status", "aria-live": "polite" });
    const loesungKnopf = el("button", { type: "button", class: "bonus-btn bonus-btn--rand bonus-btn--klein frei-loesung-knopf", text: "Lösung anzeigen", hidden: true });
    const loesungBox = el("div", { class: "bonus-loesung", hidden: true }, [
      el("p", { class: "bonus-label", text: "Lösung" }),
      el("p", { class: a.typ === "pivot" ? "bonus-loesung__text" : "bonus-loesung__formel", text: a.solution }),
      el("p", { text: a.explanation }),
    ]);
    const weiter = el("div", { class: "bonus-weiter" });

    const node = el("div", { class: "frei-aufgabe__inhalt" + (a.typ === "pivot" ? " is-pivot" : " is-formel") }, [
      el("div", { class: "frei-teil-kopf" }, [
        el("p", { class: "bonus-label", text: "Aufgabe " + nr + (a.funktion ? " · Thema: " + a.funktion : "") }),
        el("h2", { class: "bonus-titel frei-aufgabe__titel", id: "aufgabe-" + nr + "-titel", text: a.title }),
        el("div", { class: "bonus-aufgabe" }, [
          a.task.intro ? el("p", { text: a.task.intro }) : null,
          // Nur echte Schrittfolgen als nummerierte Liste – ein einzelner Schritt ist ein normaler Satz
          ...(schritte.length > 1 ? [el("ol", {}, schritte.map((s) => el("li", { text: s })))] : schritte.map((s) => el("p", { text: s }))),
        ]),
      ]),
      bereich,
      a.typ === "pivot" ? null : aktionen,
      el("div", { class: "frei-teil-unten" }, [feedback, loesungKnopf, loesungBox, tipps(eintrag), weiter]),
    ]);

    const gestartet = () => {
      if (eintrag.gestartet) return;
      eintrag.gestartet = true;
      F.track("exercise_start", null, eintrag);
    };

    // Ergebnis einer Prüfung verarbeiten (Formel oder Pivot)
    const pruefung = (ok) => {
      gestartet();
      let stand = ctx.stand();
      const vorher = L.aufgabeStatus(stand, a.id);
      const warBearbeitet = vorher.geloest || vorher.loesungAngezeigt;
      stand = L.pruefungErgebnis(stand, a.id, ok);
      ctx.setzeStand(stand);
      eintrag.versuche++;
      F.track("check", { richtig: ok, versuch: eintrag.versuche }, eintrag);
      const nachher = L.aufgabeStatus(stand, a.id);
      if (!vorher.geloest && nachher.geloest) F.track("exercise_solved", { mit_loesung: nachher.mitLoesung }, eintrag);
      if (ok) {
        loesungKnopf.hidden = true; // gelöst – die Erklärung steht in der Rückmeldung
        node.classList.add("is-geloest");
        erfolgZeigen(feedback, a, false);
        ctx.bearbeitet(a, warBearbeitet);
      } else {
        loesungKnopf.hidden = !L.loesungErlaubt(stand, a.id) || !loesungBox.hidden;
      }
    };

    loesungKnopf.addEventListener("click", () => {
      const warBearbeitet = F.bearbeitet(ctx.stand(), a);
      ctx.setzeStand(L.loesungAnzeigen(ctx.stand(), a.id));
      F.track("solution_show", null, eintrag);
      loesungKnopf.hidden = true;
      loesungBox.hidden = false;
      ctx.bearbeitet(a, warBearbeitet);
    });

    const teil = a.typ === "pivot" ? pivotTeil(bereich, a, feedback, pruefung, gestartet) : formelTeil(bereich, aktionen, a, feedback, pruefung, gestartet, nr);

    // Schon bearbeitet (Wiederkehrer): Erklärung zum Nachlesen gleich zeigen
    const s = L.aufgabeStatus(ctx.stand(), a.id);
    if (s.geloest) {
      node.classList.add("is-geloest");
      erfolgZeigen(feedback, a, true);
    }
    else if (s.loesungAngezeigt) loesungBox.hidden = false;

    let gesehen = false;
    return {
      node,
      weiter,
      // Aufgabe wurde aufgeklappt: Handy-Anpassungen brauchen echte Maße, Pivot lädt erst jetzt
      beimOeffnen() {
        if (teil && teil.beimOeffnen) teil.beimOeffnen();
        if (!gesehen) {
          gesehen = true;
          F.track("exercise_view", null, eintrag);
        }
      },
    };
  }

  /* ---------------- Formel-Aufgabe ---------------- */

  function formelTeil(bereich, aktionen, a, feedback, pruefung, gestartet, nr) {
    const sheet = E.createSheet(a.grid);
    sheet.node.setAttribute("role", "group");
    sheet.node.setAttribute("aria-label", "Tabelle. Zellen mit den Pfeiltasten wählen, zum Bearbeiten tippen oder Enter drücken.");
    sheet.node.addEventListener("pointerdown", gestartet);
    sheet.node.addEventListener("keydown", gestartet);
    bereich.appendChild(sheet.node);

    const refs = Object.keys(sheet.inputEntries);
    const feld = istMobil() && refs.length === 1 ? formelFeld(bereich, a, sheet, refs[0], gestartet) : null;

    const pruefen = el("button", { type: "button", class: "bonus-btn", text: "Prüfen" });
    const zuruecksetzen = el("button", { type: "button", class: "bonus-btn bonus-btn--rand", text: "Zurücksetzen" });
    aktionen.append(pruefen, zuruecksetzen);

    zuruecksetzen.addEventListener("click", () => {
      sheet.reset();
      refs.forEach((r) => sheet.inputEntries[r].td.classList.remove("is-correct", "is-wrong"));
      if (feld) feld.value = "";
      leeren(feedback);
    });

    pruefen.addEventListener("click", () => {
      let beantwortet = 0;
      let richtig = 0;
      refs.forEach((r) => {
        const e = sheet.inputEntries[r];
        const ergebnis = E.checkCell(e.raw, e.answer, sheet.getCellValue);
        e.td.classList.remove("is-correct", "is-wrong");
        if (ergebnis === true) { e.td.classList.add("is-correct"); beantwortet++; richtig++; }
        else if (ergebnis === false) { e.td.classList.add("is-wrong"); beantwortet++; }
      });

      if (!beantwortet) {
        // zählt nicht als Fehlversuch – es wurde ja noch nichts eingegeben
        gestartet();
        F.track("check", { richtig: false, leer: true }, { a, nr });
        const meldung = "Trag zuerst eine Formel in die gelb markierte Zelle " + refs[0] + " ein.";
        fehlerZeigen(feedback, meldung);
        E.showErrorPopup(sheet.node, meldung);
        return;
      }
      const ok = richtig === refs.length;
      if (ok) {
        E.showSuccessPopup(sheet.node);
      } else {
        const meldung = refs.length === 1 ? "Das Ergebnis stimmt noch nicht." : richtig + " von " + refs.length + " Feldern stimmen.";
        E.showErrorPopup(sheet.node, meldung);
        fehlerZeigen(feedback, meldung + " Schau dir die Tipps an – oder lass dir die Lösung anzeigen.");
      }
      pruefung(ok);
    });

    let angepasst = false;
    return {
      beimOeffnen() {
        if (angepasst) return;
        angepasst = true;
        if (istMobil()) handySpalten(sheet, a.grid.cols.length);
        zielzelleZeigen(sheet.inputEntries[refs[0]].td);
      },
    };
  }

  // Handy: leere Zusatzspalten der Engine ausblenden, Datenspalten auf die Bildschirmbreite bringen
  const ZEILENKOPF_HANDY = 34;
  const SPALTE_MIN_HANDY = 64;

  function handySpalten(sheet, anzahlDaten) {
    setTimeout(() => {
      const table = sheet.node.querySelector("table.sheet");
      const scroller = sheet.node.querySelector(".sheet-scroll");
      if (!table || !scroller || !scroller.clientWidth || table.offsetWidth <= scroller.clientWidth) return;
      const cols = [...table.querySelectorAll("colgroup col")];
      const spalten = cols.slice(1, 1 + anzahlDaten);
      cols.slice(1 + anzahlDaten).forEach((c) => (c.style.visibility = "collapse"));
      cols[0].style.width = ZEILENKOPF_HANDY + "px";
      const breiten = spalten.map((c) => parseFloat(c.style.width) || 110);
      const faktor = Math.min(1, (scroller.clientWidth - ZEILENKOPF_HANDY) / breiten.reduce((x, y) => x + y, 0));
      let summe = ZEILENKOPF_HANDY;
      spalten.forEach((c, i) => {
        const w = Math.max(SPALTE_MIN_HANDY, Math.floor(breiten[i] * faktor));
        c.style.width = w + "px";
        summe += w;
      });
      table.style.minWidth = "0";
      table.style.width = summe + "px";
    }, 0);
  }

  // Tabelle seitlich so weit scrollen, dass die gelbe Zielzelle sichtbar ist (nur falls verdeckt)
  function zielzelleZeigen(td) {
    setTimeout(() => {
      const scroller = td.closest(".sheet-scroll");
      if (!scroller || scroller.scrollWidth <= scroller.clientWidth) return;
      const verdeckt = td.getBoundingClientRect().right - scroller.getBoundingClientRect().right;
      if (verdeckt > 1) scroller.scrollLeft += verdeckt + 16;
    }, 0);
  }

  // Formelfeld fürs Handy: schreibt in die Eingabezelle (entry.raw – daraus liest auch die Prüfung)
  function formelFeld(bereich, a, sheet, ref, gestartet) {
    const entry = sheet.inputEntries[ref];
    const id = "formel-" + a.id;
    const feld = el("input", {
      id, type: "text", class: "bonus-formelfeld__eingabe", placeholder: "=",
      autocomplete: "off", autocapitalize: "off", autocorrect: "off", spellcheck: "false", enterkeyhint: "done",
    });
    bereich.appendChild(el("div", { class: "bonus-formelfeld" }, [el("label", { for: id, text: "Deine Formel in " + ref }), feld]));
    feld.addEventListener("input", () => {
      gestartet();
      entry.raw = feld.value.trim();
      entry.td.classList.remove("is-correct", "is-wrong");
      ergebnisAnzeigen(a, sheet, entry);
    });
    feld.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter") feld.blur();
    });
    sheet.node.addEventListener("focusout", () => setTimeout(() => {
      if (document.activeElement !== feld) feld.value = entry.raw;
    }, 0));
    return feld;
  }

  function ergebnisAnzeigen(a, sheet, entry) {
    const raw = entry.raw;
    let text = raw;
    let zahl = false;
    if (raw.startsWith("=") && window.ExcelFloFormula) {
      const Fo = window.ExcelFloFormula;
      const ergebnis = Fo.evaluate(raw, sheet.getCellValue);
      if (Fo.isFormulaError(ergebnis)) {
        const code = String(ergebnis.message || "").match(/#[A-ZÄÖÜ0-9\/!?]+/);
        text = code ? ({ "#N/A": "#NV", "#VALUE!": "#WERT!" }[code[0]] || code[0]) : "#WERT!";
      } else if (typeof ergebnis === "number") {
        text = zahlFormatieren(ergebnis, entry.format || formatAusFormel(a, raw));
        zahl = true;
      } else if (typeof ergebnis === "boolean") {
        text = ergebnis ? "WAHR" : "FALSCH";
      } else {
        text = ergebnis === undefined || ergebnis === null ? "0" : String(ergebnis);
      }
    }
    entry.el.textContent = text;
    entry.td.classList.toggle("cell--num", zahl);
  }

  function formatAusFormel(a, raw) {
    const funktion = (raw.match(/^=\s*([A-ZÄÖÜa-zäöü.]+)\s*\(/) || [])[1];
    if (funktion && OHNE_FORMAT.has(funktion.toUpperCase())) return null;
    const bezug = (raw.match(/\$?[A-Za-z]{1,3}\$?\d+/) || [])[0];
    const def = bezug && a.grid.cells[bezug.replace(/\$/g, "").toUpperCase()];
    if (def && def.format) return def.format;
    const mitFormat = Object.values(a.grid.cells).find((c) => c.format);
    return /SVERWEIS|MAX|MIN|SUMME|MITTELWERT/i.test(funktion || "") && mitFormat ? mitFormat.format : null;
  }

  function zahlFormatieren(n, format) {
    if (format === "currency0" && !Number.isInteger(Math.round(n * 100) / 100)) format = "currency";
    if (format === "currency") return n.toLocaleString("de-DE", { style: "currency", currency: "EUR" });
    if (format === "currency0") return n.toLocaleString("de-DE", { style: "currency", currency: "EUR", maximumFractionDigits: 0 });
    const gerundet = Number.isInteger(n) ? n : parseFloat(n.toPrecision(10));
    return String(gerundet).replace(".", ",");
  }

  /* ---------------- Pivot-Aufgabe ---------------- */

  // Das iframe entsteht erst beim Aufklappen – eine gesperrte Pivot-Aufgabe lädt so gar nicht erst
  function pivotTeil(bereich, a, feedback, pruefung, gestartet) {
    let iframe = null;

    window.addEventListener("message", (ev) => {
      if (!iframe || ev.origin !== location.origin || ev.source !== iframe.contentWindow) return;
      const m = ev.data;
      if (!m || m.quelle !== "excelflo-pivot" || m.typ !== "pruefung") return;
      const ok = !!(m.daten && m.daten.richtig);
      // Die Fehlermeldung zeigt der Pivot-Nachbau selbst; hier nur der Erfolg
      if (!ok) leeren(feedback);
      pruefung(ok);
    });

    return {
      beimOeffnen() {
        if (iframe) return;
        iframe = el("iframe", { class: "bonus-pivot", src: "pivot.html", title: "Pivot-Tabelle: " + a.title });
        iframe.addEventListener("load", () => {
          try {
            const doc = iframe.contentDocument;
            // Höhe folgt dem Inhalt (gleiche Herkunft) – keine zweite Scrollleiste
            const anpassen = () => (iframe.style.height = Math.ceil(doc.body.getBoundingClientRect().height) + "px");
            anpassen();
            new ResizeObserver(anpassen).observe(doc.body);
            doc.addEventListener("pointerdown", gestartet);
            doc.addEventListener("keydown", gestartet);
          } catch (e) {
            // feste Höhe aus portal.css bleibt
          }
        });
        bereich.appendChild(iframe);
      },
    };
  }

  /* ---------------- Tipps, Rückmeldung ---------------- */

  function tipps(eintrag) {
    const hinweise = eintrag.a.hints || [];
    if (!hinweise.length) return null;
    const details = el("details", { class: "bonus-tipps" }, [
      el("summary", { text: "Tipps anzeigen" }),
      el("ol", {}, hinweise.map((h) => el("li", { text: h }))),
    ]);
    details.addEventListener("toggle", () => {
      if (details.open) F.track("hints_open", null, eintrag);
    });
    return details;
  }

  function leeren(feedback) {
    feedback.className = "bonus-feedback";
    feedback.textContent = "";
  }

  function fehlerZeigen(feedback, meldung) {
    leeren(feedback);
    feedback.classList.add("is-error");
    feedback.appendChild(el("p", { text: meldung }));
  }

  function erfolgZeigen(feedback, a, wiederhergestellt) {
    leeren(feedback);
    feedback.classList.add("is-success");
    // Einziges Erfolgssignal der Aufgabe: die grüne Box (kein Häkchen-Icon daneben)
    feedback.appendChild(el("div", {}, [
      el("p", { class: "bonus-feedback__titel", text: wiederhergestellt ? "Schon gelöst" : "Richtig!" }),
      el("p", { text: a.erfolgTipp || a.explanation }),
    ]));
  }

  window.ExcelFloAufgabe = { bauen };
})();
