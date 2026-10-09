/* Excel.Flo – kostenlose Übungen: Startseite (index.html) mit direktem Einstieg
 *
 * Oben die Kopfkarte (Überschrift, Einleitung, „1 von 4 gelöst“ + Balken, kein Button),
 * darunter die Aufgaben als Liste – die aktuelle ist sofort aufgeklappt (Inhalt aus aufgabe.js):
 *  - bearbeitet:   schmale Zeile „✓ Gelöst“, per Klick aufklappbar (Erklärung „Schon gelöst“)
 *  - freigeschaltet, noch offen: aufgeklappt (bzw. per Klick aufklappbar)
 * Es ist immer genau eine Aufgabe offen; die Zeile gibt es nur im zugeklappten Zustand.
 *  - gesperrt:     ausgegraute Zeile mit Schloss, nicht klickbar. Freischaltung der Reihe nach,
 *                  die Aufgaben nach dem Hinweis erst nach der Anmeldung (logik.js: freigeschaltet).
 * Nach Aufgabe 3 (konfiguration.hinweisNachAufgabe) steht die Freischalt-Karte mit dem
 * Klick-Tipp-Formular, solange niemand angemeldet ist. „Weiter zu Aufgabe X →“ klappt die
 * gelöste Aufgabe zu, die nächste auf und scrollt sanft zu ihr.
 *
 * Die Sperre ist reine Anzeige im Browser (kein Login, kein Server) – wer den gespeicherten
 * Fortschritt mit den Entwicklerwerkzeugen ändert, kommt durch. Für Besucher reicht das.
 *
 * Überschrift und Einleitung stehen fest in index.html (Suchmaschinen, test/texte.test.js);
 * sind alle Aufgaben bearbeitet, ersetzt dieses Skript sie durch den Abschluss.
 */

(function () {
  "use strict";

  const F = window.ExcelFloFrei;
  const { L, el } = F;
  const reduzierteBewegung = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  const SCHLOSS_SVG =
    '<svg viewBox="0 0 24 24" width="15" height="15" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><rect x="4.5" y="10.5" width="15" height="10" rx="2"/><path d="M8 10.5V7.5a4 4 0 0 1 8 0v3"/></svg>';

  let daten = null;
  let stand = null;
  let angemeldet = false;
  // Anzeige (nicht gespeichert): die eine aufgeklappte Aufgabe
  let fokusIdx = null;
  const karten = [];
  let ctaEl = null;
  let ctaGemeldet = false;
  let zuruecksetzenEl = null;

  function start() {
    const root = document.getElementById("uebersicht");
    F.laden()
      .then((d) => {
        daten = d.daten;
        stand = d.stand;
        angemeldet = F.angemeldet();
        F.footerLinks(daten);
        F.gemerkteKampagne(daten); // utm-Parameter für das Formular merken
        mobilFormularBauen();
        const herkunft = L.kampagne(location.search, document.referrer, location.host);
        F.track("page_view", Object.assign({
          breite: window.innerWidth,
          wiederkehrer: Object.keys(stand.aufgaben).length > 0,
          bearbeitet: L.anzahlBearbeitet(daten, stand),
        }, herkunft));
        aufbauen(root);
        F.schutzSkriptLaden(daten);
      })
      .catch((err) => {
        console.error(err);
        root.textContent = "";
        root.appendChild(el("p", { class: "frei-laden", text: "Die Aufgaben konnten nicht geladen werden. Bitte lade die Seite neu." }));
      });
  }

  function aufbauen(root) {
    root.textContent = "";
    fokusIdx = L.offeneAufgabe(daten, stand, angemeldet);

    root.appendChild(el("h2", { class: "bonus-abschnitt", text: "Deine Aufgaben" }));
    const liste = el("ol", { class: "frei-liste" });
    root.appendChild(liste);

    daten.aufgaben.forEach((a, i) => {
      liste.appendChild(karteBauen(a, i));
      if (i + 1 === daten.konfiguration.hinweisNachAufgabe) {
        ctaEl = el("li", { class: "frei-liste__cta" }, [F.freischaltKarte(daten, stand, "hinweis")]);
        liste.appendChild(ctaEl);
      }
    });
    zuruecksetzenEl = zuruecksetzen(root);
    root.appendChild(zuruecksetzenEl);
    aktualisieren();

    // Wiederkehrer landen direkt bei der ersten offenen Aufgabe (bzw. bei der Freischalt-Karte)
    if (L.anzahlBearbeitet(daten, stand) > 0) {
      const ziel = fokusIdx !== null ? karten[fokusIdx].li : !ctaEl.hidden ? ctaEl : null;
      if (ziel) setTimeout(() => ziel.scrollIntoView({ block: "start" }), 50);
    }
  }

  /* ---------------- Eine Aufgabe in der Liste ---------------- */

  function karteBauen(a, i) {
    const nr = i + 1;
    const aufgabe = window.ExcelFloAufgabe.bauen(a, nr, {
      stand: () => stand,
      setzeStand: (s) => {
        stand = s;
        F.speichern(stand);
      },
      bearbeitet: (x, warBearbeitet) => {
        if (!warBearbeitet) zeit.bearbeitet(i);
        aktualisieren();
      },
    });

    const kreis = el("span", { class: "bonus-nr", "aria-hidden": "true" });
    const status = el("span", { class: "frei-zeile__status" });
    const zeile = el("button", { type: "button", class: "frei-zeile", "aria-controls": "aufgabe-" + nr + "-inhalt" }, [
      kreis,
      el("span", { class: "frei-zeile__titel" }, [el("span", { class: "frei-zeile__nr", text: "Aufgabe " + nr + " · " }), a.title]),
      status,
    ]);
    aufgabe.node.id = "aufgabe-" + nr + "-inhalt";
    const li = el("li", { class: "frei-karte", id: "aufgabe-" + nr, tabindex: "-1", "aria-labelledby": "aufgabe-" + nr + "-titel" }, [zeile, aufgabe.node]);

    // Immer genau eine Aufgabe offen: Klick auf eine Zeile öffnet sie, die bisher offene klappt zu
    // (gelöste Aufgaben zum Nachlesen; „Weiter zu Aufgabe X →“ führt zurück)
    zeile.addEventListener("click", () => {
      if (zeile.disabled) return;
      fokusIdx = i;
      aktualisieren();
    });

    karten[i] = { a, li, zeile, kreis, status, aufgabe };
    return li;
  }

  /* ---------------- Anzeige an den Stand anpassen ---------------- */

  function aktualisieren() {
    const gesamt = daten.aufgaben.length;
    const n = daten.konfiguration.hinweisNachAufgabe;

    karten.forEach((k, i) => {
      const s = L.aufgabeStatus(stand, k.a.id);
      const bearbeitet = s.geloest || s.loesungAngezeigt;
      const frei = bearbeitet || L.freigeschaltet(daten, stand, i, angemeldet);
      const fokus = i === fokusIdx;
      const offen = frei && fokus;

      k.li.classList.toggle("is-offen", offen);
      k.li.classList.toggle("is-aktiv", offen && fokus);
      k.li.classList.toggle("is-erledigt", bearbeitet);
      k.li.classList.toggle("is-gesperrt", !frei);
      k.aufgabe.node.hidden = !offen;
      // Die Zeile gibt es nur zugeklappt – aufgeklappt beginnt die Aufgabe direkt mit „AUFGABE X · THEMA“
      k.zeile.hidden = offen;
      k.zeile.disabled = !frei;
      k.zeile.setAttribute("aria-expanded", offen ? "true" : "false");
      k.kreis.innerHTML = bearbeitet ? F.CHECK_SVG : String(i + 1);

      k.status.textContent = "";
      if (s.geloest) k.status.append("✓ Gelöst");
      else if (s.loesungAngezeigt) k.status.append("Lösung angesehen");
      else if (!frei) {
        k.status.insertAdjacentHTML("beforeend", SCHLOSS_SVG);
        if (i >= n) k.status.appendChild(el("span", { class: "frei-zeile__sperrtext", text: daten.texte.gesperrtAnmeldung }));
      }

      if (offen) k.aufgabe.beimOeffnen();
      weiterBereich(k, i, bearbeitet);
    });

    // Freischalt-Karte: sobald die ersten Aufgaben bearbeitet sind – bis zur Anmeldung
    const ctaSichtbar = L.hinweisSichtbar(daten, stand) && !angemeldet;
    ctaEl.hidden = !ctaSichtbar;
    ctaEl.querySelector(".frei-hinweis__zaehler").textContent = L.platzhalter(daten.texte.hinweisZaehler, { bearbeitet: L.anzahlBearbeitet(daten, stand), gesamt });
    if (ctaSichtbar && !ctaGemeldet) {
      ctaGemeldet = true;
      F.track("teaser_view");
      F.beiSichtbarkeit(ctaEl.querySelector(".frei-formular"), () => F.track("form_view", { ort: "hinweis" }));
    }

    zuruecksetzenEl.hidden = !L.anzahlBearbeitet(daten, stand);
    if (L.anzahlBearbeitet(daten, stand) === gesamt) kopfFertig();
    else kopfUnterwegs();
    zeit.fokus(fokusIdx);
  }

  // Nach dem Bearbeiten: „Weiter zu Aufgabe X →“ zur Aufgabe, an der man gerade dran ist (erste
  // offene, freigeschaltete) – beim Nachlesen einer früheren also zurück dorthin. Gibt es keine
  // (die nächste wartet auf die Anmeldung, oder alles ist bearbeitet), kein Knopf: Darunter steht
  // dann die Freischalt-Karte bzw. der Abschluss oben.
  function weiterZiel(i) {
    const offen = L.offeneAufgabe(daten, stand, angemeldet);
    if (offen !== null && offen !== i) return offen;
    return null;
  }

  function weiterBereich(k, i, bearbeitet) {
    const w = k.aufgabe.weiter;
    w.textContent = "";
    const ziel = bearbeitet ? weiterZiel(i) : null;
    if (ziel === null) return;
    const knopf = el("button", { type: "button", class: "bonus-btn bonus-btn--gross bonus-btn--voll", text: "Weiter zu Aufgabe " + (ziel + 1) + " →" });
    knopf.addEventListener("click", () => weiter(ziel));
    w.appendChild(knopf);
  }

  function weiter(ziel) {
    fokusIdx = ziel;
    aktualisieren();
    springeZu(karten[ziel].li);
  }

  function springeZu(ziel) {
    ziel.scrollIntoView({ behavior: reduzierteBewegung ? "auto" : "smooth", block: "start" });
    ziel.focus({ preventScroll: true });
    // Absicherung: Bricht das sanfte Scrollen ab (Seitenhöhe ändert sich gerade durch das
    // Zuklappen, oder der Browser kann es nicht), direkt hinspringen
    setTimeout(() => {
      if (Math.abs(ziel.getBoundingClientRect().top) > 40) ziel.scrollIntoView({ block: "start" });
    }, 900);
  }

  /* ---------------- Kopfkarte ---------------- */

  // Wie in den Bonus-Übungen: Der Balken startet beim zuletzt gezeigten Stand und läuft sanft zum neuen
  const BALKEN_KEY = "excelflo_kostenlos_balken";
  let balkenEl = null;

  function kopfUnterwegs() {
    const gesamt = daten.aufgaben.length;
    const fertig = L.anzahlBearbeitet(daten, stand);
    const ziel = (100 * fertig) / gesamt;
    const ort = document.getElementById("kopf-fortschritt");

    if (!balkenEl) {
      let vorher = 0;
      try {
        const gemerkt = parseFloat(sessionStorage.getItem(BALKEN_KEY));
        if (isFinite(gemerkt)) vorher = gemerkt;
      } catch (e) { /* ohne Animation */ }
      balkenEl = el("div", { class: "bonus-balken", role: "progressbar", "aria-valuemin": "0" }, [
        el("span", { class: "bonus-balken__fuellung", style: "width:" + vorher + "%" }),
      ]);
      ort.replaceChildren(el("div", { class: "bonus-fortschritt" }, [el("div", { class: "bonus-fortschritt__zeile" }), balkenEl]));
    }
    ort.querySelector(".bonus-fortschritt__zeile").replaceChildren(el("span", { text: L.fortschrittText(daten, stand) }));
    balkenEl.setAttribute("aria-valuemax", String(gesamt));
    balkenEl.setAttribute("aria-valuenow", String(fertig));
    balkenEl.setAttribute("aria-label", L.fortschrittText(daten, stand));
    setTimeout(() => (balkenEl.firstChild.style.width = ziel + "%"), 60);
    try {
      sessionStorage.setItem(BALKEN_KEY, String(ziel));
    } catch (e) { /* egal */ }
  }

  function kopfFertig() {
    const t = daten.texte;
    const gesamt = daten.aufgaben.length;
    const erste = daten.aufgaben[0].funktion || daten.aufgaben[0].title;
    const letzte = daten.aufgaben[gesamt - 1].funktion || daten.aufgaben[gesamt - 1].title;
    balkenEl = null;
    document.getElementById("kopf").classList.add("bonus-kopfkarte--fertig");
    document.getElementById("seiten-titel").textContent = L.platzhalter(t.fertigTitel, { gesamt });
    document.querySelector(".frei-start__text").textContent = L.platzhalter(t.fertigText, { erste, letzte });
    document.getElementById("kopf-fortschritt").replaceChildren(
      el("div", { class: "bonus-segmente bonus-segmente--fertig", "aria-hidden": "true" }, daten.aufgaben.map(() => el("span")))
    );
    const z = F.gesamtZeit(daten);
    document.getElementById("kopf-aktion").replaceChildren(
      z
        ? el("div", { class: "bonus-zeit" }, [
            el("p", { class: "bonus-label", text: "Deine Zeit" }),
            el("p", { class: "bonus-zeit__wert", text: F.formatZeit(z) + " Min." }),
            el("p", { class: "bonus-zeit__unter", text: "für alle " + gesamt + " Aufgaben" }),
          ])
        : el("span")
    );
  }

  /* ---------------- Zeitmessung („Deine Zeit“ im Abschluss) ---------------- */

  // Gezählt wird die sichtbare Zeit, in der eine noch nicht bearbeitete Aufgabe die offene ist
  const zeit = (function () {
    let idx = null;
    let seit = null;
    const laeuft = () => idx !== null && !L.istBearbeitet(stand, daten.aufgaben[idx].id);
    const sichern = () => {
      if (idx !== null && seit !== null) F.zeitAddieren(daten.aufgaben[idx].id, Date.now() - seit);
      seit = null;
    };
    const starten = () => {
      seit = laeuft() && document.visibilityState === "visible" ? Date.now() : null;
    };
    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "visible") starten();
      else sichern();
    });
    window.addEventListener("pagehide", sichern);
    return {
      fokus(neu) {
        if (neu === idx) {
          if (seit === null) starten();
          return;
        }
        sichern();
        idx = neu;
        starten();
      },
      // vor dem nächsten aktualisieren(): Zeit dieser Aufgabe abschließen
      bearbeitet(i) {
        if (i !== idx || seit === null) return;
        F.zeitAddieren(daten.aufgaben[i].id, Date.now() - seit);
        seit = null;
      },
    };
  })();

  /* ---------------- Fortschritt zurücksetzen ---------------- */

  function zuruecksetzen(root) {
    const zeileEl = el("p", { class: "bonus-zuruecksetzen__zeile" });
    const btn = el("button", { type: "button", class: "bonus-zuruecksetzen", text: "↺ Fortschritt zurücksetzen" });
    btn.addEventListener("click", () => {
      if (!window.confirm("Fortschritt aller " + daten.aufgaben.length + " Aufgaben wirklich zurücksetzen? Das kann nicht rückgängig gemacht werden.")) return;
      stand = L.neuerStand();
      F.speichern(stand);
      F.zeitZuruecksetzen();
      // Die Anmeldung bleibt bestehen – sie gilt für die Person, nicht für den Durchgang
      location.reload();
    });
    zeileEl.appendChild(btn);
    return zeileEl;
  }

  /* ---------------- Hinweis für Smartphones ---------------- */

  function mobilFormularBauen() {
    const t = daten.texte;
    const ziel = document.getElementById("mobil-formular");
    if (!ziel) return;
    const aufklapper = el("details", { class: "frei-mobil__formular" }, [
      el("summary", { text: t.mobilFormularOeffnen }),
      el("p", { text: t.mobilFormularText }),
      F.formularBauen(daten, "mobil"),
    ]);
    let gezaehlt = false;
    aufklapper.addEventListener("toggle", () => {
      if (!aufklapper.open || gezaehlt) return;
      gezaehlt = true;
      F.track("form_view", { ort: "mobil" });
    });
    ziel.appendChild(aufklapper);
  }

  document.addEventListener("DOMContentLoaded", start);
})();
