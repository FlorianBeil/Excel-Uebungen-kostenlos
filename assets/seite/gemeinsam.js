/* Excel.Flo – kostenlose Übungen: gemeinsame Helfer für Übersicht (index.html) und
 * Übungsseite (uebung.html). Aufbau wie in den Bonus-Übungen: Übersicht mit Fortschritt und
 * Liste, je Aufgabe eine eigene Seite, nach allen Aufgaben ein Abschluss auf der Übersicht.
 *
 * Alle Entscheidungen (was ist bearbeitet, wann erscheint der Hinweis auf die Bonus-Übungen)
 * trifft logik.js – hier nur Daten laden, Speichern, Tracking, Anmeldeformular, Zeitmessung.
 *
 * Tracking (assets/geteilt/tracking.js → Supabase-Tabelle events, portal = 'kostenlos',
 * Auswertungen in supabase/kostenlos.sql). Jedes Ereignis trägt detail.geraet
 * („desktop“/„mobil“), Aufgaben-Ereignisse zusätzlich exercise_id und detail.nr:
 *   page_view        Übersicht aufgerufen   breite, wiederkehrer, bearbeitet, utm_*, herkunft
 *   exercise_view    Übungsseite geöffnet
 *   exercise_start   erste Bedienung der Aufgabe (Tabelle/Pivot angetippt oder Taste, Prüfen)
 *   check            Prüfen                 richtig, leer, versuch
 *   exercise_solved  Aufgabe gelöst         mit_loesung
 *   solution_show    Lösung angezeigt
 *   hints_open       Tipps aufgeklappt
 *   teaser_view      Freischalt-Karte nach Aufgabe 3 gezeigt
 *   form_view        Formular gesehen       ort: hinweis | abschluss | mobil (mobil = aufgeklappt)
 *   form_submit      Formular abgesendet    ort: hinweis | abschluss | mobil
 */

(function () {
  "use strict";

  const L = window.ExcelFloKostenlos;
  const DATEN_URL = "daten/aufgaben.json";
  const TRACKING_PORTAL = "kostenlos";
  const ZEITEN_KEY = "excelflo_kostenlos_zeiten";
  const KAMPAGNE_KEY = "excelflo_kostenlos_kampagne";
  // Anmeldung über das Klick-Tipp-Formular: wird beim Absenden gemerkt (Florians Wahl 2026-10-09 –
  // ohne Prüfung, ob die E-Mail stimmt) und schaltet die Aufgaben nach dem Hinweis frei
  const ANMELDUNG_KEY = "excelflo_kostenlos_angemeldet";

  const CHECK_SVG =
    '<svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>';

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

  function el(tag, attrs, kinder) {
    const node = document.createElement(tag);
    Object.entries(attrs || {}).forEach(([k, v]) => {
      if (v === null || v === undefined || v === false) return;
      if (k === "class") node.className = v;
      else if (k === "text") node.textContent = v;
      else if (k === "html") node.innerHTML = v;
      else if (k === "hidden") node.hidden = true;
      else node.setAttribute(k, v === true ? "" : v);
    });
    (kinder || []).forEach((kind) => {
      if (kind === null || kind === undefined || kind === false) return;
      node.appendChild(typeof kind === "string" ? document.createTextNode(kind) : kind);
    });
    return node;
  }

  const istPlatzhalter = (url) => !url || /^PLATZHALTER/.test(url);

  /* ---------------- Daten und Stand ---------------- */

  function laden() {
    return fetch(DATEN_URL, { cache: "no-cache" })
      .then((res) => {
        if (!res.ok) throw new Error("HTTP " + res.status);
        return res.json();
      })
      .then((daten) => {
        const probleme = L.pruefeDaten(daten);
        if (probleme.length) console.warn("daten/aufgaben.json:", probleme);
        return { daten, stand: L.ladeStand(speicher, daten) };
      });
  }

  const speichern = (stand) => L.speichereStand(speicher, stand);
  const bearbeitet = (stand, a) => L.istBearbeitet(stand, a.id);
  const uebungUrl = (a) => "uebung.html?id=" + encodeURIComponent(a.id);

  // Nächste nicht bearbeitete Aufgabe NACH der gegebenen (sonst die erste offene von vorn)
  function naechsteNach(daten, stand, a) {
    const offen = (x) => x.id !== (a && a.id) && !bearbeitet(stand, x);
    const i = a ? daten.aufgaben.findIndex((x) => x.id === a.id) : -1;
    return daten.aufgaben.slice(i + 1).find(offen) || daten.aufgaben.find(offen) || null;
  }

  // Impressum und Datenschutz im Footer (Adressen aus der Konfiguration)
  function footerLinks(daten) {
    const t = daten.texte;
    const k = daten.konfiguration;
    [["link-impressum", k.impressumUrl, t.footerImpressum], ["link-datenschutz", k.datenschutzUrl, t.footerDatenschutz]].forEach(([id, url, text]) => {
      const link = document.getElementById(id);
      if (!link) return;
      link.textContent = text;
      link.href = url;
      link.hidden = istPlatzhalter(url);
    });
  }

  /* ---------------- Tracking ---------------- */

  // window.ExcelFloTracking erst beim Senden nachschlagen – fehlt es (Blocker, Ladefehler),
  // passiert einfach nichts. Die Übungen funktionieren immer ohne Tracking.
  function track(event, detail, aufgabe) {
    const t = window.ExcelFloTracking;
    if (!t) return;
    const d = Object.assign({ geraet }, aufgabe ? { nr: aufgabe.nr } : {}, detail || {});
    t.track(TRACKING_PORTAL, aufgabe ? aufgabe.a.id : null, event, d);
  }

  /* ---------------- Zeitmessung („Deine Zeit“ im Abschluss) ---------------- */

  // Je Aufgabe die Zeit, die ihre Seite sichtbar offen war – bis sie bearbeitet ist
  function zeiten() {
    try {
      const z = JSON.parse((speicher && speicher.getItem(ZEITEN_KEY)) || "{}");
      return z && typeof z === "object" ? z : {};
    } catch (e) {
      return {};
    }
  }

  function zeitAddieren(id, ms) {
    if (!(ms > 0) || !speicher) return;
    const z = zeiten();
    z[id] = (Number(z[id]) || 0) + ms;
    try {
      speicher.setItem(ZEITEN_KEY, JSON.stringify(z));
    } catch (e) { /* egal */ }
  }

  // Nur wenn für jede Aufgabe eine Zeit vorliegt – sonst wäre die Summe geschönt
  function gesamtZeit(daten) {
    const z = zeiten();
    if (!daten.aufgaben.every((a) => Number(z[a.id]) > 0)) return null;
    return daten.aufgaben.reduce((summe, a) => summe + Number(z[a.id]), 0);
  }

  function formatZeit(ms) {
    const sek = Math.max(1, Math.round(ms / 1000));
    const h = Math.floor(sek / 3600);
    const m = Math.floor((sek % 3600) / 60);
    const s = String(sek % 60).padStart(2, "0");
    return h ? h + ":" + String(m).padStart(2, "0") + ":" + s : m + ":" + s;
  }

  function zeitZuruecksetzen() {
    try {
      if (speicher) speicher.removeItem(ZEITEN_KEY);
    } catch (e) { /* egal */ }
  }

  /* ---------------- Anmeldestatus ---------------- */

  function angemeldet() {
    try {
      return !!(speicher && speicher.getItem(ANMELDUNG_KEY));
    } catch (e) {
      return false;
    }
  }

  function anmeldungMerken() {
    try {
      if (speicher) speicher.setItem(ANMELDUNG_KEY, new Date().toISOString());
    } catch (e) { /* blockiert – dann bleibt die Aufgabe gesperrt */ }
  }

  /* ---------------- Anmeldeformular (Klick-Tipp) ---------------- */

  // Kampagnen-Parameter der Sitzung merken: Sie stehen nur beim ersten Aufruf in der
  // Adresse, das Formular wird aber erst später (oft auf einer Übungsseite) ausgefüllt.
  function gemerkteKampagne(daten) {
    let gemerkt = {};
    try {
      gemerkt = JSON.parse(sessionStorage.getItem(KAMPAGNE_KEY) || "{}") || {};
    } catch (e) { /* z.B. sessionStorage blockiert */ }

    const kt = (daten.konfiguration || {}).klicktipp || {};
    const params = new URLSearchParams(location.search);
    let neu = false;
    Object.values(kt.kampagnenFelder || {}).forEach((parameter) => {
      const wert = params.get(parameter);
      if (wert) { gemerkt[parameter] = String(wert).slice(0, 200); neu = true; }
    });
    if (neu) {
      try { sessionStorage.setItem(KAMPAGNE_KEY, JSON.stringify(gemerkt)); } catch (e) { /* dann nur dieser Aufruf */ }
    }
    return gemerkt;
  }

  // Absenden = normaler POST an Klick-Tipp (kein fetch: Klick-Tipp leitet danach selbst auf
  // die Bestätigungsseite weiter). Die Seite speichert keine Eingaben. Solange die
  // Klick-Tipp-Werte Platzhalter sind, wird nichts gesendet (Vorschau).
  function formularBauen(daten, ort) {
    const t = daten.texte;
    const k = daten.konfiguration;
    const kt = k.klicktipp || {};
    const verbunden = L.formularVerbunden(k);
    const id = (name) => "formular-" + ort + "-" + name;

    const meldung = el("p", { class: "frei-formular__meldung", role: "status", hidden: true });
    const datenschutz = istPlatzhalter(k.datenschutzUrl)
      ? document.createTextNode(t.formularDatenschutzLink)
      : el("a", { href: k.datenschutzUrl, target: "_blank", rel: "noopener", text: t.formularDatenschutzLink });

    const form = el("form", { class: "frei-formular", method: "post", action: verbunden ? kt.action : null, "accept-charset": "UTF-8" }, [
      el("div", { class: "frei-formular__felder" }, [
        el("div", { class: "frei-formular__feld" }, [
          el("label", { for: id("vorname"), text: t.formularVorname }),
          el("input", { id: id("vorname"), name: verbunden ? kt.feldVorname : "vorname", type: "text", autocomplete: "given-name", required: true, maxlength: "80" }),
        ]),
        el("div", { class: "frei-formular__feld" }, [
          el("label", { for: id("email"), text: t.formularEmail }),
          el("input", { id: id("email"), name: verbunden ? kt.feldEmail : "email", type: "email", autocomplete: "email", inputmode: "email", autocapitalize: "off", spellcheck: "false", required: true, maxlength: "200" }),
        ]),
        el("button", { type: "submit", class: "bonus-btn frei-formular__button", text: t.formularButton }),
      ]),
      el("p", { class: "frei-formular__rechtstext" }, [t.formularRechtstext + " ", datenschutz, "."]),
      meldung,
    ]);

    if (verbunden) {
      Object.entries(kt.versteckteFelder || {}).forEach(([name, wert]) => form.appendChild(el("input", { type: "hidden", name, value: String(wert) })));
      // Kampagnenfelder: Woher kam der Kontakt? Werte stehen in der Adresse der Seite (oder gemerkt).
      Object.entries(L.kampagnenFelder(kt.kampagnenFelder, location.search, gemerkteKampagne(daten)))
        .forEach(([name, wert]) => form.appendChild(el("input", { type: "hidden", name, value: wert })));
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
      anmeldungMerken(); // vor der Weiterleitung zu Klick-Tipp – zurück auf der Seite ist Aufgabe 4 offen
      // tracking.js sendet mit keepalive – das Ereignis kommt auch an, wenn die Seite gleich wechselt
      track("form_submit", { ort });
    });
    return form;
  }

  // Freischalt-Karte für die Bonus-Übungen: nach Aufgabe 3 (Übungsseite) und im Abschluss
  function freischaltKarte(daten, stand, ort) {
    const t = daten.texte;
    return el("section", { class: "frei-hinweis", id: "hinweis-" + ort, "aria-labelledby": "hinweis-" + ort + "-titel", tabindex: "-1" }, [
      el("p", {
        class: "frei-hinweis__zaehler",
        text: L.platzhalter(t.hinweisZaehler, { bearbeitet: L.anzahlBearbeitet(daten, stand), gesamt: daten.aufgaben.length }),
      }),
      el("h2", { id: "hinweis-" + ort + "-titel", text: t.hinweisTitel }),
      el("p", { class: "frei-hinweis__text", text: t.hinweisText }),
      formularBauen(daten, ort),
    ]);
  }

  // Spamschutz von Klick-Tipp (aus dem Einbettungscode). Wird bewusst erst geladen,
  // nachdem die Formulare im Seiteninhalt stehen – sonst findet das Skript sie nicht.
  // Fehlt es (Blocker, offline), bleiben die Formulare bedienbar.
  let schutzGeladen = false;
  function schutzSkriptLaden(daten) {
    const url = (daten.konfiguration.klicktipp || {}).schutzSkript;
    if (schutzGeladen || !url || istPlatzhalter(url) || !L.formularVerbunden(daten.konfiguration)) return;
    schutzGeladen = true;
    document.body.appendChild(el("script", { src: url, async: true }));
  }

  // Zählt als „gesehen“, sobald ein nennenswerter Teil sichtbar ist
  function beiSichtbarkeit(element, callback) {
    if (!element || !("IntersectionObserver" in window)) return;
    const beobachter = new IntersectionObserver(
      (eintraege) => {
        eintraege.forEach((io) => {
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

  window.ExcelFloFrei = {
    L, CHECK_SVG, el, geraet, laden, speichern, bearbeitet, uebungUrl, naechsteNach, footerLinks,
    track, angemeldet, zeitAddieren, gesamtZeit, formatZeit, zeitZuruecksetzen,
    gemerkteKampagne, formularBauen, freischaltKarte, schutzSkriptLaden, beiSichtbarkeit,
  };
})();
