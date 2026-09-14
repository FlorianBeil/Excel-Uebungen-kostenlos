/* Excel.Flo – anonyme Nutzungs-Ereignisse
 *
 * Schreibt Ereignisse (Übung geöffnet, Prüfen, Tipps, Lösung, Zurücksetzen) in die
 * Supabase-Tabelle public.events (SQL: supabase/events.sql). Per RLS darf die
 * Seite nur einfügen, nie lesen. Keine E-Mail, keine Nutzer-ID – nur eine
 * zufällige ID pro Browser-Sitzung (sessionStorage).
 *
 * Bewusst ohne supabase-js (reines fetch) und „fire and forget“: Fehler
 * (offline, Tracker-Blocker, Tabelle fehlt) werden verschluckt und dürfen die
 * Übung nie beeinflussen.
 */

(function () {
  "use strict";

  const EVENTS_URL = "https://hbhagmmbowplzjzfvuao.supabase.co/rest/v1/events";
  const SUPABASE_KEY = "sb_publishable_Ee47HbgO5Ne8PP9Jh5ugag_iE_lZcp6"; // öffentlicher Schlüssel, Zugriff über RLS begrenzt
  const SESSION_KEY = "excelflo_tracking_session";

  let sessionId = null;

  function getSessionId() {
    if (sessionId) return sessionId;
    try {
      sessionId = sessionStorage.getItem(SESSION_KEY);
    } catch (e) {
      // sessionStorage im iframe blockiert – ID gilt dann nur für diesen Seitenaufruf
    }
    if (!sessionId) {
      sessionId =
        window.crypto && window.crypto.randomUUID
          ? window.crypto.randomUUID()
          : "s-" + Date.now() + "-" + Math.random().toString(16).slice(2);
      try {
        sessionStorage.setItem(SESSION_KEY, sessionId);
      } catch (e) {}
    }
    return sessionId;
  }

  function track(portal, exerciseId, event, detail) {
    try {
      fetch(EVENTS_URL, {
        method: "POST",
        headers: {
          apikey: SUPABASE_KEY,
          "Content-Type": "application/json",
          Prefer: "return=minimal", // keine Rückgabe – die Seite hat kein Leserecht
        },
        body: JSON.stringify({
          portal,
          exercise_id: exerciseId || null,
          event,
          detail: detail || {},
          session_id: getSessionId(),
        }),
        keepalive: true, // Ereignis kommt auch an, wenn die Seite gleich verlassen wird
      }).catch(() => {});
    } catch (e) {}
  }

  window.ExcelFloTracking = { track };
})();
