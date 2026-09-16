# Excel.Flo – kostenlose Übungen

Öffentliche Seite mit fünf frei nutzbaren Excel-Übungen für Interessenten des
Excel Master Kurses. Einziges Angebot der Seite: das kostenlose Übungsportal Light
([Excel-Uebungsportal-Light](https://github.com/FlorianBeil/Excel-Uebungsportal-Light)) – nach Aufgabe 5
(und aufklappbar im Handy-Hinweis) per Klick-Tipp-Formular (Vorname + E-Mail, Double-Opt-in),
Klick-Tipp schickt danach den Link. Webinar und Kurs werden hier bewusst nicht beworben.

**Strikt getrennt vom Käufer-Portal** ([Excel-Aufgaben](https://github.com/FlorianBeil/Excel-Aufgaben),
[Pivot](https://github.com/FlorianBeil/Excel-Aufgaben-Pivot-Tabellen-)): eigenes Repo,
eigenes Deployment. Das Portal verlinkt nie hierher, diese Seite enthält alle Werbung.

## Aufbau

| Pfad | Inhalt |
|---|---|
| `index.html` | Einstiegsseite (Titel, Meta-/Vorschau-Angaben fest im HTML) |
| `daten/aufgaben.json` | Die fünf Aufgaben, alle Texte und die Konfiguration (Impressum-, Datenschutz- und Seitenadresse, Klick-Tipp-Formular) |
| `danke.html` | Bestätigungsseite nach dem Formular („bitte E-Mail bestätigen“) – in Klick-Tipp als Weiterleitung eintragen |
| `assets/seite/` | `logik.js` (Ablauf ohne Bildschirm, getestet), `seite.js` (Anzeige, Tracking), `seite.css` |
| `pivot.html` | Pivot-Nachbau, per iframe eingebunden – **erzeugt**, nicht von Hand ändern |
| `assets/pivot/datensatz.js` | Erzeugt die 730 Beispielzeilen der Pivot-Aufgabe aus dem Bauplan in der Datendatei |
| `assets/geteilt/` | **Kopie** der geteilten Logik aus dem Portal – nicht von Hand ändern |
| `assets/og-bild.jpg` | Link-Vorschaubild (aus `skripte/og-bild.html`) |
| `supabase/kostenlos.sql` | Tracking freischalten + Auswertungen |
| `test/` | `node test/logik.test.js`, `node test/texte.test.js`, `node test/datensatz.test.js` |

## Wartung

- **Geteilte Logik aktualisieren:** Doppelklick auf `geteilt-aktualisieren.bat`. Holt den
  veröffentlichten Stand (`origin/main`) aus `../excel-flo-uebungsportal` und schreibt den
  Commit in `assets/geteilt/QUELLE.txt`. Danach testen, `?v=` in `index.html` erhöhen, pushen.
- **Pivot-Nachbau aktualisieren:** `node skripte/pivot-uebernehmen.js` (holt `origin/main` aus
  `../pivot-tabelle-prototyp`). Bricht mit Meldung ab, wenn sich das Pivot-Repo so geändert hat,
  dass eine Ersetzung nicht mehr passt.
- **Vorschaubild neu erzeugen:** `skripte/og-bild-erzeugen.ps1` (Edge + Python/Pillow).
- **Klick-Tipp-Formular verbinden:** aus dem HTML-Code des Klick-Tipp-Formulars in
  `konfiguration.klicktipp` eintragen: `action` (Formular-Adresse), `feldVorname` und `feldEmail`
  (name-Attribute der Eingabefelder), `versteckteFelder` (alle `type=hidden`-Felder, name → value).
  Solange dort Platzhalter stehen, zeigt die Seite das Formular nur als Vorschau und sendet nichts.
- **Eigene Domain:** `konfiguration.seitenUrl` in der Datendatei und `canonical`, `og:url`,
  `og:image`, `twitter:image` in `index.html` ändern – `node test/texte.test.js` meldet Abweichungen.

## Tracking

Ereignisse gehen (ohne Cookies, ohne Nutzer-ID) in die Supabase-Tabelle `events` des Portals,
Bereich `portal = 'kostenlos'`. Liste der Ereignisse: Kopf von `assets/seite/seite.js`.
`supabase/kostenlos.sql` einmal im Supabase SQL Editor ausführen – danach im Table Editor:

- **Auswertung Kostenlos Übersicht** – Besuche, Light-Hinweis gesehen, Formular gesehen/abgesendet je Gerät
- **Auswertung Kostenlos Absprung** – je Gerät und Aufgabe: gesehen, gestartet, bearbeitet, ausgestiegen
- **Auswertung Kostenlos Desktop vs Mobil** – Ausstiegsquote je Aufgabe nebeneinander

Besucher aus Kampagnen mit `?utm_source=…&utm_medium=…&utm_campaign=…` verlinken – die Werte
werden beim Seitenaufruf mitgespeichert.
