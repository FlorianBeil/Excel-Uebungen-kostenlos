# Excel.Flo – kostenlose Übungen

Öffentliche Seite mit fünf frei nutzbaren Excel-Übungen für Interessenten des
Excel Master Kurses. Endet mit dem Hinweis auf das kostenlose Monats-Webinar.

**Strikt getrennt vom Käufer-Portal** ([Excel-Aufgaben](https://github.com/FlorianBeil/Excel-Aufgaben),
[Pivot](https://github.com/FlorianBeil/Excel-Aufgaben-Pivot-Tabellen-)): eigenes Repo,
eigenes Deployment. Das Portal verlinkt nie hierher, diese Seite enthält alle Werbung.

## Aufbau

| Pfad | Inhalt |
|---|---|
| `index.html` | Einstiegsseite |
| `daten/aufgaben.json` | Die fünf Aufgaben + Konfiguration (Webinar-/Kurs-URL). Einzige Stelle für Links. |
| `assets/seite/` | Ablauf, Fortschritt, Angebotsblock, Tracking-Ereignisse dieser Seite |
| `assets/pivot/` | Aus dem Pivot-Repo übernommener Pivot-Nachbau (ohne Kurs-Übungen) |
| `assets/geteilt/` | **Kopie** der geteilten Logik aus dem Portal – nicht von Hand ändern |
| `supabase/` | SQL für Tracking-Bereich und Auswertung |

## Geteilte Logik aktualisieren

Doppelklick auf `geteilt-aktualisieren.bat`. Holt den veröffentlichten Stand
(`origin/main`) aus dem lokalen Portal-Repo `../excel-flo-uebungsportal` und
schreibt den Commit in `assets/geteilt/QUELLE.txt`. Danach testen, dann pushen.

Der Pivot-Nachbau wird nicht per Skript aktualisiert (im Pivot-Repo liegt er
zusammen mit den Kurs-Übungen in einer Datei) – Änderungen dort bei Bedarf von
Hand übernehmen.
