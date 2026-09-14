-- Excel.Flo – kostenlose Übungen: Tracking freischalten und Auswertungen anlegen
--
-- Voraussetzung: Die Tabelle public.events existiert bereits (Übungsportal, supabase/events.sql).
-- Einmalig im Supabase-Dashboard ausführen:
--   SQL Editor → New query → dieses Skript einfügen → Run
-- Das Skript kann gefahrlos mehrfach ausgeführt werden.
--
-- Ereignisse der kostenlosen Seite (portal = 'kostenlos'), jedes mit detail.geraet = 'desktop' | 'mobil':
--   page_view        Seite aufgerufen                  detail: breite, wiederkehrer, bearbeitet, utm_source/medium/campaign, herkunft
--   exercise_view    Aufgabe ins Bild gescrollt        exercise_id = Aufgabe, detail.nr = 1–5
--   exercise_start   Aufgabe erstmals bedient          exercise_id, nr
--   check            Prüfen                            exercise_id, nr, richtig, leer, versuch
--   exercise_solved  Aufgabe gelöst                    exercise_id, nr, mit_loesung
--   solution_show    Lösung angezeigt                  exercise_id, nr
--   hints_open       Tipps aufgeklappt                 exercise_id, nr
--   offer_view       Angebotsblock nach Aufgabe 4 gesehen
--   webinar_click    Klick zum Webinar                 detail.ort = angebot | abschluss | mobil_hinweis
--   course_click     Klick auf „Oder direkt zum Master Kurs“
--   mail_link_click  Klick auf „Link an mich selbst schicken“
--
-- Hinweis: Die bestehenden Übersichten „Auswertung Übungsportal“ und „Auswertung Einzelklicks“
-- zeigen alle Bereiche – Zeilen dieser Seite erkennst du dort an Portal = kostenlos.

-- 1. Neuen Bereich erlauben (bisher: funktionen, pivot, powerquery, einstufungstest)
alter table public.events drop constraint if exists events_portal_check;
alter table public.events add constraint events_portal_check
  check (portal in ('funktionen', 'pivot', 'powerquery', 'einstufungstest', 'kostenlos'));

-- Alte Fassungen entfernen (Spaltennamen können sich ändern; abhängige View zuerst)
drop view if exists public."Auswertung Kostenlos Desktop vs Mobil";
drop view if exists public."Auswertung Kostenlos Absprung";
drop view if exists public."Auswertung Kostenlos Übersicht";

-- 2. Übersicht: eine Zeile pro Gerät plus Gesamt
create view public."Auswertung Kostenlos Übersicht"
with (security_invoker = true) as
with sitzungen as (
  select
    session_id,
    coalesce(
      (array_agg(detail->>'geraet' order by created_at) filter (where event = 'page_view'))[1],
      (array_agg(detail->>'geraet' order by created_at))[1]
    )                                                                                      as geraet,
    bool_or(event = 'page_view')                                                           as besuch,
    bool_or(event = 'page_view' and detail->>'wiederkehrer' = 'true')                      as wiederkehrer,
    bool_or(event = 'exercise_start')                                                      as irgendeine_gestartet,
    count(distinct detail->>'nr') filter (where event in ('exercise_solved', 'solution_show')) as bearbeitet,
    bool_or(event = 'offer_view')                                                          as angebot,
    bool_or(event = 'webinar_click')                                                       as webinar,
    bool_or(event = 'webinar_click' and detail->>'ort' = 'angebot')                        as webinar_angebot,
    bool_or(event = 'webinar_click' and detail->>'ort' = 'abschluss')                      as webinar_abschluss,
    bool_or(event = 'webinar_click' and detail->>'ort' = 'mobil_hinweis')                  as webinar_mobil,
    bool_or(event = 'course_click')                                                        as kurs,
    bool_or(event = 'mail_link_click')                                                     as mail,
    max(created_at)                                                                        as zuletzt
  from public.events
  where portal = 'kostenlos'
  group by session_id
)
select
  coalesce(geraet, 'Gesamt')                                                               as "Gerät",
  count(*) filter (where besuch)                                                           as "Besuche",
  count(*) filter (where wiederkehrer)                                                     as "davon Wiederkehrer",
  count(*) filter (where besuch and not irgendeine_gestartet)                              as "Keine Aufgabe angefangen",
  round(100.0 * count(*) filter (where besuch and not irgendeine_gestartet)
        / nullif(count(*) filter (where besuch), 0), 1)                                    as "Keine Aufgabe angefangen in %",
  count(*) filter (where bearbeitet >= 5)                                                  as "Alle 5 bearbeitet",
  count(*) filter (where angebot)                                                          as "Angebot gesehen",
  count(*) filter (where webinar)                                                          as "Webinar-Klick",
  round(100.0 * count(*) filter (where webinar)
        / nullif(count(*) filter (where besuch), 0), 1)                                    as "Webinar-Klick in % der Besuche",
  round(100.0 * count(*) filter (where webinar_angebot)
        / nullif(count(*) filter (where angebot), 0), 1)                                   as "Webinar-Klick in % Angebot gesehen",
  count(*) filter (where webinar_angebot)                                                  as "Webinar aus Angebot",
  count(*) filter (where webinar_abschluss)                                                as "Webinar aus Abschluss",
  count(*) filter (where webinar_mobil)                                                    as "Webinar aus Handy-Hinweis",
  count(*) filter (where kurs)                                                             as "Kurs-Klick",
  count(*) filter (where mail)                                                             as "Link an mich geschickt",
  max(zuletzt) at time zone 'Europe/Berlin'                                                as "Zuletzt"
from sitzungen
where geraet in ('desktop', 'mobil')
group by rollup (geraet)
order by grouping(geraet), geraet;

revoke all on public."Auswertung Kostenlos Übersicht" from anon, authenticated;

-- 3. Absprung je Aufgabe: eine Zeile pro Gerät und Aufgabe
--    Gestartet      = Aufgabe mindestens einmal bedient
--    Bearbeitet     = gelöst oder Lösung angesehen
--    Ausgestiegen   = das war die letzte Aufgabe, die in dieser Sitzung angefangen wurde
--                     (bei Aufgabe 5: angefangen, aber nicht bearbeitet)
create view public."Auswertung Kostenlos Absprung"
with (security_invoker = true) as
with ereignisse as (
  select
    session_id,
    created_at,
    event,
    exercise_id,
    detail,
    case when detail->>'nr' ~ '^[1-9]$' then (detail->>'nr')::int end                    as nr
  from public.events
  where portal = 'kostenlos'
),
geraete as (
  select
    session_id,
    coalesce(
      (array_agg(detail->>'geraet' order by created_at) filter (where event = 'page_view'))[1],
      (array_agg(detail->>'geraet' order by created_at))[1]
    )                                                                                      as geraet
  from ereignisse
  group by session_id
),
pro_aufgabe as (
  select
    session_id,
    nr,
    min(exercise_id)                                                                       as aufgabe,
    bool_or(event = 'exercise_view')                                                       as gesehen,
    bool_or(event = 'exercise_start')                                                      as gestartet,
    bool_or(event = 'exercise_solved')                                                     as geloest,
    bool_or(event = 'exercise_solved' and detail->>'mit_loesung' = 'false')                as selbst_geloest,
    bool_or(event = 'solution_show')                                                       as loesung,
    count(*) filter (where event = 'check' and coalesce(detail->>'leer', 'false') <> 'true') as pruefungen
  from ereignisse
  where nr is not null
  group by session_id, nr
),
letzte as (
  select session_id, max(nr) as letzte_nr
  from pro_aufgabe
  where gestartet
  group by session_id
)
select
  g.geraet                                                                                 as "Gerät",
  p.nr                                                                                     as "Nr.",
  min(p.aufgabe)                                                                           as "Aufgabe",
  count(*) filter (where p.gesehen)                                                        as "Gesehen",
  count(*) filter (where p.gestartet)                                                      as "Gestartet",
  count(*) filter (where p.gestartet and (p.geloest or p.loesung))                         as "Bearbeitet",
  count(*) filter (where p.selbst_geloest)                                                 as "Selbst gelöst",
  count(*) filter (where p.loesung)                                                        as "Lösung angesehen",
  count(*) filter (where p.gestartet and l.letzte_nr = p.nr
                   and (p.nr < 5 or not (p.geloest or p.loesung)))                         as "Ausgestiegen",
  round(100.0 * count(*) filter (where p.gestartet and l.letzte_nr = p.nr
                                 and (p.nr < 5 or not (p.geloest or p.loesung)))
        / nullif(count(*) filter (where p.gestartet), 0), 1)                               as "Ausstieg in %",
  round(avg(p.pruefungen) filter (where p.gestartet), 1)                                   as "Ø Prüfungen"
from pro_aufgabe p
join geraete g using (session_id)
left join letzte l using (session_id)
where g.geraet in ('desktop', 'mobil')
group by g.geraet, p.nr
order by g.geraet, p.nr;

revoke all on public."Auswertung Kostenlos Absprung" from anon, authenticated;

-- 4. Direkter Vergleich: Steigen Handy-Nutzer früher aus?
--    Zeile 0 = Seite aufgerufen, aber keine Aufgabe angefangen
create view public."Auswertung Kostenlos Desktop vs Mobil"
with (security_invoker = true) as
with ohne_start as (
  select "Gerät" as geraet, "Besuche" as basis, "Keine Aufgabe angefangen" as aus
  from public."Auswertung Kostenlos Übersicht"
  where "Gerät" in ('desktop', 'mobil')
),
zeilen as (
  select 0 as nr, 'Seite aufgerufen, nichts angefangen' as schritt, geraet, basis, aus from ohne_start
  union all
  select "Nr.", "Aufgabe", "Gerät", "Gestartet", "Ausgestiegen" from public."Auswertung Kostenlos Absprung"
)
select
  nr                                                                                       as "Nr.",
  min(schritt)                                                                             as "Schritt",
  sum(basis) filter (where geraet = 'desktop')                                             as "Desktop: Sitzungen",
  sum(basis) filter (where geraet = 'mobil')                                               as "Mobil: Sitzungen",
  round(100.0 * sum(aus) filter (where geraet = 'desktop')
        / nullif(sum(basis) filter (where geraet = 'desktop'), 0), 1)                      as "Desktop: Ausstieg in %",
  round(100.0 * sum(aus) filter (where geraet = 'mobil')
        / nullif(sum(basis) filter (where geraet = 'mobil'), 0), 1)                        as "Mobil: Ausstieg in %",
  round(100.0 * sum(aus) filter (where geraet = 'mobil')
        / nullif(sum(basis) filter (where geraet = 'mobil'), 0), 1)
  - round(100.0 * sum(aus) filter (where geraet = 'desktop')
        / nullif(sum(basis) filter (where geraet = 'desktop'), 0), 1)                      as "Mobil minus Desktop (Prozentpunkte)"
from zeilen
group by nr
order by nr;

revoke all on public."Auswertung Kostenlos Desktop vs Mobil" from anon, authenticated;
