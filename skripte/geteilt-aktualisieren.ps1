# Holt die geteilte Logik (Tabellenblatt, Formel-Engine, Design, Tracking) aus dem
# Uebungsportal-Repo in dieses Repo (Ordner assets/geteilt/).
#
# Quelle ist bewusst der VEROEFFENTLICHTE Stand (origin/main), nicht der lokale
# Arbeitsstand: unfertige Aenderungen am Portal landen so nie auf der
# kostenlosen Seite. Danach: Seite lokal testen, erst dann committen und pushen.
#
# Aufruf: Doppelklick auf geteilt-aktualisieren.bat im Hauptordner.

$ErrorActionPreference = "Stop"

$repo   = Split-Path -Parent $PSScriptRoot
$portal = Join-Path (Split-Path -Parent $repo) "excel-flo-uebungsportal"
$ziel   = Join-Path $repo "assets\geteilt"

$dateien = @(
  "assets/engine.js",
  "assets/formula-engine.js",
  "assets/engine.css",
  "assets/tracking.js",
  "assets/img/bg-green.webp",
  "assets/img/logo-white.webp",
  "assets/img/logo.svg"
)

if (-not (Test-Path (Join-Path $portal ".git"))) {
  throw "Portal-Repo nicht gefunden: $portal"
}

Write-Host "Hole aktuellen Stand des Portals von GitHub ..."
git -C $portal fetch --quiet origin
if ($LASTEXITCODE -ne 0) { throw "git fetch fehlgeschlagen" }

$commit = (git -C $portal rev-parse --short origin/main).Trim()
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("geteilt-" + [guid]::NewGuid())
New-Item -ItemType Directory -Path $tmp | Out-Null
$tar = Join-Path $tmp "geteilt.tar"

# git archive + tar statt git show: PowerShell-Pipes wuerden die Bilder beschaedigen.
git -C $portal archive --format=tar -o $tar origin/main @dateien
if ($LASTEXITCODE -ne 0) { throw "git archive fehlgeschlagen" }
tar -xf $tar -C $tmp
if ($LASTEXITCODE -ne 0) { throw "Entpacken fehlgeschlagen" }

New-Item -ItemType Directory -Force -Path (Join-Path $ziel "img") | Out-Null
foreach ($d in $dateien) {
  $relativ = $d.Substring("assets/".Length)
  Copy-Item -Force (Join-Path $tmp $d) (Join-Path $ziel $relativ)
  Write-Host "  aktualisiert: assets/geteilt/$relativ"
}
Remove-Item -Recurse -Force $tmp

$datum = Get-Date -Format "yyyy-MM-dd HH:mm"
@(
  "Kopie aus github.com/FlorianBeil/Excel-Aufgaben (Branch main)",
  "Commit: $commit",
  "Geholt: $datum",
  "",
  "Nicht von Hand aendern - wird beim naechsten Update ueberschrieben.",
  "Aenderungen an diesen Dateien immer im Portal-Repo machen."
) | Set-Content -Encoding ASCII (Join-Path $ziel "QUELLE.txt")

Write-Host ""
Write-Host "Fertig. Stand des Portals: Commit $commit"
Write-Host "Naechster Schritt: Seite testen, dann committen und pushen."
