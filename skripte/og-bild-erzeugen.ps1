# Erzeugt das Link-Vorschaubild assets/og-bild.jpg (1200x630) aus skripte/og-bild.html.
# Braucht: Microsoft Edge (vorinstalliert) und Python mit Pillow.
# Aufruf: powershell -ExecutionPolicy Bypass -File skripte\og-bild-erzeugen.ps1

$ErrorActionPreference = "Stop"

$repo = Split-Path -Parent $PSScriptRoot
$vorlage = Join-Path $PSScriptRoot "og-bild.html"
$ziel = Join-Path $repo "assets\og-bild.jpg"
$png = Join-Path ([System.IO.Path]::GetTempPath()) ("og-bild-" + [guid]::NewGuid() + ".png")

$edge = @(
  "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
  "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) { throw "Microsoft Edge nicht gefunden" }

# Leerzeichen im Pfad (z. B. "Claude Code") als %20, sonst zerfaellt die Adresse in zwei Argumente
$url = "file:///" + (($vorlage -replace "\\", "/") -replace " ", "%20")
$profil = Join-Path ([System.IO.Path]::GetTempPath()) ("og-edge-" + [guid]::NewGuid())
# Start-Process statt direktem Aufruf: Edge schreibt harmlose Meldungen nach stderr,
# die PowerShell sonst als Fehler wertet und das Skript abbricht.
$argumente = @(
  "--headless=new", "--disable-gpu", "--hide-scrollbars", "--force-device-scale-factor=1",
  "--user-data-dir=`"$profil`"", "--window-size=1200,630", "--virtual-time-budget=8000",
  "--screenshot=`"$png`"", $url
)
Start-Process -FilePath $edge -ArgumentList $argumente -Wait -WindowStyle Hidden
Start-Sleep -Milliseconds 500
if (-not (Test-Path $png)) { throw "Screenshot wurde nicht erzeugt" }

# Auf exakt 1200x630 zuschneiden und als JPG speichern (deutlich kleiner als PNG)
python -c "from PIL import Image; import sys; im=Image.open(sys.argv[1]).convert('RGB'); print('Rohbild', im.size); im=im.crop((0,0,1200,630)); im.save(sys.argv[2], 'JPEG', quality=86, optimize=True, progressive=True)" $png $ziel
if ($LASTEXITCODE -ne 0) { throw "Umwandlung fehlgeschlagen" }

Remove-Item -Force $png
Remove-Item -Recurse -Force $profil -ErrorAction SilentlyContinue
$kb = [math]::Round((Get-Item $ziel).Length / 1KB)
Write-Host "assets/og-bild.jpg geschrieben ($kb KB)"
