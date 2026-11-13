# ============================================================
# verifier.ps1 — Verification pre-installation du add-in
# Verifie les prerequis avant d'installer le .xlam
# ============================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  IATF Consultant Assistant - Verification pre-installation" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# --- 1. Excel ---
Write-Host "[1/4] Excel..." -ForegroundColor Yellow
$excel = Get-ChildItem "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" -ErrorAction SilentlyContinue
if (-not $excel) {
    $excel = Get-ChildItem "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE" -ErrorAction SilentlyContinue
}
if ($excel) {
    Write-Host "  Excel detecte : $($excel.FullName)" -ForegroundColor Green
} else {
    Write-Host "  Excel non detecte dans le chemin standard." -ForegroundColor Yellow
    Write-Host "  L'installation peut quand meme continuer." -ForegroundColor Yellow
}

# --- 2. Architecture (32/64-bit) ---
Write-Host ""
Write-Host "[2/4] Architecture Excel..." -ForegroundColor Yellow
$excelBits = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue).Platform
if (-not $excelBits) { $excelBits = "x86 (par defaut)" }
Write-Host "  Excel : $excelBits" -ForegroundColor Green
Write-Host "  Pour SQLite, installez le driver correspondant." -ForegroundColor Yellow

# --- 3. Driver SQLite ---
Write-Host ""
Write-Host "[3/4] Driver SQLite3 ODBC..." -ForegroundColor Yellow
$odbcKey64 = "HKLM:\SOFTWARE\ODBC\ODBCINST.INI\SQLite3 ODBC Driver"
$odbcKey32 = "HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI\SQLite3 ODBC Driver"
if ((Test-Path $odbcKey64) -or (Test-Path $odbcKey32)) {
    Write-Host "  Driver SQLite3 ODBC detecte." -ForegroundColor Green
} else {
    Write-Host "  Driver SQLite3 ODBC NON detecte." -ForegroundColor Yellow
    Write-Host "  L'application fonctionnera en mode XML." -ForegroundColor Yellow
    Write-Host "  Pour le mode SQL : https://www.ch-werner.de/sqliteodbc/" -ForegroundColor Yellow
}

# --- 4. Dossier APPDATA ---
Write-Host ""
Write-Host "[4/4] Dossier de donnees..." -ForegroundColor Yellow
$dataDir = Join-Path $env:APPDATA "IATFConsultantAssistant"
if (Test-Path $dataDir) {
    Write-Host "  Donnees existantes : $dataDir" -ForegroundColor Green
    $size = (Get-ChildItem $dataDir -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
    $sizeMB = [math]::Round($size / 1MB, 2)
    Write-Host "  Taille : $sizeMB MB" -ForegroundColor Green
} else {
    Write-Host "  Aucune donnee existante (premiere installation)." -ForegroundColor Green
}

# --- Resume ---
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Verification terminee. Vous pouvez lancer install.ps1" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
