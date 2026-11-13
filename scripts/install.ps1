# ============================================================
# install.ps1 — Installation du add-in IATF Consultant Assistant
# Copie le .xlam dans le dossier des add-ins Excel et l'active.
# ============================================================

param(
    [string]$XlamPath = ""
)

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  IATF Consultant Assistant - Installation" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# --- Trouver le fichier .xlam ---
if ($XlamPath -eq "") {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $XlamPath = Join-Path $scriptDir "..\build\IATFConsultantAssistant.xlam"
}

if (-not (Test-Path $XlamPath)) {
    # Chercher dans le dossier build
    $buildDir = Join-Path (Split-Path -Parent $scriptDir) "build"
    $found = Get-ChildItem $buildDir -Filter "*.xlam" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($found) {
        $XlamPath = $found.FullName
    } else {
        Write-Host "Fichier .xlam introuvable." -ForegroundColor Red
        Write-Host "Executez d'abord build.ps1 pour assembler le add-in." -ForegroundColor Yellow
        exit 1
    }
}

Write-Host "Fichier source : $XlamPath" -ForegroundColor Green

# --- Dossier de destination ---
$addInDir = Join-Path $env:APPDATA "Microsoft\AddIns"
if (-not (Test-Path $addInDir)) {
    New-Item -ItemType Directory -Path $addInDir -Force | Out-Null
}

$destPath = Join-Path $addInDir "IATFConsultantAssistant.xlam"

# --- Copier le fichier ---
Write-Host "Copie vers : $destPath" -ForegroundColor Yellow
Copy-Item $XlamPath $destPath -Force
Write-Host "Fichier copie." -ForegroundColor Green

# --- Creer le dossier de donnees ---
$dataDir = Join-Path $env:APPDATA "IATFConsultantAssistant"
if (-not (Test-Path $dataDir)) {
    New-Item -ItemType Directory -Path $dataDir -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $dataDir "templates") -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $dataDir "backup") -Force | Out-Null
    Write-Host "Dossier de donnees cree : $dataDir" -ForegroundColor Green
}

# --- Verifier le driver SQLite ---
$odbcKey = "HKLM:\SOFTWARE\ODBC\ODBCINST.INI\SQLite3 ODBC Driver"
$odbcKey32 = "HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI\SQLite3 ODBC Driver"
if (-not (Test-Path $odbcKey) -and -not (Test-Path $odbcKey32)) {
    Write-Host ""
    Write-Host "ATTENTION : Driver SQLite3 ODBC non installe." -ForegroundColor Yellow
    Write-Host "L'application fonctionnera en mode XML (sans base SQL)." -ForegroundColor Yellow
    Write-Host "Pour activer le mode SQL, installez : https://www.ch-werner.de/sqliteodbc/" -ForegroundColor Yellow
}

# --- Instructions d'activation ---
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Installation terminee !" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Activez le add-in dans Excel :"
Write-Host "  1. Ouvrez Excel"
Write-Host "  2. Fichier > Options > Add-ins"
Write-Host "  3. En bas, 'Gerer : Add-ins Excel' > cliquez 'Atteindre'"
Write-Host "  4. Cochez 'IATFConsultantAssistant'"
Write-Host "  5. Cliquez OK"
Write-Host ""
Write-Host "  Un nouvel onglet 'IATF Assistant' apparait dans le ruban."
Write-Host ""
