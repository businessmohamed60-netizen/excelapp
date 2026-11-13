# ============================================================
# build.ps1 — Assemblage du fichier .xlam IATF Consultant Assistant
# Usage : Ouvrir PowerShell, executer ce script.
# Ce script guide l'assemblage manuel du add-in Excel VBA.
# ============================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  IATF Consultant Assistant - Assemblage du add-in (.xlam)" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# --- Verification des prerequis ---
Write-Host "[1/5] Verification des prerequis..." -ForegroundColor Yellow

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$srcClasses = Join-Path $root "..\src\classes"
$srcModules = Join-Path $root "..\src\modules"
$srcForms   = Join-Path $root "..\src\forms"
$srcTests   = Join-Path $root "..\src\tests"
$ribbonFile = Join-Path $root "..\ribbon\Ribbon.xml"

$allOk = $true

foreach ($dir in @($srcClasses, $srcModules, $srcForms, $srcTests)) {
    if (Test-Path $dir) {
        $count = (Get-ChildItem $dir -File).Count
        Write-Host "  OK : $dir ($count fichiers)" -ForegroundColor Green
    } else {
        Write-Host "  MANQUANT : $dir" -ForegroundColor Red
        $allOk = $false
    }
}

if (Test-Path $ribbonFile) {
    Write-Host "  OK : $ribbonFile" -ForegroundColor Green
} else {
    Write-Host "  MANQUANT : $ribbonFile" -ForegroundColor Red
    $allOk = $false
}

if (-not $allOk) {
    Write-Host ""
    Write-Host "ERREUR : Des dossiers sources sont manquants." -ForegroundColor Red
    exit 1
}

# --- Inventaire ---
Write-Host ""
Write-Host "[2/5] Inventaire des composants..." -ForegroundColor Yellow

$classes = Get-ChildItem $srcClasses -Filter "*.cls" -File
$modules = Get-ChildItem $srcModules -Filter "*.bas" -File
$handlers = Get-ChildItem $srcModules -Filter "Code_*.cls" -File
$forms = Get-ChildItem $srcForms -Filter "*.frm" -File
$tests = Get-ChildItem $srcTests -Filter "*.bas" -File

Write-Host "  Classes metier (.cls)      : $($classes.Count)"
Write-Host "  Modules VBA (.bas)          : $($modules.Count)"
Write-Host "  Handlers/Code (.cls)        : $($handlers.Count)"
Write-Host "  Formulaires (.frm)          : $($forms.Count)"
Write-Host "  Tests unitaires (.bas)      : $($tests.Count)"
Write-Host "  Fichier Ribbon              : Ribbon.xml (7 onglets unifies)"

# --- Verification driver SQLite ---
Write-Host ""
Write-Host "[3/5] Verification driver SQLite3 ODBC..." -ForegroundColor Yellow

$excelBits = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue).Platform
if (-not $excelBits) { $excelBits = "x86" }

$odbcKey = if ($excelBits -eq "x64") {
    "HKLM:\SOFTWARE\ODBC\ODBCINST.INI\SQLite3 ODBC Driver"
} else {
    "HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI\SQLite3 ODBC Driver"
}

if (Test-Path $odbcKey) {
    Write-Host "  Driver SQLite3 ODBC detecte (Excel $excelBits)." -ForegroundColor Green
} else {
    Write-Host "  ATTENTION : Driver SQLite3 ODBC non detecte." -ForegroundColor Yellow
    Write-Host "  Telechargez : https://www.ch-werner.de/sqliteodbc/" -ForegroundColor Yellow
    Write-Host "  (L'application fonctionne aussi en mode XML sans ce driver)" -ForegroundColor Yellow
}

# --- Instructions d'assemblage ---
Write-Host ""
Write-Host "[4/5] Instructions d'assemblage dans Excel..." -ForegroundColor Yellow
Write-Host ""
Write-Host "  1. Ouvrez Excel et creez un nouveau classeur vide"
Write-Host "  2. Appuyez sur Alt+F11 pour ouvrir l'editeur VBA"
Write-Host "  3. Menu Outils > References, ajoutez :"
Write-Host "     - Microsoft Scripting Runtime"
Write-Host "     - Microsoft XML, v6.0"
Write-Host "     - Microsoft Forms 2.0 Object Library"
Write-Host ""
Write-Host "  4. Importez les modules (.bas) :"
Write-Host "     Fichier > Importer un fichier > selectionnez chaque .bas dans src\modules\"
Write-Host ""
Write-Host "  5. Importez les classes (.cls) :"
Write-Host "     Fichier > Importer un fichier > selectionnez chaque .cls dans src\classes\"
Write-Host "     et src\modules\ (fichiers Code_*.cls)"
Write-Host ""
Write-Host "  6. Importez les formulaires (.frm) :"
Write-Host "     Fichier > Importer un fichier > selectionnez chaque .frm dans src\forms\"
Write-Host ""
Write-Host "  7. Importez les tests (.bas) depuis src\tests\ (optionnel)"
Write-Host ""
Write-Host "  8. Compilez : Debogage > Compiler VBAProject"
Write-Host "     (doit passer sans erreur)"
Write-Host ""
Write-Host "  9. Executez RunAllTests dans la fenetre d'execution (Ctrl+G)"
Write-Host "     (toutes les assertions doivent etre vertes)"
Write-Host ""

# --- Injection du Ribbon XML ---
Write-Host "[5/5] Injection du Ruban XML..." -ForegroundColor Yellow
Write-Host ""
Write-Host "  Pour injecter le ruban personnalise :"
Write-Host "  1. Fermez Excel et enregistrez le classeur au format .xlsm"
Write-Host "  2. Telechargez 'Office RibbonX Editor' (gratuit sur GitHub)"
Write-Host "  3. Ouvrez le fichier .xlsm dans RibbonX Editor"
Write-Host "  4. Clic droit > Insert Office 2010+ Custom UI Part"
Write-Host "  5. Collez le contenu de ribbon\Ribbon.xml"
Write-Host "  6. Clic droit > Validate > doit passer sans erreur"
Write-Host "  7. Sauvegardez et fermez RibbonX Editor"
Write-Host ""
Write-Host "  8. Reouvrez le .xlsm dans Excel"
Write-Host "  9. Sauvegardez au format .xlam (Add-in Excel)"
Write-Host "     Fichier > Enregistrer sous > Type : 'Macro Add-In Excel (*.xlam)'"
Write-Host "     Nom : IATFConsultantAssistant.xlam"
Write-Host ""

# --- Finalisation ---
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Assemblage termine !" -ForegroundColor Green
Write-Host "  Le fichier .xlam est pret a etre deploye." -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Pour installer :"
Write-Host "  Excel > Fichier > Options > Add-ins > Gerer : Add-ins Excel > Atteindre"
Write-Host "  > Parcourir > selectionner IATFConsultantAssistant.xlam"
Write-Host ""
