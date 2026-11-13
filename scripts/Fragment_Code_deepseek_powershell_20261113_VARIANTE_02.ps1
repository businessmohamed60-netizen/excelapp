# ============================================================
# Vérification du driver SQLite3 ODBC
# ============================================================
Write-Host "Vérification du driver SQLite3 ODBC..." -ForegroundColor Cyan

$excelBits = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue).Platform
if (-not $excelBits) { $excelBits = "x86" }

$odbcKey = if ($excelBits -eq "x64") {
    "HKLM:\SOFTWARE\ODBC\ODBCINST.INI\SQLite3 ODBC Driver"
} else {
    "HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI\SQLite3 ODBC Driver"
}

if (-not (Test-Path $odbcKey)) {
    Write-Warning "Driver SQLite3 ODBC non détecté (Excel $excelBits)."
    Write-Host "Téléchargez et installez :" -ForegroundColor Yellow
    if ($excelBits -eq "x64") {
        Write-Host "  https://www.ch-werner.de/sqliteodbc/sqliteodbc_w64.exe" -ForegroundColor Yellow
    } else {
        Write-Host "  https://www.ch-werner.de/sqliteodbc/sqliteodbc.exe" -ForegroundColor Yellow
    }
    $resp = Read-Host "Continuer l'installation sans le driver ? [o/N]"
    if ($resp -ne "o") { exit 1 }
} else {
    Write-Host "Driver SQLite3 ODBC détecté." -ForegroundColor Green
}