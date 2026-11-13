# Vérifier le driver SQLite
$sqliteDriver = Get-OdbcDriver | Where-Object { $_.Name -like "*SQLite*" }
if (-not $sqliteDriver) {
    Write-Warning "Driver SQLite3 ODBC non détecté."
    Write-Host "Téléchargez-le sur https://www.ch-werner.de/sqliteodbc/ (gratuit)" -ForegroundColor Yellow
}

# Vérifier Python (optionnel, pour l'API REST)
$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) {
    Write-Host "Python non détecté — l'API REST ne sera pas disponible." -ForegroundColor Yellow
    Write-Host "Installez Python 3.x depuis https://www.python.org/ si vous voulez l'API REST." -ForegroundColor Yellow
}