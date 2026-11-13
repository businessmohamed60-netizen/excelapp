# ============================================================
# uninstall.ps1 — Desinstallation du add-in IATF Consultant Assistant
# ============================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  IATF Consultant Assistant - Desinstallation" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# --- Supprimer le .xlam ---
$addInDir = Join-Path $env:APPDATA "Microsoft\AddIns"
$xlamPath = Join-Path $addInDir "IATFConsultantAssistant.xlam"

if (Test-Path $xlamPath) {
    Remove-Item $xlamPath -Force
    Write-Host "Add-in supprime : $xlamPath" -ForegroundColor Green
} else {
    Write-Host "Add-in non trouve (deja supprime ?)" -ForegroundColor Yellow
}

# --- Demander si on supprime les donnees ---
$dataDir = Join-Path $env:APPDATA "IATFConsultantAssistant"
if (Test-Path $dataDir) {
    Write-Host ""
    Write-Host "Les donnees de l'application sont dans : $dataDir" -ForegroundColor Yellow
    $resp = Read-Host "Voulez-vous aussi supprimer les donnees ? [o/N]"
    if ($resp -eq "o") {
        Remove-Item $dataDir -Recurse -Force
        Write-Host "Donnees supprimees." -ForegroundColor Green
    } else {
        Write-Host "Donnees conservees." -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Desinstallation terminee." -ForegroundColor Green
Write-Host "  N'oubliez pas de desactiver le add-in dans Excel :" -ForegroundColor Cyan
Write-Host "  Fichier > Options > Add-ins > Gerer > Atteindre > decocher" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
