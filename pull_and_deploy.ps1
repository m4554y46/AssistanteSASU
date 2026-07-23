param(
    [string]$RepoDir = "C:\Users\micas\OneDrive\Bureau\Rapports Assistant SASU"
)

# Va dans le repo
Set-Location -LiteralPath $RepoDir

# Récupère les derniers fichiers générés par GitHub Actions
Write-Output "=== Pull du repo GitHub ==="
git pull
if (-not $?) { Write-Output "[ERR] git pull echoue"; exit 1 }

# Post-traitement : copie vers Livrables + timestamps réalistes
Write-Output "=== Post-traitement des livrables ==="
& "$RepoDir\post_traitement_livrables.ps1"

Write-Output "=== Termine ==="
