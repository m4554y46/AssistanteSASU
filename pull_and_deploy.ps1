param(
    [string]$RepoDir = "C:\Users\micas\OneDrive\Bureau\Rapports Assistant SASU"
)

$logFile = Join-Path $RepoDir "deploiement.log"

function Log($msg) {
    $line = "{0}  {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    Add-Content -Path $logFile -Value $line
    Write-Output $msg
}

try {
    Set-Location -LiteralPath $RepoDir

    Log "=== Debut deploiement ==="
    Log "=== Pull du repo GitHub ==="
    git pull 2>&1 | ForEach-Object { Log $_ }
    if ($LASTEXITCODE -ne 0) {
        Log "[ERR] git pull echoue (code $LASTEXITCODE) - nouvelle tentative dans 30s"
        Start-Sleep -Seconds 30
        git pull 2>&1 | ForEach-Object { Log $_ }
        if ($LASTEXITCODE -ne 0) {
            Log "[ERR] git pull echoue une seconde fois - arret"
            exit 1
        }
    }

    # Post-traitement : copie vers Livrables + timestamps realistes
    Log "=== Post-traitement des livrables ==="
    & "$RepoDir\post_traitement_livrables.ps1" *>&1 | ForEach-Object { Log $_ }

    Log "=== Termine ==="
}
catch {
    Log "[ERR] Exception: $($_.Exception.Message)"
    exit 1
}