param(
    [string]$RepoDir = "C:\Users\micas\OneDrive\Bureau\Rapports Assistant SASU",
    [int]$GitTimeoutSec = 180,
    [int]$MaxAttempts = 3
)

$logFile = Join-Path $RepoDir "deploiement.log"

function Log($msg) {
    $line = "{0}  {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    Add-Content -Path $logFile -Value $line
    [Console]::WriteLine($msg)
}

# Preflight reseau : evite de lancer git quand le DNS n'est pas encore monte.
# C'etait la cause racine des blocages : git partait au demarrage du PC, le
# resolver ne repondait pas, et le processus ne mourait jamais.
function Test-Connectivity {
    try {
        $null = [System.Net.Dns]::GetHostAddresses("github.com")
        return $true
    } catch {
        return $false
    }
}

# Lance git dans un job avec un timeout dur. Plus de git.exe orphelin.
function Invoke-GitPull {
    param([int]$TimeoutSec)

    $job = Start-Job -ScriptBlock {
        param($dir)
        Set-Location -LiteralPath $dir
        $env:GIT_TERMINAL_PROMPT = "0"
        $env:GCM_INTERACTIVE = "never"
        $out = git -c http.lowSpeedLimit=1000 -c http.lowSpeedTime=30 pull 2>&1
        $code = $LASTEXITCODE
        $out | ForEach-Object { "$_" }
        # Sentinelle : evite d'avoir a deviner quel element est le code de sortie
        "__GITEXIT__$code"
    } -ArgumentList $RepoDir

    $finished = Wait-Job -Job $job -Timeout $TimeoutSec

    if ($null -eq $finished) {
        Log "[WARN] git pull depasse ${TimeoutSec}s - processus arrete"
        Stop-Job -Job $job -ErrorAction SilentlyContinue
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        return 124
    }

    $output = @(Receive-Job -Job $job)
    Remove-Job -Job $job -Force -ErrorAction SilentlyContinue

    $exit = 1
    foreach ($line in $output) {
        if ("$line" -match "^__GITEXIT__(\d+)$") {
            $exit = [int]$Matches[1]
        } else {
            Log "$line"
        }
    }
    return $exit
}

try {
    Set-Location -LiteralPath $RepoDir

    Log "=== Debut deploiement ==="

    if (-not (Test-Connectivity)) {
        Log "[WARN] github.com injoignable (DNS) - run ignore, report non telecharge"
        Log "=== Termine (sans action) ==="
        exit 0
    }

    Log "=== Pull du repo GitHub ==="

    $exit = 1
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        if ($attempt -gt 1) {
            Log "[RETRY] tentative $attempt/$MaxAttempts"
            Start-Sleep -Seconds 30
        }
        $exit = Invoke-GitPull -TimeoutSec $GitTimeoutSec
        if ($exit -eq 0) { break }
        if ($exit -eq 124) {
            Log "[ERR] git pull bloque, abandon definitif (pas de boucle)"
            exit 1
        }
        Log "[ERR] git pull echoue (code $exit)"
    }

    if ($exit -ne 0) {
        Log "[ERR] git pull a echoue apres $MaxAttempts tentatives - arret"
        exit 1
    }

    # Post-traitement : copie vers Livrables + timestamps realistes
    Log "=== Post-traitement des livrables ==="
    & "$RepoDir\post_traitement_livrables.ps1" *>&1 | ForEach-Object { Log $_ }

    Log "=== Termine ==="
    exit 0
}
catch {
    Log "[ERR] Exception: $($_.Exception.Message)"
    exit 1
}
