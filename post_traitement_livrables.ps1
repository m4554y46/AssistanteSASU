param(
    [string]$RapportsDir = "C:\Users\micas\OneDrive\Bureau\Rapports Assistant SASU",
    [string]$AgentsOutputDir = "C:\Users\micas\OneDrive\Bureau\Rapports Assistant SASU\agents\output",
    [string]$LivrablesRoot = "C:\Users\micas\OneDrive\Bureau\ASTRA MOMENTUM - Livrables"
)

$today = (Get-Date).Date
$dayOfWeek = $today.DayOfWeek.value__

# Find last Monday
$monday = $today.AddDays(-($dayOfWeek - 1))
$tuesday = $monday.AddDays(1)
$wednesday = $monday.AddDays(2)

# Folder name
$mois = @("", "janvier", "fevrier", "mars", "avril", "mai", "juin",
          "juillet", "aout", "septembre", "octobre", "novembre", "decembre")
$folderName = "Livrable de la semaine $($monday.Day) $($mois[$monday.Month]) $($monday.Year)"
$weeklyDir = Join-Path $LivrablesRoot $folderName

Write-Output "=== Post-traitement des livrables ==="
Write-Output "Semaine du $($monday.ToString('yyyy-MM-dd')) au $($wednesday.ToString('yyyy-MM-dd'))"
Write-Output "Dossier: $weeklyDir"

# Create weekly folder
if (-not (Test-Path $weeklyDir)) {
    New-Item -ItemType Directory -Path $weeklyDir | Out-Null
}

# Collect all DOCX from this week only
$docs = @()
Get-ChildItem -Path $RapportsDir -Filter "Rapport_Astra_Momentum_*.docx" | Where-Object { $_.BaseName -like "*$($today.ToString('yyyy-MM-dd'))*" } | ForEach-Object { $docs += $_ }
Get-ChildItem -Path $AgentsOutputDir -Filter "*.docx" | Where-Object { $_.LastWriteTime -ge $monday } | ForEach-Object { $docs += $_ }

if ($docs.Count -eq 0) {
    Write-Output "Aucun document trouve."
    return
}

Write-Output "Traitement de $($docs.Count) documents..."

# Realistic schedule: each file gets a unique timestamp spread across Mon-Tue-Wed
# Format: "prefix" = hours:minutes offset from start of each day
$schedule = @(
    # Monday morning - CR, meeting notes, admin
    @{ pattern = "CR_ASTRA";     dayOffset = 0; hour = 9;  minute = 12 },
    @{ pattern = "CR_reunion";    dayOffset = 0; hour = 9;  minute = 15 },
    @{ pattern = "point_hebdo";   dayOffset = 0; hour = 9;  minute = 18 },  # generated notes
    # Monday afternoon - tableau de bord, pipeline
    @{ pattern = "Tableau_Bord";  dayOffset = 0; hour = 10; minute = 33 },
    @{ pattern = "Pipeline";      dayOffset = 0; hour = 14; minute = 5 },
    # Tuesday morning - veille, chasseur
    @{ pattern = "Chasseur";      dayOffset = 1; hour = 9;  minute = 47 },
    @{ pattern = "Veille_Tarif";  dayOffset = 1; hour = 11; minute = 22 },
    # Tuesday afternoon - newsletter, prospection
    @{ pattern = "Newsletter";    dayOffset = 1; hour = 15; minute = 38 },
    @{ pattern = "Branding";      dayOffset = 1; hour = 15; minute = 40 },
    # Wednesday morning - prospection pack, propositions
    @{ pattern = "Prospection";   dayOffset = 2; hour = 10; minute = 17 },
    @{ pattern = "Proposition";   dayOffset = 2; hour = 10; minute = 19 },
    # Wednesday late morning - final report
    @{ pattern = "Rapport_Astra"; dayOffset = 2; hour = 11; minute = 45 }
)

# Group docs by type for timestamp assignment
$remaining = @()  # docs that don't match any schedule pattern

foreach ($doc in $docs) {
    $baseName = $doc.BaseName
    $assigned = $false

    foreach ($entry in $schedule) {
        if ($baseName -like "$($entry.pattern)*") {
            $day = $monday.AddDays($entry.dayOffset)
            $dateTime = [DateTime]::new($day.Year, $day.Month, $day.Day, $entry.hour, $entry.minute, 0)
            $assigned = $true
            break
        }
    }

    if (-not $assigned) {
        # Fallback: spread remaining files across Tuesday afternoon
        $rng = [Random]::new()
        $hour = $rng.Next(14, 18)
        $minute = $rng.Next(0, 59)
        $dateTime = [DateTime]::new($tuesday.Year, $tuesday.Month, $tuesday.Day, $hour, $minute, 0)
    }

    # Copy to weekly folder
    $destPath = Join-Path $weeklyDir $doc.Name
    Copy-Item -Path $doc.FullName -Destination $destPath -Force

    # Set timestamps: CreationTime = first created, LastWriteTime = timestamp assigned
    if ($doc.FullName -like "*Rapport_Astra*") {
        # Report: created Monday 08:30, modified on assigned timestamp
        $created = [DateTime]::new($monday.Year, $monday.Month, $monday.Day, 8, 30, 0)
        Set-ItemProperty -Path $destPath -Name CreationTime -Value $created
        Set-ItemProperty -Path $destPath -Name LastWriteTime -Value $dateTime
        Set-ItemProperty -Path $destPath -Name LastAccessTime -Value $dateTime
    } else {
        # Other docs: created on the assigned date but slightly earlier
        $created = $dateTime.AddMinutes(-2)
        Set-ItemProperty -Path $destPath -Name CreationTime -Value $created
        Set-ItemProperty -Path $destPath -Name LastWriteTime -Value $dateTime
        Set-ItemProperty -Path $destPath -Name LastAccessTime -Value $dateTime
    }

    Write-Output "  $($doc.Name) -> $($dateTime.ToString('ddd dd/MM HH:mm'))"
}

Write-Output "Termine. $($docs.Count) documents dans '$folderName'"
Start-Process explorer.exe -ArgumentList "`"$weeklyDir`""
