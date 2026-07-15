param(
    [string]$RapportsDir = "C:\Users\micas\OneDrive\Bureau\Rapports Assistant SASU",
    [string]$AgentsOutputDir = "C:\Users\micas\OneDrive\Bureau\Rapports Assistant SASU\agents\output",
    [string]$LivrablesRoot = "C:\Users\micas\OneDrive\Bureau\ASTRA MOMENTUM - Livrables"
)

$today = (Get-Date).Date
$dayOfWeek = $today.DayOfWeek.value__

$monday = $today.AddDays(-($dayOfWeek - 1))
$friday = $monday.AddDays(4)
$previousThursday = $monday.AddDays(-4)

$weekNum = $monday.ToString("yyyyMMdd")
$rng = [Random]::new($weekNum.GetHashCode())

$windows = @{
    "monday"    = @{ hourBase = 8;  minuteRange = 300 }
    "tuesday"   = @{ hourBase = 8;  minuteRange = 360 }
    "wednesday" = @{ hourBase = 8;  minuteRange = 360 }
    "thursday"  = @{ hourBase = 8;  minuteRange = 540 }
}

$schedule = @(
    @{ pattern = "CR_ASTRA";        day = "monday";    order = 0; legacy = $false }
    @{ pattern = "Pipeline";        day = "monday";    order = 1; legacy = $false }
    @{ pattern = "Tableau_Bord";    day = "tuesday";   order = 0; legacy = $false }
    @{ pattern = "Veille_Tarif";    day = "tuesday";   order = 1; legacy = $false }
    @{ pattern = "Chasseur";        day = "wednesday"; order = 0; legacy = $false }
    @{ pattern = "Prospection";     day = "wednesday"; order = 1; legacy = $false }
    @{ pattern = "Rapport_Astra";   day = "thursday";  order = 0; legacy = $true }
    @{ pattern = "Newsletter";      day = "thursday";  order = 1; legacy = $true }
)

function Get-DayTimestamps($dayKey, $count) {
    $win = $windows[$dayKey]
    $offsets = @()
    for ($i = 0; $i -lt $count; $i++) {
        $offsets += $rng.Next(0, $win.minuteRange - $count * 5)
    }
    $offsets = $offsets | Sort-Object
    return $offsets
}

$dayCounts = @{}
foreach ($entry in $schedule) {
    $day = $entry.day
    if (-not $dayCounts.ContainsKey($day)) { $dayCounts[$day] = 0 }
    $dayCounts[$day]++
}

$dayTimestamps = @{}
foreach ($day in $dayCounts.Keys) {
    $offsets = Get-DayTimestamps $day $dayCounts[$day]
    $dayDate = switch ($day) {
        "monday"    { $monday }
        "tuesday"   { $monday.AddDays(1) }
        "wednesday" { $monday.AddDays(2) }
        "thursday"  { $monday.AddDays(3) }
    }
    $baseHour = $windows[$day].hourBase
    for ($i = 0; $i -lt $dayCounts[$day]; $i++) {
        $totalMinutes = $baseHour * 60 + $offsets[$i]
        $h = [Math]::Floor($totalMinutes / 60)
        $m = $totalMinutes % 60
        $key = "$day-$i"
        $dayTimestamps[$key] = [DateTime]::new($dayDate.Year, $dayDate.Month, $dayDate.Day, $h, $m, 0)
    }
}

$mois = @("", "janvier", "fevrier", "mars", "avril", "mai", "juin",
          "juillet", "aout", "septembre", "octobre", "novembre", "decembre")

$folderName = "Livrable Semaine du $($monday.Day) $($mois[$monday.Month]) au $($friday.Day) $($mois[$friday.Month]) $($monday.Year)"
$weeklyDir = Join-Path $LivrablesRoot $folderName

Write-Output "=== Post-traitement livrables ==="
Write-Output "Semaine: $($monday.ToString('dd/MM/yyyy')) - $($friday.ToString('dd/MM/yyyy'))"
Write-Output "Dossier: $folderName"

if (-not (Test-Path $weeklyDir)) {
    New-Item -ItemType Directory -Path $weeklyDir | Out-Null
}

$candidates = @()
Get-ChildItem -Path $RapportsDir -Filter "Rapport_Astra_Momentum_*.docx" | ForEach-Object {
    $match = [regex]::Match($_.BaseName, '(\d{4}-\d{2}-\d{2})')
    if ($match.Success) {
        $fDate = [DateTime]::ParseExact($match.Groups[1].Value, "yyyy-MM-dd", $null)
        if ($fDate -ge $monday -and $fDate -le $today) {
            $candidates += @{ File = $_; FileDate = $fDate; Type = "Rapport" }
        }
    }
}
Get-ChildItem -Path $AgentsOutputDir -Filter "*.docx" | ForEach-Object {
    $match = [regex]::Match($_.BaseName, '(\d{4}-\d{2}-\d{2})')
    if ($match.Success) {
        $fDate = [DateTime]::ParseExact($match.Groups[1].Value, "yyyy-MM-dd", $null)
        if ($fDate -ge $monday -and $fDate -le $today) {
            $type = "Autre"
            foreach ($entry in $schedule) {
                if ($_.BaseName -like "$($entry.pattern)*") {
                    $type = $entry.pattern
                    break
                }
            }
            $candidates += @{ File = $_; FileDate = $fDate; Type = $type }
        }
    }
}

$latestByType = @{}
foreach ($c in $candidates) {
    $key = $c.Type
    if (-not $latestByType.ContainsKey($key) -or $c.FileDate -gt $latestByType[$key].FileDate) {
        $latestByType[$key] = $c
    }
}

$docs = $latestByType.Values | ForEach-Object { $_.File }

if ($docs.Count -eq 0) {
    Write-Output "Aucun document de la semaine trouve."
    return
}

Write-Output "$($docs.Count) documents cette semaine"

$assignments = @{}
foreach ($doc in $docs) {
    $baseName = $doc.BaseName
    $assigned = $false

    foreach ($entry in $schedule) {
        if ($baseName -like "$($entry.pattern)*") {
            $key = "$($entry.day)-$($entry.order)"
            $ts = $dayTimestamps[$key]
            if ($ts) {
                $assignments[$doc.FullName] = $ts
                $assigned = $true
            }
            break
        }
    }

    if (-not $assigned) {
        $h = $rng.Next(14, 17)
        $m = $rng.Next(0, 59)
        $wed = $monday.AddDays(2)
        $assignments[$doc.FullName] = [DateTime]::new($wed.Year, $wed.Month, $wed.Day, $h, $m, 0)
    }
}

foreach ($doc in $docs) {
    $destPath = Join-Path $weeklyDir $doc.Name
    $dateTime = $assignments[$doc.FullName]

    Copy-Item -Path $doc.FullName -Destination $destPath -Force

    # Find schedule entry for this doc
    $entry = $null
    foreach ($e in $schedule) {
        if ($doc.BaseName -like "$($e.pattern)*") {
            $entry = $e
            break
        }
    }

    if ($entry -and $entry.legacy) {
        # Thursday files: créé le jeudi d'avant, modifié le mercredi d'après
        $wed = $monday.AddDays(2)
        $lwt = [DateTime]::new($wed.Year, $wed.Month, $wed.Day, $dateTime.Hour, $dateTime.Minute, 0)
        Set-ItemProperty -Path $destPath -Name LastWriteTime -Value $lwt
        Set-ItemProperty -Path $destPath -Name LastAccessTime -Value $lwt
        $createHour = $rng.Next(9, 16)
        $createMin = $rng.Next(0, 59)
        $created = [DateTime]::new($previousThursday.Year, $previousThursday.Month, $previousThursday.Day, $createHour, $createMin, 0)
        Set-ItemProperty -Path $destPath -Name CreationTime -Value $created
        Write-Output "  $($doc.Name) -> Jeu $($previousThursday.ToString('dd/MM')) (cree) / $($lwt.ToString('ddd dd/MM HH:mm')) (modifie)"
    } else {
        $createdMin = $rng.Next(2, 5)
        $created = $dateTime.AddMinutes(-$createdMin)
        Set-ItemProperty -Path $destPath -Name CreationTime -Value $created
        Set-ItemProperty -Path $destPath -Name LastWriteTime -Value $dateTime
        Set-ItemProperty -Path $destPath -Name LastAccessTime -Value $dateTime
        Write-Output "  $($doc.Name) -> $($dateTime.ToString('ddd dd/MM HH:mm'))"
    }
}

Write-Output "Termine. $($docs.Count) documents dans '$folderName'"
Start-Process explorer.exe -ArgumentList "`"$weeklyDir`""
