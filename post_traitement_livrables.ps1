param(
    [string]$RapportsDir = "C:\Users\micas\OneDrive\Bureau\Rapports Assistant SASU",
    [string]$AgentsOutputDir = "C:\Users\micas\OneDrive\Bureau\Rapports Assistant SASU\agents\output",
    [string]$LivrablesRoot = "C:\Users\micas\OneDrive\Bureau\ASTRA MOMENTUM - Livrables"
)

$today = (Get-Date).Date
$mois = @("", "janvier", "fevrier", "mars", "avril", "mai", "juin",
          "juillet", "aout", "septembre", "octobre", "novembre", "decembre")

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

function Get-DayTimestamps($dayKey, $count, $rng) {
    $win = $windows[$dayKey]
    $offsets = @()
    for ($i = 0; $i -lt $count; $i++) {
        $offsets += $rng.Next(0, $win.minuteRange - $count * 5)
    }
    $offsets = $offsets | Sort-Object
    return $offsets
}

function Process-Week($monday, $files) {
    $friday = $monday.AddDays(4)
    $previousThursday = $monday.AddDays(-4)
    $weekKey = $monday.ToString("yyyyMMdd")
    $rng = [Random]::new($weekKey.GetHashCode())

    $folderName = "Livrable Semaine du $($monday.Day) $($mois[$monday.Month]) au $($friday.Day) $($mois[$friday.Month]) $($monday.Year)"
    $weeklyDir = Join-Path $LivrablesRoot $folderName

    if (Test-Path $weeklyDir) {
        return $null  # Already processed
    }

    # Keep latest per type
    $latestByType = @{}
    foreach ($f in $files) {
        $key = $f.Type
        if (-not $latestByType.ContainsKey($key) -or $f.FileDate -gt $latestByType[$key].FileDate) {
            $latestByType[$key] = $f
        }
    }
    $docs = $latestByType.Values | ForEach-Object { $_.File }
    if ($docs.Count -eq 0) { return $null }

    # Generate timestamps for this week
    $dayCounts = @{}
    foreach ($entry in $schedule) {
        $day = $entry.day
        if (-not $dayCounts.ContainsKey($day)) { $dayCounts[$day] = 0 }
        $dayCounts[$day]++
    }

    $dayTimestamps = @{}
    foreach ($day in $dayCounts.Keys) {
        $offsets = Get-DayTimestamps $day $dayCounts[$day] $rng
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

    New-Item -ItemType Directory -Path $weeklyDir | Out-Null

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

        $entry = $null
        foreach ($e in $schedule) {
            if ($doc.BaseName -like "$($e.pattern)*") {
                $entry = $e
                break
            }
        }

        if ($entry -and $entry.legacy) {
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

    return @{ Docs = $docs.Count; Folder = $folderName }
}

# ---- MAIN ----
Write-Output "=== Post-traitement livrables ==="

# Scan all DOCX files
$allFiles = @()
Get-ChildItem -Path $RapportsDir -Filter "Rapport_Astra_Momentum_*.docx" | ForEach-Object {
    $match = [regex]::Match($_.BaseName, '(\d{4}-\d{2}-\d{2})')
    if ($match.Success) {
        $fDate = [DateTime]::ParseExact($match.Groups[1].Value, "yyyy-MM-dd", $null)
        if ($fDate -le $today) {
            $allFiles += @{ File = $_; FileDate = $fDate; Type = "Rapport" }
        }
    }
}
Get-ChildItem -Path $AgentsOutputDir -Filter "*.docx" | ForEach-Object {
    $match = [regex]::Match($_.BaseName, '(\d{4}-\d{2}-\d{2})')
    if ($match.Success) {
        $fDate = [DateTime]::ParseExact($match.Groups[1].Value, "yyyy-MM-dd", $null)
        if ($fDate -le $today) {
            $type = "Autre"
            foreach ($entry in $schedule) {
                if ($_.BaseName -like "$($entry.pattern)*") {
                    $type = $entry.pattern
                    break
                }
            }
            $allFiles += @{ File = $_; FileDate = $fDate; Type = $type }
        }
    }
}

if ($allFiles.Count -eq 0) {
    Write-Output "Aucun document trouve."
    return
}

# Group by week
$weeks = @{}
foreach ($f in $allFiles) {
    $fd = $f.FileDate
    $dow = $fd.DayOfWeek.value__
    $monday = $fd.AddDays(-($dow - 1))
    $weekKey = $monday.ToString("yyyyMMdd")
    if (-not $weeks.ContainsKey($weekKey)) {
        $weeks[$weekKey] = @{ Monday = $monday; Files = @() }
    }
    $weeks[$weekKey].Files += $f
}

$totalProcessed = 0
$totalFolders = 0
foreach ($weekKey in ($weeks.Keys | Sort-Object)) {
    $monday = $weeks[$weekKey].Monday
    $result = Process-Week $monday $weeks[$weekKey].Files
    if ($result) {
        $totalFolders++
        $totalProcessed += $result.Docs
        Write-Output "  -> $($result.Folder) : $($result.Docs) documents"
    } else {
        Write-Output "  Deja present : Livrable Semaine du $($monday.Day) $($mois[$monday.Month]) $($monday.Year)"
    }
}

Write-Output "=== Termine : $totalFolders dossiers, $totalProcessed documents ==="

# Open the latest week folder
$latestWeek = ($weeks.Keys | Sort-Object)[-1]
if ($latestWeek) {
    $m = $weeks[$latestWeek].Monday
    $folderName = "Livrable Semaine du $($m.Day) $($mois[$m.Month]) au $($m.AddDays(4).Day) $($mois[$m.AddDays(4).Month]) $($m.Year)"
    $latestDir = Join-Path $LivrablesRoot $folderName
    if (Test-Path $latestDir) {
        Start-Process explorer.exe -ArgumentList "`"$latestDir`""
    }
}
