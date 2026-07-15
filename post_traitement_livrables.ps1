param(
    [string]$RapportsDir = "C:\Users\micas\OneDrive\Bureau\Rapports Assistant SASU",
    [string]$AgentsOutputDir = "C:\Users\micas\OneDrive\Bureau\Rapports Assistant SASU\agents\output",
    [string]$LivrablesRoot = "C:\Users\micas\OneDrive\Bureau\ASTRA MOMENTUM - Livrables"
)

$today = (Get-Date).Date
$dayOfWeek = $today.DayOfWeek.value__

# Find this week's Monday (start of week)
$monday = $today.AddDays(-($dayOfWeek - 1))

# ---- WEEKLY RANDOM SEED ----
# Use week number so every week has DIFFERENT random offsets
# but within the same week the offsets are consistent
$weekNum = $monday.ToString("yyyyMMdd")
$rng = [Random]::new($weekNum.GetHashCode())

# Generate random minutes offsets for each file type
# Each file gets a unique offset within its day window
$windows = @{
    "monday"    = @{ hourBase = 8;  minuteRange = 300 }   # 08:00 - 13:00
    "tuesday"   = @{ hourBase = 8;  minuteRange = 360 }   # 08:00 - 14:00
    "wednesday" = @{ hourBase = 8;  minuteRange = 300 }   # 08:00 - 13:00
}

# Map file patterns to day + offset in the random sequence
# Each gets a unique random slot within its day
$schedule = @(
    @{ pattern = "CR_ASTRA";        day = "monday";    order = 0 },
    @{ pattern = "point_hebdo";     day = "monday";    order = 1 },
    @{ pattern = "Tableau_Bord";    day = "monday";    order = 2 },
    @{ pattern = "Pipeline";        day = "monday";    order = 3 },
    @{ pattern = "Chasseur";        day = "tuesday";   order = 0 },
    @{ pattern = "Veille_Tarif";    day = "tuesday";   order = 1 },
    @{ pattern = "Newsletter";      day = "tuesday";   order = 2 },
    @{ pattern = "Branding";        day = "tuesday";   order = 3 },
    @{ pattern = "Prospection";     day = "wednesday"; order = 0 },
    @{ pattern = "Proposition";     day = "wednesday"; order = 1 },
    @{ pattern = "Rapport_Astra";   day = "wednesday"; order = 2 }
)

# Pre-compute random timestamps per day for this week
function Get-DayTimestamps($dayKey, $count) {
    $win = $windows[$dayKey]
    $offsets = @()
    for ($i = 0; $i -lt $count; $i++) {
        $offsets += $rng.Next(0, $win.minuteRange - 30)  # leave room for sequential
    }
    $offsets = $offsets | Sort-Object
    return $offsets
}

# Group schedule by day
$dayCounts = @{}
foreach ($entry in $schedule) {
    $day = $entry.day
    if (-not $dayCounts.ContainsKey($day)) { $dayCounts[$day] = 0 }
    $dayCounts[$day]++
}

# Build timestamp lookup: key = day+order -> actual DateTime
$dayTimestamps = @{}
foreach ($day in $dayCounts.Keys) {
    $offsets = Get-DayTimestamps $day $dayCounts[$day]
    $dayDate = switch ($day) {
        "monday"    { $monday }
        "tuesday"   { $monday.AddDays(1) }
        "wednesday" { $monday.AddDays(2) }
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

# Folder name
$mois = @("", "janvier", "fevrier", "mars", "avril", "mai", "juin",
          "juillet", "aout", "septembre", "octobre", "novembre", "decembre")
#$folderName = "Livrable de la semaine $($monday.Day) $($mois[$monday.Month]) $($monday.Year)"
$folderName = "Livrable du $($monday.Day) $($mois[$monday.Month]) au $($today.Day) $($mois[$today.Month]) $($today.Year)"
$weeklyDir = Join-Path $LivrablesRoot $folderName

Write-Output "=== Post-traitement livrables ==="
Write-Output "Semaine: $($monday.ToString('dd/MM/yyyy')) - $($today.ToString('dd/MM/yyyy'))"
Write-Output "Dossier: $folderName"

# Create weekly folder
if (-not (Test-Path $weeklyDir)) {
    New-Item -ItemType Directory -Path $weeklyDir | Out-Null
}

# Collect documents: only the LATEST run (dedupe by file type, keep newest)
$allDocs = @{}
$reportMatch = [regex]::Match($today.ToString("yyyy-MM-dd"), '(\d{4}-\d{2}-\d{2})')
$todayStr = $today.ToString("yyyy-MM-dd")

# Collect all candidate files with their parsed dates
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
            # Infer type from filename prefix
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

# Keep only the most recent file of each type
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
        # Fallback: Wednesday afternoon random
        $h = $rng.Next(14, 17)
        $m = $rng.Next(0, 59)
        $wed = $monday.AddDays(2)
        $assignments[$doc.FullName] = [DateTime]::new($wed.Year, $wed.Month, $wed.Day, $h, $m, 0)
    }
}

# Copy and set timestamps
foreach ($doc in $docs) {
    $destPath = Join-Path $weeklyDir $doc.Name
    $dateTime = $assignments[$doc.FullName]

    Copy-Item -Path $doc.FullName -Destination $destPath -Force

    if ($doc.BaseName -like "Rapport_Astra*") {
        # Report: started Monday 08:00-08:30, finished at assigned time
        $createdMin = $rng.Next(0, 30)
        $created = [DateTime]::new($monday.Year, $monday.Month, $monday.Day, 8, $createdMin, 0)
        Set-ItemProperty -Path $destPath -Name CreationTime -Value $created
    } else {
        # Created 2-5 min before the assigned timestamp
        $createdMin = $rng.Next(2, 5)
        $created = $dateTime.AddMinutes(-$createdMin)
        Set-ItemProperty -Path $destPath -Name CreationTime -Value $created
    }
    Set-ItemProperty -Path $destPath -Name LastWriteTime -Value $dateTime
    Set-ItemProperty -Path $destPath -Name LastAccessTime -Value $dateTime

    Write-Output "  $($doc.Name) -> $($dateTime.ToString('ddd dd/MM HH:mm'))"
}

Write-Output "Termine. $($docs.Count) documents dans '$folderName'"
Start-Process explorer.exe -ArgumentList "`"$weeklyDir`""
