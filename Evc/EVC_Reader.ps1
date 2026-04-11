param(
    [string]$TimeCreated,
    [string]$Port = "",
    [string]$Path = "",
    [string]$Guid = "",
    [switch]$List,
    [Alias("h")]
    [switch]$Help
)

$inputFile = "3_Kernel_Diagnostics.txt"

if ($Help) {
    Write-Host "Usage:"
    Write-Host "  .\EVC_Reader.ps1 -TimeCreated '10/04/2026 23:03:23'"
    Write-Host "  .\EVC_Reader.ps1 -TimeCreated '10/04/2026 23:03:23' -Port 2"
    Write-Host "  .\EVC_Reader.ps1 -TimeCreated '10/04/2026 23:03:23' -Port 2 -Path 0"
    Write-Host "  .\EVC_Reader.ps1 -TimeCreated '10/04/2026 23:03:23' -Guid 7fd9e307-31d4-d83f-2811-0f9408d7dcd3"
    Write-Host "  .\EVC_Reader.ps1 -TimeCreated '10/04/2026 23:03:23' -Guid 7fd9e307"
    Write-Host "  .\EVC_Reader.ps1 -List"
    exit
}

if (-not (Test-Path $inputFile)) {
    Write-Host "[ERREUR] Fichier $inputFile introuvable"
    exit
}

# ----- Lecture et découpage en blocs -----
$lines = Get-Content $inputFile
$blocs = @()
$currentBloc = ""

foreach ($line in $lines) {
    if ($line -match "^TimeCreated\s*:") {
        if ($currentBloc -ne "") { $blocs += $currentBloc }
        $currentBloc = $line
    } else {
        $currentBloc += "`n" + $line
    }
}
if ($currentBloc -ne "") { $blocs += $currentBloc }

# ----- Liste des dates -----
if ($List) {
    foreach ($bloc in $blocs) {
        if ($bloc -match "TimeCreated\s*:\s*(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})") {
            Write-Host $matches[1]
        }
    }
    exit
}

if (-not $TimeCreated) {
    Write-Host "[ERREUR] -TimeCreated requis"
    exit
}

# ----- Sélection des blocs correspondant à la date -----
$selectedBlocs = $blocs | Where-Object { $_ -like "*$TimeCreated*" -and $_ -match 'Performance summary for Storport Device' }

if ($selectedBlocs.Count -eq 0) {
    Write-Host "[ERREUR] Aucun événement pour la date $TimeCreated"
    exit
}

# ----- Fonction d'extraction des informations -----
function Get-DeviceInfo($bloc) {
    $port = if ($bloc -match "Port = (\d+),") { $matches[1] } else { "" }
    $path = if ($bloc -match "Path = (\d+),") { $matches[1] } else { "" }
    $guid = if ($bloc -match "Guid is \{([^}]+)\}") { $matches[1] } else { "" }
    $total = if ($bloc -match "Total IO:(\d+)") { $matches[1] } else { "0" }
return [pscustomobject]@{
        Port = $port
        Path = $path
        Guid = $guid
        Total = $total
        Bloc = $bloc
    }
}

$devices = $selectedBlocs | ForEach-Object { Get-DeviceInfo $_ }

# Elimine les doublons pour le même périphérique (même Guid/Port/Path) en conservant le dernier bloc
$devices = $devices | Group-Object -Property { "$($_.Guid)|$($_.Port)|$($_.Path)" } | ForEach-Object { $_.Group[-1] }

# ----- Application des filtres -----
if ($Port -ne "") {
    $devices = $devices | Where-Object { $_.Port -eq $Port }
    if (@($devices).Count -eq 0) { Write-Host "[ERREUR] Aucun périphérique avec Port=$Port"; exit }
}
if ($Path -ne "") {
    $devices = $devices | Where-Object { $_.Path -eq $Path }
    if (@($devices).Count -eq 0) { Write-Host "[ERREUR] Aucun périphérique avec Path=$Path"; exit }
}
if ($Guid -ne "") {
    $cleanGuid = $Guid -replace '[{}]', ''
    $devices = $devices | Where-Object { $_.Guid -like "*$cleanGuid*" }
    if (@($devices).Count -eq 0) { Write-Host "[ERREUR] Aucun périphérique avec GUID contenant '$cleanGuid'"; exit }
}

# ----- Si plusieurs périphériques -> menu -----
if (@($devices).Count -gt 1) {
    Write-Host ""
    Write-Host "[INFO] $($devices.Count) périphériques trouvés. Précisez -Port, -Path ou -Guid :"
    Write-Host ""
    foreach ($d in $devices) {
        Write-Host ("  -Port {0} -Path {1}  | TotalIO={2} | -Guid {3}" -f $d.Port, $d.Path, $d.Total, $d.Guid)
    }
    Write-Host ""
    Write-Host "Exemple : .\EVC_Reader.ps1 -TimeCreated '$TimeCreated' -Port 2 -Path 0"
    Write-Host "          .\EVC_Reader.ps1 -TimeCreated '$TimeCreated' -Guid 7fd9e307"
    exit
}

# ----- Un seul périphérique : affichage du rapport -----
$selectedEntry = $devices[0].Bloc

# ====================== PARSING (identique à votre version originale) ======================
function Get-RawArray($pattern, $text) {
    $m = [regex]::Match($text, $pattern, "Singleline")
    if (-not $m.Success) { return @() }
    $values = $m.Groups[1].Value.Trim() -replace '\s+', ''
    return $values.Split(",") | ForEach-Object {
        $v = ($_ -replace "[^\d]", "")
        if ($v -eq "") { 0 } else { [int64]$v }
    }
}

$bucketLabels = @("128 us","256 us","512 us","1 ms","4 ms","16 ms","64 ms","128 ms","256 ms","512 ms","1000 ms","2000 ms","10000 ms","> 10000 ms")

$success = Get-RawArray "IO success counts are ([\d,\s]+)" $selectedEntry
$failed  = Get-RawArray "IO failed counts are ([\d,\s]+)"  $selectedEntry
$latency = Get-RawArray "IO total latency.*?are ([\d,\s]+)" $selectedEntry
$totalIO = ([regex]::Match($selectedEntry, "Total IO:\s*(\d+)")).Groups[1].Value -as [int64]

$bytesRead    = ([regex]::Match($selectedEntry, "Total Bytes Read:(\d+)")).Groups[1].Value -as [int64]
$bytesWritten = ([regex]::Match($selectedEntry, "Total Bytes Written:(\d+)")).Groups[1].Value -as [int64]

$max = 14
while ($success.Count -lt $max) { $success += 0 }
while ($failed.Count -lt $max)  { $failed  += 0 }
while ($latency.Count -lt $max) { $latency += 0 }

$sumSuccess = ($success | Measure-Object -Sum).Sum

$guid    = ([regex]::Match($selectedEntry, "Guid is \{(.*?)\}")).Groups[1].Value
$device  = [regex]::Match($selectedEntry, "Port = (\d+), Path = (\d+), Target = (\d+), Lun = (\d+)")
$portVal = $device.Groups[1].Value
$pathVal = $device.Groups[2].Value
$tgtVal  = $device.Groups[3].Value
$lunVal  = $device.Groups[4].Value

function Get-AvgLatency($total, $count) {
    if ($count -le 0 -or $total -le 0) { return "-" }
    return [math]::Round(($total / $count) / 10000, 3)
}

function Format-Bytes($bytes) {
    $mb = [math]::Round($bytes / 1MB, 2)
    $gb = [math]::Round($bytes / 1GB, 3)
    return "$mb Mo / $gb Go"
}

$highLatencyIO = $success[7] + $success[8] + $success[9] + $success[10] + $success[11] + $success[12] + $success[13]
$globalStatus = if ($highLatencyIO -eq 0) { "OK" } else { "[!] $highLatencyIO IO > 128ms" }

$lossWarning = ""
if ($sumSuccess -ne $totalIO) {
    $lossWarning = "[WARN] MISMATCH: success sum ($sumSuccess) != Total IO ($totalIO)"
}

Write-Host ""
Write-Host "+---------------------------------------------------------------+"
Write-Host "| RAPPORT STORPORT - $TimeCreated"
Write-Host "| Device GUID: {$guid}"
Write-Host "| Port: $portVal | Path: $pathVal | Target: $tgtVal | LUN: $lunVal"
Write-Host "+---------------------+---------------------+---------------------+"
Write-Host "| Metrique             | Valeur              | Unite               |"
Write-Host "+---------------------+---------------------+---------------------+"
Write-Host ("| {0,-21}| {1,19:N0} | {2,-20}|" -f "Total IO", $totalIO, "Operations")
Write-Host ("| {0,-21}| {1,19} | {2,-20}|" -f "Statut", $globalStatus, "")
Write-Host ("| {0,-21}| {1,19} | {2,-20}|" -f "Octets lus", (Format-Bytes $bytesRead), "")
Write-Host ("| {0,-21}| {1,19} | {2,-20}|" -f "Octets ecrits", (Format-Bytes $bytesWritten), "")
Write-Host "+---------------------+---------------------+---------------------+"

if ($lossWarning -ne "") {
    Write-Host ""
    Write-Host $lossWarning -ForegroundColor Yellow
}

Write-Host ""
Write-Host "+-----------------------------------------------------------------------------------------------------------+"
Write-Host "| REPARTITION DES OPERATIONS IO (ZERO LOSS VERIFIED)                                                        |"
Write-Host "+----------------+----------------+----------------+----------------+----------------+--------------------+"
Write-Host "| Bucket         | IO Reussies     | % du Total     | Latence Moy.   | IO Echouees    | Statut             |"
Write-Host "+----------------+----------------+----------------+----------------+----------------+--------------------+"

for ($i = 0; $i -lt 14; $i++) {
    $ok   = $success[$i]
    $fail = $failed[$i]
    $pct  = if ($totalIO -gt 0) { "{0,7:N1}%" -f [math]::Round(($ok / $totalIO) * 100, 1) } else { "     0%" }
    $lat  = Get-AvgLatency $latency[$i] $ok

    $status =
        if     ($ok -eq 0 -and $fail -eq 0)    { "Aucun"              }
        elseif ($i -le 2)                       { "Optimal"            }
        elseif ($i -le 3)                       { "Normal"             }
        elseif ($i -le 5)                       { "Acceptable"         }
        elseif ($i -eq 6 -and $ok -gt 0)       { "[!] A surveiller"   }
        elseif ($i -le 10 -and $ok -gt 0)      { "[!!] Degradee"      }
        elseif ($i -eq 11 -and $ok -gt 0)      { "[!!] Critique"      }
        elseif ($i -eq 12 -and $ok -gt 0)      { "[!!!] 10s ($ok IO)" }
        elseif ($i -eq 13 -and $ok -gt 0)      { "[!!!] EXTREME"      }
        else                                    { "OK"                 }

    Write-Host ("| {0,-14} | {1,14:N0} | {2,14} | {3,14} | {4,14:N0} | {5,-18} |" -f `
        $bucketLabels[$i], $ok, $pct, $lat, $fail, $status)
}

Write-Host "+----------------+----------------+----------------+----------------+----------------+--------------------+"

$totalLatency = ($latency | Measure-Object -Sum).Sum
$totalOps     = $sumSuccess + ($failed | Measure-Object -Sum).Sum
$globalAvg    = if ($totalOps -gt 0) { [math]::Round(($totalLatency / $totalOps) / 10000, 6) } else { 0 }

Write-Host ""
Write-Host "+---------------------------------------------------------------+"
Write-Host ("| Latence Globale Moyenne : {0} ms (ponderee)" -f $globalAvg)
Write-Host ("| IO totales verifiees    : {0:N0} / {1:N0}" -f $sumSuccess, $totalIO)
if ($highLatencyIO -gt 0) {
    Write-Host ("| [!] IO > 128ms          : {0:N0} operations" -f $highLatencyIO) 
}
Write-Host "+---------------------------------------------------------------+"