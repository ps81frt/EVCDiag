<#
.SYNOPSIS
    Collecte et analyse les crashes, erreurs systeme, logs kernel et IO disques.
.DESCRIPTION
    Ce script :
    1. Recupere les crashes d'applications (Event ID 1000/1001)
    2. Extrait les erreurs systeme critiques (Event ID 41, 1001, etc.)
    3. Collecte les logs kernel (WHEA, Dump, Storport, etc.)
    4. Liste les disques avec leurs details materiels
    5. Analyse les IO > 10000ms via awk
    6. Exporte tout dans EVC_Export sur le bureau
.NOTES
    Auteur : ps81frt
    Lien   : https://github.com/ps81frt/EVC
.LICENSE
    MIT
#>

param(
    [switch]$Collect,
    [string]$TimeCreated,
    [string]$Port = "",
    [string]$Path = "",
    [string]$Guid = "",
    [switch]$List,
    [switch]$ListErrors,
    [switch]$Export,
    [Alias("h")]
    [switch]$Help
)

# =============================================
# 0. DOSSIER D'EXPORT COMMUN SUR LE BUREAU
# =============================================
$outputFolder = Join-Path $env:USERPROFILE "Desktop\EVC_Export"
$inputFile    = Join-Path $outputFolder "3_Kernel_Diagnostics.txt"
if (-not (Test-Path $outputFolder)) {
    New-Item -Path $outputFolder -ItemType Directory | Out-Null
}

if ($Help) {
    Write-Host ""
    Write-Host "=== EVCDiag.ps1 ==="
    Write-Host ""
    Write-Host "  .\EVCDiag.ps1 -Collect"
    Write-Host "  .\EVCDiag.ps1 -List"
    Write-Host "  .\EVCDiag.ps1 -ListErrors"
    Write-Host "  .\EVCDiag.ps1 -TimeCreated '10/04/2026 23:03:23'"
    Write-Host "  .\EVCDiag.ps1 -TimeCreated '10/04/2026 23:03:23' -Port 2"
    Write-Host "  .\EVCDiag.ps1 -TimeCreated '10/04/2026 23:03:23' -Port 2 -Path 0"
    Write-Host "  .\EVCDiag.ps1 -TimeCreated '10/04/2026 23:03:23' -Guid 7fd9e307"
    Write-Host "  .\EVCDiag.ps1 -TimeCreated '10/04/2026 23:03:23' -Guid 7fd9e307 -Export"
    Write-Host "  .\EVCDiag.ps1 -TimeCreated '08/04/2026 20:03:00' -Guid 7fd9e307 *>> ~\desktop\report.txt"
    Write-Host ""
    Write-Host "  -Collect     : collecte tous les logs + analyse IO_Errors auto"
    Write-Host "  -ListErrors  : exporte les IO >= 10000ms -> EVC_Export\IO_Errors.txt"
    Write-Host "                 installe les outils Linux dans System32 si absents"
    Write-Host ""
    exit
}

# =============================================
# INSTALL LINUX TOOLS
# =============================================
function Install-Awk {
    $awkDest = "$env:SystemRoot\System32\awk.exe"
    if (Test-Path $awkDest) { return $awkDest }

    $awkInPath = Get-Command awk -ErrorAction SilentlyContinue
    if ($awkInPath) { return $awkInPath.Source }

    Write-Host "[INFO] awk non trouve. Telechargement depuis GitHub..."

    $zipUrl = "https://github.com/ps81frt/EVC/raw/main/Evc/LinuxToolOn-Windows.zip"
    $tmpZip = Join-Path $env:TEMP "LinuxToolOn-Windows.zip"
    $tmpDir = Join-Path $env:TEMP "LinuxTools_EVC"

    try {
        Invoke-WebRequest -Uri $zipUrl -OutFile $tmpZip -UseBasicParsing -ErrorAction Stop
        Write-Host "[INFO] Telechargement OK -> $tmpZip"
    } catch {
        Write-Host "[ERREUR] Telechargement echoue : $_"
        return $null
    }

    if (Test-Path $tmpDir) { Remove-Item $tmpDir -Recurse -Force }
    Expand-Archive -Path $tmpZip -DestinationPath $tmpDir -Force

    $allBinaries = Get-ChildItem -Path $tmpDir -Recurse -Include "*.exe","*.dll"

    if (-not $allBinaries) {
        Write-Host "[ERREUR] Aucun binaire trouve dans l'archive."
        Get-ChildItem $tmpDir -Recurse | ForEach-Object { Write-Host "    $($_.FullName)" }
        return $null
    }

    $awkDestPath  = $null
    $installErrors = @()

    foreach ($bin in $allBinaries) {
        $dest = Join-Path "$env:SystemRoot\System32" $bin.Name
        try {
            Copy-Item $bin.FullName -Destination $dest -Force
            Write-Host "[OK] $($bin.Name)"
            if ($bin.Name -eq "awk.exe") { $awkDestPath = $dest }
        } catch {
            Write-Host "[WARN] $($bin.Name) -> $_"
            $installErrors += $bin
            if ($bin.Name -eq "awk.exe") { $awkDestPath = $bin.FullName }
        }
    }

    if ($installErrors.Count -gt 0) {
        Write-Host ""
        Write-Host "[WARN] $($installErrors.Count) fichier(s) non installe(s) (relancer en admin) :"
        $installErrors | ForEach-Object { Write-Host "  - $($_.Name)" }
    }

    if (-not $awkDestPath) {
        Write-Host "[ERREUR] awk.exe introuvable dans l'archive."
        return $null
    }

    return $awkDestPath
}

# =============================================
# LIST IO ERRORS > 10000ms
# =============================================
function Invoke-ListErrors($inputFile) {
    $awkBin = Install-Awk
    if (-not $awkBin) {
        Write-Host "[ERREUR] awk introuvable."
        return
    }

    $outFile = Join-Path $outputFolder "IO_Errors.txt"

    $awkScript = @'
/^TimeCreated :/ { bloc = $0; getline; while ($0 !~ /^TimeCreated :/ && !/^$/) { bloc = bloc "\n" $0; getline } }
bloc ~ /IO success counts are/ {
    match(bloc, /IO success counts are [0-9, ]+\./)
    line = substr(bloc, RSTART, RLENGTH)
    gsub(/[^0-9,]/, "", line)
    split(line, valeurs, ",")
    if (valeurs[13] != 0 || valeurs[14] != 0) {
        match(bloc, /TimeCreated : [0-9\/: ]+/); date = substr(bloc, RSTART, RLENGTH)
        match(bloc, /Guid is \{([^}]+)\}/); guid = substr(bloc, RSTART+9, RLENGTH-10)
        print date, guid, "->", valeurs[13], "IO en 10000 ms,", valeurs[14], "IO en 10000+ ms"
    }
}
'@

    $awkScriptFile = Join-Path $env:TEMP "evc_errors.awk"
    $awkScript | Set-Content $awkScriptFile -Encoding ASCII

    $result = & $awkBin -f $awkScriptFile $inputFile 2>&1

    if ($result) {
        $result | Out-File $outFile -Encoding UTF8
        $result | ForEach-Object { Write-Host $_ }
        Write-Host ""
        Write-Host "$($result.Count) evenement(s) -> $outFile"
    } else {
        Write-Host "[OK] Aucun IO >= 10000 ms detecte."
    }
}

# =============================================
# COLLECT
# =============================================
if ($Collect) {

    # =============================================
    # 1. CRASHES D'APPLICATIONS (Event ID 1000/1001)
    # =============================================
    $appCrashFile = Join-Path $outputFolder "1_Application_Crashes.txt"
    "===== CRASHES D'APPLICATIONS (Event ID 1000) =====" | Out-File $appCrashFile -Encoding UTF8
    Get-WinEvent -LogName "Application" | Where-Object {$_.Id -eq 1000} | Sort-Object TimeCreated |
        Select-Object TimeCreated,
            @{N="Application";E={$_.Properties[0].Value}},
            @{N="Version";E={$_.Properties[1].Value}},
            @{N="Module";E={$_.Properties[3].Value}},
            @{N="Code";E={$_.Properties[6].Value}},
            @{N="Offset";E={$_.Properties[7].Value}} |
        Format-List | Out-File $appCrashFile -Append -Encoding UTF8

    "`n===== ERREURS D'APPLICATIONS (Event ID 1001) =====" | Out-File $appCrashFile -Append -Encoding UTF8
    Get-WinEvent -LogName "Application" | Where-Object {$_.Id -eq 1001} | Sort-Object TimeCreated |
        Select-Object TimeCreated, @{N="Message";E={$_.Message}} |
        Format-List | Out-File $appCrashFile -Append -Encoding UTF8

    # =============================================
    # 2. ERREURS SYSTEME CRITIQUES
    # =============================================
    $systemCrashFile = Join-Path $outputFolder "2_System_Crashes.txt"
    "===== ERREURS SYSTEME (ID 41, 1001, 7023, 7034, 157, 153) =====" | Out-File $systemCrashFile -Encoding UTF8
    Get-WinEvent -LogName "System" | Where-Object {$_.Id -in @(41,1001,7023,7034,157,153)} | Sort-Object TimeCreated |
        Select-Object TimeCreated, Id, @{N="Message";E={$_.Message}} |
        Format-List | Out-File $systemCrashFile -Append -Encoding UTF8

    # =============================================
    # 3. LOGS KERNEL (WHEA, Dump, Storport, etc.)
    # =============================================
    $kernelDiagFile = Join-Path $outputFolder "3_Kernel_Diagnostics.txt"
    $logNames = @(
        "Microsoft-Windows-Kernel-WHEA/Operational",
        "Microsoft-Windows-Kernel-WHEA/Errors",
        "Microsoft-Windows-Kernel-Dump/Operational",
        "Microsoft-Windows-Diagnostics-Performance/Operational",
        "Microsoft-Windows-Resource-Exhaustion-Detector/Operational",
        "Microsoft-Windows-Kernel-PnP/Driver Watchdog",
        "Microsoft-Windows-Fault-Tolerant-Heap/Operational",
        "Microsoft-Windows-WerKernel/Operational",
        "Microsoft-Windows-CodeIntegrity/Operational",
        "Microsoft-Windows-Security-Mitigations/KernelMode",
        "Microsoft-Windows-Kernel-Boot/Operational",
        "Microsoft-Windows-Storage-Storport/Operational",
        "Microsoft-Windows-Ntfs/Operational"
    )

    "===== LOGS KERNEL (WHEA, Dump, Storport, etc.) =====" | Out-File $kernelDiagFile -Encoding UTF8
    foreach ($logName in $logNames) {
        try {
            Get-WinEvent -LogName $logName -ErrorAction Stop | Sort-Object TimeCreated |
                Select-Object TimeCreated, LogName, @{N="Message";E={$_.Message}} |
                Format-List | Out-File $kernelDiagFile -Append -Encoding UTF8
        }
        catch {
            Write-Warning "Impossible de lire $logName : $_"
        }
    }

    # =============================================
    # 4. INFORMATIONS DISQUES + GUID STORPORT + ERREURS + LETTRES + OBJECTID
    # =============================================
    $diskInfoFile = Join-Path $outputFolder "4_Disk_Information.txt"
    "===== INFORMATIONS MATERIELLES DES DISQUES =====" | Out-File $diskInfoFile -Encoding UTF8

    $physDisks   = Get-PhysicalDisk
    $reliability = $physDisks | Get-StorageReliabilityCounter
    $volumes     = Get-Volume | Where-Object { $_.DriveLetter } | Select-Object DiskNumber, DriveLetter

    $diskInfo = foreach ($disk in $physDisks) {
        $rel = $reliability | Where-Object { $_.DeviceId -eq $disk.DeviceId }

        $objectId = if ($rel) { $rel.UniqueId } else { "N/A" }

        $storportGuid = "N/A"
        if ($rel -and $rel.UniqueId) {
            $allGuids = [regex]::Matches($rel.UniqueId, '\{[0-9A-Fa-f-]+\}')
            if ($allGuids.Count -gt 0) {
                $storportGuid = $allGuids[-1].Value.Trim('{}')
            }
        }

        $driveLetters = ($volumes | Where-Object { $_.DiskNumber -eq $disk.DeviceId -and $_.DriveLetter }).DriveLetter -join ","
        if ([string]::IsNullOrEmpty($driveLetters)) {
            $parts = Get-Partition -DiskNumber $disk.DeviceId -ErrorAction SilentlyContinue | Where-Object DriveLetter
            if ($parts) { $driveLetters = $parts.DriveLetter -join "," }
            else { $driveLetters = "Aucune" }
        }
        if ([string]::IsNullOrEmpty($driveLetters)) { $driveLetters = "Aucune" }

        [PSCustomObject]@{
            DiskNumber             = $disk.DeviceId
            DriveLetter            = $driveLetters
            Name                   = $disk.FriendlyName
            BusType                = $disk.BusType
            SerialNumber           = $disk.SerialNumber
            SizeGB                 = [math]::Round($disk.Size / 1GB, 2)
            HealthStatus           = $disk.HealthStatus
            ReadErrorsUncorrected  = if ($rel) { $rel.ReadErrorsUncorrected  } else { 0 }
            WriteErrorsUncorrected = if ($rel) { $rel.WriteErrorsUncorrected } else { 0 }
            ReadLatencyMax_ms      = if ($rel) { $rel.ReadLatencyMax  } else { 0 }
            WriteLatencyMax_ms     = if ($rel) { $rel.WriteLatencyMax } else { 0 }
            WearPercent            = if ($rel) { $rel.Wear        } else { 0 }
            Temperature_C          = if ($rel) { $rel.Temperature } else { 0 }
            ObjectId               = $objectId
            StorportGuid           = $storportGuid
        }
    }

    $diskInfo | Format-List | Out-File $diskInfoFile -Append -Encoding UTF8

    # =============================================
    # 4.1. MAP DISK -> VOLUMES (via GUID)
    # =============================================
    $partitions = Get-Partition | Where-Object { $_.DiskNumber -ne $null }
    $volumes    = Get-Volume | Where-Object { $_.DriveLetter -or $_.FileSystem -eq 'NTFS' }

    Write-Host "`n===== MAPPING DISQUE PHYSIQUE (GUID) -> VOLUMES =====" -ForegroundColor Cyan
    "`n===== MAPPING DISQUE PHYSIQUE (GUID) -> VOLUMES =====" | Out-File $diskInfoFile -Append -Encoding UTF8

    foreach ($disk in $physDisks) {
        $rel = $reliability | Where-Object { $_.DeviceId -eq $disk.DeviceId }

        $storportGuid = "N/A"
        if ($rel -and $rel.UniqueId) {
            $allGuids = [regex]::Matches($rel.UniqueId, '\{[0-9A-Fa-f-]+\}')
            if ($allGuids.Count -gt 0) {
                $storportGuid = $allGuids[-1].Value.Trim('{}')
            }
        }

        $driveNumber      = $disk.DeviceId
        $physicalDrivePath = "\\.\PhysicalDrive$driveNumber"

        $diskPartitions = $partitions | Where-Object { $_.DiskNumber -eq $driveNumber }
        $volumesOnDisk  = @()
        foreach ($part in $diskPartitions) {
            $vol = $volumes | Where-Object { $_.Partition -and $_.Partition.DiskNumber -eq $driveNumber -and $_.Partition.PartitionNumber -eq $part.PartitionNumber }
            if (-not $vol) {
                $vol = $volumes | Where-Object { $_.UniqueId -like "*$($part.Guid)*" }
            }
            if ($vol) {
                $volumesOnDisk += [PSCustomObject]@{
                    DriveLetter = if ($vol.DriveLetter) { "$($vol.DriveLetter):" } else { "Aucune" }
                    VolumeGuid  = ($vol.UniqueId -replace '.*\\','')
                    SizeGB      = [math]::Round($vol.Size / 1GB, 2)
                    FileSystem  = $vol.FileSystem
                }
            }
        }

        if ($volumesOnDisk.Count -eq 0) {
            $volumesOnDisk = [PSCustomObject]@{ DriveLetter = "Aucun volume trouve (peut-etre non monte)"; VolumeGuid = "N/A"; SizeGB = "N/A"; FileSystem = "N/A" }
        }

        Write-Host "`n[ DISQUE $driveNumber ] : $($disk.FriendlyName)" -ForegroundColor Yellow
        Write-Host "  GUID Storport  : $storportGuid"
        Write-Host "  Chemin disque  : $physicalDrivePath"
        Write-Host "  Volumes sur ce disque :"
        $volumesOnDisk | Format-Table DriveLetter, VolumeGuid, SizeGB, FileSystem -AutoSize

        "`n[ DISQUE $driveNumber ] : $($disk.FriendlyName)"     | Out-File $diskInfoFile -Append -Encoding UTF8
        "  GUID Storport  : $storportGuid"                      | Out-File $diskInfoFile -Append -Encoding UTF8
        "  Chemin disque  : $physicalDrivePath"                 | Out-File $diskInfoFile -Append -Encoding UTF8
        "  Volumes sur ce disque :"                             | Out-File $diskInfoFile -Append -Encoding UTF8
        $volumesOnDisk | Format-Table DriveLetter, VolumeGuid, SizeGB, FileSystem -AutoSize | Out-File $diskInfoFile -Append -Encoding UTF8
    }

    $candidateDisks = $diskInfo | Where-Object { $_.ReadErrorsUncorrected -gt 0 -or $_.WriteErrorsUncorrected -gt 0 }
    if (-not $candidateDisks) {
        $candidateDisks = $diskInfo | Where-Object { $_.ReadLatencyMax_ms -ge 100 -or $_.WriteLatencyMax_ms -ge 100 }
    }
    if (-not $candidateDisks) {
        $candidateDisks = $diskInfo | Sort-Object -Property @{Expression={($_.ReadLatencyMax_ms + $_.WriteLatencyMax_ms)}} -Descending | Select-Object -First 1
    }

    Write-Host "`n===== IDENTIFICATION DU(DES) DISQUE(S) POTENTIELLEMENT DEFAILLANT(S) =====" -ForegroundColor Red
    if ($candidateDisks) {
        foreach ($badDisk in $candidateDisks) {
            Write-Host " GUID Storport  : $($badDisk.StorportGuid)" -ForegroundColor Red
            Write-Host " DeviceId       : $($badDisk.DiskNumber)" -ForegroundColor Red
            Write-Host " Nom            : $($badDisk.Name)" -ForegroundColor Red
            Write-Host " Chemin disc.   : \\.\PhysicalDrive$($badDisk.DiskNumber)" -ForegroundColor Red
            Write-Host " DriveLetter(s) : $($badDisk.DriveLetter)" -ForegroundColor Red
            Write-Host " Erreurs lecture : $($badDisk.ReadErrorsUncorrected) | Erreurs ecriture : $($badDisk.WriteErrorsUncorrected)" -ForegroundColor Red
            Write-Host " Latence lecture max : $($badDisk.ReadLatencyMax_ms) ms | Latence ecriture max : $($badDisk.WriteLatencyMax_ms) ms" -ForegroundColor Red
            Write-Host " ⚠️  Verifiez immediatement ce disque si son GUID Storport correspond a l'evenement de 3_Kernel_Diagnostics.txt." -ForegroundColor Red
            Write-Host ""
        }
    } else {
        Write-Host "Aucun disque specifique ne presente de signes clairs de defaillance." -ForegroundColor Yellow
    }

    $problematic = $diskInfo | Where-Object { $_.ReadErrorsUncorrected -gt 0 -or $_.WriteErrorsUncorrected -gt 0 }
    if ($problematic) {
        "`n===== DISQUES AVEC ERREURS NON CORRIGEES =====" | Out-File $diskInfoFile -Append -Encoding UTF8
        $problematic | Format-List DiskNumber, DriveLetter, Name, ReadErrorsUncorrected, WriteErrorsUncorrected, ReadLatencyMax_ms, ObjectId, StorportGuid | Out-File $diskInfoFile -Append -Encoding UTF8
    } else {
        "`n===== AUCUN DISQUE AVEC ERREURS DETECTEES =====" | Out-File $diskInfoFile -Append -Encoding UTF8
    }

    Write-Host ""
    if ($problematic) {
        Write-Host "⚠️  Disques avec erreurs non corrigees :" -ForegroundColor Red
        $problematic | Format-List DiskNumber, DriveLetter, Name, ReadErrorsUncorrected, StorportGuid, ObjectId
    } else {
        Write-Host "✅ Aucun disque avec erreurs non corrigees detecte." -ForegroundColor Green
    }

    # =============================================
    # 5. IO ERRORS AUTO
    # =============================================
    $ioErrorsFile = Join-Path $outputFolder "IO_Errors.txt"
    Write-Host ""
    Write-Host "===== ANALYSE IO > 10000ms ====="
    Invoke-ListErrors $kernelDiagFile

    # =============================================
    # 6. OUVERTURE DES FICHIERS GENERES
    # =============================================
    notepad $appCrashFile
    notepad $systemCrashFile
    notepad $kernelDiagFile
    notepad $diskInfoFile
    if (Test-Path $ioErrorsFile) { notepad $ioErrorsFile }

    Write-Host "Diagnostics generes dans : $outputFolder`n" -ForegroundColor Green
    Write-Host "1. $appCrashFile"
    Write-Host "2. $systemCrashFile"
    Write-Host "3. $kernelDiagFile"
    Write-Host "4. $diskInfoFile"
    Write-Host "5. $ioErrorsFile"

    exit
}

# =============================================
# READER
# =============================================
if (-not (Test-Path $inputFile)) {
    Write-Host "[ERREUR] Fichier $inputFile introuvable. Lancez d'abord : .\EVCDiag.ps1 -Collect"
    exit
}

if ($ListErrors) {
    Invoke-ListErrors $inputFile
    exit
}

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

if ($List) {
    foreach ($bloc in $blocs) {
        if ($bloc -match "TimeCreated\s*:\s*(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})") {
            Write-Host $matches[1]
        }
    }
    exit
}

if (-not $TimeCreated) {
    Write-Host "[ERREUR] Aucun parametre. Utilisez -Help pour l'aide."
    exit
}

$selectedBlocs = $blocs | Where-Object { $_ -like "*$TimeCreated*" -and $_ -match 'Performance summary for Storport Device' }

if ($selectedBlocs.Count -eq 0) {
    Write-Host "[ERREUR] Aucun evenement pour la date $TimeCreated"
    exit
}

function Get-DeviceInfo($bloc) {
    $port  = if ($bloc -match "Port = (\d+),")  { $matches[1] } else { "" }
    $path  = if ($bloc -match "Path = (\d+),")  { $matches[1] } else { "" }
    $guid  = if ($bloc -match "(?:Guid is|Corresponding Class Disk Device Guid is) \{([^}]+)\}") { $matches[1] } else { "" }
    $total = if ($bloc -match "Total IO:(\d+)") { $matches[1] } else { "0" }
    return [pscustomobject]@{
        Port  = $port
        Path  = $path
        Guid  = $guid
        Total = $total
        Bloc  = $bloc
    }
}

$devices = $selectedBlocs | ForEach-Object { Get-DeviceInfo $_ }
$devices = $devices | Group-Object -Property { "$($_.Guid)|$($_.Port)|$($_.Path)" } | ForEach-Object { $_.Group[-1] }

if ($Port -ne "") {
    $devices = $devices | Where-Object { $_.Port -eq $Port }
    if (@($devices).Count -eq 0) { Write-Host "[ERREUR] Aucun peripherique avec Port=$Port"; exit }
}
if ($Path -ne "") {
    $devices = $devices | Where-Object { $_.Path -eq $Path }
    if (@($devices).Count -eq 0) { Write-Host "[ERREUR] Aucun peripherique avec Path=$Path"; exit }
}
if ($Guid -ne "") {
    $cleanGuid = $Guid -replace '[{}]', ''
    $devices = $devices | Where-Object { $_.Guid -like "*$cleanGuid*" }
    if (@($devices).Count -eq 0) { Write-Host "[ERREUR] Aucun peripherique avec GUID contenant '$cleanGuid'"; exit }
}

if (@($devices).Count -gt 1) {
    Write-Host ""
    Write-Host "[INFO] $($devices.Count) peripheriques trouves. Precisez -Port, -Path ou -Guid :"
    Write-Host ""
    foreach ($d in $devices) {
        Write-Host ("  -Port {0} -Path {1}  | TotalIO={2} | -Guid {3}" -f $d.Port, $d.Path, $d.Total, $d.Guid)
    }
    Write-Host ""
    Write-Host "Exemple : .\EVCDiag.ps1 -TimeCreated '$TimeCreated' -Port 2 -Path 0"
    Write-Host "          .\EVCDiag.ps1 -TimeCreated '$TimeCreated' -Guid 7fd9e307"
    exit
}

$selectedEntry = $devices[0].Bloc

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

$max = 14
while ($success.Count -lt $max) { $success += 0 }
while ($failed.Count -lt $max)  { $failed  += 0 }
while ($latency.Count -lt $max) { $latency += 0 }

$sumSuccess = ($success | Measure-Object -Sum).Sum

$guid    = ([regex]::Match($selectedEntry, "(?:Guid is|Corresponding Class Disk Device Guid is) \{(.*?)\}")).Groups[1].Value
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

$lossWarning = ""
if ($sumSuccess -ne $totalIO) {
    $lossWarning = "[WARN] MISMATCH: success sum ($sumSuccess) != Total IO ($totalIO)"
}

if ($lossWarning -ne "") {
    Write-Host ""
    Write-Host $lossWarning
}

$logNameValue = ([regex]::Match($selectedEntry, "LogName\s*:\s*(.+)")).Groups[1].Value
$messageValue = ([regex]::Match($selectedEntry, "Message\s*:\s*(.+)")).Groups[1].Value
$messageValue = $messageValue -replace 'whose Corresponding Class Disk Device Guid is \{[^}]+\}:?', ''
$messageValue = $messageValue.Trim()

function Write-WrappedLine($label, $text, $width) {
    $prefix = "| $label"
    $max    = $width - $prefix.Length - 1
    $words  = $text -split ' '
    $line   = $prefix
    foreach ($word in $words) {
        if ($line.Length + 1 + $word.Length -gt $width) {
            Write-Host $line
            $line = "| " + (' ' * ($label.Length)) + " $word"
        } else {
            $line += " $word"
        }
    }
    Write-Host $line
}

if ($Export) {
    $safeDate   = $TimeCreated -replace '[/: ]', '-'
    $safeGuid   = ($guid -replace '[{}]', '').Substring(0, [math]::Min(8, $guid.Length))
    $reportFile = Join-Path $outputFolder "Report_${safeDate}_${safeGuid}.txt"
    Start-Transcript -Path $reportFile -Force | Out-Null
}

Write-Host ""
Write-Host "+---------------------------------------------------------------+"
Write-Host "| TimeCreated : $TimeCreated"
if ($logNameValue) { Write-WrappedLine "LogName     :" $logNameValue 104 }
if ($messageValue) { Write-WrappedLine "Message     :" $messageValue 104 }
Write-Host "| Guid        : {$guid}"
Write-Host "+---------------------------------------------------------------+"
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

    Write-Host ("| {0,-14} | {1,14:N0} | {2,14} | {3,14} | {4,14:N0} | {5,-18} |" -f $bucketLabels[$i], $ok, $pct, $lat, $fail, $status)
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

# =============================================
# EXPORT
# =============================================
if ($Export) {
    Stop-Transcript | Out-Null
    Write-Host ""
    Write-Host "[EXPORT] $reportFile"
}

<#

.TROUVER BLOCK BUCKET ++ 10000ms +

awk '
/^TimeCreated :/ { bloc = $0; getline; while ($0 !~ /^TimeCreated :/ && !/^$/) { bloc = bloc "\n" $0; getline } }
bloc ~ /IO success counts are/ {
    match(bloc, /IO success counts are [0-9, ]+\./)
    line = substr(bloc, RSTART, RLENGTH)
    gsub(/[^0-9,]/, "", line)
    split(line, valeurs, ",")
    if (valeurs[13] != 0 || valeurs[14] != 0) {
        match(bloc, /TimeCreated : [0-9\/: ]+/); date = substr(bloc, RSTART, RLENGTH)
        match(bloc, /Guid is \{([^}]+)\}/); guid = substr(bloc, RSTART+9, RLENGTH-10)
        print date, guid, "->", valeurs[13], "IO en 10000 ms,", valeurs[14], "IO en 10000+ ms"
    }
}
' $env:USERPROFILE\Desktop\EVC_Export\3_Kernel_Diagnostics.txt

#>
