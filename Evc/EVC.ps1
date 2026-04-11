<#
.SYNOPSIS
    Génère un rapport complet des crashes, erreurs système, logs kernel et informations disques.
.DESCRIPTION
    Ce script :
    1. Récupère les crashes d'applications (Event ID 1000/1001)
    2. Extrait les erreurs système critiques (Event ID 41, 1001, etc.)
    3. Collecte les logs kernel (WHEA, Dump, Storport, etc.)
    4. Liste les disques avec leurs détails matériels
    5. Exporte tout dans des fichiers structurés sur le bureau.
#>

# =============================================
# 1. CRASHES D'APPLICATIONS (Event ID 1000/1001)
# =============================================
$appCrashFile = "$env:USERPROFILE\Desktop\1_Application_Crashes.txt"
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
# 2. ERREURS SYSTÈME CRITIQUES
# =============================================
$systemCrashFile = "$env:USERPROFILE\Desktop\2_System_Crashes.txt"
"===== ERREURS SYSTÈME (ID 41, 1001, 7023, 7034, 157, 153) =====" | Out-File $systemCrashFile -Encoding UTF8
Get-WinEvent -LogName "System" | Where-Object {$_.Id -in @(41,1001,7023,7034,157,153)} | Sort-Object TimeCreated |
    Select-Object TimeCreated, Id, @{N="Message";E={$_.Message}} |
    Format-List | Out-File $systemCrashFile -Append -Encoding UTF8

# =============================================
# 3. LOGS KERNEL (WHEA, Dump, Storport, etc.)
# =============================================
$kernelDiagFile = "$env:USERPROFILE\Desktop\3_Kernel_Diagnostics.txt"
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
$diskInfoFile = "$env:USERPROFILE\Desktop\4_Disk_Information.txt"
"===== INFORMATIONS MATÉRIELLES DES DISQUES =====" | Out-File $diskInfoFile -Encoding UTF8

$physDisks = Get-PhysicalDisk
$reliability = $physDisks | Get-StorageReliabilityCounter

# Récupération des volumes pour les lettres de lecteur
$volumes = Get-Volume | Where-Object { $_.DriveLetter } | Select-Object DiskNumber, DriveLetter
# Alternative : Get-Partition -AssignDriveLetter peut aussi être utilisé

$diskInfo = foreach ($disk in $physDisks) {
    $rel = $reliability | Where-Object { $_.DeviceId -eq $disk.DeviceId }
    
    # ObjectId complet (UniqueId) tel que renvoyé par la cmdlet
    $objectId = if ($rel) { $rel.UniqueId } else { "N/A" }
    
    # Extraction du GUID Storport (dernier GUID dans UniqueId)
    $storportGuid = "N/A"
    if ($rel -and $rel.UniqueId) {
        $allGuids = [regex]::Matches($rel.UniqueId, '\{[0-9A-Fa-f-]+\}')
        if ($allGuids.Count -gt 0) {
            $storportGuid = $allGuids[-1].Value.Trim('{}')
        }
    }
    
    # Récupération des lettres de lecteur (peuvent être multiples)
    $driveLetters = ($volumes | Where-Object { $_.DiskNumber -eq $disk.DeviceId -and $_.DriveLetter }).DriveLetter -join ","
    if ([string]::IsNullOrEmpty($driveLetters)) { 
        # Fallback : chercher via les partitions
        $parts = Get-Partition -DiskNumber $disk.DeviceId -ErrorAction SilentlyContinue | Where-Object DriveLetter
        if ($parts) { $driveLetters = $parts.DriveLetter -join "," }
        else { $driveLetters = "Aucune" }
    } 
    if ([string]::IsNullOrEmpty($driveLetters)) { $driveLetters = "Aucune" }
    
    [PSCustomObject]@{
        DiskNumber   = $disk.DeviceId
        DriveLetter  = $driveLetters
        Name         = $disk.FriendlyName
        BusType      = $disk.BusType
        SerialNumber = $disk.SerialNumber
        SizeGB       = [math]::Round($disk.Size / 1GB, 2)
        HealthStatus = $disk.HealthStatus
        ReadErrorsUncorrected = if ($rel) { $rel.ReadErrorsUncorrected } else { 0 }
        WriteErrorsUncorrected = if ($rel) { $rel.WriteErrorsUncorrected } else { 0 }
        ReadLatencyMax_ms = if ($rel) { $rel.ReadLatencyMax } else { 0 }
        WriteLatencyMax_ms = if ($rel) { $rel.WriteLatencyMax } else { 0 }
        WearPercent   = if ($rel) { $rel.Wear } else { 0 }
        Temperature_C = if ($rel) { $rel.Temperature } else { 0 }
        ObjectId      = $objectId          # UniqueId complet
        StorportGuid  = $storportGuid
    }
}

# Affichage au format liste (pour bien voir toutes les colonnes)
$diskInfo | Format-List | Out-File $diskInfoFile -Append -Encoding UTF8

# =============================================
# 4.1. MAP DISK -> VOLUMES (via GUID)
# =============================================
$diskInfoFile = "$env:USERPROFILE\Desktop\4_Disk_Information.txt"

# Récupérer les infos physiques et les compteurs
$physDisks = Get-PhysicalDisk
$reliability = $physDisks | Get-StorageReliabilityCounter

# Récupérer les informations de partitionnement et de volume
$partitions = Get-Partition | Where-Object { $_.DiskNumber -ne $null }
$volumes = Get-Volume | Where-Object { $_.DriveLetter -or $_.FileSystem -eq 'NTFS' }

Write-Host "`n===== MAPPING DISQUE PHYSIQUE (GUID) -> VOLUMES =====" -ForegroundColor Cyan
"`n===== MAPPING DISQUE PHYSIQUE (GUID) -> VOLUMES =====" | Out-File $diskInfoFile -Append -Encoding UTF8

foreach ($disk in $physDisks) {
    $rel = $reliability | Where-Object { $_.DeviceId -eq $disk.DeviceId }
    
    # Extraction du GUID Storport
    $storportGuid = "N/A"
    if ($rel -and $rel.UniqueId) {
        $allGuids = [regex]::Matches($rel.UniqueId, '\{[0-9A-Fa-f-]+\}')
        if ($allGuids.Count -gt 0) {
            $storportGuid = $allGuids[-1].Value.Trim('{}')
        }
    }
    
    # Trouver le numéro de disque logique (PhysicalDriveX) via la correspondance dans le registre
    # Une méthode plus fiable que Get-Partition seule est de chercher dans le registre
    $driveNumber = $disk.DeviceId
    $physicalDrivePath = "\\.\PhysicalDrive$driveNumber"
    
    # Lister tous les volumes sur ce disque
    $diskPartitions = $partitions | Where-Object { $_.DiskNumber -eq $driveNumber }
    $volumesOnDisk = @()
    foreach ($part in $diskPartitions) {
        $vol = $volumes | Where-Object { $_.Partition -and $_.Partition.DiskNumber -eq $driveNumber -and $_.Partition.PartitionNumber -eq $part.PartitionNumber }
        if (-not $vol) {
            # Fallback: chercher via le GUID de partition
            $vol = $volumes | Where-Object { $_.UniqueId -like "*$($part.Guid)*" }
        }
        if ($vol) {
            $volumesOnDisk += [PSCustomObject]@{
                DriveLetter  = if ($vol.DriveLetter) { "$($vol.DriveLetter):" } else { "Aucune" }
                VolumeGuid   = ($vol.UniqueId -replace '.*\\','')
                SizeGB       = [math]::Round($vol.Size / 1GB, 2)
                FileSystem   = $vol.FileSystem
            }
        }
    }
    
    if ($volumesOnDisk.Count -eq 0) {
        $volumesOnDisk = [PSCustomObject]@{ DriveLetter = "Aucun volume trouvé (peut-être non monté)"; VolumeGuid = "N/A"; SizeGB = "N/A"; FileSystem = "N/A" }
    }
    
    # Affichage des résultats
    Write-Host "`n[ DISQUE $driveNumber ] : $($disk.FriendlyName)" -ForegroundColor Yellow
    Write-Host "  GUID Storport  : $storportGuid"
    Write-Host "  Chemin disque  : $physicalDrivePath"
    Write-Host "  Volumes sur ce disque :"
    $volumesOnDisk | Format-Table DriveLetter, VolumeGuid, SizeGB, FileSystem -AutoSize
    
    # Écriture dans le fichier
    "`n[ DISQUE $driveNumber ] : $($disk.FriendlyName)" | Out-File $diskInfoFile -Append -Encoding UTF8
    "  GUID Storport  : $storportGuid" | Out-File $diskInfoFile -Append -Encoding UTF8
    "  Chemin disque  : $physicalDrivePath" | Out-File $diskInfoFile -Append -Encoding UTF8
    "  Volumes sur ce disque :" | Out-File $diskInfoFile -Append -Encoding UTF8
    $volumesOnDisk | Format-Table DriveLetter, VolumeGuid, SizeGB, FileSystem -AutoSize | Out-File $diskInfoFile -Append -Encoding UTF8
}

# Mettre en évidence le disque défaillant (DeviceId 0)
$badDiskGuid = "7fd9e307-31d4-d83f-2811-0f9408d7dcd3"
Write-Host "`n===== IDENTIFICATION DU DISQUE DÉFAILLANT =====" -ForegroundColor Red
Write-Host " GUID Storport du disque défaillant : $badDiskGuid" -ForegroundColor Red
Write-Host " Ce disque physique correspond au DeviceId 0 : WDC WD5000BEVT-22A0RT0" -ForegroundColor Red
Write-Host " Il est accessible via le chemin : \\.\PhysicalDrive0" -ForegroundColor Red
Write-Host " ⚠️  Remplacez-le immédiatement pour éviter une perte de données." -ForegroundColor Red


# Disques problématiques (avec erreurs non corrigées)
$problematic = $diskInfo | Where-Object { $_.ReadErrorsUncorrected -gt 0 -or $_.WriteErrorsUncorrected -gt 0 }
if ($problematic) {
    "`n===== DISQUES AVEC ERREURS NON CORRIGÉES =====" | Out-File $diskInfoFile -Append -Encoding UTF8
    $problematic | Format-List DiskNumber, DriveLetter, Name, ReadErrorsUncorrected, WriteErrorsUncorrected, ReadLatencyMax_ms, ObjectId, StorportGuid | Out-File $diskInfoFile -Append -Encoding UTF8
} else {
    "`n===== AUCUN DISQUE AVEC ERREURS DÉTECTÉES =====" | Out-File $diskInfoFile -Append -Encoding UTF8
}

Write-Host ""
if ($problematic) {
    Write-Host "⚠️  Disques avec erreurs non corrigées :" -ForegroundColor Red
    $problematic | Format-List DiskNumber, DriveLetter, Name, ReadErrorsUncorrected, StorportGuid, ObjectId
} else {
    Write-Host "✅ Aucun disque avec erreurs non corrigées détecté." -ForegroundColor Green
}

# =============================================
# 5. OUVERTURE DES FICHIERS GÉNÉRÉS
# =============================================
notepad $appCrashFile
notepad $systemCrashFile
notepad $kernelDiagFile
notepad $diskInfoFile

Write-Host "Diagnostics générés sur le bureau :`n" -ForegroundColor Green
Write-Host "1. $appCrashFile`n" -
Write-Host "2. $systemCrashFile`n" -
Write-Host "3. $kernelDiagFile`n" -
Write-Host "4. $diskInfoFile`n" -
Write-Host "Les fichiers Notepad sont ouverts." -ForegroundColor Green

<# 

.TROUVER BLOBK BUCKET ++ 10000ms + 

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
' 3_Kernel_Diagnostics.txt

#>
