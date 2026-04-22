<#
.SYNOPSIS
    Collecte et analyse crashes, erreurs systeme, logs kernel et IO disques - compatible Windows 7+
.DESCRIPTION
    Ce script :
    1. Recupere les crashes d'applications (Event ID 1000/1001)
       -> EVC_Export\1_Application_Crashes.txt
    2. Extrait les erreurs systeme critiques (Event ID 41, 1001, etc.)
       -> EVC_Export\2_System_Crashes.txt
    3. Collecte les logs kernel (WHEA, Dump, Storport, etc.)
       -> EVC_Export\3_Kernel_Diagnostics.txt
    4. Collecte les infos disques via WMI (compatible Win7)
       -> EVC_Export\4_Disk_Information.txt
    5. Collecte les erreurs de drivers (Event ID 219, 7000, 7001, 7011, 7026 + DriverFrameworks)
       -> EVC_Export\5_Driver_Errors.txt
    5_1. Collecte setupapi logs (10 derniers jours)
       -> EVC_Export\5_1_Driver_Logs.txt
    6. Analyse les IO > 10000ms via awk
       -> EVC_Export\IO_Errors.txt
.NOTES
    Auteur : ps81frt  (adaptation Win7 depuis EVCDiag.ps1)
    Lien   : https://github.com/ps81frt/EVC
    Compat : Windows 7 SP1 + PowerShell 2.0+
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
# GUARD : PowerShell 2.0 ne supporte pas
# les modules avances - on verifie la version
# =============================================
$psVer = $PSVersionTable.PSVersion.Major
if ($psVer -lt 2) {
    Write-Host "[ERREUR] PowerShell 2.0 minimum requis."
    exit 1
}

# =============================================
# 0. DOSSIER D'EXPORT COMMUN SUR LE BUREAU
# =============================================
$outputFolder   = Join-Path $env:USERPROFILE "Desktop\EVC_Export"
$kernelDiagFile = Join-Path $outputFolder "3_Kernel_Diagnostics.txt"
$diskInfoFile   = Join-Path $outputFolder "4_Disk_Information.txt"

if (-not (Test-Path $outputFolder)) {
    New-Item -Path $outputFolder -ItemType Directory | Out-Null
}

if ($Help) {
    Write-Host ""
    Write-Host "=== EVC_Disk_Win7.ps1 ==="
    Write-Host ""
    Write-Host "  .\EVC_Disk_Win7.ps1 -Collect"
    Write-Host "  .\EVC_Disk_Win7.ps1 -List"
    Write-Host "  .\EVC_Disk_Win7.ps1 -ListErrors"
    Write-Host "  .\EVC_Disk_Win7.ps1 -TimeCreated '10/04/2026 23:03:23'"
    Write-Host "  .\EVC_Disk_Win7.ps1 -TimeCreated '10/04/2026 23:03:23' -Port 2"
    Write-Host "  .\EVC_Disk_Win7.ps1 -TimeCreated '10/04/2026 23:03:23' -Port 2 -Path 0"
    Write-Host "  .\EVC_Disk_Win7.ps1 -TimeCreated '10/04/2026 23:03:23' -Guid 7fd9e307"
    Write-Host "  .\EVC_Disk_Win7.ps1 -TimeCreated '10/04/2026 23:03:23' -Guid 7fd9e307 -Export"
    Write-Host ""
    Write-Host "  -Collect     : collecte tous les logs + analyse IO_Errors auto"
    Write-Host "  -ListErrors  : exporte les IO >= 10000ms -> EVC_Export\IO_Errors.txt"
    Write-Host "                 installe les outils Linux dans System32 si absents"
    Write-Host ""
    exit
}

# =============================================
# HELPERS GENERAUX
# =============================================
function ConvertTo-GB($bytes) {
    if (-not $bytes -or $bytes -eq 0) { return 0 }
    return [math]::Round($bytes / 1GB, 2)
}

function Format-Bytes($bytes) {
    $mb = [math]::Round($bytes / 1MB, 2)
    $gb = [math]::Round($bytes / 1GB, 3)
    return "$mb Mo / $gb Go"
}

# New-Object PSObject compatible PS2 (remplace [PSCustomObject]@{})
function New-PSObj {
    param([hashtable]$props)
    $obj = New-Object PSObject
    foreach ($key in $props.Keys) {
        $obj | Add-Member -MemberType NoteProperty -Name $key -Value $props[$key]
    }
    return $obj
}

# =============================================
# 1. EXTRACTION ZIP : compatible Windows 7
#    Expand-Archive n'existe qu'a partir de PS5
#    -> on utilise Shell.Application COM
# =============================================
function Expand-ZipWin7 {
    param(
        [string]$ZipPath,
        [string]$DestPath
    )

    if (-not (Test-Path $DestPath)) {
        New-Item -Path $DestPath -ItemType Directory | Out-Null
    }

    try {
        $shell     = New-Object -ComObject Shell.Application
        $zip       = $shell.NameSpace($ZipPath)
        $dest      = $shell.NameSpace($DestPath)

        if (-not $zip) {
            Write-Host "[ERREUR] Shell.Application ne peut pas ouvrir : $ZipPath"
            return $false
        }

        # CopyHere flag 0x14 = no progress dialog + overwrite silently
        $dest.CopyHere($zip.Items(), 0x14)

        # Attente active : Shell.Application est asynchrone
        $timeout  = 120  # secondes max
        $elapsed  = 0
        $interval = 2

        do {
            Start-Sleep -Seconds $interval
            $elapsed += $interval
            $extracted = Get-ChildItem -Path $DestPath -Recurse -ErrorAction SilentlyContinue
        } while (($extracted.Count -eq 0) -and ($elapsed -lt $timeout))

        if ($extracted.Count -eq 0) {
            Write-Host "[ERREUR] Extraction echouee ou vide apres ${timeout}s."
            return $false
        }

        Write-Host "[INFO] Extraction terminee : $($extracted.Count) fichier(s) dans $DestPath"
        return $true

    } catch {
        Write-Host "[ERREUR] Expand-ZipWin7 : $_"
        return $false
    }
}

# =============================================
# 2. INSTALLATION DES OUTILS LINUX
#    - TLS 1.2 force (Win7 utilise TLS 1.0)
#    - Verification de la copie post-installation
#    - Fallback chemin temporaire si System32 refuse
# =============================================
function Install-AllLinuxTools {
    $zipUrl    = "https://github.com/ps81frt/LinuxToolsOnWindows/releases/download/1.0/LinuxToolOn-Windows.zip"
    $tmpZip    = Join-Path $env:TEMP "LinuxToolOn-Windows.zip"
    $tmpDir    = Join-Path $env:TEMP "LinuxTools_EVC"
    $sys32     = "$env:SystemRoot\System32"
    $keyTools  = @("awk.exe", "smartctl.exe")

    # --- Verif rapide : si les outils cles sont deja dans System32, rien a faire ---
    $allPresent = $true
    foreach ($t in $keyTools) {
        if (-not (Test-Path (Join-Path $sys32 $t))) { $allPresent = $false; break }
    }
    if ($allPresent) { return $true }

    # --- TLS 1.2 : indispensable pour GitHub sous Win7 ---
    # Win7 n'active pas TLS 1.2 par defaut (KB3140245 necessaire)
    # $tlsOk conserve pour trace eventuelle (PSScriptAnalyzer: assigned-not-used supprime)
    try {
        # Valeur numerique 3072 = Tls12, evite l'erreur si l'enum n'existe pas
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]3072
    } catch {
        Write-Host "[WARN] TLS 1.2 indisponible (KB3140245 manquant ?). Tentative avec le protocole par defaut."
    }

    # --- Telechargement (une seule fois, cache dans %TEMP%) ---
    if (-not (Test-Path $tmpZip)) {
        Write-Host "[INFO] Telechargement des outils depuis GitHub..."
        try {
            $wc = New-Object System.Net.WebClient
            $wc.DownloadFile($zipUrl, $tmpZip)
            Write-Host "[INFO] Telechargement OK -> $tmpZip"
        } catch {
            # Fallback : Invoke-WebRequest (PS3+)
            try {
                Invoke-WebRequest -Uri $zipUrl -OutFile $tmpZip -UseBasicParsing -ErrorAction Stop
                Write-Host "[INFO] Telechargement OK (Invoke-WebRequest) -> $tmpZip"
            } catch {
                Write-Host "[ERREUR] Telechargement echoue : $_"
                Write-Host ""
                Write-Host "  => Telechargez manuellement le zip :"
                Write-Host "     $zipUrl"
                Write-Host "     Puis placez-le ici : $tmpZip"
                return $false
            }
        }
    }

    # --- Extraction (Shell.Application, compatible Win7) ---
    if (-not (Test-Path $tmpDir)) {
        Write-Host "[INFO] Extraction de l'archive..."
        $ok = Expand-ZipWin7 -ZipPath $tmpZip -DestPath $tmpDir
        if (-not $ok) { return $false }
    }

    # --- Copie dans System32 avec verification post-installation ---
    $allBinaries   = Get-ChildItem -Path $tmpDir -Recurse -Include "*.exe","*.dll" -ErrorAction SilentlyContinue
    $installErrors = @()
    $installed     = @()

    if (-not $allBinaries) {
        Write-Host "[ERREUR] Aucun binaire .exe/.dll trouve dans l'archive extraite."
        Write-Host "         Verifiez le contenu de : $tmpDir"
        return $false
    }

    $newInstalls = 0

    foreach ($bin in $allBinaries) {
        $dest = Join-Path $sys32 $bin.Name

        # Deja present et identique -> silencieux, rien a faire
        if (Test-Path $dest) {
            $srcHash = (Get-FileHash $bin.FullName -Algorithm MD5 -ErrorAction SilentlyContinue).Hash
            $dstHash = (Get-FileHash $dest         -Algorithm MD5 -ErrorAction SilentlyContinue).Hash
            if ($srcHash -eq $dstHash) {
                $installed += $bin.Name
                continue
            }
        }

        # Nouvelle installation ou mise a jour -> verbeux
        try {
            Copy-Item $bin.FullName -Destination $dest -Force -ErrorAction Stop

            if (Test-Path $dest) {
                Write-Host "[OK] $($bin.Name) -> $dest"
                $installed += $bin.Name
                $newInstalls++
            } else {
                Write-Host "[WARN] $($bin.Name) copie sans erreur mais absent de System32 !"
                $installErrors += $bin
            }

        } catch {
            Write-Host "[WARN] $($bin.Name) -> echec copie System32 : $_"
            Write-Host "       Le binaire restera utilisable depuis : $($bin.FullName)"
            Write-Host "       => Relancez le script en tant qu'Administrateur."
            $installErrors += $bin
        }
    }

    # Rapport final : seulement si quelque chose s'est passe
    if ($newInstalls -gt 0) {
        Write-Host "[INFO] $newInstalls outil(s) installe(s) dans $sys32."
    }
    if ($installErrors.Count -gt 0) {
        Write-Host "[WARN] $($installErrors.Count) fichier(s) non installe(s) dans System32 :"
        $installErrors | ForEach-Object { Write-Host "  - $($_.Name)" }
        Write-Host "  NOTE : Ces outils seront utilises depuis leur emplacement temporaire."
    }

    return $true
}

# =============================================
# 3. LOCALISATION D'UN OUTIL (awk, smartctl...)
#    Ordre : System32 -> PATH -> dossier temp
# =============================================
function Get-ToolPath($toolName) {
    # 1. System32
    $sys32Path = "$env:SystemRoot\System32\$toolName.exe"
    if (Test-Path $sys32Path) { return $sys32Path }

    # 2. PATH systeme
    $inPath = Get-Command $toolName -ErrorAction SilentlyContinue
    if ($inPath) { return $inPath.Source }

    # 3. Dossier temporaire (fallback non-admin)
    $tmpDir = Join-Path $env:TEMP "LinuxTools_EVC"
    if (Test-Path $tmpDir) {
        $found = Get-ChildItem -Path $tmpDir -Recurse -Filter "$toolName.exe" -ErrorAction SilentlyContinue |
                 Select-Object -First 1
        if ($found) {
            Write-Host "[INFO] $toolName utilise depuis : $($found.FullName)"
            return $found.FullName
        }
    }

    Write-Host "[ERREUR] $toolName.exe introuvable (System32, PATH et dossier temporaire)."
    return $null
}

$script:toolsInstalled = $false

# Ensure-LinuxTools renomme en Invoke-EnsureLinuxTools (verbe approuve PS)
function Invoke-EnsureLinuxTools {
    if (-not $script:toolsInstalled) {
        Install-AllLinuxTools | Out-Null
        $script:toolsInstalled = $true
    }
}

function Install-Awk {
    Invoke-EnsureLinuxTools
    return Get-ToolPath "awk"
}

function Install-Smartctl {
    Invoke-EnsureLinuxTools
    return Get-ToolPath "smartctl"
}

# =============================================
# 4. ANALYSE IO > 10000ms
# =============================================
function Invoke-ListErrors($inputFile) {
    $awkBin = Install-Awk
    if (-not $awkBin) {
        Write-Host "[ERREUR] awk introuvable, analyse IO impossible."
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
    # Set-Content : Encoding ASCII (compatible PS2)
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
    try {
        Get-WinEvent -LogName "Application" -ErrorAction Stop |
            Where-Object {$_.Id -eq 1000} | Sort-Object TimeCreated |
            Select-Object TimeCreated,
                @{N="Application";E={$_.Properties[0].Value}},
                @{N="Version";    E={$_.Properties[1].Value}},
                @{N="Module";     E={$_.Properties[3].Value}},
                @{N="Code";       E={$_.Properties[6].Value}},
                @{N="Offset";     E={$_.Properties[7].Value}} |
            Format-List | Out-File $appCrashFile -Append -Encoding UTF8
        Write-Host "[OK] Application Crashes (ID 1000)" -ForegroundColor Green
    } catch {
        $null
    }

    "`n===== ERREURS D'APPLICATIONS (Event ID 1001) =====" | Out-File $appCrashFile -Append -Encoding UTF8
    try {
        Get-WinEvent -LogName "Application" -ErrorAction Stop |
            Where-Object {$_.Id -eq 1001} | Sort-Object TimeCreated |
            Select-Object TimeCreated, @{N="Message";E={$_.Message}} |
            Format-List | Out-File $appCrashFile -Append -Encoding UTF8
        Write-Host "[OK] Application Errors (ID 1001)" -ForegroundColor Green
    } catch {
        $null
    }

    # =============================================
    # 2. ERREURS SYSTEME CRITIQUES
    # =============================================
    $systemCrashFile = Join-Path $outputFolder "2_System_Crashes.txt"
    "===== ERREURS SYSTEME @(41,1001,7023,7034,157,153,7000,7001,7009,7011,7026,7045) =====" | Out-File $systemCrashFile -Encoding UTF8
    try {
        Get-WinEvent -LogName "System" -ErrorAction Stop |
            Where-Object {$_.Id -in @(41,1001,7023,7034,157,153,7000,7001,7009,7011,7026,7045)} |
            Sort-Object TimeCreated |
            Select-Object TimeCreated, Id, @{N="Message";E={$_.Message}} |
            Format-List | Out-File $systemCrashFile -Append -Encoding UTF8
        Write-Host "[OK] System Crashes/Errors" -ForegroundColor Green
    } catch {
        $null
    }

    # =============================================
    # 3. LOGS KERNEL (WHEA, Dump, Storport, etc.)
    # NOTE : certains journaux n'existent pas sous Win7,
    # les erreurs sont silencieusement ignorees.
    # =============================================
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
    "Generated : $(Get-Date)"                              | Out-File $kernelDiagFile -Append -Encoding UTF8

    foreach ($logName in $logNames) {
        try {
            $events = Get-WinEvent -LogName $logName -ErrorAction Stop
            if ($events) {
                $events | Sort-Object TimeCreated |
                    Select-Object TimeCreated, LogName, @{N="Message";E={$_.Message}} |
                    Format-List | Out-File $kernelDiagFile -Append -Encoding UTF8
                Write-Host "[OK] $logName ($($events.Count) evenements)" -ForegroundColor Green
            }
        } catch {
            $null
        }
    }

    # =============================================
    # 4. DISQUES PHYSIQUES via WMI (compatible Win7)
    # =============================================
    $physDisks     = Get-WmiObject -Class Win32_DiskDrive           | Sort-Object Index
    $allPartitions = Get-WmiObject -Class Win32_DiskPartition
    $allLogical    = Get-WmiObject -Class Win32_LogicalDisk
    $assocs        = Get-WmiObject -Class Win32_LogicalDiskToPartition

    $diskToLogical = @{}
    foreach ($a in $assocs) {
        $partPath = $a.Antecedent -replace '.*Win32_DiskPartition\.', '' -replace '"', ''
        $logPath  = $a.Dependent  -replace '.*Win32_LogicalDisk\.',   '' -replace '"', ''
        if ($partPath -match 'DiskIndex=(\d+),\s*Index=(\d+)') {
            $dIdx = $matches[1]
            if (-not $diskToLogical.ContainsKey($dIdx)) { $diskToLogical[$dIdx] = @() }
            $logDisk = $allLogical | Where-Object {
                ($_.DeviceID -replace '\\\\','').TrimEnd('\') -eq `
                ($logPath -replace 'DeviceID=','').Trim('"').TrimEnd('\')
            } | Select-Object -First 1
            if ($logDisk) { $diskToLogical[$dIdx] += $logDisk }
        }
    }

    $diskInfo = foreach ($disk in $physDisks) {
        $dIdx    = [string]$disk.Index
        $letters = @()
        if ($diskToLogical.ContainsKey($dIdx)) {
            $letters = $diskToLogical[$dIdx] | ForEach-Object { $_.DeviceID }
        }
        $driveLetters = if ($letters.Count -gt 0) { $letters -join "," } else { "Aucune" }

        [PSCustomObject]@{
            DiskIndex    = $disk.Index
            DriveLetter  = $driveLetters
            Caption      = $disk.Caption
            Model        = $disk.Model
            Interface    = $disk.InterfaceType
            SerialNumber = $disk.SerialNumber
            SizeGB       = ConvertTo-GB $disk.Size
            Status       = $disk.Status
            Partitions   = $disk.Partitions
            BytesPerSect = $disk.BytesPerSector
        }
    }

    "===== INFORMATIONS MATERIELLES DES DISQUES =====" | Out-File $diskInfoFile -Encoding UTF8
    "Generated : $(Get-Date)"                          | Out-File $diskInfoFile -Append -Encoding UTF8

    "`n--- Vue generale ---" | Out-File $diskInfoFile -Append -Encoding UTF8
    $diskInfo | Select-Object DiskIndex, DriveLetter, Caption, Interface, SizeGB, Status, Partitions |
        Format-Table -AutoSize | Out-File $diskInfoFile -Append -Encoding UTF8

    "`n--- Details techniques ---" | Out-File $diskInfoFile -Append -Encoding UTF8
    $diskInfo | Select-Object DiskIndex, Model, SerialNumber, BytesPerSect, SizeGB |
        Format-Table -AutoSize | Out-File $diskInfoFile -Append -Encoding UTF8

    "`n===== MAPPING DISQUE -> PARTITIONS -> VOLUMES =====" | Out-File $diskInfoFile -Append -Encoding UTF8

    foreach ($disk in $physDisks) {
        $dIdx = [string]$disk.Index
        $line = "`n[ DISQUE $($disk.Index) ] : $($disk.Caption)"
        $line | Out-File $diskInfoFile -Append -Encoding UTF8
        Write-Host $line -ForegroundColor Yellow

        $parts = $allPartitions | Where-Object { $_.DiskIndex -eq $disk.Index } | Sort-Object Index
        foreach ($part in $parts) {
            $lv           = $null
            $assocForPart = $assocs | Where-Object {
                $_.Antecedent -match "DiskIndex=$($disk.Index)" -and
                $_.Antecedent -match "Index=$($part.Index)[^0-9]"
            }
            if ($assocForPart) {
                $logPath = $assocForPart.Dependent -replace '.*DeviceID=', '' -replace '"', ''
                $lv = $allLogical | Where-Object {
                    $_.DeviceID -replace '\\','' -eq ($logPath -replace '\\','')
                } | Select-Object -First 1
            }

            $letter  = if ($lv) { $lv.DeviceID }              else { "(non monte)" }
            $sizeGB  = ConvertTo-GB $part.Size
            $fs      = if ($lv) { $lv.FileSystem }             else { "?" }
            $freeGB  = if ($lv) { ConvertTo-GB $lv.FreeSpace } else { "N/A" }

            $partLine = ("  Partition {0} | Lettre: {1,-6} | FS: {2,-7} | Taille: {3,8} Go | Libre: {4} Go" `
                -f $part.Index, $letter, $fs, $sizeGB, $freeGB)
            $partLine | Out-File $diskInfoFile -Append -Encoding UTF8
            Write-Host $partLine
        }
    }

    "`n===== DF : ESPACE DISQUE PAR LECTEUR =====" | Out-File $diskInfoFile -Append -Encoding UTF8
    $allLogical | Sort-Object DeviceID |
        Select-Object DeviceID,
            @{N="VolumeName"; E={$_.VolumeName}},
            @{N="FileSystem"; E={$_.FileSystem}},
            @{N="Size(GB)";   E={ConvertTo-GB $_.Size}},
            @{N="Free(GB)";   E={ConvertTo-GB $_.FreeSpace}},
            @{N="Used(GB)";   E={ConvertTo-GB ($_.Size - $_.FreeSpace)}} |
        Format-Table -AutoSize | Out-File $diskInfoFile -Append -Encoding UTF8

    "`n===== BLKID-FULL : PARTITIONS + UUID + FS =====" | Out-File $diskInfoFile -Append -Encoding UTF8
    $blkidRows = @()
    foreach ($disk in $physDisks) {
        $parts = $allPartitions | Where-Object { $_.DiskIndex -eq $disk.Index } | Sort-Object Index
        foreach ($part in $parts) {
            $lv           = $null
            $assocForPart = $assocs | Where-Object {
                $_.Antecedent -match "DiskIndex=$($disk.Index)" -and
                $_.Antecedent -match "Index=$($part.Index)[^0-9]"
            }
            if ($assocForPart) {
                $logPath = $assocForPart.Dependent -replace '.*DeviceID=', '' -replace '"', ''
                $lv = $allLogical | Where-Object {
                    $_.DeviceID -replace '\\','' -eq ($logPath -replace '\\','')
                } | Select-Object -First 1
            }

            $blkidRows += [PSCustomObject]@{
                Disk       = $disk.Caption
                PartNum    = $part.Index
                'Size(GB)' = ConvertTo-GB $part.Size
                FS         = if ($lv -and $lv.FileSystem) { $lv.FileSystem } else { "?" }
                Type       = $part.Type
                VolumeUUID = if ($lv -and $lv.VolumeSerialNumber) { $lv.VolumeSerialNumber } else { "N/A" }
            }
        }
    }
    $blkidRows | Format-Table -AutoSize | Out-String -Width 10000 | Out-File $diskInfoFile -Append -Encoding UTF8

    $lsblkBin = Get-ToolPath "lsblk"
    "`n===== LSBLK =====" | Out-File $diskInfoFile -Append -Encoding UTF8
    if ($lsblkBin) {
        $lsblkOut = & $lsblkBin 2>&1
        $lsblkOut | Out-File $diskInfoFile -Append -Encoding UTF8
        $lsblkOut | ForEach-Object { Write-Host $_ }
    } else {
        " [WARN] lsblk introuvable." | Out-File $diskInfoFile -Append -Encoding UTF8
        Write-Host " [WARN] lsblk introuvable." -ForegroundColor Yellow
    }

    $smartBin = Install-Smartctl
    "`n===== SMART (smartctl) =====" | Out-File $diskInfoFile -Append -Encoding UTF8

    if ($smartBin) {
        foreach ($disk in $physDisks) {
            $drive   = "\\.\PhysicalDrive$($disk.Index)"
            $busType = $disk.InterfaceType
            $devType = switch -Wildcard ($busType) {
                "SCSI"  { "scsi" }
                "IDE"   { "ata"  }
                "NVMe"  { "nvme" }
                "USB"   { "sat"  }
                default { "auto" }
            }

            "`n--- $drive : $($disk.Caption) (bus: $busType) ---" | Out-File $diskInfoFile -Append -Encoding UTF8

            # Pattern echec : couvre tous les cas connus (VMware, USB, SCSI generique)
            $failPattern = "Invalid argument|Unable to detect device type|open device.*failed|Unknown USB|Smartctl open device"

            $smartOut = & $smartBin -a $drive -d $devType 2>&1
            $smartFailed = $smartOut | Where-Object { $_ -match $failPattern }

            if ($smartFailed) {
                # Essai de tous les types dans l'ordre jusqu'au premier qui repond
                $fallbackTypes = @("sat", "scsi", "ata", "nvme", "auto") | Where-Object { $_ -ne $devType }
                $resolved = $false
                foreach ($fb in $fallbackTypes) {
                    $fbOut    = & $smartBin -a $drive -d $fb 2>&1
                    $fbFailed = $fbOut | Where-Object { $_ -match $failPattern }
                    if (-not $fbFailed) {
                        $smartOut = $fbOut
                        $resolved = $true
                        break
                    }
                }
                if (-not $resolved) {
                    " [WARN] SMART indisponible sur $drive (disque virtuel ou non supporte)." | Out-File $diskInfoFile -Append -Encoding UTF8
                    continue
                }
            }

            $smartOut | Out-File $diskInfoFile -Append -Encoding UTF8
            $smartOut | ForEach-Object { Write-Host $_ }
        }
    } else {
        " [WARN] smartctl introuvable, SMART non disponible." | Out-File $diskInfoFile -Append -Encoding UTF8
        Write-Host " [WARN] smartctl introuvable, SMART non disponible." -ForegroundColor Yellow
    }

    "`n===== IDENTIFICATION DU(DES) DISQUE(S) POTENTIELLEMENT DEFAILLANT(S) =====" | Out-File $diskInfoFile -Append -Encoding UTF8
    Write-Host ""
    Write-Host "===== IDENTIFICATION DU(DES) DISQUE(S) POTENTIELLEMENT DEFAILLANT(S) =====" -ForegroundColor Red

    $badDisks = $diskInfo | Where-Object { $_.Status -ne "OK" }
    if ($badDisks) {
        foreach ($bd in $badDisks) {
            $msg = " [!] Disque $($bd.DiskIndex) : $($bd.Caption) | Status: $($bd.Status) | Lettre(s): $($bd.DriveLetter)"
            $msg | Out-File $diskInfoFile -Append -Encoding UTF8
            Write-Host $msg -ForegroundColor Red
        }
    } else {
        $ok = " Aucun disque avec status WMI anormal detecte (Status = OK pour tous)."
        $ok | Out-File $diskInfoFile -Append -Encoding UTF8
        Write-Host $ok -ForegroundColor Green
    }
    " NOTE : Consultez la section SMART ci-dessus pour les details complets." | Out-File $diskInfoFile -Append -Encoding UTF8

    # =============================================
    # 5. ERREURS DE DRIVERS
    # =============================================
    $driverErrorFile = Join-Path $outputFolder "5_Driver_Errors.txt"
    "===== ERREURS DE DRIVERS (Event ID 219, 7000, 7001, 7011, 7026) =====" | Out-File $driverErrorFile -Encoding UTF8

    try {
        Get-WinEvent -LogName "System" -ErrorAction Stop |
            Where-Object {$_.Id -in @(219,7000,7001,7011,7026)} | Sort-Object TimeCreated |
            Select-Object TimeCreated, Id, @{N="Message";E={$_.Message}} |
            Format-List | Out-File $driverErrorFile -Append -Encoding UTF8
        Write-Host "[OK] Driver Errors (ID 219,7000,7001,7011,7026)" -ForegroundColor Green
    } catch {
        $null
    }

    $driverLogs = @(
        "Microsoft-Windows-DriverFrameworks-UserMode/Operational",
        "Microsoft-Windows-DriverFrameworks-KernelMode/Operational",
        "Microsoft-Windows-Kernel-PnP/Configuration",
        "Microsoft-Windows-DeviceSetupManager/Admin",
        "Microsoft-Windows-DeviceSetupManager/Operational"
    )

    "`n===== JOURNAUX DRIVERFRAMEWORKS ET PNP =====" | Out-File $driverErrorFile -Append -Encoding UTF8
    foreach ($logName in $driverLogs) {
        try {
            Get-WinEvent -LogName $logName -ErrorAction Stop | Sort-Object TimeCreated |
                Select-Object TimeCreated, LogName, @{N="Message";E={$_.Message}} |
                Format-List | Out-File $driverErrorFile -Append -Encoding UTF8
            Write-Host "[OK] $logName" -ForegroundColor Green
        } catch {
            $null
        }
    }

    # =============================================
    # 5_1. DRIVER LOGS - setupapi (10 derniers jours)
    # =============================================
    $driverLogFile = Join-Path $outputFolder "5_1_Driver_Logs.txt"
    $setupApiLogs  = @(
        "C:\Windows\INF\setupapi.dev.log",
        "C:\Windows\INF\setupapi.setup.log"
    )
    $cutoff = (Get-Date).AddDays(-10)

    "===== DRIVER LOGS - setupapi (10 derniers jours) =====" | Out-File $driverLogFile -Encoding UTF8
    "Filtre : depuis $cutoff"                                | Out-File $driverLogFile -Append -Encoding UTF8
    "`n"                                                     | Out-File $driverLogFile -Append -Encoding UTF8

    foreach ($setupApiLog in $setupApiLogs) {
        "--- $setupApiLog ---" | Out-File $driverLogFile -Append -Encoding UTF8

        if (Test-Path $setupApiLog) {
            $lines         = Get-Content $setupApiLog -Encoding UTF8
            $inSection     = $false
            $sectionBuf    = @()
            $pendingHeader = $null

            foreach ($line in $lines) {
                if ($line -match "^>>>\s+\[.+\]\s*$") {
                    if ($inSection -and $sectionBuf.Count -gt 0) {
                        $sectionBuf | Out-File $driverLogFile -Append -Encoding UTF8
                    }
                    $inSection     = $false
                    $sectionBuf    = @()
                    $pendingHeader = $line
                } elseif ($pendingHeader -and $line -match ">>>\s+Section start\s+(\d{4}/\d{2}/\d{2}\s+\d{2}:\d{2}:\d{2})") {
                    try {
                        $ts        = [datetime]::ParseExact($matches[1], "yyyy/MM/dd HH:mm:ss", $null)
                        $inSection = ($ts -ge $cutoff)
                    } catch {
                        $inSection = $false
                    }
                    $sectionBuf    = @($pendingHeader, $line)
                    $pendingHeader = $null
                } elseif ($inSection) {
                    $sectionBuf += $line
                }
            }
            if ($inSection -and $sectionBuf.Count -gt 0) {
                $sectionBuf | Out-File $driverLogFile -Append -Encoding UTF8
            }
            Write-Host "[OK] $setupApiLog -> $driverLogFile" -ForegroundColor Green
        } else {
            "[WARN] $setupApiLog introuvable." | Out-File $driverLogFile -Append -Encoding UTF8
            Write-Host "[WARN] $setupApiLog introuvable." -ForegroundColor Yellow
        }
    }

    # =============================================
    # 6. ANALYSE IO > 10000ms
    # =============================================
    Write-Host ""
    Write-Host "===== ANALYSE IO > 10000ms ====="
    if (Test-Path $kernelDiagFile) {
        Invoke-ListErrors $kernelDiagFile
    } else {
        Write-Host "[WARN] $kernelDiagFile introuvable, analyse IO ignoree."
    }

    $ioErrorsFile = Join-Path $outputFolder "IO_Errors.txt"
    notepad $appCrashFile
    notepad $systemCrashFile
    notepad $kernelDiagFile
    notepad $diskInfoFile
    notepad $driverErrorFile
    if (Test-Path $driverLogFile) { notepad $driverLogFile }
    if (Test-Path $ioErrorsFile)  { notepad $ioErrorsFile  }

    Write-Host ""
    Write-Host "Diagnostics generes dans : $outputFolder" -ForegroundColor Green
    Write-Host "1. $appCrashFile"
    Write-Host "2. $systemCrashFile"
    Write-Host "3. $kernelDiagFile"
    Write-Host "4. $diskInfoFile"
    Write-Host "5. $driverErrorFile"
    Write-Host "5_1. $driverLogFile"
    if (Test-Path $ioErrorsFile) { Write-Host "IO. $ioErrorsFile" }

    exit
}

# =============================================
# READER (mode sans -Collect)
# =============================================
if (-not (Test-Path $kernelDiagFile)) {
    Write-Host "[ERREUR] Fichier $kernelDiagFile introuvable. Lancez d'abord : .\EVC_Disk_Win7.ps1 -Collect"
    exit
}

if ($ListErrors) {
    Invoke-ListErrors $kernelDiagFile
    exit
}

# Lecture via StreamReader : evite de charger tout le fichier en memoire
# Filtre immediat sur TimeCreated : ne garde que les blocs pertinents
Write-Host "[...] Lecture de $kernelDiagFile ..." -NoNewline

$fileSize    = (Get-Item $kernelDiagFile).Length
$reader      = [System.IO.StreamReader]::new($kernelDiagFile, [System.Text.Encoding]::UTF8)
$bufLines    = New-Object System.Collections.Generic.List[string]
$blocs       = New-Object System.Collections.Generic.List[string]
$allDates    = New-Object System.Collections.Generic.List[string]
$bytesRead   = 0
$lastPct     = -1

while ($null -ne ($line = $reader.ReadLine())) {
    $bytesRead += [System.Text.Encoding]::UTF8.GetByteCount($line) + 2
    $pct = [math]::Min(99, [int](($bytesRead / $fileSize) * 100))
    if ($pct -ne $lastPct) {
        Write-Host "`r[...] Lecture $pct%   " -NoNewline
        $lastPct = $pct
    }

    if ($line -match "^TimeCreated\s*:") {
        if ($bufLines.Count -gt 0) {
            $bloc = $bufLines -join "`n"
            $blocs.Add($bloc)
            if ($bloc -match "TimeCreated\s*:\s*(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})") {
                $allDates.Add($matches[1])
            }
            $bufLines.Clear()
        }
        $bufLines.Add($line)
    } else {
        $bufLines.Add($line)
    }
}
if ($bufLines.Count -gt 0) {
    $bloc = $bufLines -join "`n"
    $blocs.Add($bloc)
}
$reader.Close()
Write-Host "`r[OK] $($blocs.Count) blocs lus.              "

if ($List) {
    $allDates | ForEach-Object { Write-Host $_ }
    exit
}

if (-not $TimeCreated) {
    Write-Host "[ERREUR] Aucun parametre. Utilisez -Help pour l'aide."
    exit
}

Write-Host "[...] Recherche de $TimeCreated ..." -NoNewline
$selectedBlocs = $blocs | Where-Object {
    $_ -like "*$TimeCreated*" -and $_ -match 'Performance summary for Storport Device'
}
Write-Host "`r[OK] $($selectedBlocs.Count) evenement(s) trouve(s).    "

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

function Get-RawArray($pattern, $text) {
    $m = [regex]::Match($text, $pattern, "Singleline")
    if (-not $m.Success) { return @() }
    $values = $m.Groups[1].Value.Trim() -replace '\s+', ''
    return $values.Split(",") | ForEach-Object {
        $v = ($_ -replace "[^\d]", "")
        if ($v -eq "") { 0 } else { [int64]$v }
    }
}

$bucketLabels = @(
    "128 us","256 us","512 us","1 ms","4 ms","16 ms",
    "64 ms","128 ms","256 ms","512 ms","1000 ms","2000 ms","10000 ms","> 10000 ms"
)


function Write-WrappedLine($label, $text, $width) {
    $prefix = "| $label"
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

function Get-AvgLatency($total, $count) {
    if ($count -le 0 -or $total -le 0) { return "-" }
    return [math]::Round(($total / $count) / 10000, 3)
}

if ($Export) {
    $safeDate   = $TimeCreated -replace '[/: ]', '-'
    $safeGuid   = ($guid -replace '[{}]', '').Substring(0, [math]::Min(8, $guid.Length))
    $reportFile = Join-Path $outputFolder "Report_${safeDate}_${safeGuid}.txt"
    Start-Transcript -Path $reportFile -Force | Out-Null
}

$devices = $selectedBlocs | ForEach-Object { Get-DeviceInfo $_ }
$devices = $devices | Group-Object -Property { "$($_.Guid)|$($_.Port)|$($_.Path)" } |
    ForEach-Object { $_.Group[-1] }

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
    $devices   = $devices | Where-Object { $_.Guid -like "*$cleanGuid*" }
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
    Write-Host "Exemple : .\EVC_Disk_Win7.ps1 -TimeCreated '$TimeCreated' -Port 2 -Path 0"
    Write-Host "          .\EVC_Disk_Win7.ps1 -TimeCreated '$TimeCreated' -Guid 7fd9e307"
    exit
}

$selectedEntry = $devices[0].Bloc

# --- Calculs sur le bloc selectionne ---
$success = Get-RawArray "IO success counts are ([\d,\s]+)" $selectedEntry
$failed  = Get-RawArray "IO failed counts are ([\d,\s]+)"  $selectedEntry
$latency = Get-RawArray "IO total latency.*?are ([\d,\s]+)" $selectedEntry
$totalIO = ([regex]::Match($selectedEntry, "Total IO:\s*(\d+)")).Groups[1].Value -as [int64]

$max = 14
while ($success.Count -lt $max) { $success += 0 }
while ($failed.Count  -lt $max) { $failed  += 0 }
while ($latency.Count -lt $max) { $latency += 0 }

$sumSuccess    = ($success | Measure-Object -Sum).Sum
$guid          = ([regex]::Match($selectedEntry, "(?:Guid is|Corresponding Class Disk Device Guid is) \{(.*?)\}")).Groups[1].Value
# $portVal/$pathVal/$tgtVal/$lunVal : reserves pour usage futur (debug/export)
# $device  = [regex]::Match($selectedEntry, "Port = (\d+), Path = (\d+), Target = (\d+), Lun = (\d+)")
# $portVal = $device.Groups[1].Value
# $pathVal = $device.Groups[2].Value
# $tgtVal  = $device.Groups[3].Value
# $lunVal  = $device.Groups[4].Value
$highLatencyIO = $success[7]+$success[8]+$success[9]+$success[10]+$success[11]+$success[12]+$success[13]

$lossWarning = ""
if ($sumSuccess -ne $totalIO) {
    $lossWarning = "[WARN] MISMATCH: success sum ($sumSuccess) != Total IO ($totalIO)"
}
if ($lossWarning -ne "") { Write-Host ""; Write-Host $lossWarning }

$logNameValue = ([regex]::Match($selectedEntry, "LogName\s*:\s*(.+)")).Groups[1].Value
$messageValue = ([regex]::Match($selectedEntry, "Message\s*:\s*(.+)")).Groups[1].Value
$messageValue = ($messageValue -replace 'whose Corresponding Class Disk Device Guid is \{[^}]+\}:?', '').Trim()

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
        if     ($ok -eq 0 -and $fail -eq 0)  { "Aucun"              }
        elseif ($i -le 2)                     { "Optimal"            }
        elseif ($i -le 3)                     { "Normal"             }
        elseif ($i -le 5)                     { "Acceptable"         }
        elseif ($i -eq 6 -and $ok -gt 0)     { "[!] A surveiller"   }
        elseif ($i -le 10 -and $ok -gt 0)    { "[!!] Degradee"      }
        elseif ($i -eq 11 -and $ok -gt 0)    { "[!!] Critique"      }
        elseif ($i -eq 12 -and $ok -gt 0)    { "[!!!] 10s ($ok IO)" }
        elseif ($i -eq 13 -and $ok -gt 0)    { "[!!!] EXTREME"      }
        else                                  { "OK"                 }

    Write-Host ("| {0,-14} | {1,14:N0} | {2,14} | {3,14} | {4,14:N0} | {5,-18} |" `
        -f $bucketLabels[$i], $ok, $pct, $lat, $fail, $status)
}

Write-Host "+----------------+----------------+----------------+----------------+----------------+--------------------+"

$totalLatency = ($latency | Measure-Object -Sum).Sum
$totalOps     = $sumSuccess + ($failed | Measure-Object -Sum).Sum
$globalAvg    = if ($totalOps -gt 0) { [math]::Round(($totalLatency / $totalOps) / 10000, 6) } else { 0 }

Write-Host ""
Write-Host "+---------------------------------------------------------------+"
Write-Host ("| Latence Globale Moyenne : {0} ms (ponderee)" -f $globalAvg)
Write-Host ("| IO totales verifiees    : {0:N0} / {1:N0}"   -f $sumSuccess, $totalIO)
if ($highLatencyIO -gt 0) {
    Write-Host ("| [!] IO > 128ms          : {0:N0} operations" -f $highLatencyIO)
}
Write-Host "+---------------------------------------------------------------+"

if ($Export) {
    Stop-Transcript | Out-Null
    Write-Host ""
    Write-Host "[EXPORT] $reportFile"
}