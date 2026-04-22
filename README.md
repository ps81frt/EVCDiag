# EVCDiag

## Utilisation rapide

- Ouvrir PowerShell en tant qu'administrateur
- Débloquer le script si nécessaire :
  ```powershell
  Unblock-File .\EVCDiag.ps1
  ```
- Autoriser l'exécution temporaire si la politique l'empêche :
  ```powershell
  Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
  ```
- Exécuter le script :
  ```powershell
  .\EVCDiag.ps1 -Help
  ```

`EVCDiag.ps1` collecte des diagnostics Windows orientés stockage et kernel.
Le script exporte un jeu de fichiers dans `EVC_Export` contenant :

- crashs applicatifs
- erreurs système critiques
- logs kernel et stockage
- inventaire des disques physiques et GUID Storport
- analyse des I/O lentes (seuil `10000 ms`)

---

## Execution Rapide WINDOWS 10/11

Ouvrir **PowerShell en tant qu'administrateur** puis coller — télécharge, extrait et exécute automatiquement :

```powershell
&{
    $zip = "$env:TEMP\evc.zip"
    irm "https://github.com/ps81frt/EVCDiag/archive/refs/heads/main.zip" -OutFile $zip
    Expand-Archive $zip "$env:TEMP\evc" -Force
    $script = Get-ChildItem "$env:TEMP\evc" -Recurse -Filter "EVCDiag.ps1" | Select-Object -First 1
    Set-Location $script.Directory.FullName
    Unblock-File $script.FullName
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
    & $script.FullName -Collect
}
```

---

## Execution Rapide WINDOWS 7

> `Expand-Archive` et `irm` sont absents sous PS2.  
> Ce bloc utilise `WebClient` (TLS 1.2 forcé) + `Shell.Application` pour l'extraction.

Ouvrir **PowerShell en tant qu'administrateur** puis coller — télécharge, extrait et exécute automatiquement :

```powershell
&{
    # TLS 1.2 obligatoire pour GitHub sous Win7
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]3072

    $zip = "$env:TEMP\evc.zip"
    $dst = "$env:TEMP\evc"

    (New-Object System.Net.WebClient).DownloadFile(
        "https://github.com/ps81frt/EVCDiag/archive/refs/heads/main.zip",
        $zip
    )

    if (-not (Test-Path $dst)) { New-Item -Path $dst -ItemType Directory | Out-Null }
    $shell = New-Object -ComObject Shell.Application
    $shell.NameSpace($dst).CopyHere($shell.NameSpace($zip).Items(), 0x14)

    $t = 0
    do { Start-Sleep -Seconds 2; $t += 2 } while ((Get-ChildItem $dst -Recurse).Count -eq 0 -and $t -lt 60)

    $script = Get-ChildItem $dst -Recurse -Filter "EVCDiag.ps1" | Select-Object -First 1
    Set-Location $script.Directory.FullName
    Unblock-File $script.FullName
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
    & $script.FullName -Collect
}
```

---

### Fonctionnement

- collecte des événements `Application` et `System` ciblés
- collecte des journaux kernel et matériels listés dans le script
- collecte de l'inventaire disque via `Get-PhysicalDisk`, `Get-StorageReliabilityCounter`, `Get-Volume`, `Get-Partition`
- analyse des buckets de latence Storport
- génération de `IO_Errors.txt` pour les I/O lentes

### 1. Événements Application et System

- `Application` : `Event ID 1000`, `1001`
  - détecte les crashs applicatifs et les rapports d'erreur Windows.
- `System` : `Event ID 41`, `1001`, `7023`, `7034`, `157`, `153`
  - capture les redémarrages inattendus, services arrêtés brutalement et erreurs de stockage/block device.

### 2. Logs kernel et diagnostics matériels

- `Microsoft-Windows-Kernel-WHEA/Operational` — erreurs matérielles détectées par WHEA (CPU, mémoire, disque).
- `Microsoft-Windows-Kernel-WHEA/Errors` — résumé des erreurs WHEA critiques.
- `Microsoft-Windows-Kernel-Dump/Operational` — événements de génération de dump après crash.
- `Microsoft-Windows-Diagnostics-Performance/Operational` — diagnostics de performance, blocages et démarrages lents.
- `Microsoft-Windows-Resource-Exhaustion-Detector/Operational` — détection de manque de mémoire ou de ressources critiques.
- `Microsoft-Windows-Kernel-PnP/Driver Watchdog` — problèmes de drivers bloquants ou temps d'attente excessif.
- `Microsoft-Windows-Fault-Tolerant-Heap/Operational` — corruption de heap dans des composants sensibles du système.
- `Microsoft-Windows-WerKernel/Operational` — erreurs kernel remontées par Windows Error Reporting.
- `Microsoft-Windows-CodeIntegrity/Operational` — blocages de drivers non signés ou violations d'intégrité.
- `Microsoft-Windows-Security-Mitigations/KernelMode` — état des mitigations de sécurité kernel.
- `Microsoft-Windows-Kernel-Boot/Operational` — diagnostics du démarrage et du chargement des drivers.
- `Microsoft-Windows-Storage-Storport/Operational` — informations Storport : latences I/O, erreurs SCSI, GUID des périphériques.
- `Microsoft-Windows-Ntfs/Operational` — erreurs NTFS, incohérences de métadonnées et problèmes de fichiers.

### 3. Inventaire de stockage

- `Get-PhysicalDisk` — liste les disques physiques et leur état.
- `Get-StorageReliabilityCounter` — récupère les compteurs SMART / fiabilité et l'UniqueId Storport.
- `Get-Volume` + `Get-Partition` — associe volumes, lettres et GUID de partition aux disques physiques.

### 4. Analyse de latence I/O

- extraction des buckets de latence Storport
- calcul de la répartition des opérations dans des paliers jusqu'à `10000 ms`
- création de `IO_Errors.txt` pour les I/O lentes
- identification des opérations à surveiller ou critiques

### Quand l'utiliser

- lenteurs I/O intermittentes ou constantes
- BSOD / redémarrages inattendus
- erreurs de disque, corruption NTFS ou volumes non montés
- performances dégradées sur un poste utilisateur ou serveur
- enquête post-incident après une panne de stockage

### Sorties générées

Le script crée un dossier structuré :

- `C:\Users\<Utilisateur>\Desktop\EVC_Export`

Fichiers générés :

- `1_Application_Crashes.txt` : crashs d'applications et erreurs associées
- `2_System_Crashes.txt` : erreurs système critiques
- `3_Kernel_Diagnostics.txt` : logs kernel et diagnostics matériels
- `4_Disk_Information.txt` : inventaire des disques, GUID Storport, volumes
- `5_Driver_Errors.txt` : erreurs de drivers
- `5_1_Driver_Logs.txt` : logs setupapi
- `IO_Errors.txt` : I/O supérieures à `10000 ms`

### Usage

```powershell
.\EVCDiag.ps1 -Collect
.\EVCDiag.ps1 -List
.\EVCDiag.ps1 -ListErrors
.\EVCDiag.ps1 -TimeCreated '10/04/2026 23:03:23'
.\EVCDiag.ps1 -TimeCreated '10/04/2026 23:03:23' -Port 2
.\EVCDiag.ps1 -TimeCreated '10/04/2026 23:03:23' -Port 2 -Path 0
.\EVCDiag.ps1 -TimeCreated '10/04/2026 23:03:23' -Guid 7fd9e307
.\EVCDiag.ps1 -TimeCreated '10/04/2026 23:03:23' -Guid 7fd9e307 -Export
```

### Options

- `-Collect` : collecte tous les logs et lance l'analyse I/O automatique
- `-List` : liste les dates des blocs capturés dans `EVC_Export\3_Kernel_Diagnostics.txt`
- `-ListErrors` : exporte les I/O supérieures à `10000 ms` dans `EVC_Export\IO_Errors.txt`
- `-TimeCreated` : sélectionne un bloc de diagnostic par date/heure
- `-Port` : filtre par port SCSI/Storport
- `-Path` : filtre par chemin du périphérique
- `-Guid` : filtre par GUID du périphérique
- `-Export` : génère un rapport textuel dans `EVC_Export`

### Pourquoi cette approche

Plutôt que de se limiter à un simple `Get-WinEvent`, EVCDiag :

- collecte un ensemble de logs spécialisés pour le kernel et le stockage
- met en relation un GUID Storport avec un disque physique
- identifie des disques suspects à partir de compteurs SMART et de latence
- produit un jeu de fichiers facilement partageable pour une cellule SOC ou un support

### Références utiles

- https://learn.microsoft.com/windows-hardware/drivers/whea/
- https://learn.microsoft.com/windows-hardware/drivers/storage/storport
- https://learn.microsoft.com/windows/security/threat-protection/windows-defender-application-control/working-with-ci-event-ids

### Execution et upload direct de tous les fichiers depuis le terminal :

```powershell
&{
    $zip = "$env:TEMP\evc.zip"
    irm "https://github.com/ps81frt/EVCDiag/archive/refs/heads/main.zip" -OutFile $zip
    Expand-Archive $zip "$env:TEMP\evc" -Force

    $script = Get-ChildItem "$env:TEMP\evc" -Recurse -Filter "EVCDiag.ps1" | Select-Object -First 1
    Set-Location $script.Directory.FullName
    Unblock-File $script.FullName
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
    & $script.FullName -Collect

    $files = @(
        "$env:USERPROFILE\Desktop\EVC_Export\1_Application_Crashes.txt",
        "$env:USERPROFILE\Desktop\EVC_Export\2_System_Crashes.txt",
        "$env:USERPROFILE\Desktop\EVC_Export\3_Kernel_Diagnostics.txt",
        "$env:USERPROFILE\Desktop\EVC_Export\4_Disk_Information.txt",
        "$env:USERPROFILE\Desktop\EVC_Export\5_Driver_Errors.txt",
        "$env:USERPROFILE\Desktop\EVC_Export\5_1_Driver_Logs.txt",
        "$env:USERPROFILE\Desktop\EVC_Export\IO_Errors.txt"
    )

    foreach ($f in $files) {
        if (Test-Path $f) {
            curl -F "file=@$f" https://store1.gofile.io/uploadFile |
            ConvertFrom-Json |
            Select-Object -ExpandProperty data |
            Select-Object -ExpandProperty downloadPage
        }
    }
}
```

### Execution et upload direct du fichier des pilotes :

```powershell
&{
    $zip = "$env:TEMP\evc.zip"
    irm "https://github.com/ps81frt/EVCDiag/archive/refs/heads/main.zip" -OutFile $zip
    Expand-Archive $zip "$env:TEMP\evc" -Force

    $script = Get-ChildItem "$env:TEMP\evc" -Recurse -Filter "EVCDiag.ps1" | Select-Object -First 1
    Set-Location $script.Directory.FullName
    Unblock-File $script.FullName
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
    & $script.FullName -Collect

    curl -F "file=@$env:USERPROFILE\Desktop\EVC_Export\5_Driver_Errors.txt" https://store1.gofile.io/uploadFile |
    ConvertFrom-Json |
    Select-Object -ExpandProperty data |
    Select-Object -ExpandProperty downloadPage
}
```

### Upload manuel d'un fichier

```powershell
curl -F "file=@$env:USERPROFILE\Desktop\EVC_Export\5_Driver_Errors.txt" https://store1.gofile.io/uploadFile |
ConvertFrom-Json | Select-Object -ExpandProperty data | Select-Object -ExpandProperty downloadPage
```

### Upload manuel de plusieurs fichiers

```powershell
@(
    "$env:USERPROFILE\Desktop\EVC_Export\1_Application_Crashes.txt",
    "$env:USERPROFILE\Desktop\EVC_Export\2_System_Crashes.txt",
    "$env:USERPROFILE\Desktop\EVC_Export\3_Kernel_Diagnostics.txt",
    "$env:USERPROFILE\Desktop\EVC_Export\4_Disk_Information.txt",
    "$env:USERPROFILE\Desktop\EVC_Export\5_Driver_Errors.txt",
    "$env:USERPROFILE\Desktop\EVC_Export\5_1_Driver_Logs.txt",
    "$env:USERPROFILE\Desktop\EVC_Export\IO_Errors.txt"
) | Where-Object { Test-Path $_ } | ForEach-Object {
    $url = curl -F "file=@$_" https://store1.gofile.io/uploadFile |
           ConvertFrom-Json | Select-Object -ExpandProperty data | Select-Object -ExpandProperty downloadPage
    "$_ -> $url"
}
```

### Auteur

- `ps81frt`
- GitHub : https://github.com/ps81frt/EVCDiag

### Licence

Ce projet est distribué sous la licence MIT.

### Prérequis

- PowerShell sur Windows
- droits suffisants pour accéder aux journaux d'événements et au matériel de stockage
- le script installe `awk` dans `C:\Windows\System32` si nécessaire

### Notes

- le script corrèle les GUID Storport entre `4_Disk_Information.txt` et `3_Kernel_Diagnostics.txt`
- si plusieurs périphériques sont trouvés pour une même date, préciser `-Port`, `-Path` ou `-Guid`
- le diagnostic est conçu pour être envoyé tel quel à un support ou réutilisé dans une procédure GPO
