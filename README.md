# EVCDiag

## Utilisation rapide

- Ouvrir PowerShell en tant qu’administrateur
- Débloquer le script si nécessaire :
  ```powershell
  Unblock-File .\EVCDiag.ps1
  ```
- Autoriser l’exécution temporaire si la politique l’empêche :
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

### Fonctionnement

- collecte des événements `Application` et `System` ciblés
- collecte des journaux kernel et matériels listés dans le script
- collecte de l'inventaire disque via `Get-PhysicalDisk`, `Get-StorageReliabilityCounter`, `Get-Volume`, `Get-Partition`
- analyse des buckets de latence Storport
- génération de `IO_Errors.txt` pour les I/O lentes

### 
### 1. Événements Application et System

- `Application` : `Event ID 1000`, `1001`
  - détecte les crashs applicatifs et les rapports d'erreur Windows.
- `System` : `Event ID 41`, `1001`, `7023`, `7034`, `157`, `153`
  - capture les redémarrages inattendus, services arrêtés brutalement et erreurs de stockage/block device.

### 2. Logs kernel et diagnostics matériels

- `Microsoft-Windows-Kernel-WHEA/Operational`
  - erreurs matérielles détectées par WHEA (CPU, mémoire, disque).
- `Microsoft-Windows-Kernel-WHEA/Errors`
  - résumé des erreurs WHEA critiques.
- `Microsoft-Windows-Kernel-Dump/Operational`
  - événements de génération de dump après crash.
- `Microsoft-Windows-Diagnostics-Performance/Operational`
  - diagnostics de performance, blocages et démarrages lents.
- `Microsoft-Windows-Resource-Exhaustion-Detector/Operational`
  - détection de manque de mémoire ou de ressources critiques.
- `Microsoft-Windows-Kernel-PnP/Driver Watchdog`
  - problèmes de drivers bloquants ou temps d'attente excessif.
- `Microsoft-Windows-Fault-Tolerant-Heap/Operational`
  - corruption de heap dans des composants sensibles du système.
- `Microsoft-Windows-WerKernel/Operational`
  - erreurs kernel remontées par Windows Error Reporting.
- `Microsoft-Windows-CodeIntegrity/Operational`
  - blocages de drivers non signés ou violations d'intégrité.
- `Microsoft-Windows-Security-Mitigations/KernelMode`
  - état des mitigations de sécurité kernel.
- `Microsoft-Windows-Kernel-Boot/Operational`
  - diagnostics du démarrage et du chargement des drivers.
- `Microsoft-Windows-Storage-Storport/Operational`
  - informations Storport : latences I/O, erreurs SCSI, GUID des périphériques.
- `Microsoft-Windows-Ntfs/Operational`
  - erreurs NTFS, incohérences de métadonnées et problèmes de fichiers.

### 3. Inventaire de stockage

- `Get-PhysicalDisk`
  - liste les disques physiques et leur état.
- `Get-StorageReliabilityCounter`
  - récupère les compteurs SMART / fiabilité et l'UniqueId Storport.
- `Get-Volume` + `Get-Partition`
  - associe volumes, lettres et GUID de partition aux disques physiques.

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

- `-Collect` : collecte tous les logs et produit le package complet
- `-List` : liste les timestamps des blocs de diagnostics dans `3_Kernel_Diagnostics.txt`
- `-ListErrors` : extrait les I/O lentes dans `IO_Errors.txt`
- `-TimeCreated` : sélectionne le bloc Storport correspondant à un instant précis
- `-Port` : filtre l'analyse par port Storport
- `-Path` : filtre par chemin de périphérique
- `-Guid` : filtre l'analyse par GUID de périphérique
- `-Export` : génère un rapport textuel résumé dans `EVC_Export`

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
- ce collecteur est prévu pour un diagnostic ponctuel, pas pour de la supervision continue

Fichiers générés :

- `1_Application_Crashes.txt`
- `2_System_Crashes.txt`
- `3_Kernel_Diagnostics.txt`
- `4_Disk_Information.txt`
- `IO_Errors.txt`

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
UID Storport entre `4_Disk_Information.txt` et `3_Kernel_Diagnostics.txt`
- si plusieurs périphériques sont trouvés pour une même date, préciser `-Port`, `-Path` ou `-Guid`
- le diagnostic est conçu pour être envoyé tel quel à un support ou réutilisé dans une procédure GPO

### Execution et upload direct depuit le terminal exemple:

``` powershell
$zip="$env:TEMP\evc.zip"
irm "https://github.com/ps81frt/EVCDiag/archive/refs/heads/main.zip" -OutFile $zip
Expand-Archive $zip "$env:TEMP\evc" -Force

$script = Get-ChildItem "$env:TEMP\evc" -Recurse -Filter "EVCDiag.ps1" | Select-Object -First 1

cd $script.Directory.FullName

Unblock-File $script.FullName
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
& $script.FullName -Collect

curl -F "file=@$env:USERPROFILE\Desktop\EVC_Export\5_Driver_Errors.txt" https://store1.gofile.io/uploadFile |
ConvertFrom-Json |
Select-Object -ExpandProperty data |
Select-Object -ExpandProperty downloadPage
```
