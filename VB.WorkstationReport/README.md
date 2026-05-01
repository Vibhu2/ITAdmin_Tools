# VB.WorkstationReport

PSGallery module for end-to-end workstation auditing in MSP environments. Collects user profiles, printer mappings, folder redirections, OneDrive backup status, Sync Center state, and network interfaces — all orchestrated via a single `Invoke-VBWorkstationReport` call with optional Nextcloud upload.

**Blog post:** [Printer Mapping & Migration with VB.WorkstationReport](https://pwsh.in/posts/printer-mapping-powershell-vb-workstationreport/)

---

## Install

```powershell
Install-Module -Name VB.WorkstationReport -Scope CurrentUser
```

> **Note:** `VB.NextCloud` is a soft runtime dependency required when using the `-UploadToNextcloud` parameter. Install it separately:
>
> ```powershell
> Install-Module -Name VB.NextCloud -Scope CurrentUser
> ```

---

## Functions

| Function | Description |
| --- | --- |
| `Invoke-VBWorkstationReport` | Orchestrates a full workstation audit and optionally uploads results to Nextcloud |
| `Get-VBUserProfiles` | Enumerates all local user profiles with SID, path, and last use time |
| `Get-VBUserPrinters` | Lists all mapped printers per user profile (UNC and IP) |
| `Get-VBDefaultPrinter` | Returns the default printer for each user profile |
| `Get-VBUserFolderRedirections` | Detects folder redirection targets (Desktop, Documents, etc.) |
| `Get-VBOneDriveBackupStatus` | Reports OneDrive Known Folder Move backup status per user |
| `Get-VBSyncCenterStatus` | Returns Sync Center offline files configuration and status |
| `Get-VBNetworkInterfaces` | Collects network adapter configuration (IP, MAC, gateway) |
| `Mount-VBUserHive` | Mounts a user registry hive from a profile path |
| `Dismount-VBUserHive` | Safely unmounts a mounted user registry hive |
| `Add-VBUserPrinter` | Adds a new UNC or IP printer to all or targeted user profiles |
| `Set-VBUserPrinterMigration` | Migrates printer mappings from old to new paths via CSV |
| `Invoke-VBasCurrentUser` | Runs a scriptblock in the context of the logged-on user from SYSTEM |
| `Invoke-VBDiskinformation` | collects workstation Disk information usage and health information|

---

## Basic Usage

```powershell
# Full workstation audit — saves CSVs locally
Invoke-VBWorkstationReport -ComputerName WS001 -OutputPath C:\Realtime\Reports

# Full audit with Nextcloud upload
Invoke-VBWorkstationReport `
    -ComputerName WS001 `
    -OutputPath C:\Realtime\Reports `
    -NextcloudCredential (Get-Credential) `
    -NextcloudBaseUrl 'https://cloud.example.com' `
    -NextcloudDestination 'Realtime-IT/Reports'

# Enumerate user profiles on a remote machine
Get-VBUserProfiles -ComputerName WS001

# List all mapped printers across all profiles
Get-VBUserPrinters -ComputerName WS001

# Migrate printers from old server to new server via CSV
Set-VBUserPrinterMigration -ComputerName WS001 -MappingCsv C:\Migrations\printers.csv
```

---

## Links

- [PSGallery](https://www.powershellgallery.com/packages/VB.WorkstationReport)
- [GitHub](https://github.com/Vibhu2/ITAdmin_Tools/tree/main/VB.WorkstationReport)
- [Blog post](https://pwsh.in/posts/printer-mapping-powershell-vb-workstationreport/)
- [Linkedin](https://www.linkedin.com/in/vibhu-bhatnagar-02622798/)
- [License](https://github.com/Vibhu2/ITAdmin_Tools/blob/main/LICENSE)