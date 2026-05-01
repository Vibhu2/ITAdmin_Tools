# VB.NextCloud
**VB.NextCloud** is a PSGallery module that brings WebDAV-powered Nextcloud integration directly into PowerShell — turning what used to be a manual, multi-step file collection process into a single command that just works.

The idea is deceptively simple, but the impact is significant. Imagine running sysadmin scripts across dozens of remote machines — generating health reports, event logs, audit outputs — and instead of chasing those files down manually, every single one lands in a central Nextcloud repository automatically, within minutes. No RDP. No shared drives. No waiting. Just data, where you need it, when you need it.

Four functions power the entire workflow:

- **Single-file upload** (PUT) — precision control to push any individual file to an exact Nextcloud path, on demand
- **`Start-VBNextcloudUpload`** — the workhorse; a batch orchestrator that accepts any number of files, auto-creates the destination folder structure if it doesn't exist, and fires them all up in one clean call
- **Directory listing** (PROPFIND) — enumerates the full contents of any remote Nextcloud path, giving you visibility into exactly what's been collected
- **Folder listing** (PROPFIND) — returns a folder-only tree, useful for navigating and validating your repository structure programmatically

What makes this genuinely powerful is the workflow it unlocks. Patch your endpoint scripts to call `Start-VBNextcloudUpload` at the end of every run, and your Nextcloud becomes a living, self-populating repository of everything happening across your environment. By the time you sit down at your desktop, all the reports are already there — ready to parse, aggregate, and analyse.

No other tool in the PowerShell ecosystem approaches Nextcloud this way. This isn't just a file upload utility — it's the missing link between remote automation and centralised data collection, built entirely in PowerShell and available straight from the PSGallery.

---

## Install

```powershell
Install-Module -Name VB.NextCloud -Scope CurrentUser
```

---

## Functions

| Function | Description |
| --- | --- |
| `Set-VBNextcloudFile` | Uploads a single file to Nextcloud via WebDAV PUT |
| `Start-VBNextcloudUpload` | Batch upload orchestrator — uploads a folder of files with auto folder creation |
| `Get-VBNextcloudFiles` | Lists files in a Nextcloud directory via PROPFIND |
| `Get-VBNextcloudFolders` | Lists folders only in a Nextcloud path via PROPFIND |

---

## Basic Usage

```powershell
$cred = Get-Credential

# Upload a single file
Set-VBNextcloudFile `
    -LocalPath 'C:\Reports\WS001.csv' `
    -RemotePath 'Realtime-IT/Reports/WS001.csv' `
    -BaseUrl 'https://cloud.example.com' `
    -Credential $cred

# Batch upload all CSVs from a folder
Start-VBNextcloudUpload `
    -LocalFolder 'C:\Reports' `
    -RemoteFolder 'Realtime-IT/Reports' `
    -BaseUrl 'https://cloud.example.com' `
    -Credential $cred

# List files in a remote folder
Get-VBNextcloudFiles `
    -RemotePath 'Realtime-IT/Reports' `
    -BaseUrl 'https://cloud.example.com' `
    -Credential $cred

# List sub-folders only
Get-VBNextcloudFolders `
    -RemotePath 'Realtime-IT' `
    -BaseUrl 'https://cloud.example.com' `
    -Credential $cred
```

---

## Compatibility

- PowerShell 5.1+ (uses `HttpWebRequest` for full PS 5.1 compatibility)
- PowerShell 7.x supported

---

## Links

- [PSGallery](https://www.powershellgallery.com/packages/VB.NextCloud)
- [GitHub](https://github.com/Vibhu2/ITAdmin_Tools/tree/main/VB.NextCloud)
- [Blog](https://pwsh.in)
- [License](https://github.com/Vibhu2/ITAdmin_Tools/blob/main/LICENSE)