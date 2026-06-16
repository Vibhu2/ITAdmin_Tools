# Set-VBUserPrinterMigration — Usage Guide

**Module:** VB.WorkstationReport  
**Version covered:** 1.1.5  
**Author:** Vibhu Bhatnagar

---

## Overview

`Set-VBUserPrinterMigration` migrates printer mappings across user profiles on a Windows machine. It handles all four migration directions and writes changes directly to each user's registry hive — no user logon required.

| From | To | Scenario |
|------|----|----------|
| UNC path | UNC path | Print server replacement |
| UNC path | IP address | Move off print server to direct IP |
| IP address | UNC path | Move onto a print server |
| IP address | IP address | Printer port change |

> **Important:** The script must run **locally** on the target machine for any migration that involves IP destinations. Deploy via RMM (e.g. ConnectWise, NinjaRMM). Remote execution is supported for UNC→UNC migrations only.

---

## Requirements

- PowerShell 5.1+
- Administrative privileges on the target machine
- `PrintManagement` module (built-in on Windows 8 / Server 2012+)
- Printer drivers already installed for any IP destination (run `Get-PrinterDriver | Select-Object Name` to confirm spelling)

---

## CSV File Format

The CSV is the recommended input method for production deployments. Save as UTF-8.

### Columns

| Column | Required | Description |
|--------|----------|-------------|
| `OldPath` | Yes | Current printer path — UNC (`\\Server\Share`) or IP (`10.30.1.50`) |
| `NewPath` | Yes | Target printer path — UNC or IP |
| `DriverName` | Conditional | Required when `NewPath` is an IP address. Leave blank for UNC destinations. Must match the driver store exactly. |
| `PrinterName` | No | Friendly display name for the printer (IP destinations only). When omitted the name is derived from the last segment of the old UNC path. |

### Example CSV

```csv
OldPath,NewPath,DriverName,PrinterName
\\PrintServer01\HP_Floor2,10.30.1.50,HP Universal Printing PCL 6,HP Floor 2
\\PrintServer01\Canon_HR,10.30.1.51,Canon Generic Plus PCL6,Canon HR
10.30.1.60,\\PrintServer02\Ricoh_Reception,,
\\PrintServer01\Zebra_Labels,\\PrintServer02\Zebra_Labels,,
10.30.1.70,10.30.1.71,HP Universal Printing PCL 6,Reception Printer
```

**Row breakdown:**

| Row | Type | Notes |
|-----|------|-------|
| 1 | UNC → IP | DriverName required; PrinterName sets friendly display name |
| 2 | UNC → IP | DriverName required; PrinterName sets friendly display name |
| 3 | IP → UNC | DriverName blank — not needed for UNC destination |
| 4 | UNC → UNC | DriverName and PrinterName both blank |
| 5 | IP → IP | DriverName required; PrinterName keeps consistent name |

> **Tip:** Run `Get-PrinterDriver | Select-Object Name` on the target machine and copy-paste driver names directly into the CSV to avoid spelling mismatches.

---

## Use Cases

### 1. Dry Run (Always Do This First)

Before any real deployment, run with `-WhatIf` to see exactly what would change without touching anything.

```powershell
Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv' -WhatIf
```

---

### 2. Migrate All Profiles on the Machine

Omit `-Username` and `-SID` to process every non-system profile found on the machine.

```powershell
Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv'
```

With backup (recommended for production):

```powershell
Set-VBUserPrinterMigration `
    -MappingCsv   'C:\Temp\PrinterMappings.csv' `
    -BackupMappings `
    -BackupPath   'C:\Logs\PrinterBackup.csv'
```

> If backup fails for a user, that user is **skipped** — changes are never applied without the safety net.

---

### 3. Single User — By Username

Target one or more users by name. A warning is emitted for any username not found on the machine (safe to ignore when deploying fleet-wide via RMM).

```powershell
# Single user
Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv' -Username 'jdoe'

# Multiple users
Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv' -Username 'jdoe', 'asmith'

# Domain-qualified name also works
Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv' -Username 'CONTOSO\jdoe'
```

---

### 4. Single User — By SID

Useful when deploying via RMM — the script self-selects on machines where that SID exists and silently skips all others.

```powershell
Set-VBUserPrinterMigration `
    -MappingCsv 'C:\Temp\PrinterMappings.csv' `
    -SID        'S-1-5-21-123456789-1001'
```

Combine Username and SID (OR logic — either match is included):

```powershell
Set-VBUserPrinterMigration `
    -MappingCsv 'C:\Temp\PrinterMappings.csv' `
    -Username   'jdoe' `
    -SID        'S-1-5-21-123456789-1002'
```

---

### 5. UNC → UNC (Print Server Replacement)

No driver needed. Use either a CSV or hashtable.

**Via hashtable:**

```powershell
$mappings = @{
    '\\OldPrintServer\HP01'    = '\\NewPrintServer\HP01'
    '\\OldPrintServer\Canon02' = '\\NewPrintServer\Canon02'
    '\\OldPrintServer\Zebra'   = '\\NewPrintServer\Zebra'
}
Set-VBUserPrinterMigration -PrinterMappings $mappings
```

**Via CSV** (`DriverName` and `PrinterName` columns can be blank):

```csv
OldPath,NewPath,DriverName,PrinterName
\\OldPrintServer\HP01,\\NewPrintServer\HP01,,
\\OldPrintServer\Canon02,\\NewPrintServer\Canon02,,
```

---

### 6. UNC → IP (Move Off Print Server)

Driver is required. Use `-PrinterNames` or the `PrinterName` CSV column to keep a consistent display name.

**Via CSV (recommended — per-printer driver and name control):**

```csv
OldPath,NewPath,DriverName,PrinterName
\\PrintServer01\HP_Floor2,10.30.1.50,HP Universal Printing PCL 6,HP Floor 2
\\PrintServer01\Canon_HR,10.30.1.51,Canon Generic Plus PCL6,Canon HR
```

```powershell
Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv'
```

**Via hashtable (all printers share one driver):**

```powershell
$mappings = @{
    '\\PrintServer01\HP_Floor2' = '10.30.1.50'
    '\\PrintServer01\HP_Floor3' = '10.30.1.51'
}

# Without explicit names -- display name derived from UNC share segment (HP_Floor2, HP_Floor3)
Set-VBUserPrinterMigration -PrinterMappings $mappings -DriverName 'HP Universal Printing PCL 6'

# With explicit names -- display name stays consistent regardless of old UNC path
$names = @{
    '\\PrintServer01\HP_Floor2' = 'HP Floor 2'
    '\\PrintServer01\HP_Floor3' = 'HP Floor 3'
}
Set-VBUserPrinterMigration -PrinterMappings $mappings -DriverName 'HP Universal Printing PCL 6' -PrinterNames $names
```

---

### 7. IP → UNC (Move Onto Print Server)

No driver needed for the destination. `DriverName` and `PrinterName` columns can be blank.

```powershell
$mappings = @{
    '10.30.1.50' = '\\NewPrintServer\HP01'
    '10.30.1.51' = '\\NewPrintServer\Canon02'
}
Set-VBUserPrinterMigration -PrinterMappings $mappings
```

---

### 8. IP → IP (Port Change)

Driver required. Use `PrinterName` to keep the same display name before and after.

**Via CSV:**

```csv
OldPath,NewPath,DriverName,PrinterName
10.30.1.50,10.30.1.100,HP Universal Printing PCL 6,HP Floor 2
10.30.1.51,10.30.1.101,Canon Generic Plus PCL6,Canon HR
```

```powershell
Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv'
```

---

### 9. Consistent Printer Names (Key Feature)

Without `PrinterName`, a UNC→IP migration derives the display name from the last segment of the old UNC path (`\\Server\HP_Floor2` → `HP_Floor2`). This means the name can differ depending on how the printer was originally mapped.

Supply `PrinterName` to pin the display name regardless of migration direction:

```csv
OldPath,NewPath,DriverName,PrinterName
\\PrintServer\hpprinter410,10.30.1.50,HP Universal Printing PCL 6,hpprinter410
10.30.1.50,10.30.1.55,HP Universal Printing PCL 6,hpprinter410
```

The printer always appears as `hpprinter410` whether mapped via UNC or direct IP.

---

### 10. Capture Results and Export to CSV

Every action returns a result object. Capture them for reporting.

```powershell
$results = Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv'

# Export all results
$results | Export-Csv -Path 'C:\Logs\MigrationResults.csv' -NoTypeInformation -Encoding UTF8

# Review failures only
$results | Where-Object { $_.Status -eq 'Failed' } | Format-Table

# Count by action
$results | Group-Object Action | Select-Object Name, Count
```

---

### 11. RMM Deployment Pattern

Deploy the script to each machine locally. Log results back to a network share keyed by machine name.

```powershell
$results = Set-VBUserPrinterMigration `
    -MappingCsv     'C:\Temp\PrinterMappings.csv' `
    -BackupMappings `
    -BackupPath     "\\FileServer\Logs\PrinterBackup_$env:COMPUTERNAME.csv"

$results | Export-Csv `
    -Path             "\\FileServer\Logs\MigrationResults_$env:COMPUTERNAME.csv" `
    -NoTypeInformation `
    -Encoding         UTF8
```

The backup CSV is append-safe — multiple machines writing to the same file is supported.

---

## Parameters Reference

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `-MappingCsv` | String | One of these two | Path to CSV file. Takes priority over `-PrinterMappings` if both supplied. |
| `-PrinterMappings` | Hashtable | One of these two | `@{ OldPath = NewPath }` pairs inline. |
| `-DriverName` | String | When using `-PrinterMappings` with IP destinations | Driver name applied to all IP destinations in the hashtable. |
| `-PrinterNames` | Hashtable | No | `@{ OldPath = 'FriendlyName' }` — explicit display names for IP destinations when using `-PrinterMappings`. |
| `-Username` | String[] | No | Limit to specific users by name (short, DOMAIN\user, or profile folder). Warns if not found. |
| `-SID` | String[] | No | Limit to specific users by SID. Combines with `-Username` using OR logic. |
| `-BackupMappings` | Switch | No | Snapshot each user's current printers before migrating. Requires `-BackupPath`. |
| `-BackupPath` | String | With `-BackupMappings` | CSV file to write backup snapshots to. Append-safe. |
| `-ComputerName` | String[] | No | Target computer(s). Defaults to local machine. Remote only supported for UNC→UNC. |
| `-Credential` | PSCredential | No | Credentials for remote execution. |
| `-WhatIf` | Switch | No | Preview changes without applying them. |

---

## Output Object

One object is returned per user per mapping rule.

| Property | Values | Description |
|----------|--------|-------------|
| `ComputerName` | | Target machine |
| `Username` | | User profile name |
| `SID` | | User SID |
| `OldPath` | | Old printer path from mapping rule |
| `NewPath` | | New printer path from mapping rule |
| `Action` | `Migrated`, `Skipped`, `AlreadyMigrated`, `Failed` | What happened |
| `Details` | | Registry actions taken, or reason for skip/failure |
| `Status` | `Success`, `Failed` | Roll-up status |
| `Error` | | Error message (empty on success) |
| `Timestamp` | `dd-MM-yyyy HH:mm:ss` | Time of action |

---

## Common Errors

| Error | Cause | Fix |
|-------|-------|-----|
| `Driver 'X' is not installed` | Named driver not in driver store | Run `Get-PrinterDriver \| Select-Object Name` on target machine and correct spelling in CSV |
| `Machine-level printer port additions are not supported for remote targets` | Script ran against a remote machine with IP destinations | Deploy script locally on each machine via RMM |
| `DriverName is required for IP destinations` | CSV row has IP `NewPath` but blank `DriverName` | Add driver name to that CSV row |
| `-BackupPath is required when -BackupMappings is specified` | `-BackupMappings` used without `-BackupPath` | Add `-BackupPath 'C:\path\to\backup.csv'` |
| `Backup failed for 'user' — migration skipped` | Backup write failed (permissions, locked file) | Fix backup path permissions; migration is intentionally blocked without backup |

> **Note:** A user logoff/logon may be required for printer changes to fully apply in active sessions.
