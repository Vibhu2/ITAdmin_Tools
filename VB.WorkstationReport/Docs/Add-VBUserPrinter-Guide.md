# Add-VBUserPrinter — Usage Guide

## Overview

Two functions work together to install printers from a print server onto a workstation:

| Function | Purpose |
|---|---|
| `Get-VBPrinterDriverInformation` | Queries the print server — returns printer name, IP, driver name, and INF path |
| `Add-VBUserPrinter` | Installs the printer and maps it into one or all user profiles on the workstation |

---

## Step 1 — Discover printers on the print server

Run this once to get a reference list of all shared printers, their IPs, driver names, and INF paths:

```powershell
$printers = Get-VBPrinterDriverInformation -ComputerName 'DSI-DH01-DC-004' |
    Where-Object Shared |
    Select-Object PrinterName, PortName, DriverName, InfPath

$printers | Format-Table -AutoSize
```

Example output:

```
PrinterName                  PortName      DriverName                    InfPath
-----------                  --------      ----------                    -------
Atlanta Printer #2           10.204.1.21   HP LaserJet 6L                C:\Windows\System32\DriverStore\...
HP Color LaserJet 4600 PCL6  10.150.1.20   HP Color LaserJet 4600 PCL6   C:\Windows\System32\DriverStore\...
HP LaserJet 4100 Series PCL  10.150.1.22   HP LaserJet 4100 Series PCL   C:\Windows\System32\DriverStore\...
```

---

## Step 2 — Understand the InfPath → DriverPath conversion

`Get-VBPrinterDriverInformation` returns `InfPath` as a local path on the server:

```
C:\Windows\System32\DriverStore\FileRepository\prnhp001.inf_amd64_xxx\prnhp001.inf
```

`Add-VBUserPrinter` needs a UNC path the workstation can reach. Convert it via the admin share (`C$`):

```
\\DSI-DH01-DC-004\C$\Windows\System32\DriverStore\FileRepository\prnhp001.inf_amd64_xxx\prnhp001.inf
```

The conversion is a one-liner — replace the drive colon with `$`:

```powershell
$driverPath = "\\$server\" + ($info.InfPath -replace ':', '$')
```

---

## Step 3 — Install a single printer

### For one specific user

```powershell
$server = 'DSI-DH01-DC-004'

$info = Get-VBPrinterDriverInformation -ComputerName $server |
    Where-Object { $_.PrinterName -eq 'Atlanta Printer #2' }

$driverPath = "\\$server\" + ($info.InfPath -replace ':', '$')

Add-VBUserPrinter -PrinterPath $info.PortName `
                  -PrinterName $info.PrinterName `
                  -DriverName  $info.DriverName `
                  -DriverPath  $driverPath `
                  -TargetUser  'setup'
```

### For all users on the machine

Omit `-TargetUser` — every non-system profile gets the printer:

```powershell
Add-VBUserPrinter -PrinterPath $info.PortName `
                  -PrinterName $info.PrinterName `
                  -DriverName  $info.DriverName `
                  -DriverPath  $driverPath
```

---

## Step 4 — Install multiple printers for all users

```powershell
$server = 'DSI-DH01-DC-004'
$wanted = @(
    'Atlanta Printer #2'
    'HP Color LaserJet 4600 PCL6'
    'HP LaserJet M607'
)

$results = Get-VBPrinterDriverInformation -ComputerName $server |
    Where-Object { $_.PrinterName -in $wanted } |
    ForEach-Object {
        $driverPath = "\\$server\" + ($_.InfPath -replace ':', '$')
        Add-VBUserPrinter -PrinterPath $_.PortName `
                          -PrinterName $_.PrinterName `
                          -DriverName  $_.DriverName `
                          -DriverPath  $driverPath
    }

$results | Format-Table ComputerName, Username, PrinterName, Action, Status -AutoSize
```

Example output — one row per user profile per printer:

```
ComputerName     Username  PrinterName                  Action  Status
------------     --------  -----------                  ------  ------
DSI-BH01-WS-002  jdoe      Atlanta Printer #2           Added   Success
DSI-BH01-WS-002  setup     Atlanta Printer #2           Added   Success
DSI-BH01-WS-002  jdoe      HP Color LaserJet 4600 PCL6  Added   Success
DSI-BH01-WS-002  setup     HP Color LaserJet 4600 PCL6  Added   Success
DSI-BH01-WS-002  jdoe      HP LaserJet M607             Added   Success
DSI-BH01-WS-002  setup     HP LaserJet M607             Added   Success
```

---

## Complete copy/paste workflow

Single printer, configurable user target:

```powershell
# ---- Config ----
$server      = 'DSI-DH01-DC-004'
$printerName = 'Atlanta Printer #2'  # printer to install
$targetUser  = 'setup'               # set to $null to install for all users

# ---- Resolve printer info ----
$info = Get-VBPrinterDriverInformation -ComputerName $server |
    Where-Object { $_.PrinterName -eq $printerName }

if (-not $info) { throw "Printer '$printerName' not found on $server" }

$driverPath = "\\$server\" + ($info.InfPath -replace ':', '$')

# ---- Build params and install ----
$params = @{
    PrinterPath = $info.PortName
    PrinterName = $info.PrinterName
    DriverName  = $info.DriverName
    DriverPath  = $driverPath
}

if ($targetUser) { $params['TargetUser'] = $targetUser }

Add-VBUserPrinter @params |
    Format-Table ComputerName, Username, PrinterName, Action, Status -AutoSize
```

---

## Dry run (WhatIf)

Always test with `-WhatIf` before deploying — no changes are made:

```powershell
Add-VBUserPrinter @params -WhatIf
```

---

## Save results to CSV (RMM logging)

```powershell
$results | Export-Csv "\\$server\Realtime\PrinterInstall_$env:COMPUTERNAME.csv" `
    -NoTypeInformation -Encoding UTF8 -Append

# Review failures
$results | Where-Object Status -eq 'Failed' | Format-Table -AutoSize
```

---

## Notes

- **Driver install is machine-level** — `pnputil` stages the driver once for all users. The machine-level step only runs once regardless of how many user profiles are targeted.
- **`-DriverPath` is skipped if the driver is already installed** — safe to re-run on the same machine.
- **UNC path requires admin share access** — the workstation must have admin rights to `\\server\C$` to read the INF from the DriverStore.
- **Logoff/logon required for active sessions** — registry changes written to `HKU\{SID}` take effect at next logon for users already logged in.
- **`AlreadyExists` is not an error** — means both the Connections key and Devices value are already present; the printer is correctly mapped.
- **`Add-VBUserPrinter` output properties** — `ComputerName`, `Username`, `SID`, `PrinterPath`, `PrinterName`, `Action`, `SetAsDefault`, `Status`, `Timestamp` (plus `Error` on failures).

---

---

# Get-VBUserPrinterMappings — Usage Guide

## Overview

`Get-VBUserPrinterMappings` audits every user profile on a machine and reports their current printer mappings. Run it before any migration to understand what's actually mapped, and again after to verify the result.

---

## Audit the local machine

```powershell
# All users, rich output object per user
Get-VBUserPrinterMappings

# TableOutput -- one flat object per user, good for piping and Export-Csv
Get-VBUserPrinterMappings -TableOutput
```

---

## Fleet audit via pipeline

```powershell
# Audit multiple machines and export to a single CSV
'WS-ACCOUNTS-01','WS-HR-02','WS-RECEPTION-01' |
    Get-VBUserPrinterMappings -TableOutput |
    Export-Csv 'C:\Reports\PrinterAudit.csv' -NoTypeInformation -Encoding UTF8
```

---

## Find machines where a specific printer is mapped

Useful for scoping a migration — shows exactly which machines and users have the old printer:

```powershell
'WS-ACCOUNTS-01','WS-HR-02','WS-RECEPTION-01' |
    Get-VBUserPrinterMappings -TableOutput |
    Where-Object { $_.NetworkPrinters -like '*OldPrintServer*' } |
    Select-Object ComputerName, Username, NetworkPrinters
```

---

## Full print server migration workflow

Run the audit at each step to stay in control:

```powershell
# Step 1 -- Audit the fleet: understand what's mapped and build your CSV from real data
'WS001','WS002','WS003' | Get-VBUserPrinterMappings -TableOutput |
    Export-Csv 'C:\Reports\PreMigration_Audit.csv' -NoTypeInformation

# Step 2 -- Dry-run the migration on one machine
Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv' -WhatIf

# Step 3 -- Run with backup on the test machine
Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv' `
    -BackupMappings -BackupPath 'C:\Logs\PrinterBackup.csv'

# Step 4 -- Verify the test machine looks right
Get-VBUserPrinterMappings -TableOutput | Format-Table -AutoSize

# Step 5 -- Deploy to the fleet via RMM

# Step 6 -- Re-audit the fleet to confirm all machines are clean
'WS001','WS002','WS003' | Get-VBUserPrinterMappings -TableOutput |
    Where-Object { $_.NetworkPrinters -like '*OldPrintServer*' }
# Expect no results -- empty means all machines migrated cleanly
```

---

---

# Set-VBUserPrinterMigration — Usage Guide

## Overview

`Set-VBUserPrinterMigration` migrates existing printer mappings in user profiles — it moves printers from one path to another without removing and re-adding them from scratch.

Use it when:
- Decommissioning a print server (UNC → IP or UNC → UNC)
- Switching printers from shared UNC paths to direct IP connections
- Rolling back from IP to UNC
- Changing IP addresses on existing direct-IP printers

**Difference from `Add-VBUserPrinter`:**

| | `Add-VBUserPrinter` | `Set-VBUserPrinterMigration` |
|---|---|---|
| Use when | User has no printer yet — installing fresh | User already has a printer — changing its path |
| Input | Printer path + driver info | Old path + new path mapping rules |
| Scope | One printer at a time | Multiple printers in one pass via CSV |
| Per-user targeting | `-TargetUser` | `-Username`, `-SID`, or CSV Username column |

Both functions are idempotent — safe to re-run. `Set-VBUserPrinterMigration` returns `AlreadyMigrated` when the new printer is already in place and the old one is already gone.

---

## Four migration types

| Type | OldPath format | NewPath format | DriverName required? |
|---|---|---|---|
| UNC → IP | `\\Server\ShareName` | `10.x.x.x` | Yes |
| UNC → UNC | `\\OldServer\ShareName` | `\\NewServer\ShareName` | No |
| IP → UNC | `10.x.x.x` | `\\Server\ShareName` | No |
| IP → IP | `10.x.x.x` | `10.y.y.y` | No — uses existing entry |

---

## CSV formats

Two CSV formats are supported. Rich format is recommended for all new deployments.

### Rich format (recommended)

Columns: `ComputerName`, `Username`, `OldPath`, `NewPath`, `DriverName`, `DriverPath`, `DefaultPrinter`

```
ComputerName,Username,OldPath,NewPath,DriverName,DriverPath,DefaultPrinter
WS-ACCOUNTS-01,jsmith,\\PrintServer\HP-LJ,10.10.5.20,HP LaserJet 400 M401,\\FileServer\Drivers\HP\hpljm401.inf,
WS-ACCOUNTS-01,*,\\PrintServer\Canon-HR,10.10.5.21,Canon iR-ADV C5235,\\FileServer\Drivers\Canon\cnz7dkd.inf,10.10.5.21\Canon-HR
*,*,\\OldServer\Zebra,\\NewServer\Zebra,,, 
```

Column details:

- `ComputerName` — machine name to match against `$env:COMPUTERNAME`. Use `*` to apply to every machine the script runs on.
- `Username` — profile to target. Use `*` (or leave blank) to apply to all profiles on the machine. Named users apply only to that specific profile.
- `OldPath` — current printer path (UNC or IP).
- `NewPath` — target printer path (UNC or IP).
- `DriverName` — required when `NewPath` is an IP address; blank for UNC destinations.
- `DriverPath` — optional path to driver source (`.inf`, `.cab`, or folder). The script installs the driver automatically if not already present.
- `DefaultPrinter` — optional. Display name of the printer to set as the user's default after migration. Leave blank to preserve the existing default.

### Legacy format

Columns: `OldPath`, `NewPath`, `DriverName` [, `DriverPath`, `PrinterName`]

```
OldPath,NewPath,DriverName
\\PrintServer\HP-Floor2,10.30.1.50,HP LaserJet 400 M401
\\PrintServer\Canon-HR,10.30.1.51,Canon Generic Plus PCL6
10.30.1.60,\\NewServer\Ricoh,
\\OldServer\Zebra,\\NewServer\Zebra,
```

Legacy format applies the same mappings to all users. Use `-Username` or `-SID` to limit which profiles are processed.

---

## ComputerName × Username behavior matrix

| ComputerName | Username | Result |
|---|---|---|
| `WS-ACCOUNTS-01` | `jsmith` | Applies only when script runs on `WS-ACCOUNTS-01`, only for profile `jsmith` |
| `WS-ACCOUNTS-01` | `*` | Applies when script runs on `WS-ACCOUNTS-01`, for every user profile |
| `*` | `jsmith` | Applies on any machine the script runs on, only for profile `jsmith` |
| `*` | `*` | Applies on any machine the script runs on, for every user profile |
| _(machine not listed)_ | — | Falls back to `ComputerName=*` rows if any exist; otherwise skips |

When a machine has both `ComputerName=WS-ACCOUNTS-01` rows and `ComputerName=*` rows, the specific machine rows take precedence — the wildcard rows are ignored for that machine.

---

## Use cases

### 1 — All users on the current machine (legacy CSV)

The simplest deployment. One CSV drives the whole machine, no user targeting needed.

```powershell
Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv'
```

### 2 — Specific users only (legacy CSV with -Username)

```powershell
Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv' `
    -Username 'jdoe', 'asmith'
```

A warning is emitted for any username not found on the machine. All other profiles are skipped.

### 3 — Target by SID (legacy CSV with -SID)

Useful for RMM scripts that must self-select on machines where the account exists:

```powershell
Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv' `
    -SID 'S-1-5-21-123456789-1001'
```

`-Username` and `-SID` can be combined — a profile matches if either condition is true.

### 4 — Machine-specific users (rich CSV, named users)

The CSV controls which machine and which users. Deploy the same CSV file to every machine via RMM — each machine self-filters on its own `ComputerName` and processes only the listed users.

```
ComputerName,Username,OldPath,NewPath,DriverName,DriverPath,DefaultPrinter
WS-ACCOUNTS-01,jsmith,\\PrintServer\HP-LJ,10.10.5.20,HP LaserJet 400 M401,\\FileServer\Drivers\HP\hpljm401.inf,10.10.5.20\HP-LJ
WS-ACCOUNTS-01,mwilliams,\\PrintServer\HP-LJ,10.10.5.20,HP LaserJet 400 M401,\\FileServer\Drivers\HP\hpljm401.inf,
WS-HR-02,abrown,\\PrintServer\HR-Colour,10.10.6.10,Xerox WorkCentre 6515,\\FileServer\Drivers\Xerox\xrxwc65.inf,10.10.6.10\HR-Colour
```

```powershell
Set-VBUserPrinterMigration -MappingCsv '\\FileServer\Scripts\PrinterMappings.csv'
```

Running on `WS-ACCOUNTS-01` processes `jsmith` and `mwilliams`. Running on `WS-HR-02` processes `abrown`. Running on any other machine does nothing (no rows match, no error).

### 5 — All users on a machine (rich CSV, wildcard Username)

Use `Username=*` to apply a mapping to every profile, without listing names:

```
ComputerName,Username,OldPath,NewPath,DriverName,DriverPath,DefaultPrinter
WS-ACCOUNTS-01,*,\\PrintServer\HP-LJ,10.10.5.20,HP LaserJet 400 M401,,
WS-ACCOUNTS-01,*,\\PrintServer\Canon-HR,10.10.5.21,Canon iR-ADV C5235,,
```

```powershell
Set-VBUserPrinterMigration -MappingCsv '\\FileServer\Scripts\PrinterMappings.csv'
```

### 6 — Fleet-wide deployment (rich CSV, ComputerName=*)

When all machines share the same mappings, use `ComputerName=*` — the script applies those rows on every machine it runs on, as a fallback when the machine is not explicitly listed.

```
ComputerName,Username,OldPath,NewPath,DriverName,DriverPath,DefaultPrinter
*,*,\\OldPrintServer\HP-LJ,\\NewPrintServer\HP-LJ,,,
*,*,\\OldPrintServer\Canon,\\NewPrintServer\Canon,,,
```

```powershell
# Deploy via RMM -- runs on every machine, applies wildcard rows
Set-VBUserPrinterMigration -MappingCsv '\\FileServer\Scripts\PrinterMappings.csv'
```

### 7 — Mixed: some users get extra printers (rich CSV, wildcard + named)

Wildcard rows apply to everyone. Named rows stack on top for specific users:

```
ComputerName,Username,OldPath,NewPath,DriverName,DriverPath,DefaultPrinter
WS-ACCOUNTS-01,*,\\PrintServer\HP-LJ,10.10.5.20,HP LaserJet 400 M401,,
WS-ACCOUNTS-01,jsmith,\\PrintServer\Color-Laser,10.10.5.25,HP Color LaserJet,,10.10.5.25\Color-Laser
```

`jsmith` gets both mappings applied. All other users on `WS-ACCOUNTS-01` get only the HP LaserJet row.

### 8 — UNC to UNC server migration (hashtable)

When migrating a print server and all printers share the same new server, a hashtable is the fastest approach — no CSV file needed:

```powershell
$mappings = @{
    '\\OldPrintServer\HP01'    = '\\NewPrintServer\HP01'
    '\\OldPrintServer\Canon02' = '\\NewPrintServer\Canon02'
    '\\OldPrintServer\Zebra'   = '\\NewPrintServer\Zebra'
}
Set-VBUserPrinterMigration -PrinterMappings $mappings
```

### 9 — UNC to IP (hashtable, all printers same driver)

When all IP printers use the same driver, supply it once via `-DriverName`:

```powershell
$mappings = @{
    '\\PrintServer\HP01' = '10.30.1.50'
    '\\PrintServer\HP02' = '10.30.1.51'
}
Set-VBUserPrinterMigration -PrinterMappings $mappings -DriverName 'HP LaserJet 400 M401'
```

For mixed drivers across printers, use `-MappingCsv` with a `DriverName` column instead.

### 10 — IP to UNC (hashtable)

```powershell
$mappings = @{
    '10.30.1.50' = '\\NewPrintServer\HP01'
    '10.30.1.51' = '\\NewPrintServer\HP02'
}
Set-VBUserPrinterMigration -PrinterMappings $mappings
```

---

## Dry run (WhatIf)

Always test first. No registry changes, no ports, no printers are created or removed:

```powershell
Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv' -WhatIf
```

WhatIf output shows every action that would be taken — port additions, printer additions, and per-user registry changes — without applying any of them.

---

## Backup before migration

```powershell
Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv' `
    -BackupMappings -BackupPath 'C:\Realtime\Reports\PrinterBackup.csv'
```

A before-snapshot of each user's printer mappings is written to the backup CSV before any changes are made. If the backup fails for a user, migration is **skipped for that user** — changes are never applied without the safety net. The backup CSV is append-safe across machines when written to a network share.

---

## RMM deployment pattern

Deploy to a fleet: the script runs locally on each machine, self-filters its rows from the CSV, and writes results to a per-machine log on a network share.

```powershell
$results = Set-VBUserPrinterMigration -MappingCsv '\\FileServer\Scripts\PrinterMappings.csv' `
    -BackupMappings -BackupPath "\\FileServer\Logs\PrinterBackup_$env:COMPUTERNAME.csv"

$results | Export-Csv -Path "\\FileServer\Logs\PrinterMigration_$env:COMPUTERNAME.csv" `
    -NoTypeInformation -Encoding UTF8

# Surface failures from collected logs
Import-Csv "\\FileServer\Logs\PrinterMigration_*.csv" |
    Where-Object { $_.Status -eq 'Failed' } |
    Format-Table ComputerName, Username, OldPath, NewPath, Details -AutoSize
```

Machine-level port and printer additions (required when `NewPath` is an IP address) only work when the script runs **locally** on the target machine. This is the expected RMM pattern.

### Consolidate results across the fleet

After the RMM job completes, pull all per-machine logs into one report:

```powershell
# Combine all machine result files into one
Get-ChildItem '\\FileServer\Logs\PrinterMigration_*.csv' |
    ForEach-Object { Import-Csv $_.FullName } |
    Export-Csv '\\FileServer\Logs\PrinterMigration_AllMachines.csv' -NoTypeInformation -Encoding UTF8 -Append

# What failed across the fleet?
Import-Csv '\\FileServer\Logs\PrinterMigration_AllMachines.csv' |
    Where-Object { $_.Status -eq 'Failed' } |
    Format-Table ComputerName, Username, OldPath, Error -AutoSize

# What was already migrated (machines that ran before)?
Import-Csv '\\FileServer\Logs\PrinterMigration_AllMachines.csv' |
    Where-Object { $_.Action -eq 'AlreadyMigrated' } |
    Select-Object ComputerName, Username | Sort-Object ComputerName

# Summary count by action across the fleet
Import-Csv '\\FileServer\Logs\PrinterMigration_AllMachines.csv' |
    Group-Object Action | Select-Object Name, Count | Sort-Object Count -Descending
```

---

## Output schema

One object is emitted per user per mapping rule processed. All objects share the same property set — safe to pipe directly to `Export-Csv`.

| Property | Description |
|---|---|
| `ComputerName` | Machine the migration ran on |
| `Username` | Profile that was processed |
| `SID` | User SID |
| `OldPath` | Old printer path from the mapping rule |
| `NewPath` | New printer path from the mapping rule |
| `Action` | `Migrated`, `Skipped`, `AlreadyMigrated`, or `Failed` |
| `Details` | Registry actions taken (on Migrated) or reason (on Skipped/Failed) |
| `Status` | `Success` or `Failed` |
| `Error` | Error message — empty on success |
| `Timestamp` | Time of action (`dd-MM-yyyy HH:mm:ss`) |

**Action values explained:**

- `Migrated` — old printer removed, new printer added in user registry. `Details` lists every registry key and value touched.
- `Skipped` — old printer was not found in this user's profile. Expected when a user does not have the printer being migrated.
- `AlreadyMigrated` — new printer already present and old printer already gone. Idempotent re-run, no changes made.
- `Failed` — an error occurred. Check the `Error` and `Details` properties.

---

## Troubleshooting

**All users return `Action: Skipped` / `Old printer not found`**

The `OldPath` in the CSV does not match what is actually in the user's registry. Diagnose with:

```powershell
# Run as the affected user (not as SYSTEM)
Get-ItemProperty 'HKCU:\Printers\Connections' | Select-Object *
Get-ItemProperty 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Devices' | Select-Object *
```

Compare the printer names shown there against the `OldPath` values in your CSV. Common mismatches:
- Server has multiple queues for the same physical printer (e.g. `HP LaserJet M607` vs `HP LaserJet 4100 Series PCL`) — verify the exact share name on the print server.
- Trailing spaces or path casing differences in the CSV.

**Printer with a `/` in its name is always skipped**

This was a known bug (fixed in `Update-VBUserPrinterRegistry` v1.0.1). The PowerShell registry provider treated `/` in printer names like `Canon iR-ADV C5030/5035 UFR II` as a path separator when using `-Path`. The fix changed all Connections key lookups to `-LiteralPath`. Verify you are on v1.0.1 or later.

**`Machine-level printer port additions are not supported for remote targets`**

The script must run locally on the target machine when any `NewPath` is an IP address. Deploy via RMM — do not call with a `-ComputerName` pointing to a remote machine for IP migrations.

**`Printer driver 'X' is not installed`**

The driver must be in the driver store before `Add-Printer` can use it. Either:
- Supply `-DriverPath` in the CSV (rich format) pointing to the `.inf` or folder — the script stages it via `pnputil` automatically.
- Pre-stage the driver manually: `pnputil /add-driver <path>\driver.inf /install`
- Run `Get-PrinterDriver | Select-Object Name` to list currently installed drivers.

**`-BackupPath` directory does not exist**

The parent directory of `BackupPath` must exist before running. The script does not create it. Create it first or use an existing path.

**Changes not visible for logged-in users**

Registry changes are written to `HKU\{SID}` — the loaded hive for active sessions or a mounted hive for offline profiles. Windows reads printer settings at logon. Users who are currently logged in will not see changes until they log off and back on.
