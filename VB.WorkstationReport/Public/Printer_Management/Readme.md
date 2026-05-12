### Modify Workflow 1: `Set-VBUserPrinterMigration` + `Update-VBUserPrinterRegistry` (Private)

**What it does:** Replaces existing printer mappings across user profiles. Handles 4 migration types: UNC→UNC, UNC→IP, IP→UNC, IP→IP.

**How it flows — 4 steps:**

**Step 1 — Input normalization (**`begin` **block):**\
Accepts either a `-MappingCsv` file or a `-PrinterMappings` hashtable and normalizes both into a standard `[OldPath, NewPath, DriverName]` list. Validates up front that IP destinations have a `DriverName` (you can't add an IP printer without a driver).

**Step 2 — Machine-level setup (runs once per machine, before the per-user loop):**\
For any UNC→IP mapping, before touching any user registry:

- Creates a TCP/IP printer port (`IP_10.30.1.50`) via `Add-PrinterPort` if it doesn't exist
- Creates or updates a machine-level printer (`Add-Printer` / `Set-Printer`) pointing to that port

This is why the function **must run locally** for IP destinations — `Add-PrinterPort` and `Add-Printer` don't work remotely. If you point it at a remote `ComputerName` with IP mappings, it throws a terminating error telling you to deploy via RMM instead.

**Step 3 — Profile targeting:**\
Calls `Get-VBUserProfile` to get all non-system profiles. If you supplied `-Username` or `-SID`, it filters down to just those users (with warnings for any not found).

**Step 4 — Per-user registry changes (delegated to** `Update-VBUserPrinterRegistry`**):**\
For each user profile:

1. Mounts the hive via `Mount-VBUserHive`
2. Optionally backs up current state via `Get-VBUserPrinterMappings`
3. Calls `Update-VBUserPrinterRegistry` to apply the changes
4. Always dismounts via `Dismount-VBUserHive` in a `finally` block

**Inside** `Update-VBUserPrinterRegistry` — the actual registry surgery for each mapping rule:

```shell
UNC → UNC:  rename Connections subkey + update Devices/PrinterPorts entries
UNC → IP:   delete Connections subkey + write Devices/PrinterPorts pointing to IP_x.x.x.x port
IP  → UNC:  add Connections subkey + update Devices/PrinterPorts
IP  → IP:   update Devices/PrinterPorts to new port name only
```

The UNC↔registry key name conversion is important to understand:

```shell
\\server\printer  →  ,,server,printer
```

Leading `\\` becomes `,,`, then each remaining `\` becomes `,`. That's how Windows stores UNC connections in the registry.

It also checks whether the migration was already done (idempotency): if the new printer is already there and the old one is already gone, it returns `AlreadyMigrated` without touching anything.

If the migrated printer was the user's default, it updates the `Windows\Device` value to point to the new printer.

---

### Modify Workflow 2: `Add-VBUserPrinter`

**What it does:** Adds a brand-new printer to user profiles — not replacing anything, just injecting a new one. It's the complement to `Set-VBUserPrinterMigration`.

**How it differs from** `Set-VBUserPrinterMigration`**:**

- Migration: old printer must already exist in the user's profile
- Add: writes to the user's profile regardless of what's already there
- Add also has a `-SetAsDefault` switch to immediately make it the user's default

**Flow is identical to migration at the infrastructure level** — machine-level port/printer first (if IP), then per-user hive mount → registry write → dismount in a `finally`. The idempotency check is built in: if the printer is already in the user's `Devices` key (or `Connections` key for UNC), it returns `AlreadyExists` and skips.

**Registry keys written per user for a UNC printer:**

```shell
Printers\Connections\,,server,share   ← connection marker (just the key existing is enough)
...\Devices\\\server\share = winspool,Ne00:
...\PrinterPorts\\\server\share = winspool,Ne00:,15,45
```

**For an IP printer:**

```shell
# No Connections key (that's only for UNC)
...\Devices\HP_Accounts = winspool,IP_10.30.1.50
...\PrinterPorts\HP_Accounts = winspool,IP_10.30.1.50,15,45
```

The `,15,45` suffix on `PrinterPorts` values are timeout values (15s read, 45s transmit) that Windows always appends.

---

### How the four functions relate

```shell
Get-VBUserPrinterMappings      ← read-only audit
        |
        | (backs up current state before changes)
        ↓
Set-VBUserPrinterMigration     ← replace existing printers
        |
        | (calls internally)
        ↓
Update-VBUserPrinterRegistry   ← private: does the actual registry writes
        ↑
Add-VBUserPrinter              ← add new printers (uses same hive pattern, writes directly)

Both modify functions use:
  Mount-VBUserHive   → load NTUSER.DAT for offline profiles
  Dismount-VBUserHive → safely unload (skip if user was already logged in)
```

Here are the practical real-world scenarios for each function, working from the most common to the most complex.

---

## 1. Audit first — always start here

Before you touch anything, run the audit to see what's actually mapped:

```powershell
# What printers does every user on this machine have?
Get-VBUserPrinterMappings

# Same, but clean objects only — good for piping
Get-VBUserPrinterMappings -TableOutput

# Export to CSV for a full fleet report
'WS001','WS002','WS003' | Get-VBUserPrinterMappings -TableOutput |
    Export-Csv 'C:\Reports\PrinterAudit.csv' -NoTypeInformation

# Find every machine where a specific printer is mapped
'WS001','WS002','WS003' | Get-VBUserPrinterMappings -TableOutput |
    Where-Object { $_.NetworkPrinters -like '*OldPrintServer*' }
```

---

## 2. `Add-VBUserPrinter` — deploying a new printer

**Scenario: A new printer has been installed and you need to push it to users.**

```powershell
# Add a UNC printer to every user profile on this machine
Add-VBUserPrinter -PrinterPath '\\PrintServer02\HP_Accounts'

# Add an IP printer (driver must already be installed on the machine)
Add-VBUserPrinter -PrinterPath '10.30.1.55' `
                  -PrinterName 'HP_Accounts' `
                  -DriverName  'HP LaserJet 400 M401'

# Add to one specific user only and make it their default
Add-VBUserPrinter -PrinterPath '\\PrintServer02\HP_Accounts' `
                  -TargetUser  'jdoe' `
                  -SetAsDefault

# Always dry-run first to see what would happen
Add-VBUserPrinter -PrinterPath '\\PrintServer02\HP_Accounts' -WhatIf
```

**RMM deployment pattern** — run this script on each machine via your RMM tool (NinjaRMM, Datto, etc.):

```powershell
$results = Add-VBUserPrinter -PrinterPath '\\PrintServer02\HP_Accounts'
$results | Export-Csv "\\FileServer\Logs\AddPrinter_$env:COMPUTERNAME.csv" `
    -NoTypeInformation -Append

# Surface any failures
$results | Where-Object { $_.Status -eq 'Failed' }
```

---

## 3. `Set-VBUserPrinterMigration` — the big one, for print server migrations

**Scenario: You're decommissioning OldPrintServer and moving everything to NewPrintServer, or moving from print server to direct IP printing.**

### Step 1 — Build your mapping CSV

```powershell
# UNC -> UNC (server rename, simplest case)
$csv = @"
OldPath,NewPath,DriverName
\\OldPrintServer\HP_Floor2,\\NewPrintServer\HP_Floor2,
\\OldPrintServer\Canon_HR,\\NewPrintServer\Canon_HR,
\\OldPrintServer\Zebra_Labels,\\NewPrintServer\Zebra_Labels,
"@
$csv | Out-File 'C:\Temp\PrinterMappings.csv' -Encoding UTF8

# UNC -> IP (ditching the print server entirely)
$csv = @"
OldPath,NewPath,DriverName
\\PrintServer\HP_Floor2,10.30.1.50,HP LaserJet 400 M401
\\PrintServer\Canon_HR,10.30.1.51,Canon Generic Plus PCL6
"@
$csv | Out-File 'C:\Temp\PrinterMappings.csv' -Encoding UTF8
```

To find the exact driver name on a machine:

```powershell
Get-PrinterDriver | Select-Object Name
```

### Step 2 — Always dry-run first

```powershell
Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv' -WhatIf
```

### Step 3 — Run with a backup snapshot

```powershell
Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv' `
    -BackupMappings `
    -BackupPath 'C:\Realtime\Reports\PrinterBackup.csv'
```

The backup writes a before-snapshot of every user's current printers to CSV before touching anything. If something goes wrong you have a record of what was there.

### Targeting specific users

```powershell
# Only migrate jdoe and asmith — skip everyone else on the machine
Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv' `
    -Username 'jdoe','asmith'

# Target by SID — useful in RMM scripts that self-select on matching machines
Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv' `
    -SID 'S-1-5-21-123456789-1001'
```

### Full RMM deployment pattern

Deploy this script to every machine in the fleet. Each machine logs its own results:

```powershell
$results = Set-VBUserPrinterMigration `
    -MappingCsv    'C:\Temp\PrinterMappings.csv' `
    -BackupMappings `
    -BackupPath    "\\FileServer\Logs\Backup_$env:COMPUTERNAME.csv"

$results | Export-Csv "\\FileServer\Logs\Migration_$env:COMPUTERNAME.csv" `
    -NoTypeInformation -Encoding UTF8

# Exit with a non-zero code if anything failed (for RMM alerting)
if ($results | Where-Object { $_.Status -eq 'Failed' }) { exit 1 }
```

After it runs across the fleet, consolidate the results:

```powershell
# Pull all machine results into one report
Get-ChildItem '\\FileServer\Logs\Migration_*.csv' |
    Import-Csv |
    Export-Csv '\\FileServer\Logs\Migration_AllMachines.csv' -NoTypeInformation -Append

# See what failed across the fleet
Import-Csv '\\FileServer\Logs\Migration_AllMachines.csv' |
    Where-Object { $_.Status -eq 'Failed' } |
    Format-Table ComputerName, Username, OldPath, Error -AutoSize

# See what was already migrated (machines that had been run before)
Import-Csv '\\FileServer\Logs\Migration_AllMachines.csv' |
    Where-Object { $_.Action -eq 'AlreadyMigrated' } |
    Select-Object ComputerName, Username | Sort-Object ComputerName
```

---

## 4. The full workflow for a print server migration

```shell
1. Get-VBUserPrinterMappings   → audit the fleet, build your mapping CSV from real data
2. Set-VBUserPrinterMigration -WhatIf  → dry-run on a test machine
3. Set-VBUserPrinterMigration -BackupMappings  → run on test machine with backup
4. Get-VBUserPrinterMappings   → verify the test machine looks right
5. Deploy via RMM to the fleet
6. Re-run Get-VBUserPrinterMappings across the fleet to confirm
```

**Important:** Users don't need to be logged off for the registry changes to apply — but they need to log off and back on for Windows to pick up the new printer mappings in their active session. You can combine step 5 with a scheduled logoff notice, or just let it apply naturally at next login.
