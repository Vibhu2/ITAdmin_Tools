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
