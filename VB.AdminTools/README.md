# VB.AdminTools

Miscellaneous sysadmin utilities for day-to-day Windows administration. This module contains tools that don't belong in a dedicated module — reboot detection, environment checks, and general-purpose helpers useful across any Windows environment.

---

## Install

```powershell
Install-Module -Name VB.AdminTools -Scope CurrentUser
```

---

## Functions

| Function | Description |
|---|---|
| `Get-VBPendingReboot` | Checks multiple registry locations and system flags to determine if Windows is pending a reboot |

---

## Basic Usage

```powershell
# Check local machine
Get-VBPendingReboot

# Check a remote server
Get-VBPendingReboot -ComputerName SVR01

# Check multiple machines, filter to those requiring reboot
Get-VBPendingReboot -ComputerName SVR01, SVR02, WS001 |
    Where-Object RebootRequired |
    Select-Object ComputerName, ComponentBasedServicing, WindowsUpdateReboot, UptimeDays

# Include SCCM reboot flags
Get-VBPendingReboot -ComputerName SVR01 -IncludeSCCM

# Use alternate credentials for remote access
Get-VBPendingReboot -ComputerName SVR01 -Credential (Get-Credential)
```

### Output Properties

| Property | Description |
|---|---|
| `ComputerName` | Target computer |
| `ComponentBasedServicing` | CBS pending reboot |
| `WindowsUpdateReboot` | Windows Update requires reboot |
| `PendingFileRenameOp` | File rename operations pending |
| `ServerFeatureManager` | Server role/feature install pending |
| `ComputerRenamePending` | Computer rename pending |
| `DomainJoinReboot` | Domain join/unjoin pending |
| `SCCMReboot` | SCCM client reboot pending (requires `-IncludeSCCM`) |
| `RebootRequired` | Overall reboot status (any flag set) |
| `LastBootTime` | Last system boot timestamp |
| `UptimeDays` | Uptime in days (rounded to 2 decimal places) |
| `CollectionTime` | When the check was run |
| `Status` | `Success` or `Failed` |
| `Error` | Error message if Status is Failed |

---

## Links

- [GitHub](https://github.com/Vibhu2/ITAdmin_Tools/tree/main/VB.AdminTools)
- [Blog](https://pwsh.in)
- [License](https://github.com/Vibhu2/ITAdmin_Tools/blob/main/LICENSE)
