# VB.ServerInventory

**VB.ServerInventory** is a PSGallery module purpose-built for one mission — walk up to any Windows Server and know everything about it, automatically.

Packed with **34 specialist cmdlets across 10 categories**, every function is laser-focused on a single area of the server. System information, disk health, network configuration, Active Directory health, GPO analysis, security posture, DHCP and DNS state, Remote Desktop Services, installed applications, Windows features, scheduled tasks, and shares — nothing is left uncovered. Each cmdlet does one thing and does it perfectly, returning clean structured output that feeds directly into the next stage of the pipeline.

But the real magic isn't in any individual function — it's in the orchestration.

A single master script calls all 34 cmdlets in sequence, collects every output, and assembles it into one comprehensive server report. What would otherwise take an experienced engineer hours of manual RDP sessions, registry digs, PowerShell one-liners, and MMC console clicks is reduced to a single command. Point it at a server, walk away, and come back to a complete picture.

This makes **VB.ServerInventory** invaluable in two scenarios where complete server knowledge is non-negotiable:

**Pre-migration discovery** — before moving a server to a new environment, you need to know exactly what it's running, what it's serving, what's installed, and what dependencies exist. Missing any of it can break the migration. This module misses nothing.

**MSP documentation** — onboarding a new client environment with dozens of servers is a documentation nightmare manually. With VB.ServerInventory, every server gets audited consistently, completely, and in a fraction of the time.

There is simply no reliable manual equivalent to what 34 coordinated, purpose-built cmdlets can produce in a single automated run — and that's exactly what makes this module not just useful, but essential. 

---

## Install

```powershell
Install-Module -Name VB.ServerInventory -Scope CurrentUser
```

---

## Functions by Category

### System

| Function | Description |
| --- | --- |
| `Get-VBServerSystemInfo` | OS version, hardware specs, uptime, and system role |
| `Get-VBServerDiskInfo` | Disk volumes, free space, and partition layout |
| `Get-VBServerNetworkInfo` | Network adapters, IP configuration, DNS, and gateway |

### Active Directory

| Function | Description |
| --- | --- |
| `Get-VBServerADInfo` | Domain membership, FSMO roles, and AD site |
| `Get-VBServerADHealth` | Replication status, SYSVOL, and NETLOGON health |
| `Get-VBServerDCInfo` | Domain controller roles and replication partners |

### GPO

| Function | Description |
| --- | --- |
| `Get-VBServerGPOInfo` | Applied GPOs and their order of precedence |
| `Get-VBServerGPOSettings` | Key GPO settings affecting the server |

### Security

| Function | Description |
| --- | --- |
| `Get-VBServerBitLockerStatus` | BitLocker encryption status per volume |
| `Get-VBServerFirewallStatus` | Windows Firewall profile states and active rules |
| `Get-VBServerAzureADJoin` | Azure AD join and Hybrid join status |
| `Get-VBServerLocalAdmins` | Local Administrators group membership |
| `Get-VBServerAuditPolicy` | Audit policy settings |

### Apps & Features

| Function | Description |
| --- | --- |
| `Get-VBServerInstalledApps` | Installed software from registry |
| `Get-VBServerWindowsFeatures` | Installed Windows roles and features |

### DHCP / DNS

| Function | Description |
| --- | --- |
| `Get-VBServerDHCPInfo` | DHCP server scopes, leases, and reservations |
| `Get-VBServerDNSInfo` | DNS zones, forwarders, and conditional forwarders |
| `Get-VBServerDNSRecords` | All DNS records in all forward lookup zones |

### RDS

| Function | Description |
| --- | --- |
| `Get-VBServerRDSInfo` | RDS role, licensing, and active session count |
| `Get-VBServerRDSSessions` | Active and disconnected RDS sessions |

### Services & Tasks

| Function | Description |
| --- | --- |
| `Get-VBServerServices` | Service state, start type, and account |
| `Get-VBServerScheduledTasks` | Scheduled tasks with trigger and last run status |

### Shares & Printing

| Function | Description |
| --- | --- |
| `Get-VBServerShares` | SMB shares with paths and permissions |
| `Get-VBServerPrintQueues` | Print spooler queues and installed drivers |

---

## Basic Usage

```powershell
# System summary
Get-VBServerSystemInfo -ComputerName SVR01

# Full disk report
Get-VBServerDiskInfo -ComputerName SVR01

# AD health check
Get-VBServerADHealth -ComputerName DC01

# Security posture snapshot
Get-VBServerBitLockerStatus -ComputerName SVR01
Get-VBServerFirewallStatus -ComputerName SVR01
Get-VBServerLocalAdmins -ComputerName SVR01

# Installed applications
Get-VBServerInstalledApps -ComputerName SVR01 | Sort-Object DisplayName

# Active RDS sessions
Get-VBServerRDSSessions -ComputerName RDS01
```

---

## Links

- [PSGallery](https://www.powershellgallery.com/packages/VB.ServerInventory)
- [GitHub](https://github.com/Vibhu2/ITAdmin_Tools/tree/main/VB.ServerInventory)
- [Blog](https://pwsh.in)
- [Linkedin](https://www.linkedin.com/in/vibhu-bhatnagar-02622798/)
- [License](https://github.com/Vibhu2/ITAdmin_Tools/blob/main/LICENSE)