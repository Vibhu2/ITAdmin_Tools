---
type: Reference
title: "DNSEnrichment — PowerShell Module Design Specification (v2.0 Final)"
date: 2026-05-08
tags: [powershell, dns, enrichment, module-design, production, intern-handover]
status: final
author: VB + Claude
Version:     "2.0"
Doc_status:  "Final — Production Ready"
Environment: "On-Prem | Windows | PS 7 preferred, PS 5.1 supported"
Handover:    "Intern / AI Developer"
Supersedes:  "DNSEnrichment Module Design v1.0 (2026-05-08)"
_width: wide
---
# DNSEnrichment — PowerShell Module Design Specification (Final)

> **[i] HANDOVER NOTE:** This document is the single source of truth for building the `DNSEnrichment` PowerShell module. It is written so an intern or AI developer can implement it without making design decisions. Every contract, prerequisite, output schema, and failure mode is defined here. Raise questions before writing code, not after.

***

## Contents

1. [Purpose & Scope](#1-purpose--scope)
2. [Concerns Addressed — Resolution Map](#2-concerns-addressed--resolution-map)
3. [Architecture at a Glance](#3-architecture-at-a-glance)
4. [PowerShell Version Strategy](#4-powershell-version-strategy)
5. [Module Structure](#5-module-structure)
6. [Coding Standards — Non-Negotiable](#6-coding-standards--non-negotiable)
7. [Environment Context Object](#7-environment-context-object)
8. [Prerequisite Validation — ](#8-prerequisite-validation--get-vbenrichmentcontext)`Get-VBEnrichmentContext`
9. [Storage Layer — SQLite + CSV Export](#9-storage-layer--sqlite--csv-export)
10. [Resolution Layer Functions](#10-resolution-layer-functions)
    - 10.1 [Get-VBPTRRecord](#101-get-vbptrrecord)
    - 10.2 [Get-VBDHCPLease](#102-get-vbdhcplease)
    - 10.3 [Get-VBADComputer](#103-get-vbadcomputer)
    - 10.4 [Get-VBARPEntry](#104-get-vbarpentry)
    - 10.5 [Get-VBTCPFingerprint](#105-get-vbtcpfingerprint)
    - 10.6 [Get-VBHTTPBanner](#106-get-vbhttpbanner)
    - 10.7 [Get-VBSNMPIdentity](#107-get-vbsnmpidentity)
    - 10.8 [Get-VBOUIVendor](#108-get-vbouivendor)
    - 10.9 [Get-VBRTSPBanner](#109-get-vbrtspbanner)
    - 10.10 [Get-VBmDNSRecord](#1010-get-vbmdnsrecord)
    - 10.11 [Get-VBSwitchARP](#1011-get-vbswitcharp)
11. [Classification — ](#11-classification--resolve-vbdeviceclass)`Resolve-VBDeviceClass`
12. [Orchestration — ](#12-orchestration--invoke-vbipenrichment)`Invoke-VBIPEnrichment`
13. [Progress & Transparency Model](#13-progress--transparency-model)
14. [Output Object Contracts](#14-output-object-contracts)
15. [Error Handling Contract](#15-error-handling-contract)
16. [Resolution Order — Most Reliable First](#16-resolution-order--most-reliable-first)
17. [Environment Scenarios — Capability Matrix](#17-environment-scenarios--capability-matrix)
18. [Caching, Re-Runs & DHCP Churn Handling](#18-caching-re-runs--dhcp-churn-handling)
19. [Testing Checklist](#19-testing-checklist)
20. [Build Sequence](#20-build-sequence)
21. [Decisions Locked from v3 Design Doc](#21-decisions-locked-from-v3-design-doc)
22. [Changelog](#22-changelog)

***

## 1. Purpose & Scope

The `DNSEnrichment` PowerShell module takes a list of private/RFC1918 IP addresses (typically extracted from Windows DNS debug logs) and resolves each one to:

- A **hostname** (from the most reliable available source)
- A **device class** (`Workstation`, `Server`, `Printer`, `IPPhone`, `Camera`, etc.)
- Full **enrichment metadata** — MAC, vendor, model, OS, location, open ports, source layer

It does this through a configurable 11-step pipeline (10 detection layers + ARP cache helper), running steps in priority order. It is fully transparent about what it can and cannot do in the current environment, reports prerequisite status before any probing begins, and produces structured objects that feed directly into DNS log reports, SQLite, and CSV.

**Out of scope for this module:** DNS log parsing, report generation, scheduled refresh. Those are separate modules that consume `DNSEnrichment` output.

***

## 2. Concerns Addressed — Resolution Map

This table maps every concern from the "What I'd push back on" review to the section that resolves it.

| #  | Concern                                            | Resolution                                                                                                                 | Section      |
| -- | -------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------- | ------------ |
| 1  | "You don't know what environment this will run in" | Mandatory `Get-VBEnrichmentContext` runs first; every layer reads context flags                                            | §7, §8       |
| 2  | "Silent failures are invisible"                    | Every function returns `Status` ∈ `Success`/`NoResult`/`Failed`/`Skipped` — never `$null`                                  | §14, §15     |
| 3  | "The pipeline is a black box"                      | `Write-Progress` per IP; `Write-Verbose` per layer; `LayerTrace` array on every result                                     | §13          |
| 4  | "Functions are tightly coupled"                    | Every layer is independently callable; classification is a separate function                                               | §10, §11     |
| 5  | "PS version incompatibility breaks silently"       | Compatibility-first baseline (PS 5.1), opportunistic PS 7 enhancements via `$Context.CanUseParallel` etc.                  | §4           |
| 6  | "No MAC, no OUI — silent gap"                      | Dedicated `Get-VBARPEntry` step; OUI accepts null gracefully with `SkipReason`                                             | §10.4, §10.8 |
| 7  | "SNMP community hardcoded"                         | Array of strings tried in order, recorded which succeeded                                                                  | §10.7        |
| 8  | "Running on a DC isn't exploited"                  | Context detects `IsDomainController`, `DNSIsLocal`, `DHCPIsLocal`; layers annotate source as `(local)` vs `(remote)`       | §7, §8       |
| 9  | "TCP scan is sequential & slow"                    | Parallel runspaces on PS 7 (`ForEach-Object -Parallel`); throttled at orchestrator level                                   | §4, §12      |
| 10 | "Layers must skip already-resolved IPs"            | Orchestrator gates Hostname-resolving layers on `IsResolved`; enrichment-only layers always run                            | §12, §16     |
| 11 | "No caching across runs"                           | SQLite enrichment table; on re-run only unresolved/stale IPs are re-probed; DHCP churn detected via MAC change             | §9, §18      |
| 12 | "RFC1918 isn't the only 'private' space"           | Context recognizes RFC1918, CGNAT (100.64/10), link-local (169.254/16)                                                     | §7           |
| 13 | "mDNS needs Bonjour"                               | Context detects `dns-sd.exe`; if missing, layer returns `Skipped` with documented impact — no third-party install required | §8, §10.10   |
| 14 | "Switch ARP is its own sub-project"                | Layer 10 is opt-in — only runs if `$Context.SwitchTargets` configured; gracefully skipped otherwise                        | §10.11       |
| 15 | "Can't export to CSV easily"                       | `Export-VBEnrichmentResult` accepts `-Format CSV                                                                           | Json         |

***

## 3. Architecture at a Glance

```typescript
┌────────────────────────────────────────────────────────────────┐
│                 Get-VBEnrichmentContext (MANDATORY)            │
│     Probes environment → builds context object → prints report │
│     Detects: PS version, DC role, AD/DHCP/DNS/SNMP/Bonjour     │
└────────────────────────────┬───────────────────────────────────┘
                             │ $Context flows to every function
                             ▼
┌────────────────────────────────────────────────────────────────┐
│            Invoke-VBIPEnrichment   (orchestrator)              │
│                                                                │
│   Read existing rows from SQLite → decide which IPs to probe   │
│                                                                │
│   For each unresolved IP:                                      │
│     ┌──────────────────────────────────────────────────────┐   │
│     │  PASSIVE LAYERS  (fast, no target connection)        │   │
│     │   1. AD          ⟵ highest confidence — runs first  │   │
│     │   2. DHCP        ⟵ second — provides MAC + lease    │   │
│     │   3. PTR         ⟵ third — forward-confirmed only   │   │
│     │   4. ARP cache   ⟵ collects MAC for OUI             │   │
│     ├──────────────────────────────────────────────────────┤   │
│     │  ACTIVE LAYERS   (parallel on PS 7)                  │   │
│     │   5. TCP fingerprint                                 │   │
│     │   6. HTTP banner   (gated on TCP 80/443/8080/8443)   │   │
│     │   7. SNMP                                            │   │
│     │   8. RTSP          (gated on TCP 554)                │   │
│     │   9. mDNS          (if Bonjour present)              │   │
│     │  10. Switch ARP    (if SwitchTargets configured)     │   │
│     ├──────────────────────────────────────────────────────┤   │
│     │  ENRICHMENT-ONLY                                     │   │
│     │  11. OUI vendor    (always runs if MAC known)        │   │
│     └──────────────────────────────────────────────────────┘   │
│                                                                │
│   Resolve-VBDeviceClass (once, after all signals collected)    │
│                                                                │
│   Write/Update SQLite row                                      │
└─────────────────────────────┬──────────────────────────────────┘
                              ▼
                  ┌───────────────────────┐
                  │  Export-VBEnrichment  │
                  │  CSV / JSON / Object  │
                  └───────────────────────┘
```

**Key principles:**

1. **Context first, always.** No layer ever runs without checking the context.
2. **One source of truth.** SQLite is authoritative across runs. In-memory hashtable is authoritative within a run.
3. **Independently callable.** Every `Get-VB*` function works standalone with just an IP and (optionally) a context.
4. **Skip-if-resolved gate is explicit.** The orchestrator owns the gating; layers don't know about each other.
5. **Status, never $null.** Every layer returns a structured result with a status field.

***

## 4. PowerShell Version Strategy

### Target

- **Preferred:** PowerShell 7.x (Server 2019+, Windows 10/11, side-by-side with 5.1)
- **Supported:** PowerShell 5.1 on Windows Server 2012 R2 and later
- **Not supported:** PowerShell ≤ 4.0 / Server 2008 R2 (these lack `Resolve-DnsName` and most cmdlets used)

### Strategy — Compatibility-First, Enhancement-Opportunistic

The module is **written for PS 5.1 compatibility as the baseline**. PS 7 features are used only where they provide a meaningful benefit AND a PS 5.1 fallback is provided inline. Version detection happens once in `Get-VBEnrichmentContext` and is recorded in `$Context.PSMajor`.

### Per-Feature Compatibility Matrix

| Feature                                   | PS 5.1          | PS 6+           | PS 7+           | PS 5.1 Fallback                                                                                                   |
| ----------------------------------------- | --------------- | --------------- | --------------- | ----------------------------------------------------------------------------------------------------------------- |
| `Invoke-WebRequest -SkipCertificateCheck` | No              | Yes             | Yes             | `[Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }` — must reset to `$null` in `finally` |
| `[System.Net.Sockets.TcpClient]` async    | Yes             | Yes             | Yes             | No fallback needed                                                                                                |
| `ForEach-Object -Parallel`                | No              | No              | Yes             | Sequential `foreach` — runspaces only on PS 7                                                                     |
| `Get-DhcpServerv4Lease`                   | Yes (RSAT)      | Yes (RSAT)      | Yes (RSAT)      | RSAT check in context; skip if absent                                                                             |
| `Get-ADComputer`                          | Yes (RSAT)      | Yes (RSAT)      | Yes (RSAT)      | RSAT check in context; skip if absent                                                                             |
| `Resolve-DnsName`                         | Yes (WS2012+)   | Yes             | Yes             | `[System.Net.Dns]::GetHostEntry()`                                                                                |
| `New-Object -ComObject olePrn.OleSNMP`    | Yes             | Yes             | Yes             | None — Windows-native                                                                                             |
| `PSSQLite` module                         | Yes (PSGallery) | Yes (PSGallery) | Yes (PSGallery) | Bundle `System.Data.SQLite.dll` if PSGallery blocked                                                              |
| `Test-Connection -Quiet` async            | Sync only       | Yes             | Yes             | Sequential pings                                                                                                  |

### Rules for Developers

- **Never** put `#Requires -Version 7` at the module level — it crashes 5.1 hosts.
- Functions that genuinely need PS 6/7 features check `$Context.PSMajor -lt 6` first and return `Status = 'Skipped'` with `SkipReason = 'RequiresPS6'` if needed.
- `Get-VBEnrichmentContext` sets these capability flags — read them, never re-detect in layer functions:

| Flag               | True When                                                       |
| ------------------ | --------------------------------------------------------------- |
| `CanUseParallel`   | `PSMajor -ge 7` AND IP list size > 10                           |
| `CanSkipCertCheck` | `PSMajor -ge 6` (native flag) OR PS 5.1 with workaround applied |
| `CanUsePSSQLite`   | `PSSQLite` module loadable                                      |

### Parallel Execution

When `$Context.CanUseParallel = $true`, the orchestrator runs **active probe layers** (TCP, HTTP, SNMP, RTSP, mDNS) using `ForEach-Object -Parallel -ThrottleLimit $Context.ParallelThrottleLimit` (default 10). **Passive layers** (AD, DHCP, PTR) always run sequentially — they hit shared infrastructure and parallel hammering is inappropriate. Classification always runs sequentially after parallel results are collected per IP.

***

## 5. Module Structure

```javascript
DNSEnrichment/
├── DNSEnrichment.psd1                    # Module manifest
├── DNSEnrichment.psm1                    # Root — dot-sources Public/Private/Classes
├── Classes/
│   └── VBLayerResult.ps1                 # Optional class for stronger typing (PS 5.1+)
├── Public/
│   ├── Get-VBEnrichmentContext.ps1       # Prerequisite validation
│   ├── Get-VBPTRRecord.ps1               # Layer 1
│   ├── Get-VBDHCPLease.ps1               # Layer 2
│   ├── Get-VBADComputer.ps1              # Layer 3
│   ├── Get-VBARPEntry.ps1                # ARP helper (between passive & active)
│   ├── Get-VBTCPFingerprint.ps1          # Layer 4
│   ├── Get-VBHTTPBanner.ps1              # Layer 5
│   ├── Get-VBSNMPIdentity.ps1            # Layer 6
│   ├── Get-VBOUIVendor.ps1               # Layer 7 (enrichment-only)
│   ├── Get-VBRTSPBanner.ps1              # Layer 8
│   ├── Get-VBmDNSRecord.ps1              # Layer 9
│   ├── Get-VBSwitchARP.ps1               # Layer 10
│   ├── Resolve-VBDeviceClass.ps1         # Classification
│   ├── Invoke-VBIPEnrichment.ps1         # Orchestrator
│   ├── Get-VBEnrichmentResult.ps1        # Read from SQLite
│   ├── Export-VBEnrichmentResult.ps1     # Export to CSV/JSON
│   └── Initialize-VBEnrichmentDatabase.ps1   # Create / migrate SQLite schema
├── Private/
│   ├── Write-VBEnrichmentProgress.ps1    # Internal progress writer
│   ├── New-VBLayerResult.ps1             # Standardised layer result factory
│   ├── ConvertTo-VBNormalisedMAC.ps1     # MAC normalisation (strip separators, uppercase)
│   ├── Test-VBPrivateIP.ps1              # RFC1918 + CGNAT + link-local check
│   └── Invoke-VBSqliteCommand.ps1        # Thin SQLite wrapper
├── Data/
│   └── oui.txt                           # IEEE OUI database (refresh quarterly)
├── Sql/
│   ├── 001_init.sql                      # Initial schema
│   └── 002_indexes.sql                   # Future migrations numbered
└── Tests/
    ├── *.Tests.ps1                       # One Pester file per public function
    └── Integration.Tests.ps1             # End-to-end orchestrator tests
```

**Naming convention:** All public functions are prefixed `VB`. Private helpers also prefixed `VB` but not exported. Filename matches function name.

***

## 6. Coding Standards — Non-Negotiable

These apply to every function in the module without exception.

- **No aliases.** `foreach` not `%`, `Where-Object` not `?`, `Get-ChildItem` not `gci`.
- **No** `Get-WmiObject`**.** Use `Get-CimInstance` with `-OperationTimeoutSec` always set.
- **No positional parameters** in function calls — always named.
- **No** `Write-Host`**.** Use `Write-Verbose` for operational messages, `Write-Progress` for pipeline progress, `Write-Warning` for degraded conditions.
- **No silent catches.** Every `catch` block sets `Status = 'Failed'` and captures `$_.Exception.Message` into `ErrorDetail`.
- **All functions return** `[PSCustomObject]` — never raw strings, never `$null`.
- **Every function has** `[CmdletBinding()]`**.**
- **Every parameter has an explicit type declaration.**
- **All file output uses UTF-8 with BOM** (`-Encoding utf8BOM` in PS 7, `[System.Text.Encoding]::UTF8` with preamble in PS 5.1).
- **All TCP/COM resources released in** `finally` **blocks** — never only in success paths.
- **Version bumped on every change** — patch for fixes, minor for new features, major for breaking contracts.

### Function Header Template (Mandatory)

```powershell
<#
.SYNOPSIS
    One-line description.

.DESCRIPTION
    Full description. What it does, what it returns, what it skips and why.

.PARAMETER IPAddress
    The RFC1918 / CGNAT / link-local IP address to probe.

.PARAMETER Context
    The environment context object from Get-VBEnrichmentContext.
    If not supplied, the function emits Write-Warning and runs with no
    prerequisite awareness — may produce incomplete results.

.OUTPUTS
    [PSCustomObject] — see Section 14 of the module design spec.

.EXAMPLE
    Get-VBPTRRecord -IPAddress '192.168.1.45' -Context $ctx

.NOTES
    Version:      1.0.0
    MinPSVersion: 5.1
    Author:       VB
    ChangeLog:
        1.0.0 — 2026-05-08 — Initial release
#>
```

***

## 7. Environment Context Object

The context object is the foundation of the entire module. It is produced by `Get-VBEnrichmentContext` and passed to every function. **No function calls** `Get-ADComputer` **without first checking** `$Context.ADAvailable`**.** **No function probes port 80 without checking** `$Context.NetworkProbeEnabled`**.**

### Full Schema

```powershell
[PSCustomObject]@{

    # --- PowerShell environment ---
    PSVersion             = $PSVersionTable.PSVersion       # [Version]
    PSMajor               = $PSVersionTable.PSVersion.Major # [int] 5, 6, or 7
    PSEdition             = $PSVersionTable.PSEdition       # 'Desktop' | 'Core'
    CanUseParallel        = $false                          # PSMajor -ge 7
    CanSkipCertCheck      = $false                          # PSMajor -ge 6 OR workaround applied
    CanUsePSSQLite        = $false                          # PSSQLite module loadable

    # --- Machine identity ---
    ComputerName          = $env:COMPUTERNAME
    IsDomainJoined        = $false
    IsDomainController    = $false
    DomainName            = $null

    # --- Layer availability flags ---
    DNSAvailable          = $false                          # Resolve-DnsName / GetHostEntry works
    DNSIsLocal            = $false                          # Local DNS server role detected
    DNSServer             = $null                           # Configured DNS server FQDN/IP
    DHCPAvailable         = $false                          # DhcpServer module loaded AND server reachable
    DHCPIsLocal           = $false                          # Local DHCP server role detected
    DHCPServer            = $null                           # FQDN or IP
    DHCPScopeIds          = @()                             # Auto-discovered or supplied
    ADAvailable           = $false                          # AD module loaded AND Get-ADDomain works
    ADIsLocal             = $false                          # Running on a DC
    NetworkProbeEnabled   = $true                           # Master switch for active probes (4-10)
    SNMPAvailable         = $false                          # olePrn COM creatable
    SNMPCommunityStrings  = @('public')                     # Tried in order
    OUIFileAvailable      = $false                          # oui.txt found
    OUIFilePath           = $null                           # Resolved path
    OUIFileAge            = $null                           # [TimeSpan] since last refresh
    mDNSAvailable         = $false                          # dns-sd.exe found on PATH
    RTSPProbeEnabled      = $true                           # Sub-switch for layer 8
    SwitchTargets         = @()                             # IPs of managed switches for layer 10

    # --- Storage ---
    DatabasePath          = $null                           # Path to SQLite file
    DatabaseInitialized   = $false                          # Schema present and current

    # --- Performance ---
    ParallelThrottleLimit = 10
    DefaultTimeoutMs      = @{
        TCP   = 300
        HTTP  = 3000
        SNMP  = 2000
        RTSP  = 2000
    }

    # --- Reporting ---
    PrerequisiteReport    = @()                             # [PSCustomObject[]] one row per check
    ContextBuiltAt        = (Get-Date)
    ContextDurationMs     = 0
}
```

### Private IP Recognition

`Test-VBPrivateIP` (private helper) returns `$true` for any of:

- RFC1918: `10.0.0.0/8`, `172.16.0.0/12`, `192.168.0.0/16`
- CGNAT: `100.64.0.0/10`
- Link-local: `169.254.0.0/16`
- Loopback: `127.0.0.0/8`

The orchestrator uses this to validate input IPs and skip public addresses with a warning.

***

## 8. Prerequisite Validation — `Get-VBEnrichmentContext`

**File:** `Public/Get-VBEnrichmentContext.ps1`\
\
**MinPSVersion:** 5.1\
\
**Purpose:** Build and return the environment context object. Print the prerequisite report to console (unless `-Quiet`). **This function MUST be called before any other function in the module.**

### Parameters

```powershell
param(
    [string]$DHCPServer,
    [string[]]$DHCPScopeIds,
    [string[]]$SwitchTargets         = @(),
    [string[]]$SNMPCommunityStrings  = @('public'),
    [string]$OUIFilePath,                    # default: $PSScriptRoot\..\Data\oui.txt
    [string]$DatabasePath,                   # default: $env:LOCALAPPDATA\DNSEnrichment\enrichment.db
    [int]$ParallelThrottleLimit      = 10,
    [switch]$DisableNetworkProbe,            # skips layers 4-10 entirely
    [switch]$Quiet                           # suppress console prerequisite report
)
```

### Checks Performed (in order)

| #  | Check                 | Method                                                                                       | Sets Flag                                       |
| -- | --------------------- | -------------------------------------------------------------------------------------------- | ----------------------------------------------- |
| 1  | PowerShell version    | `$PSVersionTable.PSVersion.Major`                                                            | `PSMajor`, `CanUseParallel`, `CanSkipCertCheck` |
| 2  | Domain join           | `(Get-CimInstance Win32_ComputerSystem).PartOfDomain`                                        | `IsDomainJoined`, `DomainName`                  |
| 3  | DC role               | Registry: `HKLM:\SYSTEM\CurrentControlSet\Control\ProductOptions` `ProductType` = `LanmanNT` | `IsDomainController`                            |
| 4  | DNS resolution works  | `Resolve-DnsName 127.0.0.1 -Type PTR -EA Stop` (fallback `[System.Net.Dns]::GetHostEntry`)   | `DNSAvailable`                                  |
| 5  | Local DNS role        | `Get-Service DNS -EA SilentlyContinue`                                                       | `DNSIsLocal`                                    |
| 6  | AD module             | `Import-Module ActiveDirectory -EA SilentlyContinue; Get-ADDomain -EA Stop`                  | `ADAvailable`, `ADIsLocal`                      |
| 7  | DHCP module           | `Import-Module DhcpServer -EA SilentlyContinue`                                              | `DHCPAvailable`                                 |
| 8  | DHCP server reachable | `Get-DhcpServerv4Scope -ComputerName $DHCPServer -EA Stop` (5s timeout)                      | `DHCPAvailable` confirmed                       |
| 9  | Local DHCP role       | `Get-Service DHCPServer -EA SilentlyContinue`                                                | `DHCPIsLocal`                                   |
| 10 | SNMP COM              | `New-Object -ComObject olePrn.OleSNMP` then release                                          | `SNMPAvailable`                                 |
| 11 | OUI file              | `Test-Path $OUIFilePath` AND `(Get-Item).Length -gt 1MB`                                     | `OUIFileAvailable`, `OUIFileAge`                |
| 12 | mDNS                  | `Get-Command dns-sd.exe -EA SilentlyContinue`                                                | `mDNSAvailable`                                 |
| 13 | PSSQLite              | `Get-Module PSSQLite -ListAvailable`                                                         | `CanUsePSSQLite`                                |
| 14 | SQLite database       | `Initialize-VBEnrichmentDatabase` if missing                                                 | `DatabaseInitialized`                           |

### Console Output (printed unless `-Quiet`)

```typescript
═══════════════════════════════════════════════════════════════════════
  DNSEnrichment — Environment Prerequisites
═══════════════════════════════════════════════════════════════════════
  Computer    : DC01.corp.local (Domain Controller)
  PowerShell  : 7.4.1 Core      → Parallel execution: ENABLED
  Database    : C:\...\enrichment.db (initialized, 47 rows)
  Probes      : Active probes ENABLED (network reachable)
───────────────────────────────────────────────────────────────────────

  Layer  Prerequisite          Status         Detail
  ─────  ──────────────────    ────────────   ───────────────────────────
   1     DNS (PTR)             Available      Local DNS role (fastest path)
   2     DHCP Leases           Available      Local DHCP role
   3     Active Directory      Available      Local DC — auth-free queries
   4     ARP Cache             Available      Native — `arp -a`
   5     TCP Port Scan         Available      Network probe enabled
   6     HTTP Banner           Available      PS 7 native cert bypass
   7     SNMP (olePrn)         Available      COM object created OK
   8     OUI Vendor Lookup     Available      oui.txt 4.1 MB, age 12 days
   9     RTSP Probe            Available      Network probe enabled
  10     mDNS Discovery        UNAVAILABLE    dns-sd.exe not on PATH
                                              → Layer 9 will be SKIPPED
                                              → Impact: printers/scanners
                                                may be missed if they
                                                rely on mDNS only
  11     Switch ARP            Not Configured No SwitchTargets supplied
                                              → Layer 10 will be SKIPPED
                                              → Impact: unresolved IPs
                                                won't get switch port loc

═══════════════════════════════════════════════════════════════════════
  Summary: 9 of 11 layers AVAILABLE
  Skipped : Layer 9 (mDNS), Layer 10 (Switch)
  Mode    : Parallel × 10 for active probes (4-9)
═══════════════════════════════════════════════════════════════════════
```

### `PrerequisiteReport` Entry Schema

```powershell
[PSCustomObject]@{
    Layer            = 6
    Prerequisite     = 'SNMP (olePrn)'
    Status           = 'Available'              # Available | Unavailable | NotConfigured | Degraded
    Detail           = 'COM object created OK'
    LayersAffected   = 'Layer 6 (SNMP), Layer 10 (Switch ARP)'
    SkippedLayers    = ''                       # populated when Status != Available
    Impact           = ''                       # populated when Status != Available
    Remediation      = ''                       # how to fix — populated when Status != Available
}
```

***

## 9. Storage Layer — SQLite + CSV Export

### Why SQLite

Locked from v3 design Q1: **SQLite**. Reasons:

- File-based, zero install on PS 7 with `PSSQLite` (PSGallery available per Q11).
- Survives runs without paying SQL Server licensing or admin overhead.
- CSV scales poorly past ~10k IPs; PowerShell objects don't survive between runs.
- Native `LEFT JOIN` against the parsed DNS log (in the same DB) for reports.

### Schema — `enrichment.db`

```sql
-- 001_init.sql
CREATE TABLE IF NOT EXISTS Enrichment (
    IPAddress              TEXT PRIMARY KEY,
    Hostname               TEXT,
    HostnameSource         TEXT,
    MACAddress             TEXT,
    MACAddressNormalised   TEXT,
    Vendor                 TEXT,
    DeviceClass            TEXT,
    DeviceClassSource      TEXT,
    Confidence             TEXT,
    OSClass                TEXT,
    OperatingSystem        TEXT,
    Model                  TEXT,
    Location               TEXT,
    OU                     TEXT,
    OpenPorts              TEXT,
    HTTPTitle              TEXT,
    HTTPServer             TEXT,
    SNMPDescr              TEXT,
    RTSPBanner             TEXT,
    MDNSServiceType        TEXT,
    LeaseExpiry            TEXT,                 -- ISO 8601
    StepsAttempted         INTEGER,
    StepsSucceeded         INTEGER,
    StepsNoResult          INTEGER,
    StepsSkipped           INTEGER,
    StepsFailed            INTEGER,
    LayerTraceJson         TEXT,                 -- JSON array
    IsResolved             INTEGER,
    IsUnresolved           INTEGER,
    EnrichedAt             TEXT NOT NULL,        -- ISO 8601
    EnrichmentDurationMs   INTEGER,
    FirstSeenAt            TEXT NOT NULL,        -- ISO 8601 — never updated after first insert
    UpdatedAt              TEXT NOT NULL         -- ISO 8601 — bumped every time row changes
);

CREATE INDEX IF NOT EXISTS IX_Enrichment_DeviceClass ON Enrichment(DeviceClass);
CREATE INDEX IF NOT EXISTS IX_Enrichment_Hostname    ON Enrichment(Hostname);
CREATE INDEX IF NOT EXISTS IX_Enrichment_MAC         ON Enrichment(MACAddressNormalised);

-- Audit table — captures every change for DHCP churn analysis
CREATE TABLE IF NOT EXISTS EnrichmentHistory (
    Id                     INTEGER PRIMARY KEY AUTOINCREMENT,
    IPAddress              TEXT NOT NULL,
    OldHostname            TEXT,
    NewHostname            TEXT,
    OldMACAddress          TEXT,
    NewMACAddress          TEXT,
    OldDeviceClass         TEXT,
    NewDeviceClass         TEXT,
    ChangeReason           TEXT,                 -- 'NewIP' | 'HostnameChanged' | 'MACChanged' | 'ClassChanged' | 'StaleRefresh'
    ChangedAt              TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS IX_History_IP ON EnrichmentHistory(IPAddress);

-- Schema version tracking
CREATE TABLE IF NOT EXISTS SchemaVersion (
    Version    INTEGER PRIMARY KEY,
    AppliedAt  TEXT NOT NULL
);
INSERT OR IGNORE INTO SchemaVersion (Version, AppliedAt)
VALUES (1, datetime('now'));
```

### Public Storage Functions

#### `Initialize-VBEnrichmentDatabase`

```powershell
param(
    [string]$DatabasePath,
    [PSCustomObject]$Context
)
```

- Creates the database file if missing.
- Applies migrations from `Sql/*.sql` in order, tracked by `SchemaVersion` table.
- Returns `[PSCustomObject]@{ Path; Created; Version; Status }`.
- Idempotent — safe to call every run.

#### `Get-VBEnrichmentResult`

```powershell
param(
    [string[]]$IPAddress,                # filter by IPs; null = all
    [string]$DeviceClass,                # filter
    [datetime]$Since,                    # UpdatedAt >= $Since
    [switch]$IncludeHistory,             # join EnrichmentHistory
    [PSCustomObject]$Context
)
```

Returns enrichment rows from SQLite as `[PSCustomObject[]]`. Used by reports and by the orchestrator on subsequent runs to skip already-resolved IPs.

#### `Export-VBEnrichmentResult`

```powershell
param(
    [PSCustomObject[]]$InputObject,      # accepts pipeline
    [ValidateSet('CSV','JSON','SQLite','Object')]
    [string]$Format = 'Object',
    [string]$Path,                       # output file (CSV/JSON only)
    [switch]$IncludeLayerTrace,
    [PSCustomObject]$Context
)
```

- `Object` → returns `[PSCustomObject[]]` to pipeline.
- `CSV` → UTF-8 BOM, `LayerTrace` flattened to JSON in one column.
- `JSON` → pretty-printed JSON.
- `SQLite` → upsert into the configured database.

***

## 10. Resolution Layer Functions

### Common Contract

Every layer function obeys this contract:

- **Input:** `[string]$IPAddress` (mandatory, pipeline) + `[PSCustomObject]$Context` (recommended).
- **Output:** A single `[PSCustomObject]` conforming to the Layer Result schema (§14) + layer-specific fields.
- **Behaviour if context missing:** Emits `Write-Warning`, runs best-effort, may produce incomplete results.
- **Behaviour if prerequisite unavailable:** Returns immediately with `Status = 'Skipped'`, `SkipReason` populated, no probe attempted.
- **Pipeline support:** All functions accept `ValueFromPipeline` on `$IPAddress` AND `ValueFromPipelineByPropertyName` so prior layer results flow through naturally.

### Skip-If-Resolved Discipline

Layers are categorised:

| Category               | Layers                                                                   | Behaviour                                                                                                  |
| ---------------------- | ------------------------------------------------------------------------ | ---------------------------------------------------------------------------------------------------------- |
| **Hostname-resolving** | AD, DHCP, PTR, SNMP, HTTP (when title is conclusive), RTSP, mDNS, Switch | Honour skip-if-resolved gate from orchestrator                                                             |
| **Enrichment-only**    | TCP fingerprint, ARP, OUI                                                | Always run when prerequisites met — they add fields the reports need (`OpenPorts`, `MACAddress`, `Vendor`) |

The orchestrator owns the gate. Layer functions don't know they were skipped — the orchestrator simply doesn't call them.

***

### 10.1 `Get-VBPTRRecord`

**Layer:** 3 (was 1 in v3 — see §16 for new ordering rationale)\
\
**MinPSVersion:** 5.1\
\
**Prerequisite:** `$Context.DNSAvailable`\
\
**Cost:** Single DNS query

**Logic:**

1. If `$Context.DNSAvailable -eq $false`: return `Status = 'Skipped'`, `SkipReason = 'DNSUnavailable'`.
2. Try `Resolve-DnsName -Name $IPAddress -Type PTR -ErrorAction Stop`. PS 5.1 fallback: `[System.Net.Dns]::GetHostEntry($IPAddress).HostName`.
3. **Forward-confirm:** issue `Resolve-DnsName -Name $hostname -Type A` and check that the resolved IP matches the original. Records `ForwardConfirmed = $true|$false`.
4. On success: `Hostname`, `Status = 'Success'`. If forward-confirm fails, set `Confidence = 'Low'` so classification can de-weight stale PTR records.
5. On no record: `Status = 'NoResult'`.
6. On error: `Status = 'Failed'`, `ErrorDetail`.

**Output fields (in addition to base):**

```powershell
Hostname          = [string]
PTRRecord         = [string]
ForwardConfirmed  = [bool]
```

***

### 10.2 `Get-VBDHCPLease`

**Layer:** 2\
\
**MinPSVersion:** 5.1\
\
**Prerequisite:** `$Context.DHCPAvailable`\
\
**Cost:** Single RPC/local call

**Logic:**

1. Check `DHCPAvailable` flag.
2. **One-shot cache:** on first call, enumerate ALL leases across `$Context.DHCPScopeIds` into a script-scope hashtable keyed by IP. Subsequent calls hit the hashtable. The cache is invalidated when `Get-VBEnrichmentContext` is re-run.
3. Match by IP. Return `Hostname`, `MACAddress`, `LeaseExpiry`, `ScopeId`.
4. **Lease validity:** if `LeaseExpiry < (Get-Date)`, set `Status = 'Success'` but `IsLeaseExpired = $true` so classification can flag DHCP churn risk.

**Output fields:**

```powershell
Hostname        = [string]
MACAddress      = [string]
LeaseExpiry     = [datetime]
IsLeaseExpired  = [bool]
ScopeId         = [string]
```

***

### 10.3 `Get-VBADComputer`

**Layer:** 1 (highest priority — see §16)\
\
**MinPSVersion:** 5.1\
\
**Prerequisite:** `$Context.ADAvailable`\
\
**Cost:** Single LDAP query (or zero — see caching below)

**Logic:**

1. **One-shot cache:** on first call, run `Get-ADComputer -Filter * -Properties IPv4Address, OperatingSystem, OperatingSystemVersion, DistinguishedName` once and build a hashtable keyed by `IPv4Address`. **Critical** — never query AD per-IP.
2. Look up `$IPAddress` in the cache.
3. Determine `OSClass`:
   - DC: `(Get-ADDomainController -Filter *).IPv4Address` — also cached on first call.
   - `OperatingSystem -match 'Server'` → `Server`.
   - Else → `Workstation`.
4. Extract `OU` from `DistinguishedName`.

**Output fields:**

```powershell
Hostname           = [string]
OperatingSystem    = [string]
OSClass            = [string]   # Workstation | Server | DomainController
OU                 = [string]
DistinguishedName  = [string]
```

***

### 10.4 `Get-VBARPEntry`

**Layer:** 4 (helper — runs after passive layers, before active probes)\
\
**MinPSVersion:** 5.1\
\
**Prerequisite:** None\
\
**Cost:** Native command, instant — `arp -a`

**Logic:**

1. **One-shot cache:** parse `arp -a` output once into a hashtable keyed by IP.
2. Look up `$IPAddress`. If found, return MAC.
3. If not in local ARP cache, optionally ping the IP first (1 packet, 200 ms timeout) to populate the OS ARP table, then re-check.

**Output fields:**

```powershell
MACAddress      = [string]
ARPType         = [string]   # 'static' | 'dynamic'
PingedToPopulate = [bool]
```

***

### 10.5 `Get-VBTCPFingerprint`

**Layer:** 5\
\
**MinPSVersion:** 5.1\
\
**Prerequisite:** `$Context.NetworkProbeEnabled`\
\
**Cost:** TCP connection attempts (300 ms default)

**Fingerprint port list (24 ports):**

```powershell
$FingerprintPorts = @(
    22, 23, 80, 135, 161, 443, 445, 515, 548, 554,
    631, 902, 1883, 3389, 5060, 5061, 5985, 8000,
    8080, 8443, 9100, 9443, 37777, 62078
)
```

**Logic:**

1. Check `NetworkProbeEnabled`.
2. **Within a single IP:** scan all 24 ports concurrently using async `BeginConnect` + `WaitOne`. Total wall time ≈ 300 ms regardless of port count.
3. Collect open ports.
4. Always close TCP clients in `finally`.

**Output fields:**

```powershell
OpenPorts      = [string]   # comma-separated
OpenPortsList  = [int[]]
ScanDurationMs = [int]
```

***

### 10.6 `Get-VBHTTPBanner`

**Layer:** 6\
\
**MinPSVersion:** 5.1 (with cert workaround)\
\
**Prerequisite:** `$Context.NetworkProbeEnabled` AND prior layer 5 found 80/443/8080/8443 open. Orchestrator enforces the gate.\
\
**Cost:** HTTP GET per port (3 s default)

**Logic:**

1. PS 6+: use `-SkipCertificateCheck` natively.
2. PS 5.1: `[Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }` BEFORE call, reset to `$null` in `finally` (failure to reset breaks all subsequent HTTPS in the session — critical).
3. Try ports in order: 80 → 8080 → 443 → 8443. First success wins.
4. Extract `<title>` via regex `'<title[^>]*>(.*?)</title>'` and `Server:` header.

**Output fields:**

```powershell
HTTPTitle   = [string]
HTTPServer  = [string]
HTTPPort    = [int]
HTTPScheme  = [string]   # 'http' | 'https'
```

***

### 10.7 `Get-VBSNMPIdentity`

**Layer:** 7\
\
**MinPSVersion:** 5.1\
\
**Prerequisite:** `$Context.SNMPAvailable`\
\
**Cost:** UDP 161, single packet per community string

**OIDs:**

| OID                 | Field                    |
| ------------------- | ------------------------ |
| `1.3.6.1.2.1.1.1.0` | `SNMPDescr`              |
| `1.3.6.1.2.1.1.5.0` | `Hostname` (sysName)     |
| `1.3.6.1.2.1.1.6.0` | `Location` (sysLocation) |

**Logic:**

1. Check `SNMPAvailable`.
2. Iterate `$Context.SNMPCommunityStrings` — first that responds wins. Record `CommunityUsed`.
3. `New-Object -ComObject olePrn.OleSNMP` → `.Open()` → three `.Get()` calls → `.Close()` in `finally`.
4. **Future:** when SNMPv3 is wired up (PSGallery available per v3 Q11), swap the COM block for the `SNMPv3` module — no other layer or classification logic changes.

**Output fields:**

```powershell
Hostname       = [string]
SNMPDescr      = [string]
Location       = [string]
CommunityUsed  = [string]
SNMPVersion    = [string]   # 'v1' | 'v2c' | 'v3' (future)
```

***

## 10.8 `Get-VBOUIVendor`

**Layer:** 11 *(enrichment-only — always runs if MAC available)*\
**MinPSVersion:** 5.1\
**Prerequisite:** `$Context.OUIFileAvailable` AND `$MACAddress` not null\
**Cost:** Flat-file lookup, instant

***

### OUI Data Source

- **URL:** `https://standards-oui.ieee.org/oui/oui.csv`
- **Format:** CSV — `Organization Address` column dropped on import, keep `Registry`, `Assignment`, `Organization Name` only
- **Refresh:** On load, check file creation time — if **> 30 days**, download fresh copy and delete the old one
- **Storage:** Local `oui.csv` alongside the script/module

```powershell
Invoke-WebRequest `
    -Uri 'https://standards-oui.ieee.org/oui/oui.csv' `
    -OutFile '.\oui.csv'
```

***

### Optimization

On **first call within a run**, import `oui.csv` into a hashtable keyed by 6-char OUI (`Assignment` column). Cache in `$Script:OUITable` for the session lifetime.

> **[!] WARNING:** Never query the CSV per-call — import once, lookup from hashtable every time.

***

### Logic

1. If `$MACAddress` is null → `Status = 'Skipped'`, `SkipReason = 'NoMACAvailable'`, return early
2. Normalise MAC — strip `:`, `.`, `-`, uppercase, take first 6 chars → `$MACNormalised`
3. Hashtable lookup: `$Script:OUITable[$MACNormalised]`
4. Map `Organization Name` → `VendorDeviceClass` via lookup table:

| Vendor match *(contains)* | `VendorDeviceClass` |
| ------------------------- | ------------------- |
| Yealink                   | `IPPhone`           |
| Poly / Polycom            | `IPPhone`           |
| Cisco                     | `NetworkDevice`     |
| HP / Hewlett              | `Workstation`       |
| Dell                      | `Workstation`       |
| Apple                     | `Workstation`       |
| *(no match)*              | `Unknown`           |

> **[i] INFO:** Lookup table is an ordered hashtable in the function header — extend as needed.

***

### Output Fields

| Field               | Type     | Notes                                |
| ------------------- | -------- | ------------------------------------ |
| `Vendor`            | `string` | Raw `Organization Name` from CSV     |
| `VendorDeviceClass` | `string` | Mapped class from lookup table       |
| `MACNormalised`     | `string` | 6-char uppercase OUI used for lookup |

***

### 10.9 `Get-VBRTSPBanner`

**Layer:** 8\
\
**MinPSVersion:** 5.1\
\
**Prerequisite:** `$Context.NetworkProbeEnabled` AND `$Context.RTSPProbeEnabled` AND port 554 open from layer 5\
\
**Cost:** TCP connection to 554, send `OPTIONS`, read banner

**Logic:**

1. Check prerequisites.
2. TCP connect to 554. If closed: `Status = 'NoResult'`.
3. Send `OPTIONS rtsp://${IPAddress}:554/ RTSP/1.0\r\nCSeq: 1\r\n\r\n`.
4. Read up to 1 KB or until 500 ms idle.
5. Always close stream and TCP client in `finally`.

**Output fields:**

```powershell
RTSPBanner = [string]   # first 512 chars
```

***

### 10.10 `Get-VBmDNSRecord`

**Layer:** 9\
\
**MinPSVersion:** 5.1\
\
**Prerequisite:** `$Context.mDNSAvailable`\
\
**Cost:** Passive listen via `dns-sd.exe`

**Logic:**

1. If `dns-sd.exe` not on PATH: return `Status = 'Skipped'`, `SkipReason = 'BonjourNotInstalled'`, `Impact = 'Printers, scanners, Apple devices using mDNS only may not be identified'`.
2. Run `dns-sd.exe -B <service>` for 5 s per service type, capture output.
3. Service types: `_pdl-datastream._tcp`, `_ipp._tcp`, `_scanner._tcp`, `_http._tcp`, `_afpovertcp._tcp`.
4. Match results to `$IPAddress`.

**Note:** mDNS is VLAN-bound. The context report explicitly states this. If results are always empty in a multi-VLAN environment, that's expected, not a bug.

**Output fields:**

```powershell
MDNSServiceType = [string]
MDNSServiceName = [string]
```

***

### 10.11 `Get-VBSwitchARP`

**Layer:** 10\
\
**MinPSVersion:** 5.1\
\
**Prerequisite:** `$Context.SNMPAvailable` AND `$Context.SwitchTargets.Count -gt 0`\
\
**Cost:** SNMP queries to managed switches

**OIDs:**

| OID                       | Returns                      |
| ------------------------- | ---------------------------- |
| `1.3.6.1.2.1.4.22.1.2`    | ARP: IP → MAC                |
| `1.3.6.1.2.1.17.4.3.1.2`  | MAC → port number            |
| `1.3.6.1.2.1.31.1.1.1.18` | Port description (`ifAlias`) |

**Logic:**

1. Check prerequisites.
2. **One-shot cache:** on first call, walk ARP table on every switch in `$Context.SwitchTargets` and merge into a single hashtable keyed by IP.
3. Look up `$IPAddress`. If found, walk MAC table and ifAlias for that switch.
4. Return `MACAddress`, `SwitchIP`, `SwitchPort`, `PortDescription` (used as `Location`).

**Output fields:**

```powershell
MACAddress      = [string]
SwitchIP        = [string]
SwitchPort      = [string]
PortDescription = [string]
```

***

## 11. Classification — `Resolve-VBDeviceClass`

**File:** `Public/Resolve-VBDeviceClass.ps1`\
\
**MinPSVersion:** 5.1\
\
**Input:** A signal object containing all merged layer outputs (NOT raw layer result objects).\
\
**Output:** `[PSCustomObject]` with `DeviceClass` and `Confidence`.

This function **never** calls a probe function. It is pure logic — given signals, return class.

### Input Schema

```powershell
param(
    [string]$OSClass,             # from Get-VBADComputer
    [string]$OpenPorts,           # from Get-VBTCPFingerprint
    [string]$HTTPTitle,           # from Get-VBHTTPBanner
    [string]$HTTPServer,
    [string]$SNMPDescr,           # from Get-VBSNMPIdentity
    [string]$RTSPBanner,          # from Get-VBRTSPBanner
    [string]$OUIVendor,           # from Get-VBOUIVendor
    [string]$VendorDeviceClass,
    [string]$MDNSServiceType
)
```

### Logic — Tiered `switch ($true)`

```powershell
$combined = "$HTTPTitle $HTTPServer $SNMPDescr $RTSPBanner $OUIVendor".ToLower()
$ports    = if ($OpenPorts) {
    $OpenPorts -split ',' | ForEach-Object { [int]$_.Trim() }
} else { @() }

$class = switch ($true) {

    # Tier 1 — AD is authoritative
    ($OSClass -eq 'DomainController') { 'DomainController'; break }
    ($OSClass -eq 'Server')           { 'Server';           break }
    ($OSClass -eq 'Workstation')      { 'Workstation';      break }

    # Tier 2 — Cameras (RTSP exclusive)
    ($ports -contains 554 -or
     $RTSPBanner -match 'RTSP' -or
     $combined -match 'hikvision|dahua|axis|hanwha|bosch.*security|milestone')
                                      { 'Camera';     break }

    # Tier 3 — VoIP
    ($ports -contains 5060 -or $ports -contains 5061 -or
     $combined -match 'yealink|polycom|grandstream|snom|sip.*phone|voip')
                                      { 'IPPhone';    break }

    # Tier 4 — Printers (port 9100/631 exclusive)
    ($ports -contains 9100 -or $ports -contains 631 -or
     $combined -match 'jetdirect|kyocera|ricoh|xerox|laserjet|bizhub|taskalfa|workcentre|brother|lexmark')
                                      { 'Printer';    break }

    # Tier 5 — Scanners
    ($MDNSServiceType -match '_scanner' -or $combined -match 'scanner')
                                      { 'Scanner';    break }

    # Tier 6 — VMware ESXi
    ($ports -contains 902 -or $ports -contains 9443 -or
     $combined -match 'esxi|vmware|vsphere')
                                      { 'VirtualHost'; break }

    # Tier 7 — NAS
    ($combined -match 'synology|qnap|nas|diskstation|truenas|freenas')
                                      { 'NAS';        break }

    # Tier 8 — UPS / PDU
    ($combined -match 'apc|eaton|ups|pdu|powerware|tripplite')
                                      { 'UPS';        break }

    # Tier 9 — Network infrastructure
    ($combined -match 'cisco|ubiquiti|aruba|juniper|fortinet|mikrotik|unifi|switch|router|firewall|access.point')
                                      { 'NetworkDevice'; break }

    # Tier 10 — iOS
    ($ports -contains 62078)          { 'Mobile';     break }

    # Tier 11 — IoT (MQTT)
    ($ports -contains 1883)           { 'IoT';        break }

    # Tier 12 — Windows fallback (no AD)
    ($ports -contains 3389 -and $ports -contains 135)  { 'Workstation'; break }
    ($ports -contains 5985 -and $ports -contains 445)  { 'Server';      break }

    # Tier 13 — OUI-only fallback (lowest signal)
    ($VendorDeviceClass -and $VendorDeviceClass -ne 'Unknown')
                                      { $VendorDeviceClass; break }

    default                           { 'Unknown' }
}

$confidence = switch ($class) {
    'DomainController' { 'High' }
    'Server'           { 'High' }
    'Workstation'      { 'High' }
    'Camera'           { if ($ports -contains 554) { 'High' } else { 'Medium' } }
    'IPPhone'          { 'High' }
    'Printer'          { 'High' }
    'Unknown'          { 'None' }
    default            { 'Medium' }
}

return [PSCustomObject]@{
    DeviceClass       = $class
    Confidence        = $confidence
    DeviceClassSource = ($SignalsThatMatched -join ',')   # e.g. 'OSClass,OUI'
}
```

***

## 12. Orchestration — `Invoke-VBIPEnrichment`

**File:** `Public/Invoke-VBIPEnrichment.ps1`\
\
**MinPSVersion:** 5.1\
\
**Purpose:** Run all layers in priority order against a list of IPs. Persist results to SQLite. Return one enrichment object per IP.

### Parameters

```powershell
param(
    [Parameter(Mandatory, ValueFromPipeline)]
    [string[]]$IPAddress,

    [Parameter(Mandatory)]
    [PSCustomObject]$Context,

    [switch]$SkipActiveProbes,            # skip layers 5-10 (passive only)
    [switch]$ForceRefresh,                # ignore SQLite cache, re-probe everything
    [int]$StaleThresholdHours = 168,      # 7 days — re-probe rows older than this
    [switch]$PassThru,                    # emit results immediately as they complete
    [int]$ProgressUpdateInterval = 1
)
```

### Execution Flow

```typescript
1.  Validate context (warn if missing).
2.  Validate IP list — strip publics and warn.
3.  Load existing rows from SQLite for these IPs.
4.  Decide which IPs to probe:
      - $ForceRefresh           → all IPs
      - row missing             → probe
      - row.UpdatedAt > stale   → probe
      - row.IsResolved = false  → probe
      - else                    → reuse cached row, skip probe
5.  Build one-shot caches (AD, DHCP, OUI, ARP) — once for the run.
6.  For each IP to probe (parallel on PS 7 for active layers):
      Step  1: Get-VBADComputer        → may set Hostname, OSClass
      Step  2: Get-VBDHCPLease         → may set Hostname, MAC
      Step  3: Get-VBPTRRecord         → may set Hostname (if not already)
      Step  4: Get-VBARPEntry          → may set MAC (if not already)
      [if $SkipActiveProbes — skip steps 5-10]
      Step  5: Get-VBTCPFingerprint    → always runs (enrichment)
      Step  6: Get-VBHTTPBanner        → only if 80/443/8080/8443 open
      Step  7: Get-VBSNMPIdentity      → only if 161 open OR SNMP probing forced
      Step  8: Get-VBRTSPBanner        → only if 554 open
      Step  9: Get-VBmDNSRecord        → only if Bonjour available
      Step 10: Get-VBSwitchARP         → only if SwitchTargets configured
      Step 11: Get-VBOUIVendor         → always runs if MAC known
      Step 12: Resolve-VBDeviceClass   → once, on collected signals
7.  Detect change vs existing row:
      - Hostname changed?  → write EnrichmentHistory entry, ChangeReason='HostnameChanged'
      - MAC changed?       → ChangeReason='MACChanged' (DHCP churn)
      - Class changed?     → ChangeReason='ClassChanged'
      - new row?           → ChangeReason='NewIP'
      - else               → ChangeReason='StaleRefresh', no history row
8.  Upsert into SQLite Enrichment table.
9.  If $PassThru — emit object immediately. Else collect.
10. Emit summary to Verbose:
    "Processed N. From cache: X. Newly resolved: Y. Still unknown: Z. Duration: T."
```

### Skip-If-Resolved Gate Implementation

```powershell
# Pseudo-code in orchestrator
$state = @{ IPAddress = $ip; IsResolved = $false; ... }

# Step 1 — AD
$adResult = Get-VBADComputer -IPAddress $ip -Context $ctx
if ($adResult.Status -eq 'Success') {
    $state.Hostname    = $adResult.Hostname
    $state.OSClass     = $adResult.OSClass
    $state.IsResolved  = $true
    $state.HostnameSource = 'AD'
}

# Step 2 — DHCP — runs even if resolved (to get MAC)
$dhcpResult = Get-VBDHCPLease -IPAddress $ip -Context $ctx
if ($dhcpResult.Status -eq 'Success') {
    $state.MACAddress = $dhcpResult.MACAddress
    if (-not $state.IsResolved) {
        $state.Hostname = $dhcpResult.Hostname
        $state.IsResolved = $true
        $state.HostnameSource = 'DHCP'
    }
}

# Step 3 — PTR — only if not resolved
if (-not $state.IsResolved) {
    $ptrResult = Get-VBPTRRecord -IPAddress $ip -Context $ctx
    if ($ptrResult.Status -eq 'Success' -and $ptrResult.ForwardConfirmed) {
        $state.Hostname = $ptrResult.Hostname
        $state.IsResolved = $true
        $state.HostnameSource = 'PTR'
    }
}

# ... and so on. Enrichment-only layers (TCP, OUI, ARP) always run.
```

### Parallel Mode

When `$Context.CanUseParallel`:

- Layers 1–4 (passive) run **sequentially** for each IP.
- Layers 5–9 (active probes per IP) run **in parallel across IPs** using `ForEach-Object -Parallel -ThrottleLimit $Context.ParallelThrottleLimit`.
- Caches (AD/DHCP/OUI/ARP hashtables) must be passed via `$using:` scope.
- Classification runs sequentially on collected results.

***

## 13. Progress & Transparency Model

### Live Progress (always on, unless caller sets `-ProgressAction Ignore`)

```typescript
Activity : DNSEnrichment — 23 of 47 IPs processed
Status   : 192.168.1.105 — Step 6: HTTP Banner Grab (port 80)
Percent  : 48%
SecondsRemaining : ~85
```

### Verbose Trail (when `-Verbose`)

```typescript
VERBOSE: [Context] PS 7.4.1 | DC role | AD/DHCP local | mDNS missing
VERBOSE: [192.168.1.45] Step 1 AD     → Success    Workstation | OU=Laptops
VERBOSE: [192.168.1.45] Step 2 DHCP   → Success    LAPTOP-VB01 (lease 2026-05-09 08:00)
VERBOSE: [192.168.1.45] Step 3 PTR    → Skipped    Already resolved by AD
VERBOSE: [192.168.1.45] Step 4 ARP    → Success    MAC: 00:1A:2B:3C:4D:5E
VERBOSE: [192.168.1.45] Step 5 TCP    → Success    Open: 135,445,3389
VERBOSE: [192.168.1.45] Step 6 HTTP   → NoResult   No web UI
VERBOSE: [192.168.1.45] Step 7 SNMP   → NoResult   No agent (port 161 closed)
VERBOSE: [192.168.1.45] Step 8 RTSP   → Skipped    Port 554 closed
VERBOSE: [192.168.1.45] Step 9 mDNS   → Skipped    Bonjour not installed
VERBOSE: [192.168.1.45] Step 10 Switch → Skipped   No SwitchTargets configured
VERBOSE: [192.168.1.45] Step 11 OUI   → Success    Dell Inc. → Workstation
VERBOSE: [192.168.1.45] Classify     → Workstation (High) via OSClass
VERBOSE: [192.168.1.45] DB upsert    → Updated (1 history row: HostnameChanged DESKTOP-OLD→LAPTOP-VB01)
```

### Per-IP Result Transparency — `LayerTrace`

Every result includes a `LayerTrace` array — exactly 11 entries, one per step:

```powershell
LayerTrace = @(
    [PSCustomObject]@{ Step=1;  Name='AD';     Status='Success';  DurationMs=8;   Detail='Workstation | OU=Laptops' }
    [PSCustomObject]@{ Step=2;  Name='DHCP';   Status='Success';  DurationMs=12;  Detail='LAPTOP-VB01' }
    [PSCustomObject]@{ Step=3;  Name='PTR';    Status='Skipped';  DurationMs=0;   Detail='Already resolved by AD' }
    [PSCustomObject]@{ Step=4;  Name='ARP';    Status='Success';  DurationMs=2;   Detail='00:1A:2B:3C:4D:5E (dynamic)' }
    [PSCustomObject]@{ Step=5;  Name='TCP';    Status='Success';  DurationMs=312; Detail='135,445,3389' }
    [PSCustomObject]@{ Step=6;  Name='HTTP';   Status='NoResult'; DurationMs=14;  Detail='No port open' }
    [PSCustomObject]@{ Step=7;  Name='SNMP';   Status='NoResult'; DurationMs=2010;Detail='No agent on 161' }
    [PSCustomObject]@{ Step=8;  Name='RTSP';   Status='Skipped';  DurationMs=0;   Detail='Port 554 closed' }
    [PSCustomObject]@{ Step=9;  Name='mDNS';   Status='Skipped';  DurationMs=0;   Detail='Bonjour not installed' }
    [PSCustomObject]@{ Step=10; Name='Switch'; Status='Skipped';  DurationMs=0;   Detail='No SwitchTargets' }
    [PSCustomObject]@{ Step=11; Name='OUI';    Status='Success';  DurationMs=1;   Detail='Dell Inc.' }
)
StepsAttempted = 7   # ran (excludes Skipped)
StepsSucceeded = 5
StepsNoResult  = 2
StepsSkipped   = 4
StepsFailed    = 0
```

### Stalled vs Active Detection

`Write-Progress` is updated at minimum every `$ProgressUpdateInterval` seconds. If the orchestrator is mid-probe on a slow IP, the progress bar still ticks and shows the current step. If output is truly stalled, the bar stops ticking — that's the user's signal to investigate.

***

## 14. Output Object Contracts

### Layer Result (returned by every `Get-VB*` function)

```powershell
[PSCustomObject]@{
    # Always populated
    IPAddress    = [string]
    Layer        = [int]
    LayerName    = [string]
    Status       = [string]   # Success | NoResult | Failed | Skipped
    ExecutionMs  = [int]

    # Populated when Status = Skipped
    SkipReason   = [string]
    Impact       = [string]

    # Populated when Status = Failed
    ErrorDetail  = [string]

    # Layer-specific fields above
}
```

### Final Enrichment Object (returned by `Invoke-VBIPEnrichment`)

```powershell
[PSCustomObject]@{
    # Identity
    IPAddress             = [string]
    Hostname              = [string]
    HostnameSource        = [string]   # AD | DHCP | PTR | SNMP | mDNS | Switch
    MACAddress            = [string]
    MACAddressNormalised  = [string]
    Vendor                = [string]
    DeviceClass           = [string]
    DeviceClassSource     = [string]
    Confidence            = [string]   # High | Medium | Low | None
    OSClass               = [string]
    OperatingSystem       = [string]
    Model                 = [string]
    Location              = [string]
    OU                    = [string]

    # Probe results
    OpenPorts             = [string]
    HTTPTitle             = [string]
    HTTPServer            = [string]
    SNMPDescr             = [string]
    RTSPBanner            = [string]
    MDNSServiceType       = [string]

    # Pipeline metadata
    StepsAttempted        = [int]
    StepsSucceeded        = [int]
    StepsNoResult         = [int]
    StepsSkipped          = [int]
    StepsFailed           = [int]
    LayerTrace            = [PSCustomObject[]]
    IsResolved            = [bool]
    IsUnresolved          = [bool]    # DeviceClass == 'Unknown'

    # Cache / change tracking
    EnrichedAt            = [datetime]
    EnrichmentDurationMs  = [int]
    FirstSeenAt           = [datetime]
    UpdatedAt             = [datetime]
    ChangeReason          = [string]  # NewIP | HostnameChanged | MACChanged | ClassChanged | StaleRefresh | NoChange
    FromCache             = [bool]    # True if returned from SQLite without re-probing
}
```

***

## 15. Error Handling Contract

These rules apply to every function without exception.

1. Every function wraps core logic in `try { } catch { }`.
2. `catch` blocks **never** use `continue` or `return $null` — always set `Status = 'Failed'` and `ErrorDetail = $_.Exception.Message`.
3. `Write-Error` is **never** used inside layer functions — errors are captured into the result object and surfaced by the orchestrator.
4. **All timeouts explicit.** Never rely on OS defaults.
5. **All COM objects** (olePrn) closed and released in `finally`.
6. **All TCP clients** closed in `finally`.
7. **PS 5.1 cert callback** always reset to `$null` in `finally` — leaving it set affects all subsequent HTTPS in the session.
8. If called without `$Context`: emit `Write-Warning "No context provided — running without prerequisite validation. Results may be incomplete."` and continue.
9. Failures in one layer **never** abort the orchestrator. Each layer is independently catchable.

***

## 16. Resolution Order — Most Reliable First

This ordering is intentional: most reliable signals run first, so most IPs are resolved before any active probing.

| Step | Layer                | Cost              | Resolves Hostname            | Sets `IsResolved` Gate   |
| ---- | -------------------- | ----------------- | ---------------------------- | ------------------------ |
| 1    | **Active Directory** | Zero (cached)     | Yes (auth-backed)            | ✓                        |
| 2    | **DHCP Leases**      | Zero (cached)     | Yes (auth-backed)            | ✓                        |
| 3    | **PTR Lookup**       | Tiny (DNS)        | Yes (forward-confirmed)      | ✓                        |
| 4    | **ARP cache**        | Zero              | No (provides MAC)            | —                        |
| 5    | **TCP Fingerprint**  | 300 ms            | No (provides ports)          | —                        |
| 6    | **HTTP Banner**      | 3 s               | Sometimes (model only)       | ✓ if title is conclusive |
| 7    | **SNMP**             | 2 s               | Yes (sysName)                | ✓                        |
| 8    | **RTSP**             | 2 s               | No (banner only)             | —                        |
| 9    | **mDNS**             | 5 s               | Yes (service name)           | ✓                        |
| 10   | **Switch ARP**       | SNMP × N switches | No (provides MAC + location) | —                        |
| 11   | **OUI**              | Instant           | No (provides Vendor)         | —                        |

### Why AD before PTR

- AD is auth-backed and authoritative.
- PTR records are notoriously stale — a PTR pointing at a hostname decommissioned 6 months ago is common.
- Running AD first means classification gets `OSClass` (which PTR cannot provide) and prevents being misled by stale reverse zones.

### Why TCP before HTTP/RTSP/SNMP gating

- TCP fingerprint (5) tells us which ports are open in 300 ms.
- HTTP (6), SNMP (7), RTSP (8) only run if their port is in the open set. Saves seconds per IP.

***

## 17. Environment Scenarios — Capability Matrix

Real environments vary. The module degrades gracefully — never crashes — based on detected capabilities.

| Scenario                                | AD | DHCP | PTR     | TCP | HTTP | SNMP | OUI | mDNS | Switch | Notes                                                          |
| --------------------------------------- | -- | ---- | ------- | --- | ---- | ---- | --- | ---- | ------ | -------------------------------------------------------------- |
| **DC + Windows DHCP** (default)         | ✓  | ✓    | ✓       | ✓   | ✓    | ✓    | ✓   | opt  | opt    | All layers available. Best case.                               |
| **Domain member + RSAT**                | ✓  | ✓    | ✓       | ✓   | ✓    | ✓    | ✓   | opt  | opt    | Same as DC, just slightly slower (remote AD/DHCP).             |
| **Domain member, no RSAT**              | ✗  | ✗    | ✓       | ✓   | ✓    | ✓    | ✓   | opt  | opt    | Layers 1-2 skipped. Recommend installing RSAT.                 |
| **Workgroup machine**                   | ✗  | ✗    | partial | ✓   | ✓    | ✓    | ✓   | opt  | opt    | Heavy reliance on TCP/SNMP/HTTP. R07 will show many `Unknown`. |
| **Linux DHCP / ISC / Mikrotik / Unifi** | ✓  | ✗    | ✓       | ✓   | ✓    | ✓    | ✓   | opt  | opt    | Layer 2 skipped — extension point: SSH/API plugin (future).    |
| **BIND DNS, not Windows**               | ✓  | ✓    | ✓       | ✓   | ✓    | ✓    | ✓   | opt  | opt    | `Resolve-DnsName` works against any DNS.                       |
| **No SNMP enabled anywhere**            | ✓  | ✓    | ✓       | ✓   | ✓    | ✗    | ✓   | opt  | ✗      | Layers 7+10 skipped. HTTP carries more weight.                 |
| **No network probe allowed**            | ✓  | ✓    | ✓       | ✗   | ✗    | ✗    | ✓   | ✗    | ✗      | Pass `-DisableNetworkProbe`. Passive-only enrichment.          |

`opt` **= optional, depends on Bonjour install / SwitchTargets configuration.**

The context report (Section 8) clearly shows which scenario applies and what's been disabled.

***

## 18. Caching, Re-Runs & DHCP Churn Handling

### Within a Single Run

- AD, DHCP, OUI, ARP, switch ARP — all loaded into hashtables **once** at the start of the orchestrator and passed via `$using:` to parallel runspaces.
- HTTP responses are not cached (different requests).

### Across Runs (SQLite)

On each invocation of `Invoke-VBIPEnrichment`:

1. Load existing rows for the supplied IPs from `Enrichment` table.
2. For each IP, decide:
   - **Probe** if: row missing, `UpdatedAt > $StaleThresholdHours` ago, `IsResolved = false`, OR `-ForceRefresh` set.
   - **Reuse** otherwise — return the cached row with `FromCache = $true` and skip all layers.
3. After probing, compare new result to existing row:
   - If **Hostname changed** → upsert + insert `EnrichmentHistory` row with `ChangeReason = 'HostnameChanged'`.
   - If **MAC changed** → upsert + history with `ChangeReason = 'MACChanged'` (this is the DHCP churn signal — IP got re-assigned to a different device).
   - If **DeviceClass changed** → upsert + history with `ChangeReason = 'ClassChanged'`.
   - If **nothing material changed** → update only `UpdatedAt`, no history row, `ChangeReason = 'NoChange'`.

### DHCP Churn Detection

A single IP being assigned to multiple devices over time is normal in DHCP environments. The `EnrichmentHistory` table captures every transition. Reports can:

- Show "DHCP churn" — IPs that flipped between hostnames in the last 24 h.
- Map a specific MAC across all IPs it's held over time (track-the-device queries).
- Identify "stuck" cached rows where the MAC hasn't changed in 6 months — likely a static IP.

### Refresh Strategy

Locked from v3 design Q2: **on-demand**. The script is run manually. No scheduler component in v1. To force a full refresh, pass `-ForceRefresh`.

***

## 19. Testing Checklist

Each function must pass these tests before the module ships. Pester files live in `Tests/`.

### Per-Function Tests

- [ ] Returns `[PSCustomObject]` — never `$null`, never a string.
- [ ] Returns `Status = 'Skipped'` when prerequisite is absent (mock context flag = `$false`).
- [ ] Returns `Status = 'NoResult'` when probe succeeds but finds nothing.
- [ ] Returns `Status = 'Failed'` with non-empty `ErrorDetail` when probe throws.
- [ ] `ExecutionMs` is always populated and non-negative.
- [ ] Accepts pipeline input.
- [ ] Works on PS 5.1 (test on a Windows Server 2016 box minimum).
- [ ] Works on PS 7 (test on Server 2019/2022).
- [ ] Closes all TCP/COM resources even on error path (verify with handle counter).
- [ ] PS 5.1 cert callback is reset after HTTPS use.

### Orchestrator Tests

- [ ] Processes a list of 3 IPs and returns 3 result objects.
- [ ] `StepsAttempted + StepsSkipped == 11` for every result.
- [ ] `LayerTrace` always contains exactly 11 entries.
- [ ] `-SkipActiveProbes` results in steps 5–10 all `Skipped`.
- [ ] `-PassThru` emits results before all IPs complete.
- [ ] Parallel mode (PS 7) produces identical results to sequential mode (PS 5.1).
- [ ] `-ForceRefresh` re-probes all IPs even if cached.
- [ ] Cached row with `UpdatedAt < $StaleThresholdHours` is returned without re-probing.
- [ ] Hostname change between runs creates an `EnrichmentHistory` row.

### Context Tests

- [ ] Detects DC role correctly on a DC.
- [ ] Detects member-server correctly on a non-DC.
- [ ] Reports `DNSAvailable = $false` when DNS unreachable.
- [ ] `PrerequisiteReport` always has one entry per check (never missing rows).

### Storage Tests

- [ ] `Initialize-VBEnrichmentDatabase` is idempotent.
- [ ] Schema migrations applied in order.
- [ ] CSV export round-trips through SQLite without data loss.
- [ ] `EnrichmentHistory` populated correctly on hostname/MAC/class change.

***

## 20. Build Sequence

Build in this order. Each phase produces usable output.

| Phase  | Deliverable                                                           | Dependency              |
| ------ | --------------------------------------------------------------------- | ----------------------- |
| **1**  | `Get-VBEnrichmentContext` + private helpers                           | Nothing                 |
| **2**  | `Initialize-VBEnrichmentDatabase` + storage layer                     | PSSQLite                |
| **3**  | Layers 1-3 + 4 (AD, DHCP, PTR, ARP) — passive only                    | RSAT, DNS access        |
| **4**  | `Resolve-VBDeviceClass` + `Invoke-VBIPEnrichment` (passive-only mode) | Phases 1-3              |
| **5**  | `Get-VBEnrichmentResult` + `Export-VBEnrichmentResult`                | Phase 4                 |
| **6**  | Ship a usable v1 — passive enrichment only — to production            | Phase 5                 |
| **7**  | Layer 5 (TCP) + Layer 6 (HTTP)                                        | Network access          |
| **8**  | Layer 7 (SNMP) + Layer 11 (OUI)                                       | SNMP community, oui.txt |
| **9**  | Layer 8 (RTSP) + parallel mode on PS 7                                | PS 7 environment        |
| **10** | Layer 9 (mDNS) + Layer 10 (Switch ARP)                                | Bonjour, switch SNMP    |
| **11** | DHCP churn report + history table queries                             | Phase 6                 |
| **12** | Pester test coverage to 80%+                                          | All phases              |

> **[i] INFO:** Phase 6 is a deliberate ship gate — the module is useful with passive enrichment only, and that ships before any active probing complexity is introduced. Active probes are the long tail.

***

## 21. Decisions Locked from v3 Design Doc

These are the answers from `dns-log-analysis-design-v3.md` Section 13, baked into this design:

| Q  | Decision                                                                              | Where Applied                                    |
| -- | ------------------------------------------------------------------------------------- | ------------------------------------------------ |
| 1  | Storage = **SQLite**                                                                  | §9                                               |
| 2  | Refresh = **on-demand**                                                               | §18 — no scheduler component                     |
| 3  | SNMP community = **default** `public`**, parameterised, multiple supported**          | §10.7, §7                                        |
| 4  | Reverse zones maintained = **yes**                                                    | §10.1 — PTR is layer 3 with forward-confirmation |
| 5  | DHCP servers = **single (multi assumed in sync)**                                     | §7, §10.2                                        |
| 6  | Switch port labels = **environment-dependent**                                        | §10.11 — gracefully handles missing labels       |
| 7  | After-hours definition = **N/A for this module** (consumer's concern, not enrichment) | —                                                |
| 8  | NXDOMAIN threshold = **N/A for this module**                                          | —                                                |
| 9  | Output format = **CSV + PSCustomObject + SQLite**                                     | §9 — all three supported                         |
| 10 | Parallelism = **yes on PS 7**                                                         | §4, §12                                          |
| 11 | SNMPv3 timeline = **PSGallery available** → upgrade path documented                   | §10.7 — drop-in module swap                      |

***

## 22. Changelog

| Version | Date       | Change                                                                                                                                                                                                                                                                                                                                                                  |
| ------- | ---------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 2.0     | 2026-05-08 | **Final / production-ready.** Merged v3 design decisions; added storage layer (SQLite + history); added ARP step as explicit layer 4; reordered priority (AD before PTR); added forward-confirmation on PTR; added DHCP churn detection; locked PS version strategy; added capability matrix for environment scenarios; added build sequence with ship gate at Phase 6. |
| 1.0     | 2026-05-08 | Initial module design spec.                                                                                                                                                                                                                                                                                                                                             |

***

*DNSEnrichment Module Design v2.0 (Final) — 2026-05-08 | VB + Claude*
