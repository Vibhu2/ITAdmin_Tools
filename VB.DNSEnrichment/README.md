# VB.DNSEnrichment

Resolve private/RFC1918 IP addresses to hostname, device class, MAC, vendor, and enrichment metadata through an 11-layer pipeline. Results persist in a local SQLite database for cross-run caching and DHCP churn detection.

Part of the [VBTools suite](https://github.com/Vibhu2/ITAdmin_Tools) — PowerShell modules for Windows sysadmins.

---

## Requirements

- PowerShell 5.1 or later (PS 7 enables parallel active probes)
- [PSSQLite](https://www.powershellgallery.com/packages/PSSQLite) module

```powershell
Install-Module PSSQLite -Scope CurrentUser
```

---

## Installation

```powershell
Install-Module VB.DNSEnrichment -Scope CurrentUser
```

---

## Quick Start

```powershell
# 1. Import and build context (always first — probes environment capabilities)
Import-Module VB.DNSEnrichment
$ctx = Get-VBEnrichmentContext -SNMPCommunityStrings 'public'

# 2. Initialise the SQLite database (idempotent — safe to run on every session)
Initialize-VBEnrichmentDatabase -DatabasePath $ctx.DatabasePath

# 3. Enrich one or more private IPs
$results = Invoke-VBIPEnrichment -IPAddress '192.168.1.1', '192.168.1.10' -Context $ctx

# 4. View results
$results | Select-Object IPAddress, Hostname, DeviceClass, Vendor, MACAddress, Confidence
```

---

## The Enrichment Pipeline

`Invoke-VBIPEnrichment` runs all enabled layers in order for each IP, merges the signals, classifies the device, and upserts the result into SQLite.

| Step | Layer | Type | What it provides |
|------|-------|------|-----------------|
| 1 | AD Computer | Passive | Hostname, OSClass, OU |
| 2 | DHCP Lease | Passive | Hostname, MAC, lease expiry |
| 3 | PTR Record | Passive | Hostname (with forward-confirmation) |
| 4 | ARP Cache | Passive | MAC address |
| 5 | TCP Fingerprint | Active | Open ports (24-port async scan) |
| 6 | HTTP Banner | Active | Page title, Server header |
| 7 | SNMP Identity | Active | sysDescr, sysName, sysLocation |
| 8 | RTSP Banner | Active | Camera/NVR banner |
| 9 | mDNS Record | Active | Bonjour service type and name |
| 10 | Switch ARP | Active | MAC from managed switch SNMP walk |
| 11 | OUI Vendor | Enrichment | IEEE vendor name and device class hint |

Active layers 5–10 run in parallel across IPs on PS 7 when `$ctx.CanUseParallel` is `$true`.

---

## Context Object

`Get-VBEnrichmentContext` must be called before any other function. It probes the environment, sets capability flags, and returns a `$ctx` object that every function reads from.

```powershell
$ctx = Get-VBEnrichmentContext `
    -DatabasePath         'C:\Data\enrichment.db' `
    -DHCPServer           'dhcp.corp.local' `
    -SNMPCommunityStrings 'public', 'internal' `
    -SwitchTargets        '10.0.0.254', '10.0.0.253' `
    -Quiet
```

| Parameter | Default | Purpose |
|-----------|---------|---------|
| `DatabasePath` | `%LOCALAPPDATA%\VB.DNSEnrichment\enrichment.db` | SQLite database location |
| `DHCPServer` | local machine | DHCP server to query for lease cache |
| `SNMPCommunityStrings` | `@('public')` | Community strings tried in order |
| `SwitchTargets` | `@()` | Managed switch IPs for Layer 10 |
| `Quiet` | `$false` | Suppress capability report to console |

---

## Querying Results

```powershell
# All enriched IPs
Get-VBEnrichmentResult -Context $ctx

# Specific IPs
Get-VBEnrichmentResult -IPAddress '192.168.1.1', '192.168.1.10' -Context $ctx

# Unresolved only (DeviceClass = Unknown)
Get-VBEnrichmentResult -UnresolvedOnly -Context $ctx

# Filter by device class
Get-VBEnrichmentResult -DeviceClass 'Camera' -Context $ctx

# Include change history (DHCP churn log)
Get-VBEnrichmentResult -IPAddress '192.168.1.1' -IncludeHistory -Context $ctx
```

---

## Exporting Results

```powershell
# CSV (UTF-8 with BOM — opens correctly in Excel)
$results | Export-VBEnrichmentResult -Format CSV -Path .\enrichment.csv

# JSON with full per-step layer trace
$results | Export-VBEnrichmentResult -Format JSON -Path .\enrichment.json -IncludeLayerTrace

# Pass objects through the pipeline unchanged
$results | Export-VBEnrichmentResult -Format Object | Where-Object { $_.Confidence -eq 'High' }
```

---

## Orchestrator Parameters

```powershell
Invoke-VBIPEnrichment
    -IPAddress           <string[]>   # One or more private IPs; accepts pipeline input
    -Context             <PSCustomObject>  # From Get-VBEnrichmentContext
    -SkipActiveProbes                 # Passive layers only (safe for restricted networks)
    -ForceRefresh                     # Re-probe even if cached row is fresh
    -StaleThresholdHours <int>        # Hours before a cached row is re-probed (default 168)
    -PassThru                         # Stream each result immediately instead of collecting
```

---

## Examples

### Passive-only run (no network probing)

```powershell
$ctx = Get-VBEnrichmentContext -Quiet
Initialize-VBEnrichmentDatabase -DatabasePath $ctx.DatabasePath
Invoke-VBIPEnrichment -IPAddress '10.0.0.1','10.0.0.5' -Context $ctx -SkipActiveProbes
```

### Full run from a CSV

```powershell
$ctx = Get-VBEnrichmentContext `
    -DHCPServer 'dhcp.corp.local' `
    -SNMPCommunityStrings 'public','internal' `
    -SwitchTargets '10.0.0.254'

Initialize-VBEnrichmentDatabase -DatabasePath $ctx.DatabasePath

Import-Csv .\ips.csv | Select-Object -ExpandProperty IPAddress |
    Invoke-VBIPEnrichment -Context $ctx -PassThru |
    Export-VBEnrichmentResult -Format CSV -Path .\results.csv
```

### Force-refresh stale entries

```powershell
$ctx = Get-VBEnrichmentContext -Quiet
Get-VBEnrichmentResult -UnresolvedOnly -Context $ctx |
    Select-Object -ExpandProperty IPAddress |
    Invoke-VBIPEnrichment -Context $ctx -ForceRefresh
```

---

## Device Classes

The classifier (`Resolve-VBDeviceClass`) maps collected signals to one of these classes:

`DomainController` · `Server` · `Workstation` · `NetworkDevice` · `Printer` · `Camera` · `IPPhone` · `NAS` · `UPS` · `VirtualHost` · `Scanner` · `Mobile` · `IoT` · `Unknown`

Confidence is reported as `High`, `Medium`, or `Low`.

---

## SQLite Schema

Two tables:

- **`Enrichment`** — one row per IP, upserted on each run. Contains all fields returned by `Get-VBEnrichmentResult`.
- **`EnrichmentHistory`** — append-only audit log. A row is inserted whenever `Hostname`, `MACAddress`, or `DeviceClass` changes between runs (useful for tracking DHCP churn).

---

## Individual Layer Functions

Each layer function can be called independently if you need a specific signal without running the full pipeline:

```powershell
Get-VBADComputer      -IPAddress '10.0.0.5'  -Context $ctx
Get-VBDHCPLease       -IPAddress '10.0.0.5'  -Context $ctx
Get-VBPTRRecord       -IPAddress '10.0.0.5'  -Context $ctx
Get-VBARPEntry        -IPAddress '10.0.0.5'  -Context $ctx
Get-VBTCPFingerprint  -IPAddress '10.0.0.5'  -Context $ctx
Get-VBHTTPBanner      -IPAddress '10.0.0.5'  -OpenPortsList @(80,443) -Context $ctx
Get-VBSNMPIdentity    -IPAddress '10.0.0.5'  -Context $ctx
Get-VBRTSPBanner      -IPAddress '10.0.0.5'  -Context $ctx
Get-VBmDNSRecord      -IPAddress '10.0.0.5'  -Context $ctx
Get-VBSwitchARP       -IPAddress '10.0.0.5'  -Context $ctx
Get-VBOUIVendor       -MACAddress 'AA:BB:CC:DD:EE:FF' -IPAddress '10.0.0.5' -Context $ctx
```

All layer functions return a `[PSCustomObject]` with `Status` ∈ `Success / NoResult / Failed / Skipped`.

---

## License

[MIT](https://github.com/Vibhu2/ITAdmin_Tools/blob/main/LICENSE) — © 2026 Vibhu Bhatnagar
