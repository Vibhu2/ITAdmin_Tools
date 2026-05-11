# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Module Identity

`VB.DNSEnrichment` v0.4.0 — a PowerShell module that resolves private/RFC1918 IP addresses to hostname, device class, MAC, vendor, and enrichment metadata via an 11-layer pipeline, persisting results in SQLite.

**Required dependency:** `PSSQLite` (PSGallery). Install with:
```powershell
Install-Module PSSQLite -Scope CurrentUser
```

## Running the Test Script

The test script is a standalone production validation runner, not a Pester suite. It requires at least one private IP to probe:

```powershell
# Minimal run (single IP)
.\Test-VBDNSEnrichmentModule.ps1 -IPAddress '192.168.1.1'

# Full run (multiple IPs, all options)
.\Test-VBDNSEnrichmentModule.ps1 -IPAddress '10.0.0.1','10.0.0.5' `
    -DHCPServer 'dhcp.corp.local' `
    -SNMPCommunityStrings 'public','internal' `
    -SwitchTargets '10.0.0.254' `
    -ForceRefresh

# Passive-only (no active network probes — safe for restricted networks)
.\Test-VBDNSEnrichmentModule.ps1 -IPAddress '10.0.0.1' -SkipActiveProbes
```

The test sections map 1:1 to module capabilities: Context (§1), DB init (§2), per-layer functions (§3), classifier (§4), orchestrator (§5), SQLite query (§6), export (§7), edge cases (§8-10).

## Importing the Module Manually

```powershell
Import-Module .\VB.DNSEnrichment.psd1 -Force

# Always call Get-VBEnrichmentContext first — it builds the $ctx object every other function needs
$ctx = Get-VBEnrichmentContext -SNMPCommunityStrings 'public' -Quiet

# Run the full pipeline
$results = Invoke-VBIPEnrichment -IPAddress '10.0.0.1','10.0.0.2' -Context $ctx

# Query SQLite
Get-VBEnrichmentResult -Context $ctx
Get-VBEnrichmentResult -UnresolvedOnly -Context $ctx
Get-VBEnrichmentResult -IPAddress '10.0.0.1' -IncludeHistory -Context $ctx

# Export
$results | Export-VBEnrichmentResult -Format CSV -Path .\out.csv
$results | Export-VBEnrichmentResult -Format JSON -Path .\out.json -IncludeLayerTrace
```

## Architecture

### Context Object — the Module's Foundation

`Get-VBEnrichmentContext` **must** be called first. It probes the environment, sets capability flags, and returns a `$ctx` object passed to every other function. No layer function ever re-detects capabilities — they read flags from `$ctx`:

| Flag | Controls |
|---|---|
| `ADAvailable` | Layer 1 (AD query) |
| `DHCPAvailable` | Layer 2 (DHCP lease cache) |
| `DNSAvailable` | Layer 3 (PTR lookup) |
| `NetworkProbeEnabled` | Layers 5–10 (active probes) |
| `SNMPAvailable` | Layers 7 + 10 (SNMP + Switch ARP) |
| `mDNSAvailable` | Layer 9 (dns-sd.exe browse) |
| `CanUseParallel` | PS7 `ForEach-Object -Parallel` for active layers |

### 11-Layer Pipeline

Layers run inside `Invoke-VBIPEnrichment` in this fixed order:

```
PASSIVE (sequential, no target connection):
  1. Get-VBADComputer      — one-shot AD cache; highest confidence; sets OSClass
  2. Get-VBDHCPLease       — one-shot DHCP cache; provides MAC + lease info
  3. Get-VBPTRRecord       — PTR lookup with forward-confirmation
  4. Get-VBARPEntry        — arp -a cache; provides MAC for OUI

ACTIVE (parallel across IPs on PS7):
  5. Get-VBTCPFingerprint  — 24-port async scan (~300 ms); always runs (enrichment)
  6. Get-VBHTTPBanner      — gated on ports 80/443/8080/8443 open from Layer 5
  7. Get-VBSNMPIdentity    — olePrn COM; sysDescr/sysName/sysLocation
  8. Get-VBRTSPBanner      — gated on port 554 open from Layer 5
  9. Get-VBmDNSRecord      — dns-sd.exe; requires Bonjour installed
 10. Get-VBSwitchARP       — SNMP walk on managed switches; requires SwitchTargets in $ctx
 11. Get-VBOUIVendor       — IEEE OUI CSV; always runs if MAC available (enrichment-only)

CLASSIFICATION + STORAGE (always sequential):
  Resolve-VBDeviceClass   — tiered switch logic on collected signals → DeviceClass + Confidence
  SQLite upsert           — write/update Enrichment table; write EnrichmentHistory on changes
```

**Skip-if-resolved gate:** The orchestrator tracks `IsResolved`. Hostname-resolving layers (AD, DHCP, PTR, SNMP, mDNS, Switch) are skipped for IPs already resolved. Enrichment-only layers (TCP, ARP, OUI) always run.

### SQLite Storage

Two tables:
- `Enrichment` — one row per IP (primary key `IPAddress`); upserted each run
- `EnrichmentHistory` — append-only audit log; a row is inserted when Hostname, MAC, or DeviceClass changes between runs (DHCP churn detection)

Schema migrations live in `Sql/001_init.sql` and are applied by `Initialize-VBEnrichmentDatabase`.

### Private Helpers

| Helper | Purpose |
|---|---|
| `New-VBLayerResult` | Factory for all layer return objects; enforces Status ∈ Success/NoResult/Failed/Skipped |
| `Invoke-VBSqliteCommand` | Thin PSSQLite wrapper used by all storage functions |
| `ConvertTo-VBNormalisedMAC` | Strips separators, uppercases, returns 12-char or OUI-prefix MAC |
| `Test-VBPrivateIP` | Validates RFC1918 + CGNAT (100.64/10) + link-local (169.254/16) |
| `Write-VBEnrichmentProgress` | Internal Write-Progress wrapper |

### Caching Strategy

Within a run: AD, DHCP, ARP, OUI, and Switch ARP are loaded into script-scoped hashtables on first call and reused for every IP (never queried per-IP).

Across runs: `Invoke-VBIPEnrichment` loads existing SQLite rows first. An IP is re-probed only if its row is missing, older than `$StaleThresholdHours` (default 168 h), unresolved, or `-ForceRefresh` is set.

## Coding Standards (Non-Negotiable)

- **No aliases** — `foreach` not `%`, `Where-Object` not `?`
- **No `Get-WmiObject`** — use `Get-CimInstance` with `-OperationTimeoutSec` always set
- **Named parameters only** — no positional arguments in function calls
- **No `Write-Host`** — use `Write-Verbose` for operational messages
- **No silent catches** — every `catch` sets `Status = 'Failed'` and captures `$_.Exception.Message` into `ErrorDetail`
- **All functions return `[PSCustomObject]`** — never `$null`, never raw strings
- **Every function has `[CmdletBinding()]` and explicit parameter types**
- **TCP clients and COM objects always closed in `finally` blocks** — never only on the success path
- **PS 5.1 cert callback** (`[Net.ServicePointManager]::ServerCertificateValidationCallback`) must be reset to `$null` in `finally` after any HTTPS call — leaving it set corrupts subsequent HTTPS in the session

## PS Version Compatibility

The module baseline is **PS 5.1**. PS7 features are used only where meaningful and always behind a context flag:

| Feature | PS 5.1 behaviour |
|---|---|
| `ForEach-Object -Parallel` | Sequential `foreach` (gated on `$ctx.CanUseParallel`) |
| `Invoke-WebRequest -SkipCertificateCheck` | `[Net.ServicePointManager]` callback workaround |
| `Resolve-DnsName` | Fallback to `[System.Net.Dns]::GetHostEntry()` |

Never put `#Requires -Version 7` at module level — it crashes PS 5.1 hosts.

## Classification Logic

`Resolve-VBDeviceClass` uses a 13-tier `switch ($true)` — order matters:

1. AD OSClass (DomainController / Server / Workstation) — authoritative
2. RTSP signals → Camera
3. SIP/VoIP signals → IPPhone
4. JetDirect/print ports/vendor match → Printer
5. mDNS scanner service → Scanner
6. VMware ports/strings → VirtualHost
7. NAS vendor strings → NAS
8. UPS/PDU strings → UPS
9. Network vendor strings → NetworkDevice
10. Port 62078 → Mobile (iOS lockdown)
11. Port 1883 → IoT (MQTT)
12. RDP+135 (no AD) → Workstation; WinRM+445 → Server
13. OUI vendor hint → class from vendor map
14. Default → Unknown

The classifier is pure signal-in / class-out — it never calls any probe function.

## Design Document

`Documentation/dnsenrichment-module-design-v2.md` is the authoritative spec. Consult it for the full context object schema, per-layer output contracts, error handling rules, environment capability matrix, and build sequence.
