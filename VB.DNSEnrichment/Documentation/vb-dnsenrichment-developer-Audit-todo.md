---
type: Plan
title: "VB.DNSEnrichment — Developer Fix TODO (v0.4.0 → v0.5.0)"
date: 2026-05-12
tags: [powershell, dns-enrichment, bug-fix, todo, developer-handoff]
status: draft
author: VB + Claude
---

# VB.DNSEnrichment — Developer Fix TODO

**Module version under repair:** 0.4.0
**Target version after all fixes:** 0.5.0
**Audit date:** 2026-05-12
**Current resolution rate:** 28 / 50 IPs (56%). Target: ≥ 90%.
**Repo path:** `VB.DNSEnrichment\.claude\worktrees\naughty-wescoff-22c854\VB.DNSEnrichment\`

Read every section completely before touching any code. The issues interact —
fixing them out of order can introduce regressions. Follow the priority order
exactly: Critical → High → Medium → Low.

---

## How to Read This Document

Each item contains:

- **File + line** — exact location of the defect
- **Root cause** — why it is wrong
- **Impact** — what breaks because of it
- **Fix** — exactly what to change (no rewrites permitted beyond what is described)

---

## CRITICAL — Fix First (Resolution Rate Impact)

These four issues directly cause the 22 missed IPs. Fix all four before
running any validation test.

---

### C-01 · TCP Timeout 300 ms → 1000 ms

**Files:**
- [`Public\Get-VBTCPFingerprint.ps1` line 70](Public\Get-VBTCPFingerprint.ps1)
- [`Public\Get-VBEnrichmentContext.ps1` line 510–515](Public\Get-VBEnrichmentContext.ps1)

**Root cause:**
300 ms is insufficient on any routed network. A device whose SYN/SYN-ACK
handshake takes 301 ms appears as "all ports closed." All gated layers
(HTTP, RTSP) are then skipped. 300 ms was chosen for speed; it trades
accuracy for time on real enterprise LANs.

**Impact:**
Cameras (embedded Linux, slow TCP stack), printers (JetDirect startup),
VoIP phones, and any inter-VLAN traffic all drop out. Conservative estimate:
10–15 false negatives from this alone.

**Fix — `Get-VBTCPFingerprint.ps1` line 70:**
```powershell
# BEFORE
[int]$TimeoutMs = 300

# AFTER
[int]$TimeoutMs = 1000
```

**Fix — `Get-VBEnrichmentContext.ps1` lines 510–515:**
```powershell
# BEFORE
DefaultTimeoutMs = @{
    TCP   = 300
    HTTP  = 3000
    SNMP  = 2000
    RTSP  = 2000
}

# AFTER
DefaultTimeoutMs = @{
    TCP   = 1000
    HTTP  = 5000
    SNMP  = 3000
    RTSP  = 3000
}
```

> Note: HTTP, SNMP, and RTSP timeouts are also raised here. Devices that
> previously timed out by a small margin will now succeed. The concurrent
> TCP scan design means wall time stays ≈ TimeoutMs (not multiplied by
> port count), so the cost is ~700 ms extra per IP — acceptable.

---

### C-02 · SQL Injection in SQLite Cache-Load Query

**File:** [`Public\Invoke-VBIPEnrichment.ps1` lines 158–162](Public\Invoke-VBIPEnrichment.ps1)

**Root cause:**
IP strings are interpolated directly into the SQL `IN` clause using string
formatting, not parameterised queries. Every other SQL call in this module
uses `@parameter` syntax correctly. This one does not.

```powershell
# CURRENT (vulnerable)
$placeholders = ($validIPs | ForEach-Object { "'$_'" }) -join ','
$rows = Invoke-VBSqliteCommand -DatabasePath $dbPath `
    -Query "SELECT * FROM Enrichment WHERE IPAddress IN ($placeholders)"
```

**Impact:**
Security defect and teaching anti-pattern in a student training module.
In production, a crafted IP-like string that passes `Test-VBPrivateIP` could
manipulate the query. The inconsistency also teaches students that string
interpolation into SQL is acceptable when it is not.

**Fix:**
Replace with positional parameters. SQLite supports `?1, ?2, ...` positional
placeholders, but `PSSQLite`'s `Invoke-SqliteQuery` uses named `@param`
style. Build named parameters dynamically:

```powershell
# AFTER — parameterised
$paramHash = @{}
$placeholders = for ($i = 0; $i -lt $validIPs.Count; $i++) {
    $paramHash["ip$i"] = $validIPs[$i]
    "@ip$i"
}
$inClause = $placeholders -join ','
$rows = Invoke-VBSqliteCommand -DatabasePath $dbPath `
    -Query "SELECT * FROM Enrichment WHERE IPAddress IN ($inClause)" `
    -SqlParameters $paramHash
```

---

### C-03 · mDNS Guard Condition Is Logically Inverted

**File:** [`Public\Get-VBmDNSRecord.ps1` line 86](Public\Get-VBmDNSRecord.ps1)

**Root cause:**
The condition that guards whether to run `Invoke-VBmDNSBrowse` is
double-negated and evaluates to the wrong polarity:

```powershell
# CURRENT (inverted — runs when mDNS is UNAVAILABLE, skips when AVAILABLE)
if (-not ($Context -and -not $Context.mDNSAvailable)) {
    $Script:VBmDNSCache = Invoke-VBmDNSBrowse -ServiceTypes $ServiceTypes
```

Truth-table:
| `$Context` | `mDNSAvailable` | Condition evaluates to | What happens |
|---|---|---|---|
| `$null` | — | `$true` | Browse runs (wrong — no context, unknown state) |
| exists | `$false` | `$false` | Browse skipped (correct — mDNS unavailable) |
| exists | `$true` | `$true` | Browse runs (correct) |

When `$Context` is `$null` (called without context), the browse runs
unconditionally even though availability is unknown. When context exists
and `mDNSAvailable = $true`, it works correctly — so the bug surfaces
only in the no-context call path, but that path matters for unit testing
and standalone use.

**Impact:**
Printers and scanners that advertise only via mDNS depend entirely on this
layer. Any inconsistency in the guard leaves them unresolvable.

**Fix — replace the condition with a simple positive check:**
```powershell
# AFTER
if ($Context -and $Context.mDNSAvailable) {
    $Script:VBmDNSCache = Invoke-VBmDNSBrowse -ServiceTypes $ServiceTypes
    $Script:VBmDNSCacheBuilt = $true
}
```

Also set `$Script:VBmDNSCacheBuilt = $true` only inside this block, not
unconditionally. The current code never sets it to `$true` at all — the
variable is declared and checked but the success path never assigns it.
Verify all three variables (`$Script:VBmDNSCache`, `$Script:VBmDNSCacheBuilt`)
are set correctly on both the run and skip paths.

---

### C-04 · Parallel Runspaces Rebuild All One-Shot Caches (AD, DHCP, ARP, OUI)

**File:** [`Public\Invoke-VBIPEnrichment.ps1` lines 370–465](Public\Invoke-VBIPEnrichment.ps1)

**Root cause:**
The `ForEach-Object -Parallel` block at line 373 calls `Import-Module`
inside each runspace, which gives every runspace its own fresh script
scope. `$Script:VBAdComputerCache`, `$Script:VBDhcpLeaseCache`,
`$Script:VBArpCache`, `$Script:VBOUITable` are all `$null` in each
runspace — so each runspace rebuilds them independently.

With `ParallelThrottleLimit = 10` and 50 IPs:
- `Get-ADComputer -Filter *` fires 10 times simultaneously → DC overload
- `arp.exe -a` fires 10 times simultaneously
- `oui.csv` is parsed 10 times from disk simultaneously
- The one-shot cache design provides zero benefit in parallel mode

**Impact:**
Performance regression under PS7 parallel mode. Risk of overwhelming the
domain controller. OUI lookup failures if the CSV file is held open by
multiple readers simultaneously (file lock contention on Windows).

**Fix:**
Pre-build all caches in the sequential passive pass (which already runs
before the parallel block). Capture the cache data into local variables
and pass via `$using:` into the parallel block.

```powershell
# BEFORE the parallel block, after the passive pass loop ends:

# Snapshot all script-scope caches for the parallel block
$adCacheSnapshot    = $Script:VBAdComputerCache    # already built in passive pass
$dhcpCacheSnapshot  = $Script:VBDhcpLeaseCache     # already built in passive pass
$arpCacheSnapshot   = $Script:VBArpCache           # already built in passive pass
$ouiTableSnapshot   = $Script:VBOUITable           # loaded on first OUI call

# Inside the parallel block — replace any $Script:VB* access with the snapshot:
# $using:adCacheSnapshot, $using:dhcpCacheSnapshot, etc.
```

The layer functions inside the parallel block (TCP, HTTP, SNMP, RTSP,
mDNS, Switch, OUI) do not use the AD/DHCP/ARP caches — only the
orchestrator uses them. So the parallel block steps 5–11 do not
actually access `$Script:VBAdComputerCache` at all. The real fix here
is simply:

1. Do NOT call `Import-Module` inside the parallel block if the module
   is already loaded (it will be). Check first:
   ```powershell
   if (-not (Get-Module -Name 'VB.DNSEnrichment')) {
       Import-Module $using:modPath -ErrorAction Stop
   }
   ```
2. Add a `$null` guard before the parallel block so it does not proceed
   if `$modulePath` could not be resolved:
   ```powershell
   $modulePath = (Get-Module -Name 'VB.DNSEnrichment').Path
   if (-not $modulePath) {
       throw "[Orchestrator] Cannot resolve VB.DNSEnrichment module path for parallel block."
   }
   ```
3. The OUI table (`$Script:VBOUITable`) IS used inside the parallel
   block via `Get-VBOUIVendor`. Pre-build it before entering the block
   by calling `Get-VBOUIVendor -MACAddress '000000000000' -Context $Context`
   as a warm-up call (this loads `$Script:VBOUITable` in the main scope),
   then pass via `$using:` and assign to `$Script:VBOUITable` at the
   start of the parallel block before any OUI call.

---

## HIGH — Fix Before Any Production Use

---

### H-01 · PTR Forward-Confirm Uses String Comparison — Rejects Valid Hostnames

**File:** [`Public\Get-VBPTRRecord.ps1` lines 106–118](Public\Get-VBPTRRecord.ps1)

**Root cause:**
`$resolvedIPs -contains $IPAddress` compares raw strings. If the DNS
server returns the IP in any non-identical string form (trailing space,
zero-padded octet, different case for IPv6), the match fails.

More importantly, the orchestrator at line 306 **discards the hostname
entirely** when `ForwardConfirmed = $false`:
```powershell
if ($ptrResult.Status -eq 'Success' -and $ptrResult.ForwardConfirmed) {
    $state.IsResolved = $true
}
# if ForwardConfirmed = $false → hostname is thrown away, IsResolved stays false
```

A valid PTR hostname that cannot be forward-confirmed should be kept
with `Confidence = 'Low'`, not discarded. The forward-confirm is a
quality signal, not a binary gate.

**Impact:**
Any device whose PTR → A record chain has split-brain DNS, a stale A
record, or an IPv4/IPv6 dual-stack forward record miss is treated as
completely unresolvable from PTR. Estimated 3–6 false negatives.

**Fix 1 — `Get-VBPTRRecord.ps1` line 117, normalise both sides:**
```powershell
# BEFORE
$forwardConfirmed = $resolvedIPs -contains $IPAddress

# AFTER
$parsedTarget = [System.Net.IPAddress]::Parse($IPAddress)
$forwardConfirmed = $resolvedIPs | Where-Object {
    try { [System.Net.IPAddress]::Parse($_).Equals($parsedTarget) }
    catch { $false }
} | Select-Object -First 1 | ForEach-Object { $true }
$forwardConfirmed = [bool]$forwardConfirmed
```

**Fix 2 — `Invoke-VBIPEnrichment.ps1` line 306, accept low-confidence PTR:**
```powershell
# BEFORE
if ($ptrResult.Status -eq 'Success' -and $ptrResult.ForwardConfirmed) {
    $state.Hostname       = $ptrResult.Hostname
    $state.HostnameSource = 'PTR'
    $state.IsResolved     = $true
}

# AFTER — accept PTR hostname regardless of forward-confirm; confidence tracks quality
if ($ptrResult.Status -eq 'Success' -and -not [string]::IsNullOrWhiteSpace($ptrResult.Hostname)) {
    $state.Hostname       = $ptrResult.Hostname
    $state.HostnameSource = if ($ptrResult.ForwardConfirmed) { 'PTR' } else { 'PTR-Unconfirmed' }
    $state.IsResolved     = $true
}
```

---

### H-02 · `Vendor` Column Missing from SQLite Upsert — Vendor Lost on Cache Hits

**Files:**
- [`Public\Invoke-VBIPEnrichment.ps1` lines 678–718](Public\Invoke-VBIPEnrichment.ps1) (upsert SQL)
- [`Sql\001_init.sql` line 12](Sql\001_init.sql) (schema — `Vendor TEXT` already exists)
- [`Public\Invoke-VBIPEnrichment.ps1` line 869](Public\Invoke-VBIPEnrichment.ps1) (`Invoke-VBBuildEnrichmentObject` reads `$Row.Vendor`)

**Root cause:**
The `Enrichment` table schema (defined in `001_init.sql`) includes a
`Vendor TEXT` column. The upsert SQL in the orchestrator does NOT write
this column. `Invoke-VBBuildEnrichmentObject` reads `$Row.Vendor` when
reconstructing from cache — always gets `$null`. OUI vendor data is
permanently lost after the first probe run.

**Fix — add `Vendor` to the INSERT column list and ON CONFLICT SET in
the upsert SQL block (lines 678–718):**

In the `INSERT INTO Enrichment (...)` column list, add `Vendor` after
`MACAddressNormalised`:
```sql
... MACAddress, MACAddressNormalised, Vendor, DeviceClass ...
```

In the `VALUES (...)` list, add `@vendor` in the matching position:
```sql
... @mac, @macNorm, @vendor, @deviceClass ...
```

In the `ON CONFLICT ... DO UPDATE SET` block, add:
```sql
Vendor = excluded.Vendor,
```

In the `$SqlParameters` hashtable passed to `Invoke-VBSqliteCommand`,
add:
```powershell
vendor = $state.OUIVendor
```

No schema migration needed — the column already exists in the table.

---

### H-03 · DHCP Scope Failures Are Silently Swallowed

**File:** [`Public\Get-VBDHCPLease.ps1` line 95](Public\Get-VBDHCPLease.ps1)

**Root cause:**
```powershell
$leases = Get-DhcpServerv4Lease -ScopeId $scopeId @serverSplat `
    -ErrorAction SilentlyContinue
```
`-ErrorAction SilentlyContinue` discards errors per scope. If 3 of 5
scopes fail (access denied, RPC timeout), the cache is built with 2/5
of the leases. `$Script:VBDhcpCacheBuilt = $true` is set anyway — the
caller never knows data is missing.

**Impact:**
IPs in failed scopes always return `NoResult` from DHCP with no
indication in the trace. Hostname and MAC are lost for entire subnets
with no operator warning.

**Fix:**
Switch to `-ErrorAction Stop` inside a per-scope `try/catch`, count
failures, and emit a `Write-Warning` that is recorded:

```powershell
$failedScopes = [System.Collections.Generic.List[string]]::new()
foreach ($scopeId in $scopeIds) {
    try {
        $leases = Get-DhcpServerv4Lease -ScopeId $scopeId @serverSplat -ErrorAction Stop
        foreach ($lease in $leases) {
            if (-not [string]::IsNullOrWhiteSpace($lease.IPAddress)) {
                $Script:VBDhcpLeaseCache[$lease.IPAddress.ToString()] = $lease
            }
        }
    }
    catch {
        $failedScopes.Add($scopeId)
        Write-Warning "[$LAYER_NAME] Scope $scopeId failed to enumerate: $($_.Exception.Message)"
    }
}
if ($failedScopes.Count -gt 0) {
    Write-Warning "[$LAYER_NAME] $($failedScopes.Count) of $($scopeIds.Count) scopes failed. IPs in these scopes will not be DHCP-resolved: $($failedScopes -join ', ')"
}
$Script:VBDhcpCacheBuilt = $true   # partial data is still usable
```

---

### H-04 · Switch ARP MAC Validation Discards Non-Colon Formats

**File:** [`Public\Get-VBSwitchARP.ps1` line 216](Public\Get-VBSwitchARP.ps1)

**Root cause:**
```powershell
if ($mac -and $mac -match '^[0-9A-Fa-f:]{17}$') {
```
This accepts only `AA:BB:CC:DD:EE:FF` (17 chars with colons). Older
Cisco, HP ProCurve, and Dell switches return:
- `AA-BB-CC-DD-EE-FF` (dashes)
- `AABBCCDDEEFF` (no separator)
- Decimal octets from SNMP raw encoding (rare but real)

MACs in any other format are silently dropped. The switch ARP cache
is partially empty even when the SNMP walk succeeds.

**Fix — replace the single-format regex with `ConvertTo-VBNormalisedMAC`
which already handles all formats:**
```powershell
# BEFORE
if ($mac -and $mac -match '^[0-9A-Fa-f:]{17}$') {
    $arpTable[$ip] = $mac
}

# AFTER
$macNorm = ConvertTo-VBNormalisedMAC -MACAddress $mac
if (-not [string]::IsNullOrWhiteSpace($macNorm)) {
    $arpTable[$ip] = $mac   # store original format; normalise at lookup time
}
```

---

### H-05 · `$ErrorActionPreference = 'Stop'` in Orchestrator `begin` Block

**File:** [`Public\Invoke-VBIPEnrichment.ps1` line 109](Public\Invoke-VBIPEnrichment.ps1)

**Root cause:**
Setting `$ErrorActionPreference = 'Stop'` in `begin {}` escalates all
non-terminating errors (including `Write-Warning`, benign pipeline
warnings, and minor network errors) to terminating errors in the
orchestrator's scope. Any unhandled non-terminating error from a child
function call aborts the entire orchestrator run.

The module's design relies on layer functions returning `Status = 'Failed'`
gracefully — but that recovery only works if the orchestrator's own scope
does not escalate their errors.

**Fix:**
Remove the global assignment. Apply `-ErrorAction Stop` only on the
specific cmdlet calls inside the `begin` block that genuinely must terminate
on failure (there are none — the context build and DB path resolution are
already guarded with `try/catch`):

```powershell
# REMOVE this line entirely from the begin block:
$ErrorActionPreference = 'Stop'
```

---

### H-06 · mDNS Browse Uses `Start-Sleep` Instead of `Wait-Job -Timeout`

**File:** [`Public\Get-VBmDNSRecord.ps1` lines 156–178](Public\Get-VBmDNSRecord.ps1) (`Invoke-VBmDNSBrowse`)

**Root cause:**
```powershell
$job = Start-Job -ScriptBlock { & dns-sd.exe -B $t local. 2>&1 } -ArgumentList $svcType
Start-Sleep -Seconds 5     # ← blocks the calling thread
Stop-Job -Job $job
```

`Start-Sleep` blocks the current thread for exactly 5 seconds regardless
of whether the job is ready. `Wait-Job -Timeout` would release immediately
when the job completes or after the timeout, whichever comes first. The
current code always waits the full 5 seconds even on an empty network
where `dns-sd.exe` returns immediately.

Additionally, the inner resolve loop does the same with `Start-Sleep -Seconds 3`
per discovered instance. On a network with 20 printers, the resolve phase
alone takes 60 seconds of pure sleep.

**Fix — replace `Start-Sleep` + `Stop-Job` with `Wait-Job -Timeout`:**
```powershell
# BEFORE
Start-Sleep -Seconds 5
Stop-Job -Job $job -ErrorAction SilentlyContinue
$output = Receive-Job -Job $job -ErrorAction SilentlyContinue

# AFTER
$null = Wait-Job -Job $job -Timeout 5
Stop-Job -Job $job -ErrorAction SilentlyContinue
$output = Receive-Job -Job $job -ErrorAction SilentlyContinue
```

Apply the same replacement for the inner resolve job (`-Timeout 3`).
The `Stop-Job` is still needed after `Wait-Job -Timeout` because
`Wait-Job -Timeout` does not kill the job — it only stops waiting.

---

## MEDIUM — Required for Enterprise Quality

---

### M-01 · `Data\` Folder Missing — OUI File Path Resolves to Non-Existent Location

**Files:**
- [`Public\Get-VBEnrichmentContext.ps1` lines 119–121](Public\Get-VBEnrichmentContext.ps1)
- [`Public\Get-VBOUIVendor.ps1` lines 71–76](Public\Get-VBOUIVendor.ps1)
- Module root folder (missing `Data\` directory)

**Root cause:**
`Get-VBEnrichmentContext` computes the OUI file default path as:
```powershell
$OUIFilePath = Join-Path $PSScriptRoot '..\Data\oui.csv'
```
`$PSScriptRoot` inside `Public\Get-VBEnrichmentContext.ps1` is the
`Public\` folder. `'..\Data\oui.csv'` therefore resolves to
`VB.DNSEnrichment\Data\oui.csv`.

**The `Data\` folder does not exist in the repository.** `Test-Path`
returns `$false`, the size check in `Get-VBEnrichmentContext` line 354
fails (`$ouiFileAvailable = $false`), and the context reports OUI as
`NotConfigured`. The download in `Invoke-VBLoadOUITable` will try to
create the parent and save there — but the path that `Get-VBEnrichmentContext`
resolves to and the path that `Get-VBOUIVendor` uses without context
(falling back to `$env:LOCALAPPDATA\VB.DNSEnrichment\oui.csv`) may
diverge depending on whether `$Context.OUIFilePath` is passed.

This is the primary reason OUI vendor lookup is broken — the file is
downloaded to one location but the availability check looks in another.

**Fix — three changes:**

**1. Create the `Data\` folder in the repository:**
```
VB.DNSEnrichment\
  Data\
    .gitkeep        ← empty placeholder so Git tracks the folder
```

**2. Fix the path resolution to use the module root, not `$PSScriptRoot`
(which varies depending on which file `$PSScriptRoot` is evaluated in).**

In `Get-VBEnrichmentContext.ps1` lines 119–121, replace:
```powershell
# BEFORE — $PSScriptRoot is Public\, so this is Public\..\Data\ = Data\
if (-not $OUIFilePath) {
    $OUIFilePath = Join-Path $PSScriptRoot '..\Data\oui.csv'
}

# AFTER — derive module root from the .psm1 path, which is always the module root
if (-not $OUIFilePath) {
    $moduleRoot  = Split-Path -Path (Get-Module -Name 'VB.DNSEnrichment').Path -Parent
    $OUIFilePath = Join-Path $moduleRoot 'Data\oui.csv'
}
```

**3. In `Invoke-VBLoadOUITable` (inside `Get-VBOUIVendor.ps1`), the
`$parent` folder creation already exists and is correct — it will create
`Data\` if missing. No change needed there.**

**4. Verify that `Get-VBEnrichmentContext` and `Get-VBOUIVendor` resolve
to the same path** by adding a `Write-Verbose "[OUI] Path: $ouiFilePath"`
immediately after path resolution in both functions. Confirm they match.

---

### M-02 · Context Object: `Write-Host` Must Be Replaced with `Write-Information`

**File:** [`Public\Get-VBEnrichmentContext.ps1`](Public\Get-VBEnrichmentContext.ps1)
All `Write-Host` calls inside `Format-VBEnrichmentContext` (lines 549–620).

**Root cause:**
`Write-Host` writes directly to the console. It cannot be captured,
redirected, or suppressed from a caller. The CLAUDE.md explicitly prohibits
it. In CI/CD, scheduled tasks, and remote PS sessions, `Write-Host` output
is either lost or pollutes `stdout` unexpectedly.

**Fix:**
Replace every `Write-Host` in `Format-VBEnrichmentContext` with
`Write-Information`. Add a consistent `-Tags 'VBContext'` tag so consumers
can filter:

```powershell
# BEFORE
Write-Host $line -ForegroundColor Cyan

# AFTER
Write-Information -MessageData $line -Tags 'VBContext' -InformationAction Continue
```

The `-InformationAction Continue` ensures the message appears at the
console by default (same behaviour as `Write-Host`) while remaining
redirectable. Remove all `-ForegroundColor` parameters — colour is a
host-specific concept and not supported by `Write-Information`.

If colour output is important to keep, wrap in:
```powershell
if ($Host.UI.SupportsVirtualTerminal) {
    # ANSI escape codes for colour instead of -ForegroundColor
}
```

---

### M-03 · Context Object: Prerequisite Report Must Be Actionable — Add `Severity` Field

**File:** [`Public\Get-VBEnrichmentContext.ps1`](Public\Get-VBEnrichmentContext.ps1)
All `& $addReport` calls throughout the function.

**Root cause:**
The `PrerequisiteReport` array has no `Severity` field. A consumer
cannot programmatically distinguish a cosmetic `NotConfigured` (e.g.
no Switch targets) from a critical `Unavailable` (e.g. DNS broken)
without inspecting `Impact` free text.

**Fix:**
Add `Severity` to the helper scriptblock parameter list and to the
`[PSCustomObject]` inside `$addReport`:

```powershell
# Update $addReport scriptblock signature:
$addReport = {
    param($Layer, $Prerequisite, $Status, $Detail, $LayersAffected,
          $SkippedLayers, $Impact, $Remediation, $Severity = 'Info')
    $report.Add([PSCustomObject]@{
        Layer          = $Layer
        Prerequisite   = $Prerequisite
        Status         = $Status
        Severity       = $Severity   # ← new
        Detail         = $Detail
        LayersAffected = $LayersAffected
        SkippedLayers  = $SkippedLayers
        Impact         = $Impact
        Remediation    = $Remediation
    })
}
```

Use `Severity` values: `'Critical'`, `'Warning'`, `'Info'`.
Assign `'Critical'` to DNS and AD failures. `'Warning'` to DHCP,
SNMP, OUI not configured. `'Info'` to optional layers (mDNS, Switch).

Update all `& $addReport` calls to include the Severity argument.
Update `Format-VBEnrichmentContext` to colour-code by `Severity`
instead of by `Status`.

---

### M-04 · Context Object: Database Initialisation Must Report Why It Failed

**File:** [`Public\Get-VBEnrichmentContext.ps1` lines 410–432](Public\Get-VBEnrichmentContext.ps1)

**Root cause:**
When `Initialize-VBEnrichmentDatabase` fails, the context reports
`'Unavailable'` with `$_.Exception.Message` as the detail. But the
most common failure mode — the `Sql\` migration folder not found
because the module is imported from a non-standard path — produces a
message like `"No migration files found in C:\..."` which is not
actionable without knowing the expected path.

**Fix:**
After catching the init failure, add the expected SQL folder path to
the report detail:

```powershell
catch {
    $sqlFolderHint = Join-Path (Split-Path (Get-Module 'VB.DNSEnrichment').Path -Parent) 'Sql'
    & $addReport 0 'SQLite Database' 'Unavailable' `
        "$($_.Exception.Message) | Expected SQL folder: $sqlFolderHint" `
        'Storage layer' 'All caching/persistence' `
        'Cannot persist enrichment results' `
        "Verify write access to $DatabasePath and that Sql\001_init.sql exists"
}
```

Also add a pre-check before calling `Initialize-VBEnrichmentDatabase`:
```powershell
$sqlFolder = Join-Path (Split-Path (Get-Module 'VB.DNSEnrichment').Path -Parent) 'Sql'
if (-not (Test-Path -LiteralPath $sqlFolder)) {
    & $addReport 0 'SQLite Database' 'Unavailable' `
        "Sql migration folder not found at: $sqlFolder" ...
    # set $databaseInitialized = $false and skip the Initialize call
}
```

---

### M-05 · Context Object: `CanUseParallel` Should Gate on Both PS Version AND Runspace Type

**File:** [`Public\Get-VBEnrichmentContext.ps1` line 148](Public\Get-VBEnrichmentContext.ps1)

**Root cause:**
```powershell
$canUseParallel = ($psMajor -ge 7)
```
`ForEach-Object -Parallel` is not available in PS7 when called from
within an existing `ForEach-Object -Parallel` block (nested parallel
is not supported) or from certain constrained runspaces (JEA endpoints,
some ISE hosts). The check should also verify the runspace is not
already parallel:

**Fix:**
```powershell
$canUseParallel = $false
if ($psMajor -ge 7) {
    try {
        # Test that -Parallel is actually usable in this runspace
        $null = [System.Management.Automation.Runspaces.Runspace]::DefaultRunspace
        $canUseParallel = $true
    }
    catch {
        Write-Verbose "[Context] ForEach-Object -Parallel not available in this runspace: $($_.Exception.Message)"
    }
}
```

---

### M-06 · `CollectionTime` in `New-VBLayerResult` Is Non-ISO8601 (Locale-Dependent)

**File:** [`Private\New-VBLayerResult.ps1` line 104](Private\New-VBLayerResult.ps1)

**Root cause:**
```powershell
CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
```
`dd-MM-yyyy` is ambiguous and locale-dependent. Every other timestamp
in the module uses `ToString('o')` (ISO 8601 with timezone offset).
This is the only exception — and it is wrong.

**Fix:**
```powershell
# BEFORE
CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')

# AFTER
CollectionTime = (Get-Date).ToString('o')
```

Same fix applies to `Initialize-VBEnrichmentDatabase.ps1` line 43
which also uses `'dd-MM-yyyy HH:mm:ss'`:
```powershell
# BEFORE
$collectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')

# AFTER
$collectionTime = (Get-Date).ToString('o')
```

---

### M-07 · OUI File Staleness Check Uses `CreationTime` Instead of `LastWriteTime`

**Files:**
- [`Public\Get-VBEnrichmentContext.ps1` line 357](Public\Get-VBEnrichmentContext.ps1)
- [`Public\Get-VBOUIVendor.ps1` line 207](Public\Get-VBOUIVendor.ps1)

**Root cause:**
```powershell
$ouiFileAge = (Get-Date) - $ouiItem.CreationTime
# and
elseif (((Get-Date) - $item.CreationTime).TotalDays -gt $REFRESH_DAYS)
```
`CreationTime` is the filesystem creation timestamp — reset when the
file is moved (as this code does via `Move-Item` from `.tmp`). After the
first download, `CreationTime` = now. Every subsequent run sees a 0-day-old
file and never refreshes, even after 31 days. The correct property is
`LastWriteTime`.

**Fix — both files, replace `CreationTime` with `LastWriteTime`:**
```powershell
# BEFORE
$ouiFileAge = (Get-Date) - $ouiItem.CreationTime
...
elseif (((Get-Date) - $item.CreationTime).TotalDays -gt $REFRESH_DAYS)

# AFTER
$ouiFileAge = (Get-Date) - $ouiItem.LastWriteTime
...
elseif (((Get-Date) - $item.LastWriteTime).TotalDays -gt $REFRESH_DAYS)
```

---

### M-08 · `New-VBLayerResult` Dead Code in ExtraFields Merge

**File:** [`Private\New-VBLayerResult.ps1` lines 107–114](Private\New-VBLayerResult.ps1)

**Root cause:**
```powershell
if (-not $base.Contains($key)) {
    $base[$key] = $ExtraFields[$key]   # branch A
}
else {
    $base[$key] = $ExtraFields[$key]   # branch B — identical to A
}
```
Both branches are identical. The `if` guard is inert. Either it was meant
to block overwriting base contract fields (but doesn't), or it is dead code.

**Fix:**
If the intent is to allow ExtraFields to add new fields but not
overwrite base contract fields (`IPAddress`, `Layer`, `LayerName`,
`Status`, `ExecutionMs`, `SkipReason`, `Impact`, `ErrorDetail`,
`CollectionTime`):

```powershell
$baseKeys = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@('IPAddress','Layer','LayerName','Status','ExecutionMs',
                'SkipReason','Impact','ErrorDetail','CollectionTime')
)
foreach ($key in $ExtraFields.Keys) {
    if (-not $baseKeys.Contains($key)) {
        $base[$key] = $ExtraFields[$key]
    }
    else {
        Write-Warning "[New-VBLayerResult] ExtraFields key '$key' conflicts with base contract field -- ignored."
    }
}
```

If the intent is simply to merge all ExtraFields (override allowed),
replace the `if/else` with a single line:
```powershell
foreach ($key in $ExtraFields.Keys) { $base[$key] = $ExtraFields[$key] }
```

Confirm the intended design with VB before choosing. The simpler merge
(override allowed) is what the code does today — so that is probably
the correct interpretation. Use the single-line form.

---

### M-09 · Script-Scope Caches Have No TTL — Stale Data Across Multi-Hour Sessions

**Files:** All layer files that use `$Script:VB*Cache` variables:
- `Get-VBADComputer.ps1` — `$Script:VBAdComputerCache`
- `Get-VBDHCPLease.ps1` — `$Script:VBDhcpLeaseCache`
- `Get-VBARPEntry.ps1` — `$Script:VBArpCache`
- `Get-VBmDNSRecord.ps1` — `$Script:VBmDNSCache`
- `Get-VBSwitchARP.ps1` — `$Script:VBSwitchARPCache`
- `Get-VBOUIVendor.ps1` — `$Script:VBOUITable` (30-day TTL via file age — this one is OK)

**Root cause:**
Caches are built once and never invalidated. In a scheduled task that
runs every hour, the AD and DHCP caches from the first run are reused
unchanged for all subsequent runs. New DHCP leases, ARP table changes,
and newly joined computers are invisible.

**Fix — add a companion `$Script:VB*CacheBuiltAt` timestamp for each
cache and check TTL on first access inside `begin {}` of each function:**

Pattern (apply to each layer's `begin` block):
```powershell
$CacheTTLMinutes = 60   # configurable via $Context if desired

if ($null -ne $Script:VBAdComputerCache) {
    $ageMin = ((Get-Date) - $Script:VBAdCacheBuiltAt).TotalMinutes
    if ($ageMin -gt $CacheTTLMinutes) {
        Write-Verbose "[$LAYER_NAME] AD cache is $([int]$ageMin) min old (TTL $CacheTTLMinutes min) -- rebuilding"
        $Script:VBAdComputerCache = $null
        $Script:VBAdCacheBuilt    = $false
    }
}

if ($null -eq $Script:VBAdComputerCache) {
    # ... existing build logic ...
    $Script:VBAdCacheBuiltAt = Get-Date
}
```

Add `$Script:VBAdCacheBuiltAt`, `$Script:VBDhcpCacheBuiltAt`,
`$Script:VBArpCacheBuiltAt`, `$Script:VBmDNSCacheBuiltAt`,
`$Script:VBSwitchARPCacheBuiltAt` as companion variables.
Set them immediately after each successful cache build.

---

### M-10 · Camera Classification at Tier 2 Is Too Broad — Port 554 Alone Classifies as Camera

**File:** [`Public\Resolve-VBDeviceClass.ps1` lines 147–154](Public\Resolve-VBDeviceClass.ps1)

**Root cause:**
```powershell
(
    $ports -contains 554 -or
    $RTSPBanner -match 'RTSP' -or
    $combined -match 'hikvision|dahua|...'
) { 'Camera' }
```
Port 554 is used by Cisco IOS RTSP relay, some NVR management servers,
and some VoIP systems. A device with port 554 open but no RTSP banner
and no vendor match is classified as `Camera` with `Confidence = 'High'`
(because the `High` branch triggers on `$ports -contains 554` alone).

**Fix — require either a banner OR a vendor match in addition to port 554:**
```powershell
(
    ($ports -contains 554 -and
        (-not [string]::IsNullOrWhiteSpace($RTSPBanner) -or
         $combined -match 'hikvision|dahua|axis|hanwha|bosch.*security|milestone|reolink|amcrest')
    ) -or
    $RTSPBanner -match 'RTSP' -or
    $combined -match 'hikvision|dahua|axis|hanwha|bosch.*security|milestone|reolink|amcrest'
) { 'Camera' }
```

Update the `Confidence` block to reflect:
```powershell
'Camera' {
    if (-not [string]::IsNullOrWhiteSpace($RTSPBanner) -or
        $combined -match 'hikvision|dahua|axis') { 'High' }
    elseif ($ports -contains 554) { 'Medium' }
    else { 'Low' }
}
```

---

### M-11 · Private Helper Functions Are in Public Files — Module Structure Violation

**Root cause:**
The module loader exports `$Public.BaseName` — one export per file. But
several public files contain additional unexported functions. These are
accessible inside the module but violate the documented Private/Public split:

| Unexported function | Lives in public file |
|---|---|
| `Get-VBARPTable` | `Public\Get-VBARPEntry.ps1` |
| `Invoke-VBmDNSBrowse` | `Public\Get-VBmDNSRecord.ps1` |
| `Invoke-VBSwitchSNMPWalk` | `Public\Get-VBSwitchARP.ps1` |
| `Invoke-VBLoadOUITable` | `Public\Get-VBOUIVendor.ps1` |
| `Format-VBEnrichmentContext` | `Public\Get-VBEnrichmentContext.ps1` |
| `Invoke-VBBuildEnrichmentObject` | `Public\Invoke-VBIPEnrichment.ps1` |

**Fix:**
Move each helper into a dedicated file in `Private\`:
- `Private\Get-VBARPTable.ps1`
- `Private\Invoke-VBmDNSBrowse.ps1`
- `Private\Invoke-VBSwitchSNMPWalk.ps1`
- `Private\Invoke-VBLoadOUITable.ps1`
- `Private\Format-VBEnrichmentContext.ps1`
- `Private\Invoke-VBBuildEnrichmentObject.ps1`

Remove the function definitions from the public files. The module loader
dot-sources `Private\` first, so all helpers are available when public
functions load.

> Do not rename the functions. Only move the files.

---

### M-12 · `Invoke-VBSqliteCommand` Has Identical `if/else` Branches — `NonQuery` Switch Is Inert

**File:** [`Private\Invoke-VBSqliteCommand.ps1` lines 77–83](Private\Invoke-VBSqliteCommand.ps1)

**Root cause:**
```powershell
if ($NonQuery) {
    Invoke-SqliteQuery @splat   # ← same call
}
else {
    Invoke-SqliteQuery @splat   # ← same call
}
```
Both branches call `Invoke-SqliteQuery @splat` identically. The `-NonQuery`
switch is completely inoperative. `PSSQLite`'s `Invoke-SqliteQuery`
automatically returns rows for SELECT and rows-affected for INSERT/UPDATE.
The `if/else` here is cargo code.

**Fix:**
Collapse to a single call:
```powershell
Invoke-SqliteQuery @splat
```

Remove the `$NonQuery` parameter and its `[Parameter()]` block, and update
all callers that pass `-NonQuery` to remove that parameter. The behaviour
is unchanged — `PSSQLite` already handles the distinction internally.

> Do this last in the Medium batch because it requires touching every
> caller of `Invoke-VBSqliteCommand`.

---

## LOW — Hygiene and Teaching Quality

---

### L-01 · `[object[]]$IPAddress` Parameter — Replace with `[string[]]`

**File:** [`Public\Invoke-VBIPEnrichment.ps1` line 88](Public\Invoke-VBIPEnrichment.ps1)

Replace `[object[]]$IPAddress` with `[string[]]$IPAddress`. Remove the
manual type-coercion block at lines 126–133. If the caller sends a
PSCustomObject through the pipeline, `ValueFromPipelineByPropertyName`
will bind the `IPAddress` property as a string automatically because
the aliases `IP_Address`, `IP Address`, `IP` are declared.

---

### L-02 · Missing `[ValidateNotNullOrEmpty()]` on All `$IPAddress` Parameters

**Files:** `Get-VBPTRRecord.ps1`, `Get-VBTCPFingerprint.ps1`, `Get-VBSNMPIdentity.ps1`,
`Get-VBRTSPBanner.ps1`, `Get-VBSwitchARP.ps1`, `Get-VBARPEntry.ps1`,
`Get-VBADComputer.ps1`, `Get-VBDHCPLease.ps1`.

Add `[ValidateNotNullOrEmpty()]` to every `[Parameter(Mandatory)]`
`[string]$IPAddress` declaration. Rejects empty strings at bind time
with a clear error rather than propagating to network calls.

---

### L-03 · `Invoke-VBBuildEnrichmentObject` Uses Wrong Verb and Wrong Location

**File:** [`Public\Invoke-VBIPEnrichment.ps1` line 839](Public\Invoke-VBIPEnrichment.ps1)

Covered by M-11 (move to `Private\`). Additionally rename:
```
Invoke-VBBuildEnrichmentObject → ConvertFrom-VBSqliteEnrichmentRow
```
`ConvertFrom-` is the correct approved verb for "convert an input object
into a different type." `Invoke-Build*` is not an approved pattern.

---

### L-04 · SNMP Community Strings Are Plaintext in the Context Object

**File:** [`Public\Get-VBEnrichmentContext.ps1` line 497](Public\Get-VBEnrichmentContext.ps1)

Store as `[SecureString]` array and convert to plaintext only at point
of use in `Get-VBSNMPIdentity` and `Get-VBSwitchARP`:

```powershell
# In Get-VBEnrichmentContext — store as SecureString
SNMPCommunityStrings = @($SNMPCommunityStrings | ForEach-Object {
    ConvertTo-SecureString -String $_ -AsPlainText -Force
})

# In Get-VBSNMPIdentity — convert back at use
$plaintext = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString)
)
```

---

### L-05 · TLS 1.2 Force in OUI Download Not Guarded by PS Version

**File:** [`Public\Get-VBOUIVendor.ps1` line 216](Public\Get-VBOUIVendor.ps1)

```powershell
# BEFORE — sets process-wide TLS, disables TLS 1.3 on PS 6+
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# AFTER — only needed on PS 5.1
if ($PSVersionTable.PSVersion.Major -lt 6) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}
```

---

### L-06 · Duplicate Step Descriptions Between Sequential and Parallel Orchestrator Paths

**File:** [`Public\Invoke-VBIPEnrichment.ps1`](Public\Invoke-VBIPEnrichment.ps1)

Sequential path line 541: `'SNMP unavailable (olePrn COM not present)'`
Parallel path line 423: `'SNMP unavailable'`

The two code paths have diverged in their trace detail strings. This will
continue to happen as long as the logic is duplicated.

Long-term fix: extract steps 5–11 into `Private\Invoke-VBActiveProbeSet.ps1`
called by both paths. Short-term fix: align the detail strings to be
identical between both paths.

---

## Validation Checklist

After completing all Critical and High fixes, run the following validation
before declaring the fix batch complete:

```powershell
# 1. Import module fresh
Remove-Module VB.DNSEnrichment -ErrorAction SilentlyContinue
Import-Module .\VB.DNSEnrichment.psd1 -Force -Verbose

# 2. Confirm Data\ folder and OUI file path resolve identically
$ctx = Get-VBEnrichmentContext -Quiet
$ctx.OUIFilePath   # must show VB.DNSEnrichment\Data\oui.csv

# 3. Confirm OUI file is downloaded and loaded
Get-VBOUIVendor -MACAddress '00:1A:2B:3C:4D:5E' -Context $ctx
# Must return Vendor = some vendor name, not NoResult

# 4. Run full enrichment against 5 known IPs (mix of domain-joined,
#    DHCP-only, static, printer, camera if available)
$results = '10.0.0.1','10.0.0.5','10.0.0.10','10.0.0.20','10.0.0.50' |
    Invoke-VBIPEnrichment -Context $ctx -ForceRefresh -Verbose

# 5. Confirm resolution rate has improved
$results | Select-Object IPAddress, IsResolved, Hostname, HostnameSource,
    DeviceClass, Confidence, StepsSucceeded, StepsFailed | Format-Table

# 6. Confirm no SQL errors in verbose output
# 7. Confirm OUI vendor appears on any result with a MAC address
# 8. Confirm LayerTrace for each IP shows PTR-Unconfirmed (not just Skipped)
#    for any device with a PTR record whose forward confirmation failed
```

---

## Summary Table

| ID | Priority | File | Issue | Fixed By |
|---|---|---|---|---|
| C-01 | CRITICAL | `Get-VBTCPFingerprint.ps1:70` | TCP timeout 300 ms | Raise to 1000 ms |
| C-02 | CRITICAL | `Invoke-VBIPEnrichment.ps1:158` | SQL injection | Parameterised IN clause |
| C-03 | CRITICAL | `Get-VBmDNSRecord.ps1:86` | mDNS guard inverted | Positive condition check |
| C-04 | CRITICAL | `Invoke-VBIPEnrichment.ps1:373` | Parallel cache rebuild | Module path guard + OUI pre-warm |
| H-01 | HIGH | `Get-VBPTRRecord.ps1:117` | Forward-confirm discards valid PTR | IPAddress.Parse() + accept unconfirmed |
| H-02 | HIGH | `Invoke-VBIPEnrichment.ps1:680` | Vendor not written to SQLite | Add Vendor to upsert SQL |
| H-03 | HIGH | `Get-VBDHCPLease.ps1:95` | DHCP scope errors silenced | Per-scope try/catch + warning |
| H-04 | HIGH | `Get-VBSwitchARP.ps1:216` | Switch MAC format too strict | Use ConvertTo-VBNormalisedMAC |
| H-05 | HIGH | `Invoke-VBIPEnrichment.ps1:109` | `$EAP = Stop` in begin block | Remove the assignment |
| H-06 | HIGH | `Get-VBmDNSRecord.ps1:156` | `Start-Sleep` in mDNS browse | Replace with `Wait-Job -Timeout` |
| M-01 | MEDIUM | `Get-VBEnrichmentContext.ps1:119` | OUI `Data\` folder missing | Create folder; fix path resolution |
| M-02 | MEDIUM | `Get-VBEnrichmentContext.ps1:549` | `Write-Host` in context report | Replace with `Write-Information` |
| M-03 | MEDIUM | `Get-VBEnrichmentContext.ps1` | No `Severity` in prereq report | Add Severity field to `$addReport` |
| M-04 | MEDIUM | `Get-VBEnrichmentContext.ps1:410` | DB init failure not actionable | Add SQL folder path to error detail |
| M-05 | MEDIUM | `Get-VBEnrichmentContext.ps1:148` | `CanUseParallel` too simple | Check runspace type |
| M-06 | MEDIUM | `New-VBLayerResult.ps1:104` | Non-ISO8601 timestamp | `ToString('o')` |
| M-07 | MEDIUM | `Get-VBOUIVendor.ps1:207` | OUI age uses `CreationTime` | Switch to `LastWriteTime` |
| M-08 | MEDIUM | `New-VBLayerResult.ps1:107` | Dead code in ExtraFields merge | Single-line assignment |
| M-09 | MEDIUM | All cache files | Script-scope caches have no TTL | Add `*CacheBuiltAt` + TTL check |
| M-10 | MEDIUM | `Resolve-VBDeviceClass.ps1:147` | Camera Tier 2 over-broad | Require banner OR vendor + port 554 |
| M-11 | MEDIUM | Multiple public files | Private helpers in public files | Move to `Private\` |
| M-12 | MEDIUM | `Invoke-VBSqliteCommand.ps1:77` | `NonQuery` switch is inert | Collapse to single call |
| L-01 | LOW | `Invoke-VBIPEnrichment.ps1:88` | `[object[]]$IPAddress` | Change to `[string[]]` |
| L-02 | LOW | All layer files | No `[ValidateNotNullOrEmpty()]` | Add to all `$IPAddress` params |
| L-03 | LOW | `Invoke-VBIPEnrichment.ps1:839` | Wrong verb on build helper | Rename + move to Private |
| L-04 | LOW | `Get-VBEnrichmentContext.ps1:497` | SNMP community strings plaintext | Store as `SecureString` |
| L-05 | LOW | `Get-VBOUIVendor.ps1:216` | TLS 1.2 force unguarded | PS version guard |
| L-06 | LOW | `Invoke-VBIPEnrichment.ps1` | Diverged trace detail strings | Align both code paths |
