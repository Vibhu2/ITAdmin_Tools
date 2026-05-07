---
Title:       "VB.WindowsDNSLogAnalysis — PowerShell Module Design Specification"
Version:     "1.1"
Date:        "2026-05-07"
Author:      "VB"
Doc_status:  "Final"
Environment: "On-Prem / Internal"
---

# VB.WindowsDNSLogAnalysis — PowerShell Module Design Specification

> **[i] INFO:** This document is the single source of truth for the VB.WindowsDNSLogAnalysis module.
> It is written to be handed to a developer (or your future self) with zero prior context.
> Everything needed to understand, build, and extend this module is in here — including every
> place the implementation diverged from the original design and why.

---

## Contents

1. [Background — Why This Module Exists](#1-background--why-this-module-exists)
2. [The Problem We Solved First — Parser Performance](#2-the-problem-we-solved-first--parser-performance)
3. [The Pivot to SQLite — Why and What Changed](#3-the-pivot-to-sqlite--why-and-what-changed)
4. [Module Overview](#4-module-overview)
5. [Dependencies](#5-dependencies)
6. [File Structure](#6-file-structure)
7. [Database Design](#7-database-design)
8. [Public Functions](#8-public-functions)
9. [Private Functions](#9-private-functions)
10. [Performance Architecture — The Import Pipeline](#10-performance-architecture--the-import-pipeline)
11. [Design Principles](#11-design-principles)
12. [Getting Started — First Run](#12-getting-started--first-run)
13. [Reference — Windows DNS Debug Log Format](#13-reference--windows-dns-debug-log-format)
14. [Implementation Deviations — As Built vs Design](#14-implementation-deviations--as-built-vs-design)
15. [External References](#15-external-references)

---

## 1. Background — Why This Module Exists

### The starting point

Windows DNS servers, when debug logging is enabled, write a file called `dns.log` to `%windir%\System32\dns\`. On a busy internal DNS server this file can reach **450MB–1GB per day**, containing 20–40 million lines. Each line represents a DNS packet the server sent or received.

We needed to answer operational questions from these logs:

- Which client IPs are generating the most DNS traffic?
- Which domains are being queried most frequently?
- Are there clients hitting NXDOMAIN responses — potential malware beaconing or misconfiguration?
- Which query types (A, AAAA, MX, PTR…) dominate traffic?
- What does DNS traffic look like across a time window?
- What is each machine actually using DNS for — protocol, query type, top domain?

### The original approach and why it failed

The first tool written was `Get-DNSLogSummary` — a PowerShell function that read the log file, applied regex, and returned results. The critical flaw: **it re-parsed the entire file every time it was called**. Ask two different questions, parse twice. Ask ten questions, parse ten times. On a 450MB file, every question had a cost measured in minutes.

The function was eventually optimised to run in 2–5 minutes per parse (see Section 2), but re-parsing is still the wrong model. The log file is historical data — it does not change. It should be read once, stored, and queried at will.

### The solution — parse once, store in SQLite, query forever

`VB.WindowsDNSLogAnalysis` is a PowerShell module that:

1. Parses DNS debug log files once using the full optimised parser engine
2. Stores every record in a local SQLite database
3. Exposes the data through SQL-backed query functions

After the initial import, every question is answered by a SQL query — milliseconds, not minutes. New log files are appended to the same database. Re-importing an already-imported file is detected and skipped automatically via SHA256 file hashing.

---

## 2. The Problem We Solved First — Parser Performance

Before writing this module, significant work went into making the log parser as fast as possible. That work is the foundation of the private `Invoke-VBDNSLogParser` function. A developer picking this up needs to understand *why* the parser is written the way it is — it looks unusual compared to typical PowerShell scripts.

### Performance journey — 450MB log file

| Approach | Time taken |
| :--- | ---: |
| `Get-Content` + pipeline (original) | 5+ hours |
| `StreamReader` + compiled regex | ~30–45 min |
| + `*PACKET*` pre-filter + IP cache + `IndexOf` decode | ~10–15 min |
| + Path B (deferred object construction) | ~2–5 min |
| + SQLite PRAGMA tuning + prepared stmt + single transaction | **target: 2–5 min total including DB write** |

### What was changed and why — the nine optimisations

**1. `StreamReader` instead of `Get-Content`**

`Get-Content` loads the entire file into memory as a string array before a single line is processed. `System.IO.StreamReader.ReadLine()` reads one line at a time. On a 450MB file the memory difference is enormous and processing starts immediately.

```powershell
# Never do this for large files
$lines = Get-Content $path           # loads ~450MB into RAM before loop starts

# Correct approach
$reader = [System.IO.StreamReader]::new($path)
while (($line = $reader.ReadLine()) -ne $null) { ... }
$reader.Close()
```

**2. `*PACKET*` pre-filter before regex**

A Windows DNS debug log is mostly noise — EVENT records, blank lines, header text, server startup messages. Only `PACKET` lines carry actual query data, roughly 20–30% of all lines. A cheap string check before running the expensive regex eliminates 70–80% of lines before the regex engine ever sees them.

```powershell
if ($line -notlike '*PACKET*') { continue }   # skip ~75% of all lines here
# Only PACKET lines reach the regex below
```

**3. Compiled regex**

PowerShell interprets regex patterns fresh on every call by default. Compiling the pattern once and reusing it across 30 million lines gives 2–3× faster matching.

```powershell
# In the begin{} block — compiled once
$logRegex = [regex]::new(
    '\s+(UDP|TCP)\s+(Rcv|Snd)\s+([\d.]+|[0-9a-fA-F:]+)\s+([0-9a-fA-F]+)\s+(Q|R|)\s*\[([0-9a-fA-F]+)\s+([A-Z]*)\s+(\w+)\]\s+(\w+)\s+(.+)',
    [System.Text.RegularExpressions.RegexOptions]::Compiled
)

# In the process{} loop — reused, not recompiled
$m = $logRegex.Match($line)
```

> **[i] NOTE:** The regex above reflects the actual built pattern (10 groups). See Section 14, Deviation #2 for the full explanation of why this differs from the original 8-group design assumption.

**4. Conditional timestamp parsing**

`[datetime]::ParseExact()` is expensive. In the original standalone function it was skipped when no time filter was set. In this module we always store `LogDateTime` in the database, so we always parse — but we parse it once per line using `ParseExact` with a known fixed format, which is the fastest possible path.

```powershell
$logDateTime = [datetime]::ParseExact(
    "$logDate $logTime", 'M/d/yyyy h:mm:ss tt',
    [System.Globalization.CultureInfo]::InvariantCulture
).ToString('yyyy-MM-ddTHH:mm:ss')
```

**5. IP classification cache**

Every unique client IP appears thousands of times in a DNS log. Running 3–4 regex checks per line (one for version detection, three for RFC1918 ranges) on every PACKET line was a significant cost. A hashtable means each IP is classified once on first appearance, then retrieved in O(1) for every repeat.

```powershell
if (-not $ipCache.ContainsKey($ip)) {
    $ipCache[$ip] = @{
        Version = if ($ip -match '^\d') { 'IPv4' } else { 'IPv6' }
        Private = Test-VBPrivateIP -IPAddress $ip
    }
}
$cached = $ipCache[$ip]
```

**6. `IndexOf`/`Substring` for DNS wire-format decoding**

DNS query names are stored in wire format: `(7)example(3)com(0)`. The original approach ran a second compiled regex on every PACKET line to decode this. Replacing it with a pure .NET string loop using `IndexOf` and `Substring` eliminates the regex engine entirely for this step — no allocations beyond a small label list. This change alone produced a 15–20% overall speedup.

```powershell
# Decode (7)example(3)com(0) → example.com
# Handled in private function ConvertFrom-VBDNSName
# Uses IndexOf/Substring loop, not regex
while ($pos -lt $rqLen -and $rawQueryName[$pos] -eq [char]'(') {
    $close  = $rawQueryName.IndexOf(')', $pos)
    $segLen = [int]$rawQueryName.Substring($pos + 1, $close - $pos - 1)
    if ($segLen -eq 0) { break }
    $labels.Add($rawQueryName.Substring($close + 1, $segLen))
    $pos = $close + 1 + $segLen
}
```

**7. `[string]::Concat()` for key construction**

PowerShell string interpolation (`"$a|$b|$c"`) creates intermediate strings and goes through the formatter. `[string]::Concat()` is a direct .NET call. On 30 million iterations the difference is 3–5%.

```powershell
# Slower — string interpolation
$key = "$protocol|$direction|$ip|$queryType"

# Faster — direct .NET call, use for any string built inside the hot loop
$key = [string]::Concat($protocol, '|', $direction, '|', $ip, '|', $queryType)
```

**8. `Write-Progress` at 10,000-line intervals**

`Write-Progress` triggers a console redraw. At 2,000-line intervals on a 30M-line file that is 15,000 redraws. At 10,000 lines it is 3,000 — same user experience, 5–7% less overhead.

**9. Path B — deferred object construction (most important structural decision)**

`PSCustomObject` creation involves .NET reflection, property bag construction, and PowerShell's type system. It is expensive to do inside a tight I/O loop millions of times. The hot loop instead appends a plain `object[]` array — the cheapest possible allocation. Object construction is deferred until after the disk read is complete.

In this module, Path B is even more powerful than in the standalone function: **PSCustomObject is never created during import at all**. The `object[]` row maps directly to SQLite `@parameter` binding. There is zero conversion cost between the parse result and the database insert.

```powershell
# Hot loop — cheapest possible allocation, no PSCustomObject
$buffer.Add([object[]]@($logDateTime, $threadId, $packetId, $protocol,
    $direction, $ip, $ipVersion, $isPrivate, $xid, $pktKind,
    $opcode, $flagsHex, $flagsChar, $rCode, $status, $error,
    $queryType, $queryName, $sourceFile))

# object[] is handed directly to SQLite parameter binding in Invoke-VBBulkInsert
# No PSCustomObject is ever created in the import path
```

---

## 3. The Pivot to SQLite — Why and What Changed

### Why SQLite

SQLite is a serverless, file-based relational database. No installation, no service, no configuration — just a `.db` file. It is the right choice here because:

- We need SQL (`GROUP BY`, `ORDER BY`, `WHERE`, aggregations) which PowerShell filter chains cannot match for flexibility or speed
- The data is write-once from multiple log files, read-many for analysis
- A `.db` file is portable — open it directly in DB Browser for SQLite, DBeaver, or Python
- `PSSQLite` wraps `System.Data.SQLite` with PowerShell-friendly cmdlets and is a single `Install-Module` away

### What the shift to SQLite changes in the design

The standalone `Get-DNSLogSummary` applied filters (protocol, direction, private IP exclusion) during parsing — because parsing was the only time the data was available. In this module, **filters move from parse-time to query-time**. The parser captures everything; SQL `WHERE` handles filtering. This means:

- The same imported data answers any question without re-parsing
- The `-ExcludePrivateIPs` option on `Import-VBDNSLog` is a storage decision (never import private IPs at all), not a filter decision
- All the filter parameters from `Get-DNSLogSummary` reappear on `Get-VBDNSLog`, translated to SQL `WHERE` clauses

### SQLite-specific performance decisions

These have no equivalent in the original function. They are the difference between a slow import and a fast one.

**WAL mode (Write-Ahead Logging)**

SQLite's default journal mode serialises all writes through a single lock. WAL mode allows multiple concurrent writers — essential for the PS7 parallel import strategy where each thread writes its own file's data independently.

```sql
PRAGMA journal_mode = WAL;
```

**PRAGMA tuning — set once per connection before any inserts**

```sql
PRAGMA synchronous   = NORMAL;     -- safe with WAL, much faster than FULL
PRAGMA cache_size    = -65536;     -- 64MB page cache in RAM
PRAGMA temp_store    = MEMORY;     -- temp tables and indexes stay in RAM
PRAGMA mmap_size     = 268435456;  -- 256MB memory-mapped I/O, reduces syscall overhead
```

**Single transaction per file (batch at 50k rows)**

SQLite autocommit means one transaction per `INSERT` — one disk sync per row. On 1 million rows that is 1 million disk syncs. A single `BEGIN TRANSACTION` / `COMMIT` wrapping all inserts reduces that to one. For very large files, batch every 50,000 rows to control memory.

```powershell
$tx = $conn.BeginTransaction()
# insert rows...
if (++$rowCount % 50000 -eq 0) {
    $tx.Commit()
    $tx = $conn.BeginTransaction()
    $cmd.Transaction = $tx   # re-wire command to new transaction
}
$tx.Commit()
```

> **[!] WARNING:** Do NOT call `$tx.Dispose()` between batch commits. On some builds of `System.Data.SQLite` (bundled with PSSQLite), disposing the transaction object also invalidates the `$cmd.Transaction` reference, causing subsequent inserts to silently discard rows. Let the GC collect the old transaction. See Section 14, Deviation #5 for full detail.

**Prepared statements**

The SQL `INSERT` is parsed and compiled by SQLite once. Every row reuses the same execution plan via parameter binding — faster than string-concatenated inserts and eliminates injection risk.

```powershell
$cmd = $conn.CreateCommand()
$cmd.CommandText = "INSERT INTO DNSLog (LogDateTime, Protocol, ...) VALUES (@p0, @p1, ...)"
# Parameters added once in begin{} block
# Inside loop: just set .Value per column and call ExecuteNonQuery()
```

---

## 4. Module Overview

**Module name:** `VB.WindowsDNSLogAnalysis`
**PowerShell version required:** 7.0 or later (for `ForEach-Object -Parallel`)
**Platform:** Windows (DNS debug log format is Windows Server specific)
**Purpose:** Parse Windows DNS debug log files once, store all records in SQLite, expose data via SQL-backed query functions
**Use case:** Internal IT operations — DNS traffic analysis, troubleshooting, security investigation

> **[i] INFO:** This module is for internal use only. It uses PS7-specific features (`-Parallel`) and assumes a controlled environment. It is not designed for client distribution.

### Naming convention

All public and private functions follow the VB naming standard: `VB` in the noun portion of the `Verb-Noun` name. This scopes them clearly in shared environments where multiple modules may export similar verb/noun combinations (e.g. `Get-VBDNSLog` rather than `Get-DNSLog`).

---

## 5. Dependencies

| Dependency | How to install | Why needed |
| :--- | :--- | :--- |
| PowerShell 7.0+ | `winget install Microsoft.PowerShell` | `ForEach-Object -Parallel` for multi-file import |
| PSSQLite | `Install-Module PSSQLite` | SQLite connection, `Invoke-SqliteQuery`, prepared statements |

SQLite itself is bundled inside `PSSQLite` — no separate installation needed.

Verify before first use:

```powershell
$PSVersionTable.PSVersion          # must be 7.0+
Get-Module PSSQLite -ListAvailable # must be installed
```

---

## 6. File Structure

```
VB.WindowsDNS/
├── VB.WindowsDNSLogAnalysis.psd1       # Module manifest — FunctionsToExport lists Public only
├── VB.WindowsDNSLogAnalysis.psm1       # Loader — dot-sources all Public and Private .ps1 files
│
├── Public/                             # Exported — callable by the user
│   ├── Initialize-VBDNSLogDatabase.ps1
│   ├── Import-VBDNSLog.ps1
│   ├── Get-VBDNSLog.ps1
│   ├── Invoke-VBDNSLogQuery.ps1
│   ├── Get-VBDNSLogStatistics.ps1
│   └── Export-VBDNSLogReport.ps1
│
├── Private/                            # Internal helpers — never exported, never called directly
│   ├── Invoke-VBDNSLogParser.ps1       # Hot-loop parsing engine — StreamReader + compiled regex
│   ├── ConvertFrom-VBDNSName.ps1       # Wire-format DNS name decoder — IndexOf/Substring
│   ├── Test-VBPrivateIP.ps1            # RFC1918 + IPv6 private range classifier
│   ├── Invoke-VBBulkInsert.ps1         # Transaction + prepared statement bulk insert
│   └── Get-VBImportStatus.ps1          # SHA256 hash check against ImportLog table
│
└── Documentation/
    └── Module_Design.md                # This document
```

The `VB.WindowsDNSLogAnalysis.psm1` loader pattern:

```powershell
$Private = Get-ChildItem "$PSScriptRoot\Private\*.ps1" -ErrorAction SilentlyContinue
$Public  = Get-ChildItem "$PSScriptRoot\Public\*.ps1"  -ErrorAction SilentlyContinue

foreach ($file in @($Private + $Public)) { . $file.FullName }

Export-ModuleMember -Function $Public.BaseName
```

---

## 7. Database Design

The database is a single `.db` file. The path is chosen by the caller at `Initialize-VBDNSLogDatabase` time and passed to every subsequent function via `-DatabasePath`.

### Table: `DNSLog`

This is the main data table. Every record is one line from a DNS debug log that matched the PACKET pattern. Nothing is filtered at import time (except optionally private IPs).

| Column | Type | Source | Description |
| :--- | :--- | :--- | :--- |
| `Id` | INTEGER PK | Auto | Autoincrement row identifier |
| `LogDate` | TEXT | Field 1 | Raw date from log line (M/d/yyyy) |
| `LogTime` | TEXT | Field 2 | Raw time from log line (h:mm:ss tt) |
| `LogDateTime` | TEXT | Derived | ISO8601 combined `yyyy-MM-ddTHH:mm:ss` — **indexed** — primary time filter column |
| `ThreadId` | TEXT | Field 3 | DNS server thread ID (hex) |
| `PacketId` | TEXT | Field 5 | Internal packet identifier (hex) — extracted via separate targeted match |
| `Protocol` | TEXT | Field 6 | `UDP` or `TCP` — **indexed** |
| `Direction` | TEXT | Field 7 | `Rcv` (received from client) or `Snd` (sent by server) |
| `IPAddress` | TEXT | Field 8 | Remote client IP address — **indexed** |
| `IPVersion` | TEXT | Derived | `IPv4` or `IPv6` — derived from IPAddress format |
| `IsPrivate` | INTEGER | Derived | `1` if RFC1918/private, `0` if public — derived from IPAddress |
| `TransactionId` | TEXT | Field 9 | DNS transaction ID / Xid (hex) |
| `PacketKind` | TEXT | Field 10 | `Q` (query) or `R` (response) |
| `Opcode` | TEXT | Hardcoded | Hardcoded `'Q'` — not separately encoded in debug log format; preserved for forward compatibility |
| `FlagsHex` | TEXT | Field 12 | Raw flags in hex (e.g. `8081`) |
| `FlagsChar` | TEXT | Field 13 | Flag chars: `A`=Auth Answer, `T`=Truncated, `D`=Recursion Desired, `R`=Recursion Available |
| `ResponseCode` | TEXT | Field 14 | DNS RCODE: `NOERROR`, `NXDOMAIN`, `SERVFAIL`, `REFUSED`, `FORMERR`, `NOTIMPL` — **indexed** |
| `Status` | TEXT | Derived | `Success` if RCODE is NOERROR, otherwise `Error` |
| `Error` | TEXT | Derived | RCODE string if Status is Error, empty string if Success |
| `QueryType` | TEXT | Field 15 | DNS record type: `A`, `AAAA`, `MX`, `PTR`, `SOA`, `NS`, `CNAME`, `TXT`, `SRV`, `ANY`… — **indexed** |
| `QueryName` | TEXT | Field 16 | Fully qualified domain name, decoded from wire format — **indexed** |
| `SourceFile` | TEXT | Added | Full path of the log file this record came from — **indexed** |
| `ImportedAt` | TEXT | Added | ISO8601 timestamp of when this row was inserted |

> **[!] IMPORTANT — PacketKind and ResponseCode:**
> `ResponseCode` is only meaningful on **response** rows (`PacketKind = 'R'`). Query rows (`PacketKind = 'Q'`) always carry `NOERROR` in the log regardless of the actual outcome. All NXDOMAIN/error counts in reports are correctly scoped to `PacketKind = 'R'` only. If a DNS server is configured to log only inbound queries (no responses), all NXDOMAIN-related report columns will show zero — this is a server logging configuration issue, not a data issue. Enable bidirectional logging with `Set-DnsServerDiagnostics -All $true` to capture response rows.

### Table: `ImportLog`

One row per imported log file. Used by `Get-VBImportStatus` to detect already-imported files and skip them. Also provides an audit trail.

| Column | Type | Description |
| :--- | :--- | :--- |
| `Id` | INTEGER PK | Autoincrement |
| `FilePath` | TEXT UNIQUE | Full path of the log file — UNIQUE prevents duplicate imports |
| `FileName` | TEXT | File name only (for display) |
| `FileHash` | TEXT | SHA256 hash — dedup by content, not just path |
| `RecordCount` | INTEGER | Number of rows inserted from this file |
| `ErrorCount` | INTEGER | Number of lines that failed to parse |
| `DurationSeconds` | REAL | Parse + insert time for this file |
| `ImportedAt` | TEXT | ISO8601 timestamp of the import |
| `ServerName` | TEXT | DNS server name, parsed from file path if determinable |

### Indexes

All indexes are created by `Initialize-VBDNSLogDatabase`. They exist on the columns most commonly used in `WHERE`, `ORDER BY`, and `GROUP BY` clauses.

```sql
CREATE INDEX IF NOT EXISTS idx_LogDateTime  ON DNSLog (LogDateTime);
CREATE INDEX IF NOT EXISTS idx_IPAddress    ON DNSLog (IPAddress);
CREATE INDEX IF NOT EXISTS idx_Protocol     ON DNSLog (Protocol);
CREATE INDEX IF NOT EXISTS idx_ResponseCode ON DNSLog (ResponseCode);
CREATE INDEX IF NOT EXISTS idx_QueryType    ON DNSLog (QueryType);
CREATE INDEX IF NOT EXISTS idx_QueryName    ON DNSLog (QueryName);
CREATE INDEX IF NOT EXISTS idx_SourceFile   ON DNSLog (SourceFile);
```

Without indexes, SQLite scans every row for every query — fine for 10,000 rows, unacceptable for 10 million.

---

## 8. Public Functions

### `Initialize-VBDNSLogDatabase`

**Why it exists:** The database must be set up with the correct schema, indexes, and journal mode before any data can be imported. This is the one-time setup step.

**What it does:**
- Creates the `.db` file at the specified path (or connects to an existing one)
- Creates the `DNSLog` table if it does not exist
- Creates the `ImportLog` table if it does not exist
- Creates all 7 indexes
- Sets `PRAGMA journal_mode = WAL` permanently on the database
- Validates the `-DatabasePath` argument before opening (see input guards below)

**Input validation guards (added post-design — see Section 14, Deviation #7):**

| Guard | Triggered when |
| :--- | :--- |
| Folder path detection | `-DatabasePath` resolves to an existing directory |
| Non-SQLite file detection (magic byte check) | `-DatabasePath` points to a file that does not have `SQLite format 3` header |
| Parent directory existence | `-DatabasePath` specifies a directory that does not exist yet |

**Idempotent:** All `CREATE TABLE` and `CREATE INDEX` statements use `IF NOT EXISTS`. Safe to re-run against an existing database — nothing is changed or lost.

**Parameters:**

| Parameter | Type | Required | Description |
| :--- | :--- | :--- | :--- |
| `-DatabasePath` | string | Yes | Full path to the `.db` file to create or verify |

**Returns:** Nothing. Emits `Verbose` messages confirming each object created.

**Usage:**
```powershell
Initialize-VBDNSLogDatabase -DatabasePath 'C:\Realtime\dns_analysis.db'
```

---

### `Import-VBDNSLog`

**Why it exists:** This is the main workhorse — the function that reads log files and populates the database. It is the most complex function in the module and the most performance-critical. Everything in Section 2 was built specifically to make this function fast.

**What it does:**
1. Resolves the input path to a list of `.log` files (single file, wildcard, or directory)
2. For each file, calls `Get-VBImportStatus` to check the SHA256 hash against `ImportLog` — skips if already imported
3. Uses `ForEach-Object -Parallel` (PS7) to process multiple files concurrently
4. Per thread: dot-sources all Private `.ps1` files (required — PS7 parallel runspaces do not inherit module functions; see Section 14, Deviation #4)
5. Per thread: calls `Invoke-VBDNSLogParser` to parse the file into an `object[]` buffer
6. Per thread: calls `Invoke-VBBulkInsert` to write the buffer to SQLite and write the `ImportLog` audit row

**Parameters:**

| Parameter | Type | Required | Default | Description |
| :--- | :--- | :--- | :--- | :--- |
| `-InputPath` | string | Yes | — | Path to log file, directory, or wildcard (`C:\logs\*.log`) |
| `-DatabasePath` | string | Yes | — | Path to the SQLite `.db` file (must be initialised first) |
| `-Recurse` | switch | No | false | Search subdirectories for log files |
| `-ThrottleLimit` | int | No | 8 | Max parallel threads (`ForEach-Object -Parallel`) |
| `-ExcludePrivateIPs` | switch | No | false | Skip private IP records at parse time — reduces storage on busy internal servers |
| `-Force` | switch | No | false | Re-import files even if hash exists in `ImportLog` |

**Returns:** A summary object per file:
```powershell
FileName       : dns_20260501.log
RecordCount    : 1847293
ErrorCount     : 12
DurationSec    : 187.4
Status         : Success
```

**Usage:**
```powershell
# Import a single file
Import-VBDNSLog -InputPath 'C:\DNS\Logs\dns_20260501.log' -DatabasePath 'C:\Realtime\dns_analysis.db'

# Import all logs in a directory, 4 parallel threads
Import-VBDNSLog -InputPath 'C:\DNS\Logs\' -DatabasePath 'C:\Realtime\dns_analysis.db' -ThrottleLimit 4

# Force re-import of a specific file
Import-VBDNSLog -InputPath 'C:\DNS\Logs\dns.log' -DatabasePath 'C:\Realtime\dns_analysis.db' -Force
```

> **[!] WARNING:** Run `Initialize-VBDNSLogDatabase` before calling `Import-VBDNSLog` for the first time. Calling it against a non-existent or uninitialised database will error.

---

### `Get-VBDNSLog`

**Why it exists:** Most operational questions about DNS traffic follow a predictable pattern — filter by IP, domain, query type, time window, error status. This function exposes those filters as named parameters so the common cases require no SQL knowledge.

**What it does:** Builds a parameterised `SELECT` query against `DNSLog` based on the provided parameters and returns matching rows as `[PSCustomObject[]]`. All filtering happens in SQLite — no PowerShell-side filtering after the query.

**Parameters:**

| Parameter | Type | Description |
| :--- | :--- | :--- |
| `-DatabasePath` | string | Path to the SQLite `.db` file |
| `-IPAddress` | string | Filter to a specific client IP |
| `-QueryName` | string | Filter by domain name (exact or SQL `LIKE` pattern with `%`) |
| `-QueryType` | string | Filter by record type: `A`, `AAAA`, `MX`, `PTR`… |
| `-Protocol` | ValidateSet | `UDP`, `TCP`, or `All` (default `All`) |
| `-Direction` | ValidateSet | `Rcv`, `Snd`, or `All` (default `All`) |
| `-PacketKind` | ValidateSet | `Q`, `R`, or `All` (default `All`) |
| `-ResponseCode` | string | `NOERROR`, `NXDOMAIN`, `SERVFAIL`… |
| `-Status` | ValidateSet | `Success` or `Error` |
| `-DateFrom` | datetime | Start of time window (inclusive) |
| `-DateTo` | datetime | End of time window (inclusive) |
| `-ExcludePrivateIPs` | switch | Exclude rows where `IsPrivate = 1` |
| `-SourceFile` | string | Limit to records from a specific imported file |
| `-Limit` | int | Max rows to return (default: 10,000) |

**Returns:** `[PSCustomObject[]]` with the same properties as the `DNSLog` table columns.

**Usage:**
```powershell
# All NXDOMAIN responses in the last 24 hours
Get-VBDNSLog -DatabasePath 'C:\Realtime\dns_analysis.db' `
    -ResponseCode NXDOMAIN -DateFrom (Get-Date).AddDays(-1)

# All queries from a specific client IP
Get-VBDNSLog -DatabasePath 'C:\Realtime\dns_analysis.db' -IPAddress '10.209.1.165' -PacketKind Q

# All AAAA queries matching a domain pattern
Get-VBDNSLog -DatabasePath 'C:\Realtime\dns_analysis.db' -QueryType AAAA -QueryName '%realtime-it.com%'
```

---

### `Invoke-VBDNSLogQuery`

**Why it exists:** `Get-VBDNSLog` covers the predictable cases. `Invoke-VBDNSLogQuery` is the escape hatch for anything requiring full SQL — aggregations, joins with `ImportLog`, complex `GROUP BY`, or any question not anticipated by the filter parameters.

**What it does:** Opens a connection to the database and executes the provided SQL string directly. Returns a `DataTable` that can be piped to `Export-Csv`, `Out-GridView`, `ConvertTo-Json`, or `Export-VBDNSLogReport`.

**Parameters:**

| Parameter | Type | Required | Description |
| :--- | :--- | :--- | :--- |
| `-DatabasePath` | string | Yes | Path to the SQLite `.db` file |
| `-Query` | string | Yes | SQL query to execute |
| `-Parameters` | hashtable | No | Named parameters (prevents SQL injection) |

**Returns:** `[System.Data.DataTable]`

**Usage:**
```powershell
# Top NXDOMAIN domains (beaconing indicator)
Invoke-VBDNSLogQuery -DatabasePath 'C:\Realtime\dns_analysis.db' -Query @"
    SELECT QueryName, COUNT(*) AS Hits, COUNT(DISTINCT IPAddress) AS UniqueClients
    FROM DNSLog
    WHERE ResponseCode = 'NXDOMAIN'
    GROUP BY QueryName
    ORDER BY Hits DESC
    LIMIT 20
"@

# Always use -Parameters for variable values
$ip = '10.209.1.165'
Invoke-VBDNSLogQuery -DatabasePath 'C:\Realtime\dns_analysis.db' `
    -Query "SELECT QueryName, COUNT(*) AS Hits FROM DNSLog WHERE IPAddress = @ip AND PacketKind = 'Q' GROUP BY QueryName ORDER BY Hits DESC LIMIT 20" `
    -Parameters @{ ip = $ip }
```

---

### `Get-VBDNSLogStatistics`

**Why it exists:** The first questions asked of any DNS dataset are always the same — who is talking the most, what record types dominate, what is the error rate? Pre-building these as named report types means they are available immediately without writing any SQL.

**What it does:** Executes a pre-defined SQL query matched to the `-ReportType` value and returns the result as `[PSCustomObject[]]`. Each report type is self-contained and optionally time-bounded.

**Parameters:**

| Parameter | Type | Required | Description |
| :--- | :--- | :--- | :--- |
| `-DatabasePath` | string | Yes | Path to the SQLite `.db` file |
| `-ReportType` | ValidateSet | Yes | See table below |
| `-Top` | int | No | **Optional.** Omit to return all results; specify to limit (e.g. `-Top 25`). Applies to ranking reports only. |
| `-DateFrom` | datetime | No | Restrict analysis to a time window |
| `-DateTo` | datetime | No | Restrict analysis to a time window |

**Report types:**

| ReportType | What it returns | Key columns |
| :--- | :--- | :--- |
| `TopTalkers` | Client IPs ranked by query count | `IPAddress`, `IPVersion`, `IsPrivate`, `Queries`, `Responses`, `NXDOMAIN`, `OtherErrors` |
| `TopDomains` | Domain names ranked by query count | `QueryName`, `Queries`, `Responses`, `NXDOMAIN`, `NOERROR`, `OtherErrors` |
| `TalkerDetail` | Per-IP query type and protocol breakdown | `IPAddress`, `IPVersion`, `IsPrivate`, `Queries`, `A_Queries`, `AAAA_Queries`, `PTR_Queries`, `MX_Queries`, `SRV_Queries`, `TXT_Queries`, `Other_Queries`, `UDP_Queries`, `TCP_Queries`, `TopDomain` |
| `QueryTypeBreakdown` | Count and percentage per QueryType (A, AAAA, MX…) | `QueryType`, `Count`, `Percentage` |
| `ErrorRate` | Count and percentage per ResponseCode | `ResponseCode`, `Count`, `Percentage` |
| `Timeline` | Query count grouped by hour | `Hour`, `TotalPackets`, `Queries`, `Responses`, `Errors` |
| `PrivateVsPublic` | Row counts split by IsPrivate | `AddressType`, `Count`, `Percentage` |
| `DirectionSplit` | Rcv vs Snd counts | `Direction`, `Count`, `Percentage` |
| `ImportSummary` | One row per imported file from ImportLog | `FileName`, `FilePath`, `RecordCount`, `ErrorCount`, `DurationSeconds`, `ImportedAt`, `ServerName`, `FileHash` |

**Summary output:** `TopTalkers` and `TopDomains` automatically print a summary line after results showing total unique IPs/domains, private/public split (TopTalkers), and total query count. The summary is written via `Write-Host` and does not enter the pipeline — export operations are unaffected.

**`TalkerDetail` use case:** Identify what each machine is actually using DNS for. High `PTR_Queries` = reverse lookups (often monitoring tools or printers). High `SRV_Queries` = domain-joined services. High `TCP_Queries` = large responses, possible zone transfer attempts.

**`ErrorRate` note:** Only counts response packets (`PacketKind = 'R'`). If the DNS server only logs inbound queries, this report will show all NOERROR.

**Usage:**
```powershell
$db = 'C:\Realtime\dns_analysis.db'

Get-VBDNSLogStatistics -DatabasePath $db -ReportType TopTalkers | Format-Table -AutoSize
Get-VBDNSLogStatistics -DatabasePath $db -ReportType TopTalkers -Top 25 | Format-Table -AutoSize
Get-VBDNSLogStatistics -DatabasePath $db -ReportType TalkerDetail -Top 10 | Format-Table -AutoSize
Get-VBDNSLogStatistics -DatabasePath $db -ReportType ErrorRate
Get-VBDNSLogStatistics -DatabasePath $db -ReportType Timeline -DateFrom (Get-Date).AddDays(-7)
```

---

### `Export-VBDNSLogReport`

**Why it exists:** Query results need to be shared — in tickets, with management, in spreadsheets. This function handles all output formatting so the query functions stay clean and return only data objects.

**What it does:** Accepts pipeline input from `Get-VBDNSLog`, `Invoke-VBDNSLogQuery`, or `Get-VBDNSLogStatistics` and writes to a file in the specified format.

**Parameters:**

| Parameter | Type | Required | Description |
| :--- | :--- | :--- | :--- |
| `-InputObject` | object | Yes (pipeline) | Result set from any query function |
| `-OutputPath` | string | Yes | Destination file path |
| `-Format` | ValidateSet | No | `CSV` (default), `HTML`, `XLSX` |
| `-Title` | string | No | Report title for HTML and XLSX headers |

**Returns:** Nothing. Writes the file and emits a `Verbose` message with the row count and path.

**Usage:**
```powershell
Get-VBDNSLogStatistics -DatabasePath 'C:\Realtime\dns_analysis.db' -ReportType TopDomains |
    Export-VBDNSLogReport -OutputPath 'C:\Reports\TopDomains.csv'

Get-VBDNSLog -DatabasePath 'C:\Realtime\dns_analysis.db' -ResponseCode NXDOMAIN |
    Export-VBDNSLogReport -OutputPath 'C:\Reports\NXDomains.xlsx' -Format XLSX -Title 'NXDOMAIN Report'
```

---

## 9. Private Functions

These functions are internal implementation details. They are never exported and never called directly by the user. They are called by the public functions above. If you are extending the module, these are what you modify to change parsing or storage behaviour.

---

### `Invoke-VBDNSLogParser`

**Why it exists:** The parsing engine is the most performance-critical code in the module. Isolating it in its own private function means it can be called from inside a PS7 parallel thread, independently unit tested, and modified without touching import logic.

**What it does:** Opens a `StreamReader` on the given file, reads every line, and applies the full optimised parsing pipeline — pre-filter → compiled regex → IP cache → `ConvertFrom-VBDNSName` → `object[]` buffer. Returns a `PSCustomObject` wrapper containing the buffer and error count.

**Key implementation rules:**
- The `begin{}` block compiles the regex once and initialises the IP cache and buffer — shared across all lines in the file
- The hot loop does the minimum possible work per line: pre-filter first, regex second, cache lookup third, string ops last
- `object[]` is used throughout — no `PSCustomObject`, no hashtables, no named properties inside the loop
- `ConvertFrom-VBDNSName` is called per PACKET line but is itself a tight loop with no regex engine
- The function returns a `PSCustomObject` wrapper `{ Buffer, ErrorCount }` — **not** the raw `List` directly (see Section 14, Deviation #3 for why)

**Parameters:**

| Parameter | Type | Description |
| :--- | :--- | :--- |
| `-FilePath` | string | Full path to the log file to parse |
| `-ExcludePrivateIPs` | switch | Skip rows where the IP is private |

**Returns:** `[PSCustomObject]@{ Buffer = [List[object[]]]; ErrorCount = [int] }`

**`object[]` column order (fixed contract — must stay in sync with `Invoke-VBBulkInsert`):**

```
Index  Column
  0    LogDateTime   (ISO8601 string)
  1    LogDate
  2    LogTime
  3    ThreadId
  4    PacketId
  5    Protocol
  6    Direction
  7    IPAddress
  8    IPVersion
  9    IsPrivate      (0 or 1 as integer)
 10    TransactionId
 11    PacketKind
 12    Opcode
 13    FlagsHex
 14    FlagsChar
 15    ResponseCode
 16    Status
 17    Error
 18    QueryType
 19    QueryName
 20    SourceFile
```

> **[!] WARNING:** The column order in `object[]` is a contract between `Invoke-VBDNSLogParser` and `Invoke-VBBulkInsert`. If you add or reorder columns, update the `INSERT` parameter binding in `Invoke-VBBulkInsert` at exactly the same time or rows will be inserted into the wrong columns silently.

---

### `ConvertFrom-VBDNSName`

**Why it exists:** DNS query names in the debug log are encoded in wire format: `(7)example(3)com(0)`. Every PACKET line has one. Decoding this correctly and efficiently matters — it is called on every parsed line. Using a regex here was measured to be 15–20% slower than the `IndexOf`/`Substring` approach, hence the unusual-looking implementation.

**What it does:** Takes a raw wire-format string and returns a decoded FQDN string. Uses a `while` loop with `IndexOf` to locate each length prefix `(n)`, extracts the label of that length, builds a label list, then joins with `.`.

**Input:** `(7)example(3)com(0)`
**Output:** `example.com`

**Edge cases handled:**
- `(0)` terminator — stops the loop (not included as a label)
- Single-label names (`(9)localhost(0)` → `localhost`)
- Already-decoded names with no `(n)` prefix — returned as-is
- Empty or null input — returns empty string

---

### `Test-VBPrivateIP`

**Why it exists:** IP classification (is this address private?) is needed for every PACKET line. The IP cache in `Invoke-VBDNSLogParser` means the actual classification logic runs only once per unique IP, but that first call must be correct. Centralising this logic here means it can be updated in one place if ranges need to change.

**What it does:** Takes an IP address string and returns `$true` if it falls within a private or reserved range.

**Ranges checked:**

| Range | Standard | Address family |
| :--- | :--- | :--- |
| `10.0.0.0/8` | RFC1918 | IPv4 |
| `172.16.0.0/12` | RFC1918 | IPv4 |
| `192.168.0.0/16` | RFC1918 | IPv4 |
| `127.0.0.0/8` | Loopback | IPv4 |
| `169.254.0.0/16` | Link-local | IPv4 |
| `::1` | Loopback | IPv6 |
| `fe80::/10` | Link-local | IPv6 |
| `fc00::/7` | Unique local | IPv6 |

**Implementation note:** Uses pre-compiled regex patterns (module-level constants, not recompiled per call) for IPv4 ranges and `StartsWith` checks for IPv6 ranges.

---

### `Invoke-VBBulkInsert`

**Why it exists:** The SQLite insert strategy — PRAGMA tuning, prepared statement, single transaction, batch commits — is identical for every file regardless of content. Centralising it here means `Import-VBDNSLog` stays clean and the insert logic is tuned in one place.

**What it does:**
1. Opens a `SQLiteConnection` to the database
2. Applies PRAGMA tuning values on the connection
3. Prepares a single `INSERT INTO DNSLog` statement with 21 named parameters
4. Begins a transaction
5. Iterates the `List[object[]]` buffer — for each row, binds the array values to parameters by index and calls `ExecuteNonQuery()`
6. Commits every 50,000 rows, begins a new transaction, re-wires `$cmd.Transaction` to new transaction
7. Final commit after the loop
8. Writes one `ImportLog` audit row on the same connection
9. Closes the connection

**Parameters:**

| Parameter | Type | Description |
| :--- | :--- | :--- |
| `-DatabasePath` | string | Path to the SQLite `.db` file |
| `-Buffer` | `List[object[]]` | Parsed rows from `Invoke-VBDNSLogParser` |
| `-BatchSize` | int | Rows per transaction commit (default: 50,000) |
| `-SourceFile` | string | Full path of log file — written to ImportLog |
| `-RecordCount` | int | Row count — written to ImportLog |
| `-ErrorCount` | int | Parse error count — written to ImportLog |
| `-FileHash` | string | SHA256 hash — written to ImportLog |
| `-DurationSeconds` | double | Total parse+insert time — written to ImportLog |

**Returns:** Row count inserted as `[int]`.

> **[i] INFO:** Each PS7 parallel thread creates its own `SQLiteConnection`. WAL mode allows concurrent connections from different threads. Never pass a connection object across thread boundaries in PS7 parallel blocks — this causes unpredictable behaviour.

> **[i] INFO:** The `ImportLog` write is done inside `Invoke-VBBulkInsert` (not in `Import-VBDNSLog`) so it runs on the same connection as the bulk insert and only succeeds if the bulk insert completed without throwing. See Section 14, Deviation #6.

---

### `Get-VBImportStatus`

**Why it exists:** Importing the same file twice creates duplicate rows in `DNSLog`. We need a reliable pre-check before starting an expensive parse-and-insert operation. SHA256 hashing is used rather than path comparison because log files are often renamed or rotated.

**What it does:** Computes the SHA256 hash of the given file and queries `ImportLog` for a matching `FileHash`. Returns a status object indicating whether the file has already been imported.

**Why SHA256 and not just the file path?** Path-based dedup breaks when log files are renamed (e.g. `dns.log` → `dns_20260501.log`). A content hash detects the same data regardless of filename. It also correctly handles the opposite case — a new file at the same path with different content gets a different hash and is imported.

**Parameters:**

| Parameter | Type | Description |
| :--- | :--- | :--- |
| `-FilePath` | string | Full path to the log file to check |
| `-DatabasePath` | string | Path to the SQLite `.db` file |

**Returns:**
```powershell
[PSCustomObject]@{
    FilePath        = 'C:\DNS\dns.log'
    FileHash        = 'abc123def456...'
    AlreadyImported = $true        # or $false
    ImportedAt      = '2026-05-01T10:30:00'   # null if not yet imported
}
```

---

## 10. Performance Architecture — The Import Pipeline

This section documents the full end-to-end flow of a single file being imported. Understand this before modifying any parsing or insert code.

```
Import-VBDNSLog
│
├── Get-VBImportStatus
│     SHA256 hash check vs ImportLog
│     → already imported: skip file entirely
│
└── ForEach-Object -Parallel  (one thread per file, -ThrottleLimit N)
    │
    ├── Dot-source Private/*.ps1           ← required: PS7 runspaces start empty
    │
    ├── Invoke-VBDNSLogParser              ← the hot read loop
    │     StreamReader.ReadLine()           line-by-line, flat memory usage
    │       │
    │       ├── '*PACKET*' pre-filter      skip ~75% of all lines here
    │       │
    │       ├── $logRegex.Match()          compiled regex, ~25% of lines reach this
    │       │
    │       ├── $ipCache lookup            O(1) after first occurrence per IP
    │       │     └── Test-VBPrivateIP     called once per unique IP only
    │       │
    │       ├── ConvertFrom-VBDNSName      IndexOf/Substring — no regex engine
    │       │
    │       ├── RCODE → Status/Error       inline string compare, no function call
    │       │
    │       └── $buffer.Add(object[])     no PSCustomObject, no reflection overhead
    │     returns PSCustomObject{ Buffer, ErrorCount }
    │
    └── Invoke-VBBulkInsert                ← the hot write loop
          PRAGMA tuning (WAL, cache, mmap, sync=NORMAL)
          Prepare INSERT statement once
          BEGIN TRANSACTION
            for each object[] in buffer:
              bind array[index] to @param by position
              ExecuteNonQuery()
            COMMIT every 50,000 rows → new transaction → re-wire cmd.Transaction
          Final COMMIT
          Write row to ImportLog (same connection)
```

**What never happens in the import path:**
- No `PSCustomObject` creation (only at query time, never at import time)
- No `Get-Content` or full-file memory loads
- No PowerShell-side filtering (all filtering happens at query time via SQL `WHERE`)
- No shared state between parallel threads (each has its own connection and buffer)
- No DB reads during import (IP cache keeps classification entirely in RAM)
- No `$tx.Dispose()` between batch commits (invalidates `cmd.Transaction` on some SQLite builds)

---

## 11. Design Principles

These are the cross-cutting rules applied throughout the module. New functions and modifications must follow them.

**Output-presentation separation.** `Get-` functions return data objects only. `Export-VBDNSLogReport` handles all file output. No `Get-` function writes to disk or formats for display. This makes every function independently testable and composable via the pipeline.

**Cache expensive repeated operations.** IP classification, compiled regex, `Get-Date` calls — if the same computation produces the same result more than once, compute it once. On a 1 million line file from one server there may be only 200 unique client IPs. Classify 200 times, not 1 million.

**Apply cheap filters before expensive ones.** `*PACKET*` string check before regex. Protocol/direction filter before `ConvertFrom-VBDNSName`. All filters before buffer append. Order matters inside a hot loop running 30 million iterations.

**Fail gracefully, never crash the pipeline.** Parse errors emit a failure object into the result stream rather than throwing. A multi-file import should complete all files even if one has corrupted lines or a permissions error. `ErrorCount` in `ImportLog` records how many lines failed per file.

**Pre-size collections.** `List[object[]]::new(500000)` avoids repeated internal array resizing. Adjust the initial capacity if working with files significantly smaller or larger than 450MB.

**Never share connections across threads.** Each PS7 parallel thread creates its own `SQLiteConnection`. WAL mode supports this pattern. Passing a connection object across parallel thread boundaries causes unpredictable behaviour and potential data corruption.

**Measure before optimising.** Every optimisation in Section 2 was justified by profiling the actual bottleneck, not guessing. Do not add complexity to code that has not been measured as a bottleneck.

**Validate early with actionable messages.** `Initialize-VBDNSLogDatabase` checks for common user mistakes (folder path, wrong file type, missing parent directory) before any database operation. Raw SQLite error messages (`"unable to open database file"`) do not tell the user what to do. Guard messages do.

---

## 12. Getting Started — First Run

Step-by-step from a fresh machine:

```powershell
# 1. Verify PowerShell 7
$PSVersionTable.PSVersion       # must be 7.x

# 2. Install PSSQLite
Install-Module PSSQLite -Scope CurrentUser

# 3. Import the module
Import-Module 'C:\Users\Vibhu.Bhatnagar\Documents\GitHub\ITAdmin_Tools\VB.WindowsDNS\VB.WindowsDNSLogAnalysis.psd1'

# 4. Create the database — run once
Initialize-VBDNSLogDatabase -DatabasePath 'C:\Realtime\dns_analysis.db'

# 5. Import your first log file
Import-VBDNSLog -InputPath 'C:\DNS\Logs\dns.log' `
    -DatabasePath 'C:\Realtime\dns_analysis.db' -Verbose

# 6. Import an entire directory (parallel, 4 threads)
Import-VBDNSLog -InputPath 'C:\DNS\Logs\' `
    -DatabasePath 'C:\Realtime\dns_analysis.db' -ThrottleLimit 4

# 7. Query with filters
Get-VBDNSLog -DatabasePath 'C:\Realtime\dns_analysis.db' `
    -ResponseCode NXDOMAIN -DateFrom (Get-Date).AddDays(-1)

# 8. Pre-built statistics
Get-VBDNSLogStatistics -DatabasePath 'C:\Realtime\dns_analysis.db' -ReportType TopTalkers -Top 20 | Format-Table -AutoSize
Get-VBDNSLogStatistics -DatabasePath 'C:\Realtime\dns_analysis.db' -ReportType TalkerDetail | Format-Table -AutoSize
Get-VBDNSLogStatistics -DatabasePath 'C:\Realtime\dns_analysis.db' -ReportType ErrorRate

# 9. Raw SQL for anything else
Invoke-VBDNSLogQuery -DatabasePath 'C:\Realtime\dns_analysis.db' -Query @"
    SELECT QueryName, COUNT(*) AS Hits
    FROM DNSLog
    WHERE ResponseCode = 'NXDOMAIN'
    GROUP BY QueryName
    ORDER BY Hits DESC
    LIMIT 20
"@

# 10. Export results
Get-VBDNSLogStatistics -DatabasePath 'C:\Realtime\dns_analysis.db' -ReportType TopDomains |
    Export-VBDNSLogReport -OutputPath 'C:\Reports\TopDomains.csv'
```

---

## 13. Reference — Windows DNS Debug Log Format

A Windows DNS debug log line looks like this:

```
10/14/2025 12:25:16 AM 2C94 PACKET  00000000015117E0 UDP Rcv 10.209.1.165    7419   Q [0001   D   NOERROR] A      (8)automate(11)realtime-it(3)com(0)
```

Field breakdown:

| Position | Example value | Maps to column | Notes |
| :--- | :--- | :--- | :--- |
| 1 | `10/14/2025` | LogDate | Format: M/d/yyyy |
| 2 | `12:25:16 AM` | LogTime | Format: h:mm:ss tt |
| 3 | `2C94` | ThreadId | Hex thread ID |
| 4 | `PACKET` | — | Literal — used as the pre-filter string |
| 5 | `00000000015117E0` | PacketId | Extracted via separate targeted match (`'PACKET\s+([0-9a-fA-F]+)\s+'`) |
| 6 | `UDP` | Protocol | UDP or TCP |
| 7 | `Rcv` | Direction | Rcv or Snd |
| 8 | `10.209.1.165` | IPAddress | Remote client IP address |
| 9 | `7419` | TransactionId | DNS Xid (hex) |
| 10 | `Q` | PacketKind | Q = query, R = response; may be absent on some lines (defaults to Q) |
| 11–13 | `[0001 D NOERROR]` | FlagsHex + FlagsChar + ResponseCode | Bracket-enclosed block — 3 separate captures |
| 14 | `A` | QueryType | DNS record type |
| 15 | `(8)automate(11)realtime-it(3)com(0)` | QueryName | Wire-format encoded FQDN |

> **[!] NOTE:** There is no standalone Opcode field between PacketKind and `[` in the actual log format. Earlier documentation assumed this field existed — it does not. The `Opcode` database column is hardcoded to `'Q'`. See Section 14, Deviation #2.

**Flag character codes (field 12):**

| Char | Meaning |
| :--- | :--- |
| `A` | Authoritative Answer |
| `T` | Truncated Response |
| `D` | Recursion Desired |
| `R` | Recursion Available |

**RCODE values (field 13) and their Status mapping:**

| RCODE | Meaning | Status in DB |
| :--- | :--- | :--- |
| `NOERROR` | Successful response | `Success` |
| `SERVFAIL` | Server internal failure | `Error` |
| `NXDOMAIN` | Non-existent domain | `Error` |
| `REFUSED` | Query refused by policy | `Error` |
| `FORMERR` | Malformed query | `Error` |
| `NOTIMPL` | Opcode not supported | `Error` |

---

## 14. Implementation Deviations — As Built vs Design

This section documents every place where the built module differs from the original v1.0 design and **why** each deviation was made. It is the authoritative record of the build decisions made during implementation.

### Summary

| # | Area | Design Said | Built Instead | Reason |
|---|---|---|---|---|
| 1 | Function names | `Get-DNSLog`, `Import-DNSLog`, etc. | `Get-VBDNSLog`, `Import-VBDNSLog`, etc. | VB naming convention |
| 2 | Regex — field order and groups | 8-group pattern with separate Opcode group | 10-group pattern; Opcode hardcoded; PacketId separate match | Real log format has no Opcode between PacketKind and `[` |
| 3 | `Invoke-VBDNSLogParser` return type | Returns raw `List[object[]]` | Returns `PSCustomObject` wrapper `{ Buffer, ErrorCount }` | `Add-Member` pipeline bug destroys `List` type |
| 4 | Parallel thread function availability | Not addressed | Private functions dot-sourced into each runspace | PS7 parallel runspaces do not inherit module functions |
| 5 | Transaction dispose between batch commits | `$tx.Dispose()` called between commits | `$tx.Dispose()` removed | Disposing invalidates `cmd.Transaction` on some SQLite builds |
| 6 | ImportLog write location | Written in `Import-VBDNSLog` after parallel block | Written inside `Invoke-VBBulkInsert` on same connection | Thread safety; ImportLog row only written if insert succeeded |
| 7 | `Initialize-VBDNSLogDatabase` input validation | No guards | Three guards added (folder, non-SQLite file, missing parent) | All three triggered in real use; SQLite error messages are cryptic |
| 8 | Progress reporting — parse phase | `Write-Progress` with `-PercentComplete -1` | Real percentage from `BaseStream.Position / fileSize` | Indeterminate bar appeared frozen on 700-second imports |
| 9 | Progress reporting — insert phase | Not mentioned | `Write-Progress` at every 50k-row batch commit | 700-second silent insert phase appeared frozen |
| 10 | Post-design additions | — | `TalkerDetail` report type added | User-requested per-IP query type and protocol breakdown |
| 11 | `-Top` parameter behaviour | Required, default 20 | Optional; omit = all results | Confusion over explicit `-All` switch; simpler to make Top optional |

---

### Deviation 1 — Function Naming: `VB` in the Noun

**Design:** `Initialize-DNSLogDatabase`, `Import-DNSLog`, `Get-DNSLog`, `Invoke-DNSLogQuery`,
`Get-DNSLogStatistics`, `Export-DNSLogReport`

**Built:** `Initialize-VBDNSLogDatabase`, `Import-VBDNSLog`, `Get-VBDNSLog`, `Invoke-VBDNSLogQuery`,
`Get-VBDNSLogStatistics`, `Export-VBDNSLogReport`

Same change applied to all 5 private functions.

**Why:** VB coding standard — all functions in VB-authored modules carry `VB` in the noun portion of the `Verb-Noun` name (same pattern as `Get-VBDisk`, `Get-VBServerInventory`, etc.). This scopes them clearly in shared environments where multiple modules may export similar verb/noun combinations.

**Impact:** Module manifest `FunctionsToExport` updated. No logic changes.

---

### Deviation 2 — Regex: Real Log Format vs Design Assumption

**Design regex (8 groups):**
```
(UDP|TCP)\s+(Rcv|Snd)\s+([\d.]+|[0-9a-fA-F:]+)\s+(\w+)\s+(Q|R)\s+\[([^\]]+)\]\s+(\w+)\s+(.+)
```
This assumed an **Opcode field** (`Q`, `N`, `U`) between `PacketKind` and the `[` bracket block, and grouped the entire bracket block as a single capture.

**Real log line:**
```
10/14/2025 12:25:16 AM 2C94 PACKET  00000000015117E0 UDP Rcv 10.209.1.165    7419   Q [0001   D   NOERROR] A      (8)automate(11)realtime-it(3)com(0)
```

After `PacketKind` (`Q`), the line goes **directly into `[FlagsHex FlagsChar ResponseCode]`**. There is no separate Opcode character between `Q` and `[`.

**Built regex (10 groups):**
```
\s+(UDP|TCP)\s+(Rcv|Snd)\s+([\d.]+|[0-9a-fA-F:]+)\s+([0-9a-fA-F]+)\s+(Q|R|)\s*\[([0-9a-fA-F]+)\s+([A-Z]*)\s+(\w+)\]\s+(\w+)\s+(.+)
```

Key differences:
- `[0-9a-fA-F]+` for TransactionId instead of `\w+`
- PacketKind `(Q|R|)` allows empty match — some lines omit it; blank defaults to `'Q'`
- `\s*` before `[` — whitespace between PacketKind and bracket is variable
- Bracket block split into 3 separate groups: FlagsHex, FlagsChar, ResponseCode
- FlagsChar is `[A-Z]*` (may be empty: `[0001   NOERROR]` with no flag chars)
- PacketId extracted separately via `'PACKET\s+([0-9a-fA-F]+)\s+'` match

**Opcode:** Hardcoded to `'Q'` — not separately encoded in debug log lines at this position. Column preserved in schema for forward compatibility.

**Impact:** This was the root cause of `RecordCount = 0` on initial import. Fixing this produced 3,003,774 parsed rows.

---

### Deviation 3 — `Invoke-VBDNSLogParser` Return Type

**Design:** Function returns raw `List[object[]]` directly.

**Built:** Function returns a `PSCustomObject` wrapper:
```powershell
return [PSCustomObject]@{
    Buffer     = $buffer     # List[object[]]
    ErrorCount = $errorCount
}
```

**Why — the `Add-Member` pipeline destruction bug:**

The original build attempted to attach `ErrorCount` to the `List` using `Add-Member`:
```powershell
$buffer | Add-Member -NotePropertyName '_ErrorCount' -NotePropertyValue $errorCount -Force
return $buffer
```

`Add-Member` takes pipeline input. When you pipe a `List[object[]]`, PowerShell **iterates the collection** — it sends each `object[]` element through the pipeline, not the `List` itself. `Add-Member` operates on individual `object[]` elements. The return value becomes a plain `object[]` of `object[]` items, losing the `List[object[]]` type. When `Invoke-VBBulkInsert` then iterated `$buffer`, calling `$row[$i]` on each "row" returned `.ToString()` on the array — the literal string `"System.Object[]"` — rather than the field value. This is why all text columns contained `System.Object[]` in the database.

The `PSCustomObject` wrapper keeps the `List` in a named property. It is never piped. Buffer type is preserved intact.

**Impact:** Critical correctness fix.

---

### Deviation 4 — Parallel Thread Function Availability

**Design:** Did not address how private functions are made available inside `ForEach-Object -Parallel` blocks.

**Built:** `$PSScriptRoot` is captured before the parallel block, and each thread dot-sources all Private `.ps1` files at startup:

```powershell
$privateDir = Join-Path $PSScriptRoot '..\Private'

$filesToProcess | ForEach-Object -Parallel {
    $privateDir = $using:privateDir
    Get-ChildItem -Path $privateDir -Filter '*.ps1' | ForEach-Object { . $_.FullName }
    # Now Invoke-VBDNSLogParser, Invoke-VBBulkInsert, etc. are available
    ...
} -ThrottleLimit $ThrottleLimit
```

**Why:** PS7 `ForEach-Object -Parallel` creates fresh runspaces that do not inherit the calling session's module imports or function definitions. Each thread starts with only PS built-ins.

Without this, every thread threw `"The term 'Invoke-VBDNSLogParser' is not recognized"` and no files were processed.

**Note:** The `$using:` scope modifier is required to pass outer variables into the parallel block.

---

### Deviation 5 — Transaction `Dispose()` Bug

**Design:** Batch commits with `$tx.Dispose()` between commits.

**Built:** `Dispose()` removed:
```powershell
if ($rowsInserted % $BatchSize -eq 0) {
    $tx.Commit()
    $tx = $conn.BeginTransaction()
    $cmd.Transaction = $tx   # re-wire command to new transaction
}
```

**Why:** On certain builds of `System.Data.SQLite` bundled with PSSQLite, calling `$tx.Dispose()` between batch commits invalidates the `$cmd.Transaction` reference — the command held a reference to the now-disposed transaction. Subsequent `ExecuteNonQuery()` calls either threw or silently discarded rows. The symptom was only 1 row in the database despite 3,003,774 being parsed.

The `$cmd.Transaction = $tx` re-assignment after each `BeginTransaction()` is essential and was not in the design.

---

### Deviation 6 — `ImportLog` Write Location

**Design:** `ImportLog` audit row written inside `Import-VBDNSLog` after the parallel block completes.

**Built:** `ImportLog` write is done inside `Invoke-VBBulkInsert`, on the same connection used for the bulk insert, passed via extra parameters (`-SourceFile`, `-RecordCount`, `-ErrorCount`, `-FileHash`, `-DurationSeconds`).

**Why:** `Import-VBDNSLog` runs inside a `ForEach-Object -Parallel` block. Writing the audit row on the same connection that just finished the bulk insert is simpler and avoids opening a second connection per file just for the audit. It also means the `ImportLog` row is only written if the bulk insert succeeded — if the insert throws, the catch block rolls back and the ImportLog row is never written.

---

### Deviation 7 — `Initialize-VBDNSLogDatabase` Input Validation Guards

**Design:** No validation of `-DatabasePath` beyond basic existence checks.

**Built:** Three guards added before any database operation:

**Guard 1 — Folder path detection**
```powershell
if (Test-Path $DatabasePath -PathType Container) {
    throw "...must be a path to a .db FILE, not a folder..."
}
```
*Triggered by:* User passed the module folder path as `-DatabasePath`.

**Guard 2 — Non-SQLite file detection (magic byte check)**
```powershell
$magic  = [System.IO.File]::ReadAllBytes($DatabasePath) | Select-Object -First 16
$header = [System.Text.Encoding]::ASCII.GetString($magic).TrimEnd([char]0)
if (-not $header.StartsWith('SQLite format 3')) {
    throw "...is NOT a SQLite database...pass it to Import-VBDNSLog -InputPath instead..."
}
```
*Triggered by:* User passed the `.txt` DNS log file as `-DatabasePath`.

**Guard 3 — Parent directory existence**
```powershell
$parentDir = Split-Path $DatabasePath -Parent
if (-not [string]::IsNullOrWhiteSpace($parentDir) -and -not (Test-Path $parentDir)) {
    throw "Parent directory does not exist: '$parentDir'..."
}
```
*Triggered by:* User specified a path inside a directory that had not been created.

**Why:** All three occurred in real use during initial setup. Raw SQLite messages (`"unable to open database file"`, `"file is not a database"`) do not name the fix.

---

### Deviation 8 — Progress Reporting: Parse Phase

**Design:** `Write-Progress` with `-PercentComplete -1` (indeterminate bar).

**Built:** Real percentage from bytes read:
```powershell
$fileSize = (Get-Item $FilePath).Length

# Inside loop every 10,000 lines:
$pct = [math]::Min(99, [math]::Round($reader.BaseStream.Position / $fileSize * 100))
Write-Progress -Activity "Parsing DNS Log: $fileName" `
    -Status "Lines: $($lineCount.ToString('N0')) | PACKET rows: $($packetCount.ToString('N0')) | $pct%" `
    -PercentComplete $pct
```

**Why:** `-PercentComplete -1` displays an indeterminate bar that does not advance. On a 700-second import the bar appeared to do nothing. `StreamReader.BaseStream.Position` gives the current byte offset without additional I/O — dividing by file size gives an accurate percentage at zero cost.

---

### Deviation 9 — Progress Reporting: Insert Phase

**Design:** No progress reporting mentioned for `Invoke-VBBulkInsert`.

**Built:** `Write-Progress` added at every batch commit (every 50,000 rows):
```powershell
Write-Progress -Activity "Inserting DNS records: $insertFile" `
    -Status "$($rowsInserted.ToString('N0')) of $($totalRows.ToString('N0')) rows committed ($pct%)" `
    -PercentComplete $pct
```
Bar cleared with `-Completed` after the `ImportLog` row is written.

**Why:** The insert phase for 3 million rows took approximately 700 seconds with no visible output. From the user's perspective the process appeared frozen. Batch commit points are natural update moments — the work has already paused for a commit, so the `Write-Progress` call adds negligible overhead.

---

### Post-Design Addition — `TalkerDetail` Report Type

**Not in original design.** Added to `Get-VBDNSLogStatistics` after initial release.

**What it returns:** Per-IP breakdown of query types and protocol, plus the top queried domain per IP. Columns: `IPAddress`, `IPVersion`, `IsPrivate`, `Queries`, `A_Queries`, `AAAA_Queries`, `PTR_Queries`, `MX_Queries`, `SRV_Queries`, `TXT_Queries`, `Other_Queries`, `UDP_Queries`, `TCP_Queries`, `TopDomain`.

**TopDomain** is derived via a correlated subquery:
```sql
(SELECT d2.QueryName FROM DNSLog d2
 WHERE d2.IPAddress = DNSLog.IPAddress AND d2.PacketKind = 'Q' [AND time filter]
 GROUP BY d2.QueryName ORDER BY COUNT(*) DESC LIMIT 1) AS TopDomain
```

The time filter (`$innerWhere`) is derived from the outer `WHERE` clause with `WHERE` replaced by `AND` so it can be appended inline.

**Why added:** Operational need to understand per-machine DNS behaviour — not just how much traffic an IP generates (TopTalkers) but what it is doing: is it doing reverse lookups (monitoring/printers), domain-joined service discovery (SRV), large TCP responses (zone transfer attempts)?

---

### Post-Design Change — `-Top` Parameter Made Optional

**Design:** `-Top` with default value of 20 (always applied).

**Built:** `-Top` has no default value. Detected via `$PSBoundParameters.ContainsKey('Top')`. If omitted, all results are returned. If specified, `LIMIT @Top` is appended.

**Why:** An explicit `-All` switch was added first to return all results, then removed after user feedback — it caused confusion alongside `-Top`. Making `-Top` optional (omit = all results) achieves the same goal without an extra switch. Summary always shown for `TopTalkers`/`TopDomains` regardless of whether `-Top` is specified.

---

### What Was NOT Changed

Everything else in the design was implemented as specified:

- **Database schema** — identical: `DNSLog` and `ImportLog` tables with all columns and all 7 indexes
- **WAL mode + PRAGMA tuning** — applied exactly as designed
- **Prepared statements** — single `INSERT` compiled once, parameters bound per row
- **Batch size** — 50,000 rows per transaction commit
- **StreamReader** — line-by-line, no `Get-Content`
- **`*PACKET*` pre-filter** — Stage 1, unchanged
- **IP classification cache** — one classification per unique IP
- **`ConvertFrom-VBDNSName`** — `IndexOf`/`Substring` loop, no regex
- **SHA256 dedup** — content hash, not filename
- **`-ExcludePrivateIPs`** — storage decision applied at parse time
- **`-Force`** — deletes `ImportLog` row before re-importing
- **`Get-VBDNSLog`** — all filter parameters, all filtering in SQLite `WHERE`
- **`Invoke-VBDNSLogQuery`** — `-Parameters` hashtable, returns `DataTable`
- **`Export-VBDNSLogReport`** — CSV/HTML/XLSX, pipeline-first design
- **`object[]` column order** — 21 columns (indexes 0–20), fixed contract
- **Error handling** — try/catch per line in parser, never crashes pipeline
- **Per-thread connection isolation** — each parallel thread creates its own connection

---

## 15. External References

- [DNS Logging and Diagnostics — Microsoft Learn](https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2012-r2-and-2012/dn800669(v=ws.11))
- [Enable DNS Logging — Microsoft Learn](https://learn.microsoft.com/en-us/windows-server/networking/dns/dns-logging-and-diagnostics)
- [Windows Server DNS: Read the DNS Debug Log — TechNet Wiki](https://social.technet.microsoft.com/wiki/contents/articles/13640.windows-server-dns-read-the-dns-debug-log.aspx)
- [DNS Server Debug Logging — Andi Bellstedt](https://www.andibellstedt.com/posts/004_dnsserver-debug-logging/)
- [How to Enable DNS Query Logging — Windows OS Hub](https://woshub.com/enable-dns-query-logging-parse-logfile/)
- [Parsing Microsoft DNS Server Logs — Stephen Reese](https://www.rsreese.com/parsing-microsoft-dns-server-logs/)
- [PSSQLite Module — GitHub](https://github.com/RamblingCookieMonster/PSSQLite)
- [SQLite PRAGMA documentation](https://www.sqlite.org/pragma.html)
- [SQLite WAL mode](https://www.sqlite.org/wal.html)

---

*VB.WindowsDNSLogAnalysis Module Design Specification v1.1 — 2026-05-07 | Author: VB*
