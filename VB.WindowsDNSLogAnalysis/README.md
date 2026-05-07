# VB.WindowsDNSLogAnalysis

> **Parse once. Store in SQLite. Query forever.**
> A PowerShell module that turns raw Windows DNS debug logs into a structured, SQL-queryable database — with 9 built-in reports, parallel import, and export to CSV, HTML, or Excel.

---

## The Problem with DNS Debug Logs

Windows DNS debug logs are plain text — massive, sprawling, and nearly impossible to analyze at scale. A single busy DNS server can produce **450 MB of log data** every few hours. Grepping through that with `Select-String` works once. It doesn't scale, it can't be correlated, and it can't answer questions like:

- *Which machines are generating the most DNS traffic right now?*
- *What percentage of my queries are returning NXDOMAIN?*
- *Is there unusual overnight activity that might indicate beaconing?*
- *Which IP is hammering my DNS server with TXT queries?*

**VB.WindowsDNSLogAnalysis** solves this by importing your logs into SQLite once, then letting you query, filter, and export the data in seconds — with full parameterized SQL support and pre-built security-focused reports.

---

## Features

| Capability | Detail |
|---|---|
| **Fast Parsing** | StreamReader + compiled regex — processes 450 MB in 2–5 minutes |
| **SQLite Backend** | Single portable `.db` file, no server required |
| **SHA-256 Deduplication** | Re-importing the same file is safely skipped |
| **Parallel Import** | `ForEach-Object -Parallel` for multi-file batch imports (PS7+) |
| **9 Built-in Reports** | TopTalkers, TopDomains, TalkerDetail, ErrorRate, Timeline, and more |
| **Raw SQL Access** | Full parameterized query support via `Invoke-VBDNSLogQuery` |
| **Flexible Export** | CSV, HTML, and XLSX output via `Export-VBDNSLogReport` |
| **Security-Focused** | Reports designed to surface beaconing, exfiltration indicators, and recon |

---

## Requirements

- PowerShell 7.0 or later
- [PSSQLite](https://github.com/RamblingCookieMonster/PSSQLite) module

---

## Installation

```powershell
# 1. Install the SQLite dependency
Install-Module PSSQLite -Scope CurrentUser

# 2. Clone this repo
git clone https://github.com/Vibhu2/ITAdmin_Tools.git

# 3. Import the module
Import-Module '.\VB.WindowsDNSLogAnalysis\VB.WindowsDNSLogAnalysis.psd1'
```

---

## Quick Start

Five commands from nothing to queryable DNS data:

```powershell
# Step 1 — Import the module
Import-Module '.\VB.WindowsDNSLogAnalysis\VB.WindowsDNSLogAnalysis.psd1'

# Step 2 — Create the database (run once)
Initialize-VBDNSLogDatabase -DatabasePath 'C:\Realtime\dns_analysis.db'

# Step 3 — Import a log file
Import-VBDNSLog -InputPath 'C:\DNS\Logs\dns.log' -DatabasePath 'C:\Realtime\dns_analysis.db'

# Step 4 — Run your first report
$db = 'C:\Realtime\dns_analysis.db'
Get-VBDNSLogStatistics -DatabasePath $db -ReportType TopTalkers | Format-Table -AutoSize
```

That's it. Your DNS logs are now a queryable database.

---

## Module Functions

| Function | Purpose |
|---|---|
| `Initialize-VBDNSLogDatabase` | Create the SQLite schema and indexes. Run once per database. |
| `Import-VBDNSLog` | Parse DNS debug log files and load records into SQLite |
| `Get-VBDNSLog` | Filter and retrieve individual records by IP, domain, type, date range |
| `Get-VBDNSLogStatistics` | Run one of 9 pre-built analytical reports |
| `Invoke-VBDNSLogQuery` | Execute raw parameterized SQL against the database |
| `Export-VBDNSLogReport` | Export any result set to CSV, HTML, or XLSX |

---

## Importing Logs

```powershell
# Single file
Import-VBDNSLog -InputPath 'C:\DNS\Logs\dns.log' -DatabasePath $db

# Entire directory with 4 parallel threads
Import-VBDNSLog -InputPath 'C:\DNS\Logs\' -DatabasePath $db -ThrottleLimit 4

# Recurse subdirectories (useful for archive folders)
Import-VBDNSLog -InputPath 'D:\DNS_Archive\' -DatabasePath $db -Recurse

# Force re-import (bypass SHA-256 dedup check)
Import-VBDNSLog -InputPath 'C:\DNS\Logs\dns.log' -DatabasePath $db -Force

# Skip private IPs at import time (permanent — cannot be queried later)
Import-VBDNSLog -InputPath 'C:\DNS\Logs\dns.log' -DatabasePath $db -ExcludePrivateIPs
```

> **Tip:** Use `-ThrottleLimit 4` when importing a folder with many log files. Each file runs in its own parallel thread under PowerShell 7's `ForEach-Object -Parallel`.

---

## The 9 Built-in Reports

Set your database path once, then run any report:

```powershell
$db = 'C:\Realtime\dns_analysis.db'
```

---

### 1. TopTalkers — Who is generating the most DNS traffic?

```powershell
# All IPs
Get-VBDNSLogStatistics -DatabasePath $db -ReportType TopTalkers | Format-Table -AutoSize

# Top 25, last 24 hours
Get-VBDNSLogStatistics -DatabasePath $db -ReportType TopTalkers -Top 25 `
    -DateFrom (Get-Date).AddDays(-1) | Format-Table -AutoSize
```

**Output columns:** `IPAddress`, `IPVersion`, `IsPrivate`, `Queries`, `Responses`, `NXDOMAIN`, `OtherErrors`

A summary line at the bottom gives you total unique IPs, private vs. public split, and total query count.

---

### 2. TopDomains — What are machines looking up?

```powershell
# Top 50 domains this week
Get-VBDNSLogStatistics -DatabasePath $db -ReportType TopDomains -Top 50 `
    -DateFrom (Get-Date).AddDays(-7) | Format-Table -AutoSize
```

**Output columns:** `QueryName`, `Queries`, `Responses`, `NXDOMAIN`, `NOERROR`, `OtherErrors`

Domains with high NXDOMAIN counts and many unique clients are a classic beaconing indicator.

---

### 3. TalkerDetail — What is each machine actually doing?

```powershell
# Per-IP breakdown with query type, protocol, and top domain
Get-VBDNSLogStatistics -DatabasePath $db -ReportType TalkerDetail -Top 10 | Format-Table -AutoSize
```

**Output columns:** `IPAddress`, `IPVersion`, `IsPrivate`, `Queries`, `A_Queries`, `AAAA_Queries`, `PTR_Queries`, `MX_Queries`, `SRV_Queries`, `TXT_Queries`, `Other_Queries`, `UDP_Queries`, `TCP_Queries`, `TopDomain`

**What to look for:**
- High `PTR_Queries` → reverse lookups, common from monitoring tools or printers
- High `SRV_Queries` → domain-joined services doing service discovery
- High `TXT_Queries` → potential DNS-based data exfiltration
- High `TCP_Queries` → large responses; possible zone transfer attempts

---

### 4. QueryTypeBreakdown — What types of records are being requested?

```powershell
Get-VBDNSLogStatistics -DatabasePath $db -ReportType QueryTypeBreakdown | Format-Table -AutoSize
```

**Output columns:** `QueryType`, `Count`, `Percentage`

High `TXT` percentages can indicate DNS tunneling. High `ANY` queries warrant investigation — they're often used in reconnaissance and amplification attacks.

---

### 5. ErrorRate — How many queries are failing?

```powershell
# Last 7 days
Get-VBDNSLogStatistics -DatabasePath $db -ReportType ErrorRate `
    -DateFrom (Get-Date).AddDays(-7) | Format-Table -AutoSize
```

**Output columns:** `ResponseCode`, `Count`, `Percentage`

> Note: Only response packets (`PacketKind = R`) carry meaningful response codes. If your log only captures inbound queries, all rows will show NOERROR.

---

### 6. Timeline — When is the traffic happening?

```powershell
# Last 48 hours, by hour
Get-VBDNSLogStatistics -DatabasePath $db -ReportType Timeline `
    -DateFrom (Get-Date).AddDays(-2) | Format-Table -AutoSize
```

**Output columns:** `Hour`, `TotalPackets`, `Queries`, `Responses`, `Errors`

Look for spikes during off-hours — overnight DNS activity with consistent intervals is a textbook beaconing pattern.

---

### 7. PrivateVsPublic — Sanity check your traffic mix

```powershell
Get-VBDNSLogStatistics -DatabasePath $db -ReportType PrivateVsPublic | Format-Table -AutoSize
```

**Output columns:** `AddressType`, `Count`, `Percentage`

On an internal DNS server, almost all traffic should be from RFC1918 private addresses. A significant public IP percentage warrants investigation.

---

### 8. DirectionSplit — Is your logging configured correctly?

```powershell
Get-VBDNSLogStatistics -DatabasePath $db -ReportType DirectionSplit | Format-Table -AutoSize
```

**Output columns:** `Direction`, `Count`, `Percentage`

A healthy bidirectional log shows roughly equal `Rcv` and `Snd` counts. All `Rcv` and no `Snd` means only inbound queries are being captured — you're missing response data.

---

### 9. ImportSummary — What's in the database?

```powershell
Get-VBDNSLogStatistics -DatabasePath $db -ReportType ImportSummary | Format-Table -AutoSize
```

**Output columns:** `FileName`, `FilePath`, `RecordCount`, `ErrorCount`, `DurationSeconds`, `ImportedAt`, `ServerName`, `FileHash`

Verify which log files are loaded, when they were imported, how many records each contributed, and spot files with high parse error counts.

---

## Filtering Individual Records

```powershell
# All NXDOMAIN responses in the last 24 hours
Get-VBDNSLog -DatabasePath $db -ResponseCode NXDOMAIN -DateFrom (Get-Date).AddDays(-1)

# All queries from a specific IP
Get-VBDNSLog -DatabasePath $db -IPAddress '10.209.1.165' -PacketKind Q

# All AAAA queries matching a domain pattern
Get-VBDNSLog -DatabasePath $db -QueryType AAAA -QueryName '%microsoft.com%'

# Errors from public IPs in the last 7 days, capped at 5000 rows
Get-VBDNSLog -DatabasePath $db -Status Error -ExcludePrivateIPs `
    -DateFrom (Get-Date).AddDays(-7) -Limit 5000
```

---

## Raw SQL Queries

Need something the built-in reports don't cover? Use `Invoke-VBDNSLogQuery` with full parameterized SQL:

```powershell
# Top NXDOMAIN domains — beaconing indicator
Invoke-VBDNSLogQuery -DatabasePath $db -Query @"
    SELECT QueryName, COUNT(*) AS Hits, COUNT(DISTINCT IPAddress) AS UniqueClients
    FROM DNSLog
    WHERE ResponseCode = 'NXDOMAIN'
    GROUP BY QueryName
    ORDER BY Hits DESC
    LIMIT 20
"@

# What is one specific IP looking up? (parameterized — safe from injection)
$ip = '10.209.1.165'
Invoke-VBDNSLogQuery -DatabasePath $db `
    -Query "SELECT QueryName, COUNT(*) AS Hits FROM DNSLog WHERE IPAddress = @ip AND PacketKind = 'Q' GROUP BY QueryName ORDER BY Hits DESC LIMIT 20" `
    -Parameters @{ ip = $ip }
```

> Always use `-Parameters` for variable values. Never interpolate user-controlled strings directly into SQL.

---

## Exporting Reports

```powershell
# Export to CSV
Get-VBDNSLogStatistics -DatabasePath $db -ReportType TopTalkers |
    Export-VBDNSLogReport -OutputPath 'C:\Reports\top_talkers.csv'

# Export to Excel
Get-VBDNSLogStatistics -DatabasePath $db -ReportType TopDomains |
    Export-VBDNSLogReport -OutputPath 'C:\Reports\top_domains.xlsx' -Format XLSX -Title 'Top Domains'

# Export all 9 reports in one run
$outDir = 'C:\Reports\DNS'
$date   = Get-Date -Format 'yyyy-MM-dd'
New-Item -Path $outDir -ItemType Directory -Force | Out-Null

$reports = @('TopTalkers','TopDomains','TalkerDetail','QueryTypeBreakdown',
             'ErrorRate','Timeline','PrivateVsPublic','DirectionSplit','ImportSummary')

foreach ($r in $reports) {
    $data = Get-VBDNSLogStatistics -DatabasePath $db -ReportType $r
    $data | Export-VBDNSLogReport -OutputPath "$outDir\${date}_${r}.csv"
    $data | Export-VBDNSLogReport -OutputPath "$outDir\${date}_${r}.xlsx" -Format XLSX -Title $r
    Write-Host "Exported: $r ($($data.Count) rows)"
}
```

---

## Enabling Full DNS Debug Logging on the Server

If your data only shows `PacketKind = Q` (queries, no responses), enable bidirectional logging on the DNS server:

```powershell
# On the DNS server — requires DnsServer module
Set-DnsServerDiagnostics -All $true

# Or via dnscmd
dnscmd /config /loglevel 0x8100F331
```

Alternatively: DNS Manager → Server Properties → Debug Logging → tick **Incoming** + **Outgoing** + **Responses**.

---

## Module Structure

```
VB.WindowsDNSLogAnalysis/
├── VB.WindowsDNSLogAnalysis.psd1        # Module manifest
├── VB.WindowsDNSLogAnalysis.psm1        # Module root
├── Public/
│   ├── Initialize-VBDNSLogDatabase.ps1
│   ├── Import-VBDNSLog.ps1
│   ├── Get-VBDNSLog.ps1
│   ├── Get-VBDNSLogStatistics.ps1
│   ├── Invoke-VBDNSLogQuery.ps1
│   └── Export-VBDNSLogReport.ps1
├── Private/
│   ├── Invoke-VBDNSLogParser.ps1        # StreamReader + regex parse engine
│   ├── Invoke-VBBulkInsert.ps1          # Batched SQLite inserts
│   ├── ConvertFrom-VBDNSName.ps1        # DNS name normalization
│   ├── Get-VBImportStatus.ps1           # SHA-256 dedup check
│   └── Test-VBPrivateIP.ps1             # RFC1918 classification
└── Documentation/
    ├── VB.WindowsDnsLogAnalysisUsage.md
    └── Module_Design.md
```

---

## Version History

| Version | Date | Notes |
|---|---|---|
| 1.1.0 | 2026-05-07 | Added `TalkerDetail` report; `-Top` now optional; fixed `ResponseCode` scoping to response packets only; parse and insert progress bars |
| 1.0.0 | 2026-05-07 | Initial release — streaming parser, SHA-256 dedup, parallel import, 9 reports, CSV/HTML/XLSX export |

---

## Author

**Vibhu Bhatnagar** — Internal IT  
Part of the [ITAdmin_Tools](https://github.com/Vibhu2/ITAdmin_Tools) collection.

---

*PowerShell 7+ · PSSQLite · Windows DNS Debug Logs*
