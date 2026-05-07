---
type: Guide
title: "VB.WindowsDNSLogAnalysis — Module Usage Guide"
date: 2026-05-07
tags: [dns, powershell, sqlite, windows, logging, analysis, vb-windowsdnsloganalysis]
status: final
author: VB + Claude
---

# VB.WindowsDNSLogAnalysis — Module Usage Guide

> Windows DNS debug log analysis module. Parse once, store in SQLite, query forever.
> Module path: `C:\Users\Vibhu.Bhatnagar\Documents\GitHub\ITAdmin_Tools\VB.WindowsDNS\VB.WindowsDNSLogAnalysis.psd1`

---

## Quick Start

```powershell
# 1. Install dependency (once only)
Install-Module PSSQLite -Scope CurrentUser

# 2. Import module
Import-Module 'C:\Users\Vibhu.Bhatnagar\Documents\GitHub\ITAdmin_Tools\VB.WindowsDNS\VB.WindowsDNSLogAnalysis.psd1'

# 3. Create database (once only)
Initialize-VBDNSLogDatabase -DatabasePath 'C:\Realtime\dns_analysis.db'

# 4. Import a log file
Import-VBDNSLog -InputPath 'C:\DNS\Logs\dns.log' -DatabasePath 'C:\Realtime\dns_analysis.db'

# 5. Query
Get-VBDNSLogStatistics -DatabasePath 'C:\Realtime\dns_analysis.db' -ReportType TopTalkers
```

---

## Module Functions

| Function | Purpose |
|---|---|
| `Initialize-VBDNSLogDatabase` | Create the SQLite database with schema and indexes. Run once. |
| `Import-VBDNSLog` | Parse DNS debug log files and load into the database |
| `Get-VBDNSLog` | Filter and retrieve records by IP, domain, type, date, etc. |
| `Get-VBDNSLogStatistics` | Pre-built reports (9 report types) |
| `Invoke-VBDNSLogQuery` | Run raw SQL queries |
| `Export-VBDNSLogReport` | Export results to CSV, HTML, or XLSX |

---

## Import Parameters

```powershell
# Single file
Import-VBDNSLog -InputPath 'C:\DNS\Logs\dns.log' -DatabasePath 'C:\Realtime\dns_analysis.db'

# Entire directory, 4 parallel threads
Import-VBDNSLog -InputPath 'C:\DNS\Logs\' -DatabasePath 'C:\Realtime\dns_analysis.db' -ThrottleLimit 4

# Recurse subdirectories
Import-VBDNSLog -InputPath 'D:\DNS_Archive\' -DatabasePath 'C:\Realtime\dns_analysis.db' -Recurse

# Force re-import (ignore SHA256 dedup)
Import-VBDNSLog -InputPath 'C:\DNS\Logs\dns.log' -DatabasePath 'C:\Realtime\dns_analysis.db' -Force

# Skip private IPs at import time (permanent -- cannot be queried later)
Import-VBDNSLog -InputPath 'C:\DNS\Logs\dns.log' -DatabasePath 'C:\Realtime\dns_analysis.db' -ExcludePrivateIPs
```

---

## Report Types — One Example Each

Set the database path variable once for all examples below:

```powershell
$db = 'C:\Realtime\dns_analysis.db'
```

---

### 1. TopTalkers — Which IPs generate the most DNS traffic?

```powershell
# All IPs, no limit
Get-VBDNSLogStatistics -DatabasePath $db -ReportType TopTalkers | Format-Table -AutoSize

# Top 25 only
Get-VBDNSLogStatistics -DatabasePath $db -ReportType TopTalkers -Top 25 | Format-Table -AutoSize

# Last 24 hours only
Get-VBDNSLogStatistics -DatabasePath $db -ReportType TopTalkers -DateFrom (Get-Date).AddDays(-1) | Format-Table -AutoSize
```

**Columns:** `IPAddress`, `IPVersion`, `IsPrivate`, `Queries`, `Responses`, `NXDOMAIN`, `OtherErrors`

**Summary printed automatically at bottom:** Unique IP count, private/public split, total queries.

---

### 2. TopDomains — Which domains are queried most?

```powershell
# All domains
Get-VBDNSLogStatistics -DatabasePath $db -ReportType TopDomains | Format-Table -AutoSize

# Top 50 domains this week
Get-VBDNSLogStatistics -DatabasePath $db -ReportType TopDomains -Top 50 `
    -DateFrom (Get-Date).AddDays(-7) | Format-Table -AutoSize
```

**Columns:** `QueryName`, `Queries`, `Responses`, `NXDOMAIN`, `NOERROR`, `OtherErrors`

**Summary printed automatically at bottom:** Unique domain count, total queries, NXDOMAIN total.

---

### 3. TalkerDetail — Per-IP breakdown of query types and protocol

```powershell
# All IPs with query type and protocol breakdown + top domain per IP
Get-VBDNSLogStatistics -DatabasePath $db -ReportType TalkerDetail | Format-Table -AutoSize

# Top 10 busiest IPs with detail
Get-VBDNSLogStatistics -DatabasePath $db -ReportType TalkerDetail -Top 10 | Format-Table -AutoSize
```

**Columns:** `IPAddress`, `IPVersion`, `IsPrivate`, `Queries`, `A_Queries`, `AAAA_Queries`,
`PTR_Queries`, `MX_Queries`, `SRV_Queries`, `TXT_Queries`, `Other_Queries`,
`UDP_Queries`, `TCP_Queries`, `TopDomain`

**Use case:** Identify what each machine is actually using DNS for. High PTR count = reverse
lookups (often monitoring or printers). High SRV = domain-joined services. High TCP = large
responses, possible zone transfer attempts.

---

### 4. QueryTypeBreakdown — A vs AAAA vs MX vs PTR across all traffic

```powershell
Get-VBDNSLogStatistics -DatabasePath $db -ReportType QueryTypeBreakdown | Format-Table -AutoSize
```

**Columns:** `QueryType`, `Count`, `Percentage`

**Use case:** Understand the overall DNS traffic composition. Unusually high TXT can indicate
DNS-based data exfiltration. High ANY queries can indicate reconnaissance.

---

### 5. ErrorRate — NOERROR vs NXDOMAIN vs SERVFAIL breakdown

```powershell
Get-VBDNSLogStatistics -DatabasePath $db -ReportType ErrorRate | Format-Table -AutoSize

# Last 7 days only
Get-VBDNSLogStatistics -DatabasePath $db -ReportType ErrorRate `
    -DateFrom (Get-Date).AddDays(-7) | Format-Table -AutoSize
```

**Columns:** `ResponseCode`, `Count`, `Percentage`

**Note:** Only counts response packets (`PacketKind = R`). Query packets never carry meaningful
response codes. If your log only captures queries (inbound direction), all rows will show NOERROR.

---

### 6. Timeline — Traffic volume by hour

```powershell
# Full history
Get-VBDNSLogStatistics -DatabasePath $db -ReportType Timeline | Format-Table -AutoSize

# Last 48 hours
Get-VBDNSLogStatistics -DatabasePath $db -ReportType Timeline `
    -DateFrom (Get-Date).AddDays(-2) | Format-Table -AutoSize
```

**Columns:** `Hour`, `TotalPackets`, `Queries`, `Responses`, `Errors`

**Use case:** Identify traffic spikes, overnight activity (potential malware beaconing),
or unusually quiet periods that might indicate a logging gap.

---

### 7. PrivateVsPublic — RFC1918 vs public IP split

```powershell
Get-VBDNSLogStatistics -DatabasePath $db -ReportType PrivateVsPublic | Format-Table -AutoSize
```

**Columns:** `AddressType`, `Count`, `Percentage`

**Use case:** Quick sanity check — on an internal DNS server most traffic should be private IPs.
A significant public IP percentage warrants investigation.

---

### 8. DirectionSplit — Received vs Sent packet split

```powershell
Get-VBDNSLogStatistics -DatabasePath $db -ReportType DirectionSplit | Format-Table -AutoSize
```

**Columns:** `Direction`, `Count`, `Percentage`

**Use case:** Confirm logging configuration. A healthy bidirectional log shows roughly equal
Rcv and Snd counts. All Rcv and no Snd means only inbound queries are being logged.

---

### 9. ImportSummary — Audit of what has been imported

```powershell
Get-VBDNSLogStatistics -DatabasePath $db -ReportType ImportSummary | Format-Table -AutoSize
```

**Columns:** `FileName`, `FilePath`, `RecordCount`, `ErrorCount`, `DurationSeconds`, `ImportedAt`,
`ServerName`, `FileHash`

**Use case:** Verify which log files are loaded, when they were imported, how many records each
contributed, and detect any with high parse error counts.

---

## Get-VBDNSLog — Filtered Record Queries

```powershell
# All NXDOMAIN responses in the last 24 hours
Get-VBDNSLog -DatabasePath $db -ResponseCode NXDOMAIN -DateFrom (Get-Date).AddDays(-1)

# All queries from a specific IP
Get-VBDNSLog -DatabasePath $db -IPAddress '10.209.1.165' -PacketKind Q

# All AAAA queries matching a domain pattern
Get-VBDNSLog -DatabasePath $db -QueryType AAAA -QueryName '%microsoft.com%'

# Errors from public IPs in the last 7 days, limit 5000
Get-VBDNSLog -DatabasePath $db -Status Error -ExcludePrivateIPs `
    -DateFrom (Get-Date).AddDays(-7) -Limit 5000
```

---

## Invoke-VBDNSLogQuery — Raw SQL

```powershell
# Top NXDOMAIN domains (beaconing indicator)
Invoke-VBDNSLogQuery -DatabasePath $db -Query @"
    SELECT QueryName, COUNT(*) AS Hits, COUNT(DISTINCT IPAddress) AS UniqueClients
    FROM DNSLog
    WHERE ResponseCode = 'NXDOMAIN'
    GROUP BY QueryName
    ORDER BY Hits DESC
    LIMIT 20
"@

# Always use -Parameters for variable values
$ip = '10.209.1.165'
Invoke-VBDNSLogQuery -DatabasePath $db `
    -Query  "SELECT QueryName, COUNT(*) AS Hits FROM DNSLog WHERE IPAddress = @ip AND PacketKind = 'Q' GROUP BY QueryName ORDER BY Hits DESC LIMIT 20" `
    -Parameters @{ ip = $ip }
```

---

## All Reports — Export Everything in One Run

Runs all 9 report types and exports each to both CSV and XLSX. Adjust paths as needed.

```powershell
$db      = 'C:\Realtime\dns_analysis.db'
$outDir  = 'C:\Reports\DNS'
$date    = Get-Date -Format 'yyyy-MM-dd'

# Create output directory if it doesn't exist
New-Item -Path $outDir -ItemType Directory -Force | Out-Null

# ── Helper: export one report to CSV and XLSX ──────────────────────────────
function Export-Report {
    param($Data, $Name)
    $base = "$outDir\${date}_${Name}"
    $Data | Export-VBDNSLogReport -OutputPath "$base.csv"
    $Data | Export-VBDNSLogReport -OutputPath "$base.xlsx" -Format XLSX -Title $Name
    Write-Host "  Exported: $Name ($($Data.Count) rows)" -ForegroundColor Green
}

Write-Host "`nDNS Log Analysis — $date" -ForegroundColor Cyan
Write-Host "Database : $db"
Write-Host "Output   : $outDir`n"

# 1. TopTalkers
Export-Report -Name 'TopTalkers' -Data (
    Get-VBDNSLogStatistics -DatabasePath $db -ReportType TopTalkers
)

# 2. TopDomains
Export-Report -Name 'TopDomains' -Data (
    Get-VBDNSLogStatistics -DatabasePath $db -ReportType TopDomains
)

# 3. TalkerDetail
Export-Report -Name 'TalkerDetail' -Data (
    Get-VBDNSLogStatistics -DatabasePath $db -ReportType TalkerDetail
)

# 4. QueryTypeBreakdown
Export-Report -Name 'QueryTypeBreakdown' -Data (
    Get-VBDNSLogStatistics -DatabasePath $db -ReportType QueryTypeBreakdown
)

# 5. ErrorRate
Export-Report -Name 'ErrorRate' -Data (
    Get-VBDNSLogStatistics -DatabasePath $db -ReportType ErrorRate
)

# 6. Timeline (last 7 days)
Export-Report -Name 'Timeline_7Days' -Data (
    Get-VBDNSLogStatistics -DatabasePath $db -ReportType Timeline `
        -DateFrom (Get-Date).AddDays(-7)
)

# 7. PrivateVsPublic
Export-Report -Name 'PrivateVsPublic' -Data (
    Get-VBDNSLogStatistics -DatabasePath $db -ReportType PrivateVsPublic
)

# 8. DirectionSplit
Export-Report -Name 'DirectionSplit' -Data (
    Get-VBDNSLogStatistics -DatabasePath $db -ReportType DirectionSplit
)

# 9. ImportSummary
Export-Report -Name 'ImportSummary' -Data (
    Get-VBDNSLogStatistics -DatabasePath $db -ReportType ImportSummary
)

Write-Host "`nAll reports exported to: $outDir" -ForegroundColor Cyan
Get-ChildItem $outDir -Filter "${date}_*" | Select-Object Name, Length, LastWriteTime | Format-Table -AutoSize
```

**Output files created:**
```
2026-05-07_TopTalkers.csv / .xlsx
2026-05-07_TopDomains.csv / .xlsx
2026-05-07_TalkerDetail.csv / .xlsx
2026-05-07_QueryTypeBreakdown.csv / .xlsx
2026-05-07_ErrorRate.csv / .xlsx
2026-05-07_Timeline_7Days.csv / .xlsx
2026-05-07_PrivateVsPublic.csv / .xlsx
2026-05-07_DirectionSplit.csv / .xlsx
2026-05-07_ImportSummary.csv / .xlsx
```

---

## DNS Debug Logging — Enable Full Capture on the Server

If your log only shows `PacketKind = Q` (queries only), responses are not being logged.
Enable full bidirectional logging on the DNS server:

```powershell
# On the DNS server (requires DnsServer module)
Set-DnsServerDiagnostics -All $true

# Or via dnscmd
dnscmd /config /loglevel 0x8100F331
```

Or via DNS Manager → server Properties → Debug Logging → tick **Incoming** + **Outgoing** + **Responses**.

---

*VB.WindowsDNSLogAnalysis Usage Guide v1.1 — 2026-05-07 | Author: VB + Claude*
