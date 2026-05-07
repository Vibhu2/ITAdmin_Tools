# ============================================================
# FUNCTION : Get-VBDNSLogStatistics
# VERSION  : 1.2.0
# CHANGED  : 2026-05-07 -- Initial build
#            2026-05-07 -- Fix ResponseCode scoping to R rows only
#            2026-05-07 -- Remove -All/-IncludeSummary; -Top optional;
#                          summary always shown for TopTalkers/TopDomains
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Pre-built SQL report queries for common DNS traffic analysis
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Executes a pre-built statistical report query against the DNSLog database.

.DESCRIPTION
    Provides 8 named report types covering the most common first-look questions
    asked of a DNS dataset. Each report type runs a self-contained SQL query and
    returns the result as PSCustomObject[].

    For TopTalkers and TopDomains:
        - Omit -Top to return ALL results (no limit)
        - Specify -Top N to limit to the top N rows
        - A summary table is always printed to the console after the data rows.
          The summary never enters the pipeline -- it is safe to pipe results
          to Export-VBDNSLogReport or Format-Table without interference.

    Report types:
        TopTalkers          -- Client IPs ranked by query count
        TopDomains          -- Domain names ranked by query count
        TalkerDetail        -- Per-IP breakdown: query types, protocol split, top domain
        QueryTypeBreakdown  -- Count and percentage per QueryType (A, AAAA, MX...)
        ErrorRate           -- Count and percentage per ResponseCode (responses only)
        Timeline            -- Query count grouped by hour
        PrivateVsPublic     -- Row counts split by IsPrivate flag
        DirectionSplit      -- Rcv vs Snd packet counts
        ImportSummary       -- One row per imported file from ImportLog

.PARAMETER DatabasePath
    Full path to the SQLite .db file.

.PARAMETER ReportType
    The pre-built report to run. See description for available types.

.PARAMETER Top
    Limit results to the top N rows for TopTalkers and TopDomains.
    Omit this parameter to return all results with no limit.

.PARAMETER DateFrom
    Restrict analysis to records with LogDateTime >= this value.

.PARAMETER DateTo
    Restrict analysis to records with LogDateTime <= this value.

.EXAMPLE
    # All IPs, no limit, with summary
    Get-VBDNSLogStatistics -DatabasePath 'C:\DNS\dns.db' -ReportType TopTalkers | Format-Table -AutoSize

.EXAMPLE
    # Top 10 IPs only, with summary
    Get-VBDNSLogStatistics -DatabasePath 'C:\DNS\dns.db' -ReportType TopTalkers -Top 10 | Format-Table -AutoSize

.EXAMPLE
    # Export all domains to CSV -- summary prints to screen only, not in CSV
    Get-VBDNSLogStatistics -DatabasePath 'C:\DNS\dns.db' -ReportType TopDomains |
        Export-VBDNSLogReport -OutputPath 'C:\Reports\TopDomains.csv'

.EXAMPLE
    Get-VBDNSLogStatistics -DatabasePath 'C:\DNS\dns.db' -ReportType ErrorRate

.EXAMPLE
    Get-VBDNSLogStatistics -DatabasePath 'C:\DNS\dns.db' -ReportType Timeline `
        -DateFrom (Get-Date).AddDays(-7)

.OUTPUTS
    [PSCustomObject[]] Result rows from the selected report query.
    Summary (TopTalkers/TopDomains only) is written to the console host and
    never enters the PowerShell pipeline.

.NOTES
    Version  : 1.2.0
    Author   : Vibhu Bhatnagar
    Modified : 2026-05-07
    Category : Public

    Requires: PSSQLite module (Install-Module PSSQLite)
#>

function Get-VBDNSLogStatistics {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath,

        [Parameter(Mandatory = $true)]
        [ValidateSet(
            'TopTalkers',
            'TopDomains',
            'TalkerDetail',
            'QueryTypeBreakdown',
            'ErrorRate',
            'Timeline',
            'PrivateVsPublic',
            'DirectionSplit',
            'ImportSummary'
        )]
        [string]$ReportType,

        # Optional -- omit to return all rows (no LIMIT)
        [int]$Top,

        [datetime]$DateFrom,

        [datetime]$DateTo
    )

    if (-not (Test-Path $DatabasePath)) {
        throw "Get-VBDNSLogStatistics: Database not found at '$DatabasePath'."
    }

    # Build time window WHERE clause and parameter hashtable
    $timeFilters = [System.Collections.Generic.List[string]]::new()
    $sqlParams   = @{}

    if ($PSBoundParameters.ContainsKey('DateFrom')) {
        $timeFilters.Add('LogDateTime >= @DateFrom')
        $sqlParams['DateFrom'] = $DateFrom.ToString('yyyy-MM-ddTHH:mm:ss')
    }
    if ($PSBoundParameters.ContainsKey('DateTo')) {
        $timeFilters.Add('LogDateTime <= @DateTo')
        $sqlParams['DateTo'] = $DateTo.ToString('yyyy-MM-ddTHH:mm:ss')
    }

    $timeWhere = if ($timeFilters.Count -gt 0) {
        'WHERE ' + [string]::Join(' AND ', $timeFilters)
    } else { '' }

    # LIMIT clause -- only applied when -Top is explicitly provided
    $limitClause = if ($PSBoundParameters.ContainsKey('Top')) {
        $sqlParams['Top'] = $Top
        'LIMIT @Top'
    } else { '' }

    # ----------------------------------------------------------------
    # Build report query
    # ----------------------------------------------------------------
    $query = switch ($ReportType) {

        'TopTalkers' {
@"
SELECT
    IPAddress,
    IPVersion,
    IsPrivate,
    SUM(CASE WHEN PacketKind = 'Q' THEN 1 ELSE 0 END)                                        AS Queries,
    SUM(CASE WHEN PacketKind = 'R' THEN 1 ELSE 0 END)                                        AS Responses,
    SUM(CASE WHEN PacketKind = 'R' AND ResponseCode = 'NXDOMAIN' THEN 1 ELSE 0 END)          AS NXDOMAIN,
    SUM(CASE WHEN PacketKind = 'R' AND ResponseCode NOT IN ('NOERROR','NXDOMAIN')
             AND ResponseCode != ''                               THEN 1 ELSE 0 END)          AS OtherErrors
FROM DNSLog
$timeWhere
GROUP BY IPAddress
ORDER BY Queries DESC
$limitClause
"@
        }

        'TopDomains' {
@"
SELECT
    QueryName,
    SUM(CASE WHEN PacketKind = 'Q' THEN 1 ELSE 0 END)                                        AS Queries,
    SUM(CASE WHEN PacketKind = 'R' THEN 1 ELSE 0 END)                                        AS Responses,
    SUM(CASE WHEN PacketKind = 'R' AND ResponseCode = 'NXDOMAIN' THEN 1 ELSE 0 END)          AS NXDOMAIN,
    SUM(CASE WHEN PacketKind = 'R' AND ResponseCode = 'NOERROR'  THEN 1 ELSE 0 END)          AS NOERROR,
    SUM(CASE WHEN PacketKind = 'R' AND ResponseCode NOT IN ('NOERROR','NXDOMAIN')
             AND ResponseCode != ''                               THEN 1 ELSE 0 END)          AS OtherErrors
FROM DNSLog
$timeWhere
GROUP BY QueryName
ORDER BY Queries DESC
$limitClause
"@
        }

        'TalkerDetail' {
            # Build the inner time filter for the TopDomain subquery using the same params.
            # The outer $timeWhere uses named params (@DateFrom/@DateTo) which are already
            # in $sqlParams -- the correlated subquery can reference them the same way.
            $innerWhere = if ($timeWhere) {
                $timeWhere -replace 'WHERE ', 'AND '
            } else { '' }
@"
SELECT
    IPAddress,
    IPVersion,
    IsPrivate,
    SUM(CASE WHEN PacketKind = 'Q'                        THEN 1 ELSE 0 END) AS Queries,
    SUM(CASE WHEN PacketKind = 'Q' AND QueryType = 'A'    THEN 1 ELSE 0 END) AS A_Queries,
    SUM(CASE WHEN PacketKind = 'Q' AND QueryType = 'AAAA' THEN 1 ELSE 0 END) AS AAAA_Queries,
    SUM(CASE WHEN PacketKind = 'Q' AND QueryType = 'PTR'  THEN 1 ELSE 0 END) AS PTR_Queries,
    SUM(CASE WHEN PacketKind = 'Q' AND QueryType = 'MX'   THEN 1 ELSE 0 END) AS MX_Queries,
    SUM(CASE WHEN PacketKind = 'Q' AND QueryType = 'SRV'  THEN 1 ELSE 0 END) AS SRV_Queries,
    SUM(CASE WHEN PacketKind = 'Q' AND QueryType = 'TXT'  THEN 1 ELSE 0 END) AS TXT_Queries,
    SUM(CASE WHEN PacketKind = 'Q' AND QueryType NOT IN ('A','AAAA','PTR','MX','SRV','TXT')
                                                          THEN 1 ELSE 0 END) AS Other_Queries,
    SUM(CASE WHEN PacketKind = 'Q' AND Protocol  = 'UDP'  THEN 1 ELSE 0 END) AS UDP_Queries,
    SUM(CASE WHEN PacketKind = 'Q' AND Protocol  = 'TCP'  THEN 1 ELSE 0 END) AS TCP_Queries,
    (
        SELECT d2.QueryName
        FROM   DNSLog d2
        WHERE  d2.IPAddress  = DNSLog.IPAddress
          AND  d2.PacketKind = 'Q'
          $innerWhere
        GROUP BY d2.QueryName
        ORDER BY COUNT(*) DESC
        LIMIT 1
    ) AS TopDomain
FROM DNSLog
$timeWhere
GROUP BY IPAddress
ORDER BY Queries DESC
$limitClause
"@
        }

        'QueryTypeBreakdown' {
@"
SELECT
    QueryType,
    COUNT(*) AS Count,
    ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) AS Percentage
FROM DNSLog
$timeWhere
GROUP BY QueryType
ORDER BY Count DESC
"@
        }

        'ErrorRate' {
@"
SELECT
    ResponseCode,
    COUNT(*) AS Count,
    ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) AS Percentage
FROM DNSLog
$(if ($timeWhere) { $timeWhere + ' AND ' } else { 'WHERE ' })PacketKind = 'R'
GROUP BY ResponseCode
ORDER BY Count DESC
"@
        }

        'Timeline' {
@"
SELECT
    STRFTIME('%Y-%m-%dT%H:00:00', LogDateTime) AS Hour,
    COUNT(*) AS TotalPackets,
    SUM(CASE WHEN PacketKind = 'Q' THEN 1 ELSE 0 END) AS Queries,
    SUM(CASE WHEN PacketKind = 'R' THEN 1 ELSE 0 END) AS Responses,
    SUM(CASE WHEN PacketKind = 'R' AND ResponseCode != 'NOERROR'
             AND ResponseCode != ''                    THEN 1 ELSE 0 END) AS Errors
FROM DNSLog
$timeWhere
GROUP BY Hour
ORDER BY Hour ASC
"@
        }

        'PrivateVsPublic' {
@"
SELECT
    CASE IsPrivate WHEN 1 THEN 'Private' ELSE 'Public' END AS AddressType,
    COUNT(*) AS Count,
    ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) AS Percentage
FROM DNSLog
$timeWhere
GROUP BY IsPrivate
ORDER BY Count DESC
"@
        }

        'DirectionSplit' {
@"
SELECT
    Direction,
    COUNT(*) AS Count,
    ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) AS Percentage
FROM DNSLog
$timeWhere
GROUP BY Direction
ORDER BY Count DESC
"@
        }

        'ImportSummary' {
@"
SELECT
    FileName,
    FilePath,
    RecordCount,
    ErrorCount,
    DurationSeconds,
    ImportedAt,
    ServerName,
    FileHash
FROM ImportLog
ORDER BY ImportedAt DESC
"@
        }
    }

    Write-Verbose "Get-VBDNSLogStatistics: Running '$ReportType' report"

    $invokeParams = @{ DataSource = $DatabasePath; Query = $query }
    if ($sqlParams.Count -gt 0) { $invokeParams.SqlParameters = $sqlParams }

    $results = Invoke-SqliteQuery @invokeParams

    Write-Verbose "Get-VBDNSLogStatistics: '$ReportType' returned $($results.Count) row(s)"

    # ----------------------------------------------------------------
    # Emit data rows into the pipeline.
    # All report types are fully composable -- pipe to Format-Table,
    # Export-VBDNSLogReport, or any other cmdlet freely.
    # For TopTalkers/TopDomains the summary prints via Write-Host
    # after the rows are emitted. Because Format-Table buffers before
    # rendering, the summary will appear above the table when piped
    # inline -- this is a PS pipeline limitation and is acceptable.
    # ----------------------------------------------------------------
    $results

    if ($ReportType -notin 'TopTalkers', 'TopDomains') { return }

    # Summary query -- same time window, no LIMIT
    $summaryParams = @{}
    if ($PSBoundParameters.ContainsKey('DateFrom')) {
        $summaryParams['DateFrom'] = $DateFrom.ToString('yyyy-MM-ddTHH:mm:ss')
    }
    if ($PSBoundParameters.ContainsKey('DateTo')) {
        $summaryParams['DateTo'] = $DateTo.ToString('yyyy-MM-ddTHH:mm:ss')
    }

    if ($ReportType -eq 'TopTalkers') {

        $s = (Invoke-SqliteQuery -DataSource $DatabasePath -SqlParameters $summaryParams -Query @"
SELECT
    COUNT(DISTINCT IPAddress)                                  AS UniqueIPs,
    SUM(CASE WHEN IsPrivate = 1 THEN 1 ELSE 0 END)             AS PrivateIPs,
    SUM(CASE WHEN IsPrivate = 0 THEN 1 ELSE 0 END)             AS PublicIPs,
    SUM(CASE WHEN PacketKind = 'Q' THEN 1 ELSE 0 END)          AS TotalQueries,
    SUM(CASE WHEN PacketKind = 'R' THEN 1 ELSE 0 END)          AS TotalResponses
FROM DNSLog
$timeWhere
"@)[0]

        $shown    = $results.Count
        $limitMsg = if ($PSBoundParameters.ContainsKey('Top')) { "(limited to top $Top)" } else { '(all IPs)' }

        Write-Host ''
        Write-Host ''
        Write-Host ('  {0}' -f ('─' * 52)) -ForegroundColor DarkGray
        Write-Host ('  {0,-30}  {1}' -f 'TopTalkers Summary', '') -ForegroundColor Cyan
        Write-Host ('  {0}' -f ('─' * 52)) -ForegroundColor DarkGray
        Write-Host ('  {0,-30}  {1,10}' -f 'Unique IP addresses',  $s.UniqueIPs.ToString('N0'))
        Write-Host ('  {0,-30}  {1,10}' -f '  Private (RFC1918)',   $s.PrivateIPs.ToString('N0'))
        Write-Host ('  {0,-30}  {1,10}' -f '  Public',             $s.PublicIPs.ToString('N0'))
        Write-Host ('  {0,-30}  {1,10}' -f 'Total queries (Q)',    $s.TotalQueries.ToString('N0'))
        Write-Host ('  {0,-30}  {1,10}' -f 'Total responses (R)',  $s.TotalResponses.ToString('N0'))
        Write-Host ('  {0,-30}  {1,10}' -f "Rows shown $limitMsg", $shown.ToString('N0'))
        Write-Host ('  {0}' -f ('─' * 52)) -ForegroundColor DarkGray
        Write-Host ''

    } elseif ($ReportType -eq 'TopDomains') {

        $s = (Invoke-SqliteQuery -DataSource $DatabasePath -SqlParameters $summaryParams -Query @"
SELECT
    COUNT(DISTINCT QueryName)                                                          AS UniqueDomains,
    SUM(CASE WHEN PacketKind = 'Q' THEN 1 ELSE 0 END)                                 AS TotalQueries,
    SUM(CASE WHEN PacketKind = 'R' AND ResponseCode = 'NOERROR'  THEN 1 ELSE 0 END)   AS TotalNOERROR,
    SUM(CASE WHEN PacketKind = 'R' AND ResponseCode = 'NXDOMAIN' THEN 1 ELSE 0 END)   AS TotalNXDOMAIN
FROM DNSLog
$timeWhere
"@)[0]

        $shown    = $results.Count
        $limitMsg = if ($PSBoundParameters.ContainsKey('Top')) { "(limited to top $Top)" } else { '(all domains)' }

        Write-Host ''
        Write-Host ('  {0}' -f ('─' * 52)) -ForegroundColor DarkGray
        Write-Host ('  {0,-30}  {1}' -f 'TopDomains Summary', '') -ForegroundColor Cyan
        Write-Host ('  {0}' -f ('─' * 52)) -ForegroundColor DarkGray
        Write-Host ('  {0,-30}  {1,10}' -f 'Unique domains queried',  $s.UniqueDomains.ToString('N0'))
        Write-Host ('  {0,-30}  {1,10}' -f 'Total queries (Q)',       $s.TotalQueries.ToString('N0'))
        Write-Host ('  {0,-30}  {1,10}' -f 'Total NOERROR (R)',       $s.TotalNOERROR.ToString('N0'))
        Write-Host ('  {0,-30}  {1,10}' -f 'Total NXDOMAIN (R)',      $s.TotalNXDOMAIN.ToString('N0'))
        Write-Host ('  {0,-30}  {1,10}' -f "Rows shown $limitMsg",    $shown.ToString('N0'))
        Write-Host ('  {0}' -f ('─' * 52)) -ForegroundColor DarkGray
        Write-Host ''
    }
}
