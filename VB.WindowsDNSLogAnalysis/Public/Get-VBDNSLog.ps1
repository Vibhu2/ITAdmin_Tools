# ============================================================
# FUNCTION : Get-VBDNSLog
# VERSION  : 1.0.0
# CHANGED  : 2026-05-07 -- Initial build
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Query the DNSLog table with named filter parameters
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Queries the DNSLog table with common operational filters and returns matching rows.

.DESCRIPTION
    Builds a parameterised SELECT query against DNSLog based on the provided parameters
    and returns matching rows as PSCustomObject[]. All filtering happens in SQLite --
    no PowerShell-side post-filtering after the query returns.

    For anything requiring full SQL (aggregations, GROUP BY, complex joins), use
    Invoke-VBDNSLogQuery instead.

.PARAMETER DatabasePath
    Full path to the SQLite .db file.

.PARAMETER IPAddress
    Filter to a specific client IP address (exact match).

.PARAMETER QueryName
    Filter by domain name. Supports SQL LIKE patterns with % wildcard.

.PARAMETER QueryType
    Filter by DNS record type (A, AAAA, MX, PTR, SOA, NS, CNAME, TXT, SRV, ANY...).

.PARAMETER Protocol
    Filter by transport protocol. Values: UDP, TCP, All (default: All).

.PARAMETER Direction
    Filter by packet direction. Values: Rcv, Snd, All (default: All).

.PARAMETER PacketKind
    Filter by packet kind. Values: Q (query), R (response), All (default: All).

.PARAMETER ResponseCode
    Filter by DNS response code (NOERROR, NXDOMAIN, SERVFAIL, REFUSED, FORMERR, NOTIMPL).

.PARAMETER Status
    Filter by derived status. Values: Success, Error.

.PARAMETER DateFrom
    Start of time window (inclusive). Compared against LogDateTime.

.PARAMETER DateTo
    End of time window (inclusive). Compared against LogDateTime.

.PARAMETER ExcludePrivateIPs
    Exclude rows where IsPrivate = 1.

.PARAMETER SourceFile
    Limit results to records imported from a specific log file (full path).

.PARAMETER Limit
    Maximum number of rows to return. Default: 10,000.

.EXAMPLE
    Get-VBDNSLog -DatabasePath 'C:\DNS\dns_analysis.db' -ResponseCode NXDOMAIN -DateFrom (Get-Date).AddDays(-1)

.EXAMPLE
    Get-VBDNSLog -DatabasePath 'C:\DNS\dns_analysis.db' -IPAddress '10.150.1.60' -PacketKind Q

.EXAMPLE
    Get-VBDNSLog -DatabasePath 'C:\DNS\dns_analysis.db' -QueryType AAAA -QueryName '%realtime-it.com%'

.OUTPUTS
    [PSCustomObject[]] Rows from the DNSLog table matching the specified filters.

.NOTES
    Version  : 1.0.0
    Author   : Vibhu Bhatnagar
    Modified : 2026-05-07
    Category : Public

    Requires: PSSQLite module (Install-Module PSSQLite)
#>

function Get-VBDNSLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath,

        [string]$IPAddress,

        [string]$QueryName,

        [string]$QueryType,

        [ValidateSet('UDP', 'TCP', 'All')]
        [string]$Protocol = 'All',

        [ValidateSet('Rcv', 'Snd', 'All')]
        [string]$Direction = 'All',

        [ValidateSet('Q', 'R', 'All')]
        [string]$PacketKind = 'All',

        [string]$ResponseCode,

        [ValidateSet('Success', 'Error')]
        [string]$Status,

        [datetime]$DateFrom,

        [datetime]$DateTo,

        [switch]$ExcludePrivateIPs,

        [string]$SourceFile,

        [int]$Limit = 10000
    )

    if (-not (Test-Path $DatabasePath)) {
        throw "Get-VBDNSLog: Database not found at '$DatabasePath'."
    }

    $whereClauses = [System.Collections.Generic.List[string]]::new()
    $sqlParams    = @{}

    if ($PSBoundParameters.ContainsKey('IPAddress')) {
        $whereClauses.Add('IPAddress = @IPAddress')
        $sqlParams['IPAddress'] = $IPAddress
    }

    if ($PSBoundParameters.ContainsKey('QueryName')) {
        if ($QueryName -like '*%*') {
            $whereClauses.Add('QueryName LIKE @QueryName')
        } else {
            $whereClauses.Add('QueryName = @QueryName')
        }
        $sqlParams['QueryName'] = $QueryName
    }

    if ($PSBoundParameters.ContainsKey('QueryType')) {
        $whereClauses.Add('QueryType = @QueryType')
        $sqlParams['QueryType'] = $QueryType
    }

    if ($Protocol -ne 'All') {
        $whereClauses.Add('Protocol = @Protocol')
        $sqlParams['Protocol'] = $Protocol
    }

    if ($Direction -ne 'All') {
        $whereClauses.Add('Direction = @Direction')
        $sqlParams['Direction'] = $Direction
    }

    if ($PacketKind -ne 'All') {
        $whereClauses.Add('PacketKind = @PacketKind')
        $sqlParams['PacketKind'] = $PacketKind
    }

    if ($PSBoundParameters.ContainsKey('ResponseCode')) {
        $whereClauses.Add('ResponseCode = @ResponseCode')
        $sqlParams['ResponseCode'] = $ResponseCode
    }

    if ($PSBoundParameters.ContainsKey('Status')) {
        $whereClauses.Add('Status = @Status')
        $sqlParams['Status'] = $Status
    }

    if ($PSBoundParameters.ContainsKey('DateFrom')) {
        $whereClauses.Add('LogDateTime >= @DateFrom')
        $sqlParams['DateFrom'] = $DateFrom.ToString('yyyy-MM-ddTHH:mm:ss')
    }

    if ($PSBoundParameters.ContainsKey('DateTo')) {
        $whereClauses.Add('LogDateTime <= @DateTo')
        $sqlParams['DateTo'] = $DateTo.ToString('yyyy-MM-ddTHH:mm:ss')
    }

    if ($ExcludePrivateIPs) {
        $whereClauses.Add('IsPrivate = 0')
    }

    if ($PSBoundParameters.ContainsKey('SourceFile')) {
        $whereClauses.Add('SourceFile = @SourceFile')
        $sqlParams['SourceFile'] = $SourceFile
    }

    $whereStr = if ($whereClauses.Count -gt 0) {
        'WHERE ' + [string]::Join(' AND ', $whereClauses)
    } else { '' }

    $query = @"
SELECT
    Id, LogDateTime, LogDate, LogTime, ThreadId, PacketId,
    Protocol, Direction, IPAddress, IPVersion, IsPrivate,
    TransactionId, PacketKind, Opcode, FlagsHex, FlagsChar,
    ResponseCode, Status, Error, QueryType, QueryName,
    SourceFile, ImportedAt
FROM DNSLog
$whereStr
ORDER BY LogDateTime DESC
LIMIT @Limit
"@

    $sqlParams['Limit'] = $Limit

    Write-Verbose "Get-VBDNSLog: Executing query with $($whereClauses.Count) filter(s), LIMIT=$Limit"

    $results = Invoke-SqliteQuery -DataSource $DatabasePath -Query $query -SqlParameters $sqlParams

    Write-Verbose "Get-VBDNSLog: Returned $($results.Count) row(s)"

    return $results
}
