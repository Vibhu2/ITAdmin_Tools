# ============================================================
# FUNCTION : Invoke-VBDNSLogQuery
# VERSION  : 1.0.0
# CHANGED  : 2026-05-07 -- Initial build
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Execute raw SQL queries against the DNSLogDB database
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Executes a raw SQL query against the DNSLogDB SQLite database and returns a DataTable.

.DESCRIPTION
    The escape hatch for anything Get-VBDNSLog does not cover. Executes the provided SQL
    string directly against the database and returns the result as a System.Data.DataTable,
    which can be piped to Export-Csv, Out-GridView, ConvertTo-Json, or Export-VBDNSLogReport.

    Always use -Parameters for values that come from variables to prevent SQL injection.

.PARAMETER DatabasePath
    Full path to the SQLite .db file.

.PARAMETER Query
    SQL query string to execute. Use @name placeholders for parameterised values.

.PARAMETER Parameters
    Hashtable of named parameter values. Keys match @name placeholders in -Query
    (without the @ prefix).

.EXAMPLE
    Invoke-VBDNSLogQuery -DatabasePath 'C:\DNS\dns_analysis.db' -Query @"
        SELECT IPAddress, COUNT(*) AS QueryCount
        FROM DNSLog
        WHERE PacketKind = 'Q'
        GROUP BY IPAddress
        ORDER BY QueryCount DESC
        LIMIT 10
    "@

.EXAMPLE
    Invoke-VBDNSLogQuery -DatabasePath 'C:\DNS\dns_analysis.db' `
        -Query "SELECT * FROM DNSLog WHERE IPAddress = @ip AND QueryType = @qt" `
        -Parameters @{ ip = '10.150.1.60'; qt = 'AAAA' }

.EXAMPLE
    Invoke-VBDNSLogQuery -DatabasePath 'C:\DNS\dns_analysis.db' -Query @"
        SELECT QueryName, COUNT(*) AS Hits
        FROM DNSLog WHERE ResponseCode = 'NXDOMAIN'
        GROUP BY QueryName ORDER BY Hits DESC LIMIT 20
    "@ | Export-VBDNSLogReport -OutputPath 'C:\Reports\NXDomains.csv'

.OUTPUTS
    [System.Data.DataTable] Result set from the executed query.

.NOTES
    Version  : 1.0.0
    Author   : Vibhu Bhatnagar
    Modified : 2026-05-07
    Category : Public

    Requires: PSSQLite module (Install-Module PSSQLite)
#>

function Invoke-VBDNSLogQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath,

        [Parameter(Mandatory = $true)]
        [string]$Query,

        [hashtable]$Parameters = @{}
    )

    if (-not (Test-Path $DatabasePath)) {
        throw "Invoke-VBDNSLogQuery: Database not found at '$DatabasePath'."
    }

    Write-Verbose "Invoke-VBDNSLogQuery: Executing query against '$DatabasePath'"

    $invokeParams = @{
        DataSource = $DatabasePath
        Query      = $Query
    }

    if ($Parameters.Count -gt 0) {
        $invokeParams.SqlParameters = $Parameters
    }

    $result = Invoke-SqliteQuery @invokeParams

    Write-Verbose "Invoke-VBDNSLogQuery: Returned $($result.Count) row(s)"

    return $result
}
