function Invoke-VBSqliteCommand {
<#
.SYNOPSIS
    Thin wrapper over PSSQLite's Invoke-SqliteQuery for VB.DNSEnrichment.

.DESCRIPTION
    Centralises SQLite access so every call uses the same connection options,
    parameter passing convention, and error handling. The wrapper is intentionally
    minimal -- it does not pool connections or open transactions; PSSQLite manages
    the connection lifecycle per call.

    Two modes:
        -Query        Execute SQL and return rows ([PSCustomObject[]])
        -NonQuery     Execute SQL with no result set; returns rows-affected ([int])

.PARAMETER DatabasePath
    Full path to the SQLite database file.

.PARAMETER Query
    SQL text to execute.

.PARAMETER SqlParameters
    Hashtable of named parameters referenced as @name in the SQL text.

.PARAMETER NonQuery
    Switch -- treat the statement as INSERT/UPDATE/DELETE/DDL and return the
    rows-affected count rather than result rows.

.OUTPUTS
    [PSCustomObject[]] when -NonQuery is omitted.
    [int]              when -NonQuery is supplied.

.EXAMPLE
    Invoke-VBSqliteCommand -DatabasePath $db -Query 'SELECT * FROM Enrichment WHERE IPAddress = @ip' `
        -SqlParameters @{ ip = '192.168.1.45' }

.EXAMPLE
    Invoke-VBSqliteCommand -DatabasePath $db -NonQuery `
        -Query 'INSERT INTO SchemaVersion (Version, AppliedAt) VALUES (@v, @t)' `
        -SqlParameters @{ v = 1; t = (Get-Date).ToString('o') }

.NOTES
    Version:      1.0.0
    MinPSVersion: 5.1
    Author:       VB
    ChangeLog:
        1.0.0 -- 2026-05-10 -- Initial release
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DatabasePath,

        [Parameter(Mandatory)]
        [string]$Query,

        [Parameter()]
        [hashtable]$SqlParameters = @{},

        [Parameter()]
        [switch]$NonQuery
    )

    if (-not (Get-Module -Name PSSQLite)) {
        Import-Module PSSQLite -ErrorAction Stop
    }

    $splat = @{
        DataSource = $DatabasePath
        Query      = $Query
    }

    if ($SqlParameters.Count -gt 0) {
        $splat['SqlParameters'] = $SqlParameters
    }

    if ($NonQuery) {
        # PSSQLite returns affected-rows count when the statement is non-SELECT
        Invoke-SqliteQuery @splat
    }
    else {
        Invoke-SqliteQuery @splat
    }
}
