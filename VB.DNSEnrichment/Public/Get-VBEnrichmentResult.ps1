function Get-VBEnrichmentResult {
<#
.SYNOPSIS
    Read enrichment rows from the SQLite database.

.DESCRIPTION
    Returns enrichment rows from the Enrichment table as PSCustomObject arrays.
    All parameters are optional filters -- omit all to return every row.

    Used by reports and by the orchestrator on subsequent runs to determine
    which IPs already have fresh enrichment data.

.PARAMETER IPAddress
    Filter to one or more specific IP addresses.

.PARAMETER DeviceClass
    Filter by device class (e.g. 'Printer', 'Camera', 'Workstation').

.PARAMETER Since
    Return only rows where UpdatedAt >= this datetime.

.PARAMETER UnresolvedOnly
    Return only rows where IsResolved = 0 (hostname could not be determined).

.PARAMETER IncludeHistory
    Also return associated EnrichmentHistory rows in a History property on each result.

.PARAMETER Context
    Environment context from Get-VBEnrichmentContext. Provides DatabasePath.

.PARAMETER DatabasePath
    Override the database path directly (alternative to supplying Context).

.OUTPUTS
    [PSCustomObject[]] -- one object per matching row.

.EXAMPLE
    $ctx = Get-VBEnrichmentContext
    Get-VBEnrichmentResult -Context $ctx

.EXAMPLE
    Get-VBEnrichmentResult -DeviceClass 'Printer' -Context $ctx

.EXAMPLE
    Get-VBEnrichmentResult -IPAddress '192.168.1.45','192.168.1.46' -Context $ctx

.NOTES
    Version:      1.0.0
    MinPSVersion: 5.1
    Author:       VB
    ChangeLog:
        1.0.0 -- 2026-05-11 -- Initial release
#>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter()]
        [string[]]$IPAddress,

        [Parameter()]
        [string]$DeviceClass,

        [Parameter()]
        [datetime]$Since,

        [Parameter()]
        [switch]$UnresolvedOnly,

        [Parameter()]
        [switch]$IncludeHistory,

        [Parameter()]
        [PSCustomObject]$Context,

        [Parameter()]
        [string]$DatabasePath
    )

    # Resolve database path
    $dbPath = $DatabasePath
    if (-not $dbPath -and $Context) { $dbPath = $Context.DatabasePath }
    if (-not $dbPath) { $dbPath = Join-Path $env:LOCALAPPDATA 'VB.DNSEnrichment\enrichment.db' }

    if (-not (Test-Path -LiteralPath $dbPath)) {
        Write-Warning "[Get-VBEnrichmentResult] Database not found at $dbPath -- run Initialize-VBEnrichmentDatabase first."
        return
    }

    try {
        # Build WHERE clauses dynamically
        $whereClauses = New-Object System.Collections.Generic.List[string]
        $params       = @{}

        if ($IPAddress -and $IPAddress.Count -gt 0) {
            $placeholders = ($IPAddress | ForEach-Object { $i = $params.Count; $params["p$i"] = $_; "@p$i" }) -join ','
            $whereClauses.Add("IPAddress IN ($placeholders)")
        }

        if ($DeviceClass) {
            $params['dc'] = $DeviceClass
            $whereClauses.Add('DeviceClass = @dc')
        }

        if ($Since) {
            $params['since'] = $Since.ToString('o')
            $whereClauses.Add('UpdatedAt >= @since')
        }

        if ($UnresolvedOnly) {
            $whereClauses.Add('IsResolved = 0')
        }

        $where = if ($whereClauses.Count -gt 0) { 'WHERE ' + ($whereClauses -join ' AND ') } else { '' }
        $sql   = "SELECT * FROM Enrichment $where ORDER BY IPAddress"

        $sqlParams = @{}
        foreach ($key in $params.Keys) { $sqlParams[$key] = $params[$key] }

        $rows = Invoke-VBSqliteCommand -DatabasePath $dbPath -Query $sql -SqlParameters $sqlParams

        foreach ($row in $rows) {
            $obj = ConvertFrom-VBSqliteEnrichmentRow -Row $row -FromCache $true

            if ($IncludeHistory) {
                try {
                    $history = Invoke-VBSqliteCommand -DatabasePath $dbPath `
                        -Query 'SELECT * FROM EnrichmentHistory WHERE IPAddress = @ip ORDER BY ChangedAt DESC' `
                        -SqlParameters @{ ip = $row.IPAddress }
                    $obj | Add-Member -NotePropertyName 'History' -NotePropertyValue $history -Force
                }
                catch {
                    $obj | Add-Member -NotePropertyName 'History' -NotePropertyValue @() -Force
                }
            }

            $obj
        }
    }
    catch {
        Write-Warning "[Get-VBEnrichmentResult] Query failed: $($_.Exception.Message)"
    }
}
