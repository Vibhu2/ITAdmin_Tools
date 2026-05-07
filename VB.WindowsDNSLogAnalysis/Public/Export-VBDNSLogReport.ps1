# ============================================================
# FUNCTION : Export-VBDNSLogReport
# VERSION  : 1.0.0
# CHANGED  : 2026-05-07 -- Initial build
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Write query results to CSV, HTML, or XLSX report files
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Writes DNS log query results to a CSV, HTML, or XLSX report file.

.DESCRIPTION
    Accepts pipeline input from Get-VBDNSLog, Invoke-VBDNSLogQuery, or
    Get-VBDNSLogStatistics and writes the result set to a formatted file.

    Formats supported:
        CSV   -- Default. UTF-8 CSV. Compatible with Excel and most tools.
        HTML  -- Self-contained HTML table with inline CSS. Opens in any browser.
        XLSX  -- Excel workbook. Requires ImportExcel module (Install-Module ImportExcel).

.PARAMETER InputObject
    The result set to export. Accepts pipeline input from any query function.

.PARAMETER OutputPath
    Destination file path. The parent directory must exist.

.PARAMETER Format
    Output format: CSV (default), HTML, or XLSX.

.PARAMETER Title
    Optional report title for HTML page title/heading and Excel worksheet name.

.EXAMPLE
    Get-VBDNSLogStatistics -DatabasePath 'C:\DNS\dns_analysis.db' -ReportType TopDomains |
        Export-VBDNSLogReport -OutputPath 'C:\Reports\TopDomains.csv'

.EXAMPLE
    Get-VBDNSLog -DatabasePath 'C:\DNS\dns_analysis.db' -ResponseCode NXDOMAIN |
        Export-VBDNSLogReport -OutputPath 'C:\Reports\NXDomains.html' -Format HTML -Title 'NXDOMAIN Report'

.EXAMPLE
    Get-VBDNSLogStatistics -DatabasePath 'C:\DNS\dns_analysis.db' -ReportType TopTalkers |
        Export-VBDNSLogReport -OutputPath 'C:\Reports\TopTalkers.xlsx' -Format XLSX -Title 'Top Talkers'

.OUTPUTS
    None. Writes the output file and emits a Verbose message with row count and path.

.NOTES
    Version  : 1.0.0
    Author   : Vibhu Bhatnagar
    Modified : 2026-05-07
    Category : Public

    XLSX format requires: ImportExcel module (Install-Module ImportExcel)
    CSV and HTML have no additional dependencies.
#>

function Export-VBDNSLogReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [ValidateSet('CSV', 'HTML', 'XLSX')]
        [string]$Format = 'CSV',

        [string]$Title = 'DNS Log Report'
    )

    begin {
        $rows = [System.Collections.Generic.List[object]]::new()

        $outputDir = Split-Path $OutputPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($outputDir) -and -not (Test-Path $outputDir)) {
            throw "Export-VBDNSLogReport: Output directory does not exist: '$outputDir'"
        }

        if ($Format -eq 'XLSX') {
            if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
                throw "Export-VBDNSLogReport: XLSX format requires the ImportExcel module. Run: Install-Module ImportExcel"
            }
        }
    }

    process {
        foreach ($row in $InputObject) {
            $rows.Add($row)
        }
    }

    end {
        if ($rows.Count -eq 0) {
            Write-Warning "Export-VBDNSLogReport: No data to export."
            return
        }

        switch ($Format) {

            'CSV' {
                $rows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Force
                Write-Verbose "Export-VBDNSLogReport: CSV written -- $($rows.Count) rows to '$OutputPath'"
            }

            'HTML' {
                $generatedAt = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                $headers     = ($rows[0].PSObject.Properties.Name)
                $headerHtml  = ($headers | ForEach-Object { "<th>$_</th>" }) -join ''

                $rowsHtml = ($rows | ForEach-Object {
                    $row   = $_
                    $cells = ($headers | ForEach-Object { "<td>$($row.$_)</td>" }) -join ''
                    "<tr>$cells</tr>"
                }) -join "`n"

                $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$Title</title>
    <style>
        body   { font-family: Consolas, 'Courier New', monospace; font-size: 13px; margin: 24px; background: #f5f5f5; }
        h1     { color: #1a1a2e; font-size: 18px; margin-bottom: 4px; }
        p.meta { color: #666; font-size: 11px; margin-top: 0; margin-bottom: 16px; }
        table  { border-collapse: collapse; width: 100%; background: #fff; box-shadow: 0 1px 3px rgba(0,0,0,.15); }
        th     { background: #1a1a2e; color: #fff; padding: 8px 12px; text-align: left; font-size: 12px; }
        td     { padding: 6px 12px; border-bottom: 1px solid #e0e0e0; white-space: nowrap; }
        tr:nth-child(even) td { background: #f9f9f9; }
        tr:hover td { background: #e8f0fe; }
    </style>
</head>
<body>
    <h1>$Title</h1>
    <p class="meta">Generated: $generatedAt &nbsp;|&nbsp; Rows: $($rows.Count)</p>
    <table>
        <thead><tr>$headerHtml</tr></thead>
        <tbody>
$rowsHtml
        </tbody>
    </table>
</body>
</html>
"@
                [System.IO.File]::WriteAllText($OutputPath, $html, [System.Text.Encoding]::UTF8)
                Write-Verbose "Export-VBDNSLogReport: HTML written -- $($rows.Count) rows to '$OutputPath'"
            }

            'XLSX' {
                $sheetName = if ($Title.Length -gt 31) { $Title.Substring(0, 31) } else { $Title }

                $rows | Export-Excel -Path $OutputPath `
                    -WorksheetName $sheetName `
                    -AutoSize `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -Title $Title `
                    -TitleBold `
                    -TitleSize 14 `
                    -Force

                Write-Verbose "Export-VBDNSLogReport: XLSX written -- $($rows.Count) rows to '$OutputPath'"
            }
        }
    }
}
