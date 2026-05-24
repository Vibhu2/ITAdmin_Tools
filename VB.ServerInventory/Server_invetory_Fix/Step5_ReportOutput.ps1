#region -----------------------------------------------------------------------
# STEP 5 : SECTION-WISE REPORT OUTPUT + CSV EXPORT
# v1.2
# Changes from v1.1:
#   - Restored on-screen data display with sensible defaults:
#       single-record sections  → Format-List
#       multi-record sections   → Format-Table -AutoSize (first $ScreenRowLimit rows)
#   - $ScreenRowLimit controls how many rows print on screen (default 10)
#   - $SingleRecordSections list controls which sections use Format-List
#   - CSV export unchanged — always exports full data regardless of screen limit
#-------------------------------------------------------------------------------

#-- Helper: flatten nested arrays / hashtables / PSCustomObjects so Export-Csv
#   writes real values instead of 'System.Object[]' type-name strings.
function ConvertTo-FlatObject {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline)]
        [object]$InputObject
    )
    process {
        if ($null -eq $InputObject) { return }

        $props = [ordered]@{}
        foreach ($prop in $InputObject.PSObject.Properties) {
            $val = $prop.Value
            $props[$prop.Name] = switch ($true) {
                ($val -is [System.Collections.IEnumerable] -and $val -isnot [string]) {
                    ($val | ForEach-Object { $_ }) -join '; '
                    break
                }
                ($val -is [hashtable] -or $val -is [PSCustomObject]) {
                    $val | ConvertTo-Json -Compress -Depth 3
                    break
                }
                default { $val }
            }
        }
        [PSCustomObject]$props
    }
}

#-- On-screen display config
#   Sections that return a single flat object are better read as a vertical list.
#   Everything else gets a table. Adjust these two variables to taste.
$ScreenRowLimit     = 10     # max rows to print on screen per section (CSV always gets all)
$SingleRecordSections = @(
    'SystemInfo',
    'AzureADJoinStatus',
    'DNSServerInfo',
    'DHCPInformation',
    'DHCPDetailedInfo',
    'NetworkInformation',
    'WindowsUpdateInfo',
    'ActiveDirectory',
    'BitLockerRecovery',
    'RDSUsers'
)

#-- Report banner
Write-Host "`n"
Write-Host "################################################################" -ForegroundColor White
Write-Host "  SERVER INVENTORY REPORT : $hostname"                            -ForegroundColor White
Write-Host "  Generated : $(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss')"        -ForegroundColor White
Write-Host "################################################################" -ForegroundColor White

foreach ($section in $InventoryResults) {

    # Resolve record count — prefer upstream value, fall back to live count
    $recordCount = if ($null -ne $section.RecordCount -and $section.RecordCount -ne '') {
        $section.RecordCount
    } else {
        @($section.Data).Count
    }

    $statusColor = if ($section.Status -eq 'Success') { 'Green' } else { 'Red' }

    Write-Host "`n================================================================" -ForegroundColor Cyan
    Write-Host "  SECTION  : $($section.Section)"                                  -ForegroundColor Yellow
    Write-Host "  Status   : $($section.Status)   |   Records: $recordCount"       -ForegroundColor $statusColor
    Write-Host "================================================================"   -ForegroundColor Cyan

    # Failed section — show error, skip CSV
    if ($section.Status -eq 'Failed') {
        Write-Host "  ERROR: $($section.Error)" -ForegroundColor Red
        continue
    }

    # Empty data — nothing to export
    if (-not $section.Data) {
        Write-Host "  (No data returned for this section)" -ForegroundColor DarkYellow
        continue
    }

    # On-screen display
    if ($section.Section -in $SingleRecordSections) {
        # Single flat object — vertical list is easier to read
        $section.Data | Format-List
    } else {
        # Multi-record — table, capped at $ScreenRowLimit rows
        $displayData = @($section.Data) | Select-Object -First $ScreenRowLimit
        $displayData | Format-Table -AutoSize
        if ($recordCount -gt $ScreenRowLimit) {
            Write-Host "  ... $($recordCount - $ScreenRowLimit) more row(s) not shown. See CSV for full data." -ForegroundColor DarkGray
        }
    }

    # CSV export — flatten nested objects before writing
    $csvPath = Join-Path $OutputPath "$($section.Section).csv"
    try {
        $section.Data |
            ConvertTo-FlatObject |
            Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
        Write-Host "  CSV : $csvPath" -ForegroundColor DarkGreen
    }
    catch {
        Write-Host "  CSV export FAILED : $_" -ForegroundColor Red
    }
}

Write-Host "`n################################################################" -ForegroundColor White
Write-Host "  END OF REPORT"                                                    -ForegroundColor White
Write-Host "  Output folder : $OutputPath"                                      -ForegroundColor White
Write-Host "################################################################`n" -ForegroundColor White
#endregion
