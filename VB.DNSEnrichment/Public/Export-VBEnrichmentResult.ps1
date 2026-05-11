function Export-VBEnrichmentResult {
<#
.SYNOPSIS
    Export enrichment results to CSV, JSON, or return as pipeline objects.

.DESCRIPTION
    Consumes enrichment PSCustomObjects (from Invoke-VBIPEnrichment or
    Get-VBEnrichmentResult) and writes them to a file or returns them to
    the pipeline.

    Format behaviour:
        Object  -- returns [PSCustomObject[]] to pipeline (default)
        CSV     -- UTF-8 BOM, LayerTrace flattened to a JSON string column
        JSON    -- pretty-printed JSON array, UTF-8 BOM

    LayerTrace is always excluded from CSV/JSON by default unless -IncludeLayerTrace
    is supplied.

.PARAMETER InputObject
    Enrichment objects to export. Accepts pipeline input.

.PARAMETER Format
    Output format: Object | CSV | JSON. Default: Object.

.PARAMETER Path
    Output file path. Required when Format is CSV or JSON.

.PARAMETER IncludeLayerTrace
    Include the LayerTrace column (CSV) or property (JSON). Excluded by default
    to keep exports readable.

.PARAMETER Context
    Environment context from Get-VBEnrichmentContext. Not required for this
    function -- accepted for pipeline consistency.

.OUTPUTS
    None when Format is CSV or JSON (writes to file).
    [PSCustomObject[]] when Format is Object.

.EXAMPLE
    Get-VBEnrichmentResult -Context $ctx | Export-VBEnrichmentResult -Format CSV -Path 'C:\Reports\enrichment.csv'

.EXAMPLE
    Invoke-VBIPEnrichment -IPAddress $ips -Context $ctx |
        Export-VBEnrichmentResult -Format JSON -Path 'C:\Reports\enrichment.json' -IncludeLayerTrace

.EXAMPLE
    $results = Get-VBEnrichmentResult -Context $ctx | Export-VBEnrichmentResult

.NOTES
    Version:      1.0.0
    MinPSVersion: 5.1
    Author:       VB
    ChangeLog:
        1.0.0 -- 2026-05-11 -- Initial release
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [PSCustomObject[]]$InputObject,

        [Parameter()]
        [ValidateSet('CSV', 'JSON', 'Object')]
        [string]$Format = 'Object',

        [Parameter()]
        [string]$Path,

        [Parameter()]
        [switch]$IncludeLayerTrace,

        [Parameter()]
        [PSCustomObject]$Context
    )

    begin {
        if ($Format -ne 'Object' -and -not $Path) {
            throw "Export-VBEnrichmentResult: -Path is required when -Format is '$Format'."
        }
        $buffer = New-Object System.Collections.Generic.List[PSCustomObject]
    }

    process {
        foreach ($item in $InputObject) {
            $buffer.Add($item)
        }
    }

    end {
        if ($buffer.Count -eq 0) {
            Write-Warning "[Export-VBEnrichmentResult] No objects to export."
            return
        }

        switch ($Format) {

            'Object' {
                foreach ($item in $buffer) { $item }
            }

            'CSV' {
                # Flatten LayerTrace to JSON string; remove it if not requested
                $flat = foreach ($item in $buffer) {
                    $row = [ordered]@{}
                    foreach ($prop in $item.PSObject.Properties) {
                        if ($prop.Name -eq 'LayerTrace') {
                            if ($IncludeLayerTrace) {
                                $row['LayerTraceJson'] = $prop.Value | ConvertTo-Json -Compress -Depth 3
                            }
                        }
                        elseif ($prop.Name -eq 'History') {
                            # Skip -- only present when IncludeHistory used
                        }
                        else {
                            $row[$prop.Name] = $prop.Value
                        }
                    }
                    [PSCustomObject]$row
                }

                # PS 5.1: Export-Csv -Encoding utf8 does NOT include BOM.
                # Write manually using StreamWriter with UTF8 BOM encoding.
                $utf8Bom  = New-Object System.Text.UTF8Encoding $true
                $csvLines = $flat | ConvertTo-Csv -NoTypeInformation

                $writer = $null
                try {
                    $writer = New-Object System.IO.StreamWriter($Path, $false, $utf8Bom)
                    foreach ($line in $csvLines) {
                        $writer.WriteLine($line)
                    }
                    Write-Verbose "[Export] CSV written to $Path ($($buffer.Count) rows)"
                }
                finally {
                    if ($writer) { $writer.Close() }
                }
            }

            'JSON' {
                $toSerialise = if ($IncludeLayerTrace) {
                    $buffer.ToArray()
                }
                else {
                    $buffer | ForEach-Object {
                        $props = $_.PSObject.Properties |
                            Where-Object { $_.Name -ne 'LayerTrace' -and $_.Name -ne 'History' }
                        $obj = [ordered]@{}
                        foreach ($p in $props) { $obj[$p.Name] = $p.Value }
                        [PSCustomObject]$obj
                    }
                }

                $json    = $toSerialise | ConvertTo-Json -Depth 5
                $utf8Bom = New-Object System.Text.UTF8Encoding $true
                $writer  = $null
                try {
                    $writer = New-Object System.IO.StreamWriter($Path, $false, $utf8Bom)
                    $writer.Write($json)
                    Write-Verbose "[Export] JSON written to $Path ($($buffer.Count) objects)"
                }
                finally {
                    if ($writer) { $writer.Close() }
                }
            }
        }
    }
}
