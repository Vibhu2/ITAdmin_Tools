function Format-VBEnrichmentContext {
<#
.SYNOPSIS
    Render the context object as a formatted console report.

.DESCRIPTION
    Internal display helper invoked by Get-VBEnrichmentContext when -Quiet is not
    supplied. Writes the report via Write-Information using the VBContext tag.
    Returns nothing to the pipeline.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Context
    )

    $line = ('=' * 71)
    $sep  = ('-' * 71)

    Write-Information -MessageData '' -Tags 'VBContext' -InformationAction Continue
    Write-Information -MessageData $line -Tags 'VBContext' -InformationAction Continue
    Write-Information -MessageData '  VB.DNSEnrichment -- Environment Prerequisites' -Tags 'VBContext' -InformationAction Continue
    Write-Information -MessageData $line -Tags 'VBContext' -InformationAction Continue

    $machineLine = "  Computer    : $($Context.ComputerName)"
    if ($Context.IsDomainController) { $machineLine += ' (Domain Controller)' }
    elseif ($Context.IsDomainJoined) { $machineLine += " (Domain: $($Context.DomainName))" }
    else                             { $machineLine += ' (Workgroup)' }
    Write-Information -MessageData $machineLine -Tags 'VBContext' -InformationAction Continue

    $parallelText = if ($Context.CanUseParallel) { 'ENABLED' } else { 'DISABLED (PS < 7)' }
    Write-Information -MessageData "  PowerShell  : $($Context.PSVersion) $($Context.PSEdition)  -> Parallel execution: $parallelText" -Tags 'VBContext' -InformationAction Continue

    $dbText = if ($Context.DatabaseInitialized) { "$($Context.DatabasePath) (initialized)" } else { "$($Context.DatabasePath) (NOT INITIALIZED)" }
    Write-Information -MessageData "  Database    : $dbText" -Tags 'VBContext' -InformationAction Continue

    $probeText = if ($Context.NetworkProbeEnabled) { 'Active probes ENABLED' } else { 'Active probes DISABLED' }
    Write-Information -MessageData "  Probes      : $probeText" -Tags 'VBContext' -InformationAction Continue

    Write-Information -MessageData $sep -Tags 'VBContext' -InformationAction Continue
    Write-Information -MessageData '' -Tags 'VBContext' -InformationAction Continue
    Write-Information -MessageData '  Layer  Prerequisite          Status         Detail' -Tags 'VBContext' -InformationAction Continue
    Write-Information -MessageData '  -----  -------------------   ------------   -------------------------------' -Tags 'VBContext' -InformationAction Continue

    $available     = 0
    $unavailable   = 0
    $notConfigured = 0
    $skipped       = 0
    $skippedLayers = New-Object System.Collections.Generic.List[string]

    foreach ($row in $Context.PrerequisiteReport) {
        switch ($row.Status) {
            'Available'     { $available++;     break }
            'Unavailable'   { $unavailable++;   break }
            'NotConfigured' { $notConfigured++; break }
            'Skipped'       { $skipped++;       break }
        }

        $layerText = if ($row.Layer -gt 0) { ([string]$row.Layer).PadLeft(2) } else { '--' }
        $preq      = $row.Prerequisite.PadRight(20)
        $statusPad = $row.Status.PadRight(13)
        $detail    = if ($row.Detail) { $row.Detail } else { '' }

        Write-Information -MessageData ('   {0}    {1}  {2}  {3}' -f $layerText, $preq, $statusPad, $detail) -Tags 'VBContext' -InformationAction Continue

        if ($row.Status -ne 'Available' -and $row.SkippedLayers) {
            $skippedLayers.Add($row.SkippedLayers)
            if ($row.Impact) {
                Write-Information -MessageData "                                              -> Impact: $($row.Impact)" -Tags 'VBContext' -InformationAction Continue
            }
            if ($row.Remediation) {
                Write-Information -MessageData "                                              -> Fix: $($row.Remediation)" -Tags 'VBContext' -InformationAction Continue
            }
        }
    }

    Write-Information -MessageData '' -Tags 'VBContext' -InformationAction Continue
    Write-Information -MessageData $line -Tags 'VBContext' -InformationAction Continue
    Write-Information -MessageData "  Summary: $available available, $unavailable unavailable, $notConfigured not configured" -Tags 'VBContext' -InformationAction Continue
    if ($skippedLayers.Count -gt 0) {
        Write-Information -MessageData "  Skipped: $($skippedLayers -join ', ')" -Tags 'VBContext' -InformationAction Continue
    }
    $modeText = if ($Context.CanUseParallel) { "Parallel x $($Context.ParallelThrottleLimit) on active probes" } else { 'Sequential (PS 5.1)' }
    Write-Information -MessageData "  Mode   : $modeText" -Tags 'VBContext' -InformationAction Continue
    Write-Information -MessageData "  Built  : $($Context.ContextDurationMs) ms" -Tags 'VBContext' -InformationAction Continue
    Write-Information -MessageData $line -Tags 'VBContext' -InformationAction Continue
    Write-Information -MessageData '' -Tags 'VBContext' -InformationAction Continue
}
