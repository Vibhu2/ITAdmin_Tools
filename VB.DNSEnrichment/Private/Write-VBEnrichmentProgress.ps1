function Write-VBEnrichmentProgress {
<#
.SYNOPSIS
    Internal Write-Progress wrapper used by the orchestrator for per-IP step display.

.DESCRIPTION
    Wraps Write-Progress with a fixed activity prefix and the design's status
    format -- "<IP> -- Step <N>: <LayerName>". Computes percent-complete and
    seconds-remaining from $Current/$Total/$ElapsedMs.

    Defined here in Round 1 so the orchestrator (Round 2) can use it directly
    without backfilling. Safe to call repeatedly -- Write-Progress is idempotent
    on the same -Id.

.PARAMETER Activity
    The progress activity title. Defaults to 'VB.DNSEnrichment'.

.PARAMETER Current
    Number of IPs processed so far (0-based or 1-based, caller's choice).

.PARAMETER Total
    Total number of IPs in the run.

.PARAMETER IPAddress
    The IP currently being processed.

.PARAMETER StepNumber
    Layer number 1-11.

.PARAMETER LayerName
    Short layer name (AD, DHCP, PTR, ARP, TCP, HTTP, SNMP, OUI, RTSP, mDNS, Switch).

.PARAMETER ElapsedMs
    Wall-clock milliseconds elapsed since the run started -- used to estimate
    seconds remaining.

.PARAMETER Id
    Progress bar Id. Defaults to 1.

.PARAMETER Completed
    Switch -- emit a -Completed Write-Progress to clear the bar.

.OUTPUTS
    None.

.EXAMPLE
    Write-VBEnrichmentProgress -Current 23 -Total 47 -IPAddress '192.168.1.105' `
        -StepNumber 6 -LayerName 'HTTP' -ElapsedMs 18500

.NOTES
    Version:      1.0.0
    MinPSVersion: 5.1
    Author:       VB
    ChangeLog:
        1.0.0 -- 2026-05-10 -- Initial release
#>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$Activity = 'VB.DNSEnrichment',

        [Parameter()]
        [int]$Current = 0,

        [Parameter()]
        [int]$Total = 0,

        [Parameter()]
        [string]$IPAddress,

        [Parameter()]
        [int]$StepNumber,

        [Parameter()]
        [string]$LayerName,

        [Parameter()]
        [long]$ElapsedMs = 0,

        [Parameter()]
        [int]$Id = 1,

        [Parameter()]
        [switch]$Completed
    )

    if ($Completed) {
        Write-Progress -Id $Id -Activity $Activity -Completed
        return
    }

    $percent = 0
    if ($Total -gt 0) {
        $percent = [math]::Min(100, [math]::Floor(($Current / $Total) * 100))
    }

    $status = "$Current of $Total IPs"
    if ($IPAddress) {
        $status = "$IPAddress -- Step $StepNumber`: $LayerName"
    }

    $secondsRemaining = -1
    if ($Current -gt 0 -and $Total -gt 0 -and $ElapsedMs -gt 0) {
        $msPerItem    = $ElapsedMs / $Current
        $remainingMs  = $msPerItem * ($Total - $Current)
        $secondsRemaining = [int]([math]::Ceiling($remainingMs / 1000))
    }

    $splat = @{
        Id              = $Id
        Activity        = "$Activity -- $Current of $Total IPs processed"
        Status          = $status
        PercentComplete = $percent
    }
    if ($secondsRemaining -ge 0) {
        $splat['SecondsRemaining'] = $secondsRemaining
    }

    Write-Progress @splat
}
