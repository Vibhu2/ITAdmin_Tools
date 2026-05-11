function New-VBLayerResult {
<#
.SYNOPSIS
    Build a standardised layer-result PSCustomObject for any Get-VB* layer function.

.DESCRIPTION
    Every layer function in VB.DNSEnrichment must return an object with a consistent
    base shape (IPAddress, Layer, LayerName, Status, ExecutionMs, SkipReason, Impact,
    ErrorDetail, CollectionTime). This factory enforces that shape and merges in any
    layer-specific fields supplied via -ExtraFields.

    Status values:
        Success    -- probe succeeded and returned data
        NoResult   -- probe ran but found nothing (e.g. port closed, no PTR record)
        Failed     -- probe threw -- ErrorDetail populated
        Skipped    -- prerequisite missing -- SkipReason populated, no probe attempted

.PARAMETER IPAddress
    The IP being processed.

.PARAMETER Layer
    The layer number (1-11) per the design spec.

.PARAMETER LayerName
    Short human-readable layer name -- AD, DHCP, PTR, ARP, TCP, HTTP, SNMP, OUI, RTSP, mDNS, Switch.

.PARAMETER Status
    Success | NoResult | Failed | Skipped.

.PARAMETER ExecutionMs
    Wall-clock time the layer spent, in milliseconds.

.PARAMETER SkipReason
    When Status='Skipped' -- short reason code (e.g. 'DNSUnavailable', 'NoMACAvailable').

.PARAMETER Impact
    When Status='Skipped' -- one-line description of what is lost by skipping.

.PARAMETER ErrorDetail
    When Status='Failed' -- the exception message.

.PARAMETER ExtraFields
    Hashtable of layer-specific fields to merge into the result.

.OUTPUTS
    [PSCustomObject]

.EXAMPLE
    New-VBLayerResult -IPAddress '192.168.1.45' -Layer 3 -LayerName 'PTR' `
        -Status 'Success' -ExecutionMs 12 -ExtraFields @{
            Hostname = 'pc01.corp.local'
            ForwardConfirmed = $true
        }

.NOTES
    Version:      1.0.0
    MinPSVersion: 5.1
    Author:       VB
    ChangeLog:
        1.0.0 -- 2026-05-10 -- Initial release
#>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)]
        [string]$IPAddress,

        [Parameter(Mandatory)]
        [int]$Layer,

        [Parameter(Mandatory)]
        [string]$LayerName,

        [Parameter(Mandatory)]
        [ValidateSet('Success', 'NoResult', 'Failed', 'Skipped')]
        [string]$Status,

        [Parameter()]
        [int]$ExecutionMs = 0,

        [Parameter()]
        [string]$SkipReason,

        [Parameter()]
        [string]$Impact,

        [Parameter()]
        [string]$ErrorDetail,

        [Parameter()]
        [hashtable]$ExtraFields = @{}
    )

    # Use ordered hashtable so property order is stable across all layer outputs
    $base = [ordered]@{
        IPAddress      = $IPAddress
        Layer          = $Layer
        LayerName      = $LayerName
        Status         = $Status
        ExecutionMs    = $ExecutionMs
        SkipReason     = $SkipReason
        Impact         = $Impact
        ErrorDetail    = $ErrorDetail
        CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
    }

    foreach ($key in $ExtraFields.Keys) {
        if (-not $base.Contains($key)) {
            $base[$key] = $ExtraFields[$key]
        }
        else {
            $base[$key] = $ExtraFields[$key]
        }
    }

    [PSCustomObject]$base
}
