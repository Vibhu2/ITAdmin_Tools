function ConvertFrom-VBSqliteEnrichmentRow {
<#
.SYNOPSIS
    Reconstruct an enrichment PSCustomObject from a SQLite row.
    Private helper used by Invoke-VBIPEnrichment.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Row,

        [Parameter()]
        [bool]$FromCache = $false
    )

    $layerTrace = @()
    if ($Row.LayerTraceJson) {
        try {
            $layerTrace = $Row.LayerTraceJson | ConvertFrom-Json
        }
        catch { }
    }

    [PSCustomObject]@{
        IPAddress            = $Row.IPAddress
        Hostname             = $Row.Hostname
        HostnameSource       = $Row.HostnameSource
        MACAddress           = $Row.MACAddress
        MACAddressNormalised = $Row.MACAddressNormalised
        Vendor               = $Row.Vendor
        DeviceClass          = $Row.DeviceClass
        DeviceClassSource    = $Row.DeviceClassSource
        Confidence           = $Row.Confidence
        OSClass              = $Row.OSClass
        OperatingSystem      = $Row.OperatingSystem
        Model                = $Row.Model
        Location             = $Row.Location
        OU                   = $Row.OU
        OpenPorts            = $Row.OpenPorts
        HTTPTitle            = $Row.HTTPTitle
        HTTPServer           = $Row.HTTPServer
        SNMPDescr            = $Row.SNMPDescr
        RTSPBanner           = $Row.RTSPBanner
        MDNSServiceType      = $Row.MDNSServiceType
        StepsAttempted       = [int]$Row.StepsAttempted
        StepsSucceeded       = [int]$Row.StepsSucceeded
        StepsNoResult        = [int]$Row.StepsNoResult
        StepsSkipped         = [int]$Row.StepsSkipped
        StepsFailed          = [int]$Row.StepsFailed
        LayerTrace           = $layerTrace
        IsResolved           = [bool]$Row.IsResolved
        IsUnresolved         = [bool]$Row.IsUnresolved
        EnrichedAt           = if ($Row.EnrichedAt)  { [datetime]$Row.EnrichedAt }  else { $null }
        EnrichmentDurationMs = [int]$Row.EnrichmentDurationMs
        FirstSeenAt          = if ($Row.FirstSeenAt) { [datetime]$Row.FirstSeenAt } else { $null }
        UpdatedAt            = if ($Row.UpdatedAt)   { [datetime]$Row.UpdatedAt }   else { $null }
        ChangeReason         = 'NoChange'
        FromCache            = $FromCache
    }
}
