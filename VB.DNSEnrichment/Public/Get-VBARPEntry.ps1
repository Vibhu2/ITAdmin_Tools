function Get-VBARPEntry {
<#
.SYNOPSIS
    Look up a MAC address for an IP via the local ARP cache (Layer 4).

.DESCRIPTION
    On the first call within a session, parses the output of 'arp -a' into a
    script-scope hashtable keyed by IP address. All subsequent calls perform a
    hashtable lookup with zero I/O cost.

    If the IP is not found in the initial ARP table, the function optionally
    sends a single ICMP ping (200 ms timeout) to populate the OS ARP table,
    then re-reads the cache for that IP only.

    No network prerequisite -- arp.exe is always available on Windows.

.PARAMETER IPAddress
    The RFC1918 / CGNAT / link-local IP address to look up.

.PARAMETER Context
    Environment context object from Get-VBEnrichmentContext. Accepted for
    pipeline consistency; no flags required by this layer.

.OUTPUTS
    [PSCustomObject] -- base layer result fields plus:
        MACAddress       [string]
        MACNormalised    [string]
        ARPType          [string]   static | dynamic
        PingedToPopulate [bool]

.EXAMPLE
    $ctx = Get-VBEnrichmentContext
    Get-VBARPEntry -IPAddress '192.168.1.45' -Context $ctx

.EXAMPLE
    '192.168.1.45','192.168.1.46' | Get-VBARPEntry -Context $ctx

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
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [string]$IPAddress,

        [Parameter()]
        [PSCustomObject]$Context
    )

    begin {
        $LAYER_NUM  = 4
        $LAYER_NAME = 'ARP'

        $CacheTTLMinutes = 60
        if ($null -ne $Script:VBArpCache) {
            $ageMin = ((Get-Date) - $Script:VBArpCacheBuiltAt).TotalMinutes
            if ($ageMin -gt $CacheTTLMinutes) {
                Write-Verbose "[$LAYER_NAME] ARP cache is $([int]$ageMin) min old (TTL $CacheTTLMinutes min) -- rebuilding"
                $Script:VBArpCache = $null
            }
        }

        if ($null -eq $Script:VBArpCache) {
            $Script:VBArpCache      = Get-VBARPTable
            $Script:VBArpCacheBuiltAt = Get-Date
            Write-Verbose "[$LAYER_NAME] ARP cache built: $($Script:VBArpCache.Count) entries"
        }
    }

    process {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        try {
            $entry = $Script:VBArpCache[$IPAddress]
            $pinged = $false

            if ($null -eq $entry) {
                # Ping to populate OS ARP table then re-check this one IP
                $pinged = $true
                Write-Verbose "[$LAYER_NAME] $IPAddress not in ARP cache -- pinging to populate"
                $null = Test-Connection -ComputerName $IPAddress -Count 1 -ErrorAction SilentlyContinue

                # Re-read only arp -a for the specific IP (fast)
                $refreshed = Get-VBARPTable -IPFilter $IPAddress
                if ($refreshed.Count -gt 0) {
                    $entry = $refreshed[$IPAddress]
                    # Merge into session cache
                    if ($null -ne $entry) {
                        $Script:VBArpCache[$IPAddress] = $entry
                    }
                }
            }

            if ($null -eq $entry) {
                $sw.Stop()
                return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                    -Status 'NoResult' -ExecutionMs $sw.ElapsedMilliseconds
            }

            $macNormalised = ConvertTo-VBNormalisedMAC -MACAddress $entry.MAC

            $sw.Stop()
            New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Success' -ExecutionMs $sw.ElapsedMilliseconds `
                -ExtraFields @{
                    MACAddress       = $entry.MAC
                    MACNormalised    = $macNormalised
                    ARPType          = $entry.Type
                    PingedToPopulate = $pinged
                }
        }
        catch {
            $sw.Stop()
            New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Failed' -ExecutionMs $sw.ElapsedMilliseconds `
                -ErrorDetail $_.Exception.Message
        }
    }
}
