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
        [string]$IPAddress,

        [Parameter()]
        [PSCustomObject]$Context
    )

    begin {
        $LAYER_NUM  = 4
        $LAYER_NAME = 'ARP'

        if ($null -eq $Script:VBArpCache) {
            $Script:VBArpCache = Get-VBARPTable
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

function Get-VBARPTable {
<#
.SYNOPSIS
    Parse arp -a output into a hashtable keyed by IP address.
    Private helper used by Get-VBARPEntry.
#>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$IPFilter
    )

    $table = @{}

    $arpArgs = if ($IPFilter) { @($IPFilter) } else { @() }
    $arpLines = & arp.exe -a @arpArgs 2>$null

    if ($LASTEXITCODE -ne 0 -and -not $IPFilter) {
        Write-Warning "[ARP] arp.exe returned exit code $LASTEXITCODE"
        return $table
    }

    # arp -a output format: "  <IP>          <MAC>         <type>"
    # e.g.:  192.168.1.1     00-11-22-33-44-55     dynamic
    foreach ($line in $arpLines) {
        if ($line -match '^\s+([\d\.]+)\s+([\w\-]+)\s+(static|dynamic)\s*$') {
            $ip  = $Matches[1].Trim()
            $mac = $Matches[2].Trim()
            $type = $Matches[3].Trim()
            # Skip multicast/broadcast entries (01-xx, ff-ff-ff-ff-ff-ff)
            if ($mac -match '^01-' -or $mac -eq 'ff-ff-ff-ff-ff-ff') { continue }
            $table[$ip] = [PSCustomObject]@{
                MAC  = $mac
                Type = $type
            }
        }
    }

    $table
}
