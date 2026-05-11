function Get-VBmDNSRecord {
<#
.SYNOPSIS
    Discover mDNS services for an IP using dns-sd.exe (Layer 9).

.DESCRIPTION
    Runs dns-sd.exe -B for each of 5 service types, captures output for 5 seconds
    per type, and attempts to match results to the supplied IP address. A matching
    service name is returned as MDNSServiceType and MDNSServiceName.

    Service types probed:
        _pdl-datastream._tcp   -- printers (raw/JetDirect over TCP)
        _ipp._tcp              -- IPP printers
        _scanner._tcp          -- scanners
        _http._tcp             -- generic HTTP services (NAS, cameras, APs)
        _afpovertcp._tcp       -- Apple Filing Protocol (Macs, NAS)

    mDNS is VLAN-bound. If the probe machine and target are on different VLANs,
    results will always be empty -- this is expected, not a bug. The context report
    states this explicitly.

    dns-sd.exe -B lists services by name, not by IP. The function resolves each
    discovered service name via dns-sd.exe -L to get its IP and matches against
    the target. This is inherently slow (5 s listen per type) and best run on
    the same VLAN as the target devices.

    Prerequisites: $Context.mDNSAvailable must be $true (dns-sd.exe on PATH).

.PARAMETER IPAddress
    The RFC1918 / CGNAT / link-local IP address to match.

.PARAMETER Context
    Environment context from Get-VBEnrichmentContext. Provides mDNSAvailable.

.OUTPUTS
    [PSCustomObject] -- base layer result fields plus:
        MDNSServiceType [string]   e.g. '_ipp._tcp'
        MDNSServiceName [string]   service instance name

.EXAMPLE
    $ctx = Get-VBEnrichmentContext
    Get-VBmDNSRecord -IPAddress '192.168.1.50' -Context $ctx

.EXAMPLE
    '192.168.1.50' | Get-VBmDNSRecord -Context $ctx

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
        $LAYER_NUM  = 9
        $LAYER_NAME = 'mDNS'

        if (-not $Context) {
            Write-Warning "[$LAYER_NAME] No context provided -- running without prerequisite validation."
        }

        $ServiceTypes = @(
            '_pdl-datastream._tcp'
            '_ipp._tcp'
            '_scanner._tcp'
            '_http._tcp'
            '_afpovertcp._tcp'
        )

        # One-shot cache per run: browse all service types once and store results
        # keyed by IP so multiple IPs share the same dns-sd browse results
        if ($null -eq $Script:VBmDNSCache) {
            $Script:VBmDNSCache       = @{}
            $Script:VBmDNSCacheBuilt  = $false

            if (-not ($Context -and -not $Context.mDNSAvailable)) {
                Write-Verbose "[$LAYER_NAME] Building mDNS service cache (one-shot browse)..."
                $Script:VBmDNSCache = Invoke-VBmDNSBrowse -ServiceTypes $ServiceTypes
                $Script:VBmDNSCacheBuilt = $true
                Write-Verbose "[$LAYER_NAME] mDNS cache built: $($Script:VBmDNSCache.Count) IP entries"
            }
        }
    }

    process {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        if ($Context -and -not $Context.mDNSAvailable) {
            $sw.Stop()
            return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Skipped' -ExecutionMs $sw.ElapsedMilliseconds `
                -SkipReason 'BonjourNotInstalled' `
                -Impact 'Printers, scanners, and Apple devices using mDNS only may not be identified'
        }

        try {
            $entry = $Script:VBmDNSCache[$IPAddress]

            if ($null -eq $entry) {
                $sw.Stop()
                return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                    -Status 'NoResult' -ExecutionMs $sw.ElapsedMilliseconds
            }

            $sw.Stop()
            New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Success' -ExecutionMs $sw.ElapsedMilliseconds `
                -ExtraFields @{
                    MDNSServiceType = $entry.ServiceType
                    MDNSServiceName = $entry.ServiceName
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

function Invoke-VBmDNSBrowse {
<#
.SYNOPSIS
    Browse mDNS service types and resolve discovered services to IPs.
    Private helper used only by Get-VBmDNSRecord.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$ServiceTypes
    )

    # Result hashtable: IP -> @{ ServiceType; ServiceName }
    $ipMap = @{}

    foreach ($svcType in $ServiceTypes) {
        Write-Verbose "[mDNS] Browsing $svcType for 5 seconds..."

        # dns-sd -B runs until killed -- use a job to enforce 5 s timeout
        $job = Start-Job -ScriptBlock {
            param($t)
            & dns-sd.exe -B $t local. 2>&1
        } -ArgumentList $svcType

        Start-Sleep -Seconds 5
        Stop-Job -Job $job -ErrorAction SilentlyContinue
        $output = Receive-Job -Job $job -ErrorAction SilentlyContinue
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue

        # Parse browse output -- lines look like:
        # Browsing for _ipp._tcp.local.
        # DATE: ---Sun 10 May 2026---
        # Timestamp  A/D Flags IF Domain     ServiceType    InstanceName
        # 09:12:34.123  Add  3  8 local.  _ipp._tcp.        HP LaserJet Pro M404n
        foreach ($line in $output) {
            if ($line -match '^\d{2}:\d{2}:\d{2}\.\d+\s+Add\s+\d+\s+\d+\s+\S+\s+\S+\s+(.+)$') {
                $instanceName = $Matches[1].Trim()

                # Resolve the instance to an IP via dns-sd -L
                $resolveJob = Start-Job -ScriptBlock {
                    param($name, $type)
                    & dns-sd.exe -L $name $type local. 2>&1
                } -ArgumentList $instanceName, $svcType

                Start-Sleep -Seconds 3
                Stop-Job -Job $resolveJob -ErrorAction SilentlyContinue
                $resolveOutput = Receive-Job -Job $resolveJob -ErrorAction SilentlyContinue
                Remove-Job -Job $resolveJob -Force -ErrorAction SilentlyContinue

                # Resolved output contains a line with the IPv4 address
                foreach ($rLine in $resolveOutput) {
                    if ($rLine -match '(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})') {
                        $resolvedIP = $Matches[1]
                        if (-not $ipMap.ContainsKey($resolvedIP)) {
                            $ipMap[$resolvedIP] = @{
                                ServiceType = $svcType
                                ServiceName = $instanceName
                            }
                        }
                        break
                    }
                }
            }
        }
    }

    $ipMap
}
