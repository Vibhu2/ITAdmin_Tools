function Get-VBSwitchARP {
<#
.SYNOPSIS
    Look up a device's MAC address, switch port, and port description via SNMP (Layer 10).

.DESCRIPTION
    On the first call within a session, walks the ARP, MAC-address, and interface
    alias tables on every switch in $Context.SwitchTargets and builds a script-scope
    hashtable keyed by IP. All subsequent calls are hashtable lookups with zero I/O.

    SNMP OIDs walked per switch:
        1.3.6.1.2.1.4.22.1.2    ipNetToMediaPhysAddress  -- IP -> MAC (ARP table)
        1.3.6.1.2.1.17.4.3.1.2  dot1dTpFdbPort           -- MAC -> bridge port number
        1.3.6.1.2.1.31.1.1.1.18 ifAlias                  -- port number -> description

    The COM object is always released in a finally block. If a switch times out or
    refuses the community string it is skipped with Write-Warning; remaining switches
    are still queried.

    Prerequisites:
        $Context.SNMPAvailable must be $true (olePrn COM available)
        $Context.SwitchTargets.Count -gt 0 (at least one switch IP configured)

.PARAMETER IPAddress
    The RFC1918 / CGNAT / link-local IP address to look up.

.PARAMETER Context
    Environment context from Get-VBEnrichmentContext. Provides SNMPAvailable,
    SNMPCommunityStrings, SwitchTargets, and DefaultTimeoutMs.SNMP.

.OUTPUTS
    [PSCustomObject] -- base layer result fields plus:
        MACAddress      [string]
        SwitchIP        [string]
        SwitchPort      [string]
        PortDescription [string]   ifAlias value -- used as Location

.EXAMPLE
    $ctx = Get-VBEnrichmentContext -SwitchTargets '192.168.1.254','192.168.2.254'
    Get-VBSwitchARP -IPAddress '192.168.1.45' -Context $ctx

.EXAMPLE
    '192.168.1.45' | Get-VBSwitchARP -Context $ctx

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
        $LAYER_NUM  = 10
        $LAYER_NAME = 'Switch'

        if (-not $Context) {
            Write-Warning "[$LAYER_NAME] No context provided -- running without prerequisite validation."
        }

        $communityStrings = if ($Context -and $Context.SNMPCommunityStrings.Count -gt 0) {
            $Context.SNMPCommunityStrings
        } else { @('public') }

        $timeoutMs = if ($Context -and $Context.DefaultTimeoutMs) {
            $Context.DefaultTimeoutMs.SNMP
        } else { 2000 }

        $switchTargets = if ($Context) { $Context.SwitchTargets } else { @() }

        # One-shot cache: build ARP+port map from all switches on first call
        if ($null -eq $Script:VBSwitchARPCache) {
            $Script:VBSwitchARPCache      = @{}
            $Script:VBSwitchARPCacheBuilt = $false

            $skipReasons = @()
            if ($Context -and -not $Context.SNMPAvailable) { $skipReasons += 'SNMPUnavailable' }
            if ($switchTargets.Count -eq 0)                { $skipReasons += 'NoSwitchTargets' }

            if ($skipReasons.Count -eq 0) {
                Write-Verbose "[$LAYER_NAME] Building switch ARP cache from $($switchTargets.Count) switch(es)..."
                foreach ($switchIP in $switchTargets) {
                    $switchEntries = Invoke-VBSwitchSNMPWalk -SwitchIP $switchIP `
                        -CommunityStrings $communityStrings -TimeoutMs $timeoutMs
                    foreach ($key in $switchEntries.Keys) {
                        if (-not $Script:VBSwitchARPCache.ContainsKey($key)) {
                            $Script:VBSwitchARPCache[$key] = $switchEntries[$key]
                        }
                    }
                }
                $Script:VBSwitchARPCacheBuilt = $true
                Write-Verbose "[$LAYER_NAME] Switch ARP cache built: $($Script:VBSwitchARPCache.Count) IP entries"
            }
        }
    }

    process {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        if ($Context -and -not $Context.SNMPAvailable) {
            $sw.Stop()
            return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Skipped' -ExecutionMs $sw.ElapsedMilliseconds `
                -SkipReason 'SNMPUnavailable' `
                -Impact 'No switch port location -- unresolved IPs will lack physical location'
        }

        if ($switchTargets.Count -eq 0) {
            $sw.Stop()
            return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Skipped' -ExecutionMs $sw.ElapsedMilliseconds `
                -SkipReason 'NoSwitchTargets' `
                -Impact 'No switch port location -- re-run Get-VBEnrichmentContext with -SwitchTargets'
        }

        try {
            $entry = $Script:VBSwitchARPCache[$IPAddress]

            if ($null -eq $entry) {
                $sw.Stop()
                return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                    -Status 'NoResult' -ExecutionMs $sw.ElapsedMilliseconds
            }

            $sw.Stop()
            New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Success' -ExecutionMs $sw.ElapsedMilliseconds `
                -ExtraFields @{
                    MACAddress      = $entry.MACAddress
                    SwitchIP        = $entry.SwitchIP
                    SwitchPort      = $entry.SwitchPort
                    PortDescription = $entry.PortDescription
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

function Invoke-VBSwitchSNMPWalk {
<#
.SYNOPSIS
    Walk ARP, MAC, and ifAlias tables on a single managed switch via SNMP.
    Private helper used only by Get-VBSwitchARP.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SwitchIP,

        [Parameter(Mandatory)]
        [string[]]$CommunityStrings,

        [Parameter()]
        [int]$TimeoutMs = 2000
    )

    # OIDs to walk
    $OID_ARP_TABLE   = '1.3.6.1.2.1.4.22.1.2'    # ipNetToMediaPhysAddress: IP -> MAC
    $OID_FDB_PORT    = '1.3.6.1.2.1.17.4.3.1.2'  # dot1dTpFdbPort: MAC -> port index
    $OID_IF_ALIAS    = '1.3.6.1.2.1.31.1.1.1.18' # ifAlias: port index -> description

    $result   = @{}  # keyed by IP
    $snmp     = $null
    $community = $null

    # Try community strings to find one that works
    foreach ($cs in $CommunityStrings) {
        $snmp = $null
        try {
            $snmp = New-Object -ComObject olePrn.OleSNMP -ErrorAction Stop
            $snmp.Timeout = $TimeoutMs
            $snmp.Open($SwitchIP, $cs, 2, $TimeoutMs)
            # Test with a quick get -- if it throws, community is wrong
            $null = $snmp.Get('1.3.6.1.2.1.1.1.0')
            $community = $cs
            break
        }
        catch {
            if ($snmp) {
                try { $snmp.Close() } catch { }
                try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($snmp) | Out-Null } catch { }
                $snmp = $null
            }
        }
    }

    if (-not $community) {
        Write-Warning "[$SwitchIP] No SNMP community string succeeded -- skipping this switch"
        return $result
    }

    try {
        # Walk ARP table: each entry is OID suffix = ifIndex.IP, value = MAC hex string
        $arpTable  = @{}   # IP -> MAC
        $arpWalk   = $snmp.GetTree($OID_ARP_TABLE)
        if ($arpWalk) {
            foreach ($key in $arpWalk.Keys) {
                # OID suffix format: <ifIndex>.<a>.<b>.<c>.<d>  e.g. 1.192.168.1.45
                if ($key -match '\.(\d+\.\d+\.\d+\.\d+)$') {
                    $ip  = $Matches[1]
                    $mac = $arpWalk[$key]
                    if ($mac -and $mac -match '^[0-9A-Fa-f:]{17}$') {
                        $arpTable[$ip] = $mac
                    }
                }
            }
        }

        # Walk FDB port table: MAC -> bridge port index
        $macPortTable = @{}  # MAC (no separators, uppercase) -> port index
        $fdbWalk      = $snmp.GetTree($OID_FDB_PORT)
        if ($fdbWalk) {
            foreach ($key in $fdbWalk.Keys) {
                # OID suffix is MAC as decimal octets: 0.26.43.60.77.94
                if ($key -match '((\d+\.){5}\d+)$') {
                    $octets    = $Matches[1] -split '\.'
                    $macHex    = ($octets | ForEach-Object { '{0:X2}' -f [int]$_ }) -join ''
                    $portIndex = $fdbWalk[$key]
                    $macPortTable[$macHex] = $portIndex
                }
            }
        }

        # Walk ifAlias table: port index -> description string
        $ifAliasTable = @{}  # port index -> alias
        $aliasWalk    = $snmp.GetTree($OID_IF_ALIAS)
        if ($aliasWalk) {
            foreach ($key in $aliasWalk.Keys) {
                if ($key -match '\.(\d+)$') {
                    $ifIndex = $Matches[1]
                    $alias   = $aliasWalk[$key]
                    if (-not [string]::IsNullOrWhiteSpace($alias)) {
                        $ifAliasTable[$ifIndex] = $alias
                    }
                }
            }
        }

        # Merge: for each IP in ARP table, look up port and description
        foreach ($ip in $arpTable.Keys) {
            $mac       = $arpTable[$ip]
            $macNorm   = ($mac -replace '[:\-\.]', '').ToUpperInvariant()
            $portIndex = $macPortTable[$macNorm]
            $portAlias = if ($portIndex) { $ifAliasTable[$portIndex.ToString()] } else { $null }

            $result[$ip] = [PSCustomObject]@{
                MACAddress      = $mac
                SwitchIP        = $SwitchIP
                SwitchPort      = $portIndex
                PortDescription = $portAlias
            }
        }
    }
    catch {
        Write-Warning "[$SwitchIP] SNMP walk failed: $($_.Exception.Message)"
    }
    finally {
        if ($snmp) {
            try { $snmp.Close() } catch { }
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($snmp) | Out-Null } catch { }
            $snmp = $null
        }
    }

    $result
}
