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
                    $ip      = $Matches[1]
                    $mac     = $arpWalk[$key]
                    $macNorm = ConvertTo-VBNormalisedMAC -MACAddress $mac
                    if (-not [string]::IsNullOrWhiteSpace($macNorm)) {
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
