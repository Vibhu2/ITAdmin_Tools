function Get-VBSNMPIdentity {
<#
.SYNOPSIS
    Query a device via SNMP v1/v2c to retrieve sysDescr, sysName, and sysLocation (Layer 7).

.DESCRIPTION
    Uses the Windows-native olePrn.OleSNMP COM object (no third-party module required).
    Iterates $Context.SNMPCommunityStrings in order -- first string that gets a
    response wins. Records which community string succeeded in CommunityUsed.

    OIDs queried:
        1.3.6.1.2.1.1.1.0  sysDescr    -> SNMPDescr
        1.3.6.1.2.1.1.5.0  sysName     -> Hostname
        1.3.6.1.2.1.1.6.0  sysLocation -> Location

    The COM object is always released in a finally block on every code path.
    Timeout is enforced by setting the OleSNMP Timeout property before Open().

    Prerequisites: $Context.SNMPAvailable must be $true.

.PARAMETER IPAddress
    The RFC1918 / CGNAT / link-local IP address to query.

.PARAMETER Context
    Environment context from Get-VBEnrichmentContext. Provides SNMPAvailable,
    SNMPCommunityStrings, and DefaultTimeoutMs.SNMP.

.PARAMETER TimeoutMs
    SNMP response timeout in milliseconds per community string attempt.
    Defaults to $Context.DefaultTimeoutMs.SNMP (2000).

.OUTPUTS
    [PSCustomObject] -- base layer result fields plus:
        Hostname      [string]   sysName OID value
        SNMPDescr     [string]   sysDescr OID value
        Location      [string]   sysLocation OID value
        CommunityUsed [string]   the community string that succeeded
        SNMPVersion   [string]   'v1' (olePrn always uses v1/v2c)

.EXAMPLE
    $ctx = Get-VBEnrichmentContext -SNMPCommunityStrings 'public','readonly'
    Get-VBSNMPIdentity -IPAddress '192.168.1.200' -Context $ctx

.EXAMPLE
    '192.168.1.200','192.168.1.201' | Get-VBSNMPIdentity -Context $ctx

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
        [PSCustomObject]$Context,

        [Parameter()]
        [int]$TimeoutMs = 2000
    )

    begin {
        $LAYER_NUM  = 7
        $LAYER_NAME = 'SNMP'

        if (-not $Context) {
            Write-Warning "[$LAYER_NAME] No context provided -- running without prerequisite validation."
        }

        if ($PSBoundParameters.ContainsKey('TimeoutMs') -eq $false -and $Context -and $Context.DefaultTimeoutMs) {
            $TimeoutMs = $Context.DefaultTimeoutMs.SNMP
        }

        $communityStrings = if ($Context -and $Context.SNMPCommunityStrings.Count -gt 0) {
            @($Context.SNMPCommunityStrings | ForEach-Object {
                [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($_))
            })
        }
        else {
            @('public')
        }

        # OleSNMP Timeout property is in milliseconds
        $snmpTimeoutMs = $TimeoutMs
    }

    process {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        if ($Context -and -not $Context.SNMPAvailable) {
            $sw.Stop()
            return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Skipped' -ExecutionMs $sw.ElapsedMilliseconds `
                -SkipReason 'SNMPUnavailable' `
                -Impact 'No SNMP sysName or sysLocation -- printers and network devices may be unidentified'
        }

        $snmp = $null
        $lastError = $null

        foreach ($community in $communityStrings) {
            try {
                $snmp = New-Object -ComObject olePrn.OleSNMP -ErrorAction Stop
                $snmp.Timeout = $snmpTimeoutMs
                $snmp.Open($IPAddress, $community, 2, $snmpTimeoutMs)

                $sysDescr    = $snmp.Get('1.3.6.1.2.1.1.1.0')
                $sysName     = $snmp.Get('1.3.6.1.2.1.1.5.0')
                $sysLocation = $snmp.Get('1.3.6.1.2.1.1.6.0')

                $snmpHostname  = if ([string]::IsNullOrWhiteSpace($sysName))     { $null } else { $sysName.Trim() }
                $snmpDescr     = if ([string]::IsNullOrWhiteSpace($sysDescr))    { $null } else { $sysDescr.Trim() }
                $snmpLocation  = if ([string]::IsNullOrWhiteSpace($sysLocation)) { $null } else { $sysLocation.Trim() }

                $sw.Stop()
                return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                    -Status 'Success' -ExecutionMs $sw.ElapsedMilliseconds `
                    -ExtraFields @{
                        Hostname      = $snmpHostname
                        SNMPDescr     = $snmpDescr
                        Location      = $snmpLocation
                        CommunityUsed = $community
                        SNMPVersion   = 'v1'
                    }
            }
            catch {
                $lastError = $_.Exception.Message
                Write-Verbose "[$LAYER_NAME] $IPAddress community '$community' failed: $lastError"
            }
            finally {
                if ($snmp) {
                    try { $snmp.Close() }   catch { }
                    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($snmp) | Out-Null } catch { }
                    $snmp = $null
                }
            }
        }

        # All community strings exhausted
        $sw.Stop()
        if ($lastError -match 'timeout|timed out|no response' -or $lastError -match '0x80004005') {
            # Timeout = device not running SNMP or port filtered
            New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'NoResult' -ExecutionMs $sw.ElapsedMilliseconds
        }
        else {
            New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Failed' -ExecutionMs $sw.ElapsedMilliseconds `
                -ErrorDetail $lastError
        }
    }
}
