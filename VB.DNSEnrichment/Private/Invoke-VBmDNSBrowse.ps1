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

        $null = Wait-Job -Job $job -Timeout 5
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

                $null = Wait-Job -Job $resolveJob -Timeout 3
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
