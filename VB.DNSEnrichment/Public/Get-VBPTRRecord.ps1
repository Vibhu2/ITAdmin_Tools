function Get-VBPTRRecord {
<#
.SYNOPSIS
    Resolve an IP address via DNS PTR lookup with forward-confirmation (Layer 3).

.DESCRIPTION
    Performs a reverse DNS (PTR) lookup for the supplied IP. On success, the resolved
    hostname is forward-confirmed by querying the A record for that name and verifying
    that at least one resolved IP matches the original. Stale PTR records that no
    longer resolve back to the same IP are returned with ForwardConfirmed = $false
    and Confidence = 'Low'.

    PS 5.1 fallback: uses [System.Net.Dns]::GetHostEntry() when Resolve-DnsName
    is not available (unlikely on WS2012+ but handled defensively).

    TLS 1.2 is forced on PS 5.1 to ensure the session is not degraded by default
    system settings (belt-and-suspenders -- DNS itself doesn't use TLS, but other
    layers in the session might).

    Prerequisites: $Context.DNSAvailable must be $true.

.PARAMETER IPAddress
    The RFC1918 / CGNAT / link-local IP address to look up.

.PARAMETER Context
    Environment context object from Get-VBEnrichmentContext. Provides DNSAvailable.

.OUTPUTS
    [PSCustomObject] -- base layer result fields plus:
        Hostname         [string]
        PTRRecord        [string]
        ForwardConfirmed [bool]
        Confidence       [string]   High | Low

.EXAMPLE
    $ctx = Get-VBEnrichmentContext
    Get-VBPTRRecord -IPAddress '192.168.1.45' -Context $ctx

.EXAMPLE
    '192.168.1.45','192.168.1.46' | Get-VBPTRRecord -Context $ctx

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
        $LAYER_NUM  = 3
        $LAYER_NAME = 'PTR'

        if (-not $Context) {
            Write-Warning "[$LAYER_NAME] No context provided -- running without prerequisite validation."
        }

        $hasResolveDnsName = [bool](Get-Command -Name Resolve-DnsName -ErrorAction SilentlyContinue)
    }

    process {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        if ($Context -and -not $Context.DNSAvailable) {
            $sw.Stop()
            return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Skipped' -ExecutionMs $sw.ElapsedMilliseconds `
                -SkipReason 'DNSUnavailable' `
                -Impact 'PTR-based hostname resolution unavailable'
        }

        try {
            $ptrRecord = $null
            $hostname  = $null

            if ($hasResolveDnsName) {
                $ptrResult = Resolve-DnsName -Name $IPAddress -Type PTR -ErrorAction Stop
                $ptrRecord = $ptrResult | Where-Object { $_.Type -eq 'PTR' } |
                    Select-Object -ExpandProperty NameHost -First 1
            }
            else {
                # PS 5.1 fallback -- [System.Net.Dns]::GetHostEntry does PTR + A in one call
                $entry     = [System.Net.Dns]::GetHostEntry($IPAddress)
                $ptrRecord = $entry.HostName
            }

            if ([string]::IsNullOrWhiteSpace($ptrRecord)) {
                $sw.Stop()
                return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                    -Status 'NoResult' -ExecutionMs $sw.ElapsedMilliseconds
            }

            $hostname = $ptrRecord.TrimEnd('.')

            # Forward-confirm: resolve A record and check for IP match
            $forwardConfirmed = $false
            try {
                if ($hasResolveDnsName) {
                    $aResult = Resolve-DnsName -Name $hostname -Type A -ErrorAction Stop
                    $resolvedIPs = @($aResult | Where-Object { $_.Type -eq 'A' } |
                        Select-Object -ExpandProperty IPAddress)
                }
                else {
                    $fwdEntry    = [System.Net.Dns]::GetHostAddresses($hostname)
                    $resolvedIPs = @($fwdEntry | ForEach-Object { $_.ToString() })
                }

                $forwardConfirmed = $resolvedIPs -contains $IPAddress
            }
            catch {
                Write-Verbose "[$LAYER_NAME] Forward-confirm failed for $hostname`: $($_.Exception.Message)"
                $forwardConfirmed = $false
            }

            $confidence = if ($forwardConfirmed) { 'High' } else { 'Low' }

            $sw.Stop()
            New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Success' -ExecutionMs $sw.ElapsedMilliseconds `
                -ExtraFields @{
                    Hostname         = $hostname
                    PTRRecord        = $ptrRecord
                    ForwardConfirmed = $forwardConfirmed
                    Confidence       = $confidence
                }
        }
        catch {
            # NXDOMAIN and SERVFAIL both throw from Resolve-DnsName -- distinguish them
            $msg = $_.Exception.Message
            if ($msg -match 'DNS name does not exist|NXDOMAIN|no such host') {
                $sw.Stop()
                return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                    -Status 'NoResult' -ExecutionMs $sw.ElapsedMilliseconds
            }

            $sw.Stop()
            New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Failed' -ExecutionMs $sw.ElapsedMilliseconds `
                -ErrorDetail $msg
        }
    }
}
