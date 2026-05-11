function Get-VBTCPFingerprint {
<#
.SYNOPSIS
    Scan 24 well-known ports concurrently to fingerprint a device (Layer 5).

.DESCRIPTION
    Probes all 24 ports in parallel using async BeginConnect + WaitOne on a
    single IP. Total wall time is approximately equal to the timeout value
    regardless of how many ports are scanned -- the connections race concurrently.

    All TcpClient objects are released in a finally block on every code path.

    Port list (24 ports) as defined in the module design spec:
        22    SSH          5060  SIP
        23    Telnet       5061  SIP/TLS
        80    HTTP         5985  WinRM HTTP
        135   RPC          8000  Alt HTTP
        161   SNMP         8080  Alt HTTP
        443   HTTPS        8443  Alt HTTPS
        445   SMB          9100  JetDirect
        515   LPD          9443  VMware
        548   AFP          37777 Dahua NVR
        554   RTSP         62078 iOS
        631   IPP          1883  MQTT
        902   VMware       3389  RDP

    Prerequisites: $Context.NetworkProbeEnabled must be $true.

.PARAMETER IPAddress
    The RFC1918 / CGNAT / link-local IP address to scan.

.PARAMETER Context
    Environment context from Get-VBEnrichmentContext. Provides
    NetworkProbeEnabled and DefaultTimeoutMs.TCP.

.PARAMETER TimeoutMs
    Per-port connection timeout in milliseconds. Defaults to
    $Context.DefaultTimeoutMs.TCP (300 ms). Override here for slower networks.

.OUTPUTS
    [PSCustomObject] -- base layer result fields plus:
        OpenPorts     [string]   comma-separated open port numbers
        OpenPortsList [int[]]    array of open port numbers
        ScanDurationMs [int]

.EXAMPLE
    $ctx = Get-VBEnrichmentContext
    Get-VBTCPFingerprint -IPAddress '192.168.1.45' -Context $ctx

.EXAMPLE
    '192.168.1.45','192.168.1.46' | Get-VBTCPFingerprint -Context $ctx

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
        [PSCustomObject]$Context,

        [Parameter()]
        [int]$TimeoutMs = 300
    )

    begin {
        $LAYER_NUM  = 5
        $LAYER_NAME = 'TCP'

        if (-not $Context) {
            Write-Warning "[$LAYER_NAME] No context provided -- running without prerequisite validation."
        }

        # Resolve timeout from context if not explicitly overridden
        if ($PSBoundParameters.ContainsKey('TimeoutMs') -eq $false -and $Context -and $Context.DefaultTimeoutMs) {
            $TimeoutMs = $Context.DefaultTimeoutMs.TCP
        }

        $FingerprintPorts = @(
            22, 23, 80, 135, 161, 443, 445, 515, 548, 554,
            631, 902, 1883, 3389, 5060, 5061, 5985, 8000,
            8080, 8443, 9100, 9443, 37777, 62078
        )
    }

    process {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        if ($Context -and -not $Context.NetworkProbeEnabled) {
            $sw.Stop()
            return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Skipped' -ExecutionMs $sw.ElapsedMilliseconds `
                -SkipReason 'NetworkProbeDisabled' `
                -Impact 'No port-based device class signals available'
        }

        $clients = @()
        $handles = @()

        try {
            # Launch all connections concurrently
            foreach ($port in $FingerprintPorts) {
                $client = New-Object System.Net.Sockets.TcpClient
                $clients += $client
                try {
                    $handle = $client.BeginConnect($IPAddress, $port, $null, $null)
                    $handles += [PSCustomObject]@{ Port = $port; Client = $client; Handle = $handle }
                }
                catch {
                    # BeginConnect itself failed (bad IP format etc) -- skip this port
                    $handles += [PSCustomObject]@{ Port = $port; Client = $client; Handle = $null }
                }
            }

            # Wait for all and collect open ports
            $openPorts = New-Object System.Collections.Generic.List[int]
            foreach ($entry in $handles) {
                if ($null -eq $entry.Handle) { continue }
                try {
                    $completed = $entry.Handle.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
                    if ($completed -and $entry.Client.Connected) {
                        $openPorts.Add($entry.Port)
                    }
                }
                catch {
                    # Connection refused or timeout -- port closed, ignore
                }
            }

            $sw.Stop()

            if ($openPorts.Count -eq 0) {
                return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                    -Status 'NoResult' -ExecutionMs $sw.ElapsedMilliseconds `
                    -ExtraFields @{
                        OpenPorts      = ''
                        OpenPortsList  = @()
                        ScanDurationMs = [int]$sw.ElapsedMilliseconds
                    }
            }

            $sortedPorts = @($openPorts | Sort-Object)
            New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Success' -ExecutionMs $sw.ElapsedMilliseconds `
                -ExtraFields @{
                    OpenPorts      = $sortedPorts -join ','
                    OpenPortsList  = $sortedPorts
                    ScanDurationMs = [int]$sw.ElapsedMilliseconds
                }
        }
        catch {
            $sw.Stop()
            New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Failed' -ExecutionMs $sw.ElapsedMilliseconds `
                -ErrorDetail $_.Exception.Message
        }
        finally {
            # Always release all TcpClient objects
            foreach ($client in $clients) {
                try { $client.Close() } catch { }
                try { $client.Dispose() } catch { }
            }
        }
    }
}
