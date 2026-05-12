function Get-VBRTSPBanner {
<#
.SYNOPSIS
    Grab the RTSP banner from a device by sending an OPTIONS request to TCP 554 (Layer 8).

.DESCRIPTION
    Opens a raw TCP connection to port 554, sends a minimal RTSP OPTIONS request,
    reads up to 1 KB of response (or until 500 ms idle), and returns the first 512
    characters of the banner. The presence of an RTSP banner is a high-confidence
    signal for IP cameras and NVRs.

    The orchestrator gates this layer: it is only called when TCP layer 5 found
    port 554 open AND $Context.RTSPProbeEnabled is $true.

    All stream and TCP client resources are released in a finally block on every
    code path.

    Prerequisites: $Context.NetworkProbeEnabled and $Context.RTSPProbeEnabled.

.PARAMETER IPAddress
    The RFC1918 / CGNAT / link-local IP address to probe.

.PARAMETER Context
    Environment context from Get-VBEnrichmentContext. Provides
    NetworkProbeEnabled, RTSPProbeEnabled, and DefaultTimeoutMs.RTSP.

.PARAMETER TimeoutMs
    Read idle timeout in milliseconds. Defaults to $Context.DefaultTimeoutMs.RTSP (2000).

.OUTPUTS
    [PSCustomObject] -- base layer result fields plus:
        RTSPBanner [string]   first 512 chars of the RTSP response

.EXAMPLE
    $ctx = Get-VBEnrichmentContext
    Get-VBRTSPBanner -IPAddress '192.168.1.200' -Context $ctx

.EXAMPLE
    '192.168.1.200' | Get-VBRTSPBanner -Context $ctx

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
        $LAYER_NUM  = 8
        $LAYER_NAME = 'RTSP'

        if (-not $Context) {
            Write-Warning "[$LAYER_NAME] No context provided -- running without prerequisite validation."
        }

        if ($PSBoundParameters.ContainsKey('TimeoutMs') -eq $false -and $Context -and $Context.DefaultTimeoutMs) {
            $TimeoutMs = $Context.DefaultTimeoutMs.RTSP
        }
    }

    process {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        if ($Context -and (-not $Context.NetworkProbeEnabled -or -not $Context.RTSPProbeEnabled)) {
            $sw.Stop()
            return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Skipped' -ExecutionMs $sw.ElapsedMilliseconds `
                -SkipReason 'RTSPProbeDisabled' `
                -Impact 'No RTSP banner -- cameras may be classified by OUI only'
        }

        $client = $null
        $stream = $null

        try {
            $client = New-Object System.Net.Sockets.TcpClient
            $connectResult = $client.BeginConnect($IPAddress, 554, $null, $null)
            $connected = $connectResult.AsyncWaitHandle.WaitOne($TimeoutMs, $false)

            if (-not $connected -or -not $client.Connected) {
                $sw.Stop()
                return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                    -Status 'NoResult' -ExecutionMs $sw.ElapsedMilliseconds
            }

            $client.EndConnect($connectResult)
            $stream = $client.GetStream()
            $stream.ReadTimeout  = $TimeoutMs
            $stream.WriteTimeout = $TimeoutMs

            # Send minimal RTSP OPTIONS request
            $request = "OPTIONS rtsp://${IPAddress}:554/ RTSP/1.0`r`nCSeq: 1`r`n`r`n"
            $requestBytes = [System.Text.Encoding]::ASCII.GetBytes($request)
            $stream.Write($requestBytes, 0, $requestBytes.Length)

            # Read response -- up to 1 KB, stop on idle or timeout
            $buffer    = New-Object byte[] 1024
            $totalRead = 0
            try {
                $totalRead = $stream.Read($buffer, 0, $buffer.Length)
            }
            catch [System.IO.IOException] {
                # Read timeout is expected -- we got partial data or nothing
            }

            if ($totalRead -eq 0) {
                $sw.Stop()
                return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                    -Status 'NoResult' -ExecutionMs $sw.ElapsedMilliseconds
            }

            $banner = [System.Text.Encoding]::ASCII.GetString($buffer, 0, $totalRead)
            # Trim to 512 chars and strip control characters
            $banner = ($banner -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', '').Substring(
                0, [math]::Min(512, $banner.Length)
            ).Trim()

            $sw.Stop()
            New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Success' -ExecutionMs $sw.ElapsedMilliseconds `
                -ExtraFields @{ RTSPBanner = $banner }
        }
        catch {
            $sw.Stop()
            New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Failed' -ExecutionMs $sw.ElapsedMilliseconds `
                -ErrorDetail $_.Exception.Message
        }
        finally {
            if ($stream) { try { $stream.Close() } catch { } }
            if ($client) { try { $client.Close() } catch { } ; try { $client.Dispose() } catch { } }
        }
    }
}
