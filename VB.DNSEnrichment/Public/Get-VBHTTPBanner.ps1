function Get-VBHTTPBanner {
<#
.SYNOPSIS
    Grab HTTP page title and Server header from a device's web interface (Layer 6).

.DESCRIPTION
    Tries ports in order: 80 -> 8080 -> 443 -> 8443. Stops at the first successful
    response. Extracts the HTML <title> element and the Server: response header.

    PS 6+ uses Invoke-WebRequest -SkipCertificateCheck natively.
    PS 5.1 applies the ServerCertificateValidationCallback workaround before each
    HTTPS request and always resets it to $null in a finally block -- failure to
    reset breaks all subsequent HTTPS calls in the session.

    The orchestrator gates this layer: it is only called if TCP layer 5 found at
    least one of 80, 8080, 443, or 8443 open. This function itself does not
    re-check open ports; it tries each port unconditionally (caller-gates).

    Prerequisites: $Context.NetworkProbeEnabled must be $true.

.PARAMETER IPAddress
    The RFC1918 / CGNAT / link-local IP address to probe.

.PARAMETER OpenPortsList
    Array of open ports from Get-VBTCPFingerprint. Used to skip ports known to
    be closed without attempting a connection. Optional -- if omitted all four
    ports are tried.

.PARAMETER Context
    Environment context from Get-VBEnrichmentContext. Provides
    NetworkProbeEnabled, CanSkipCertCheck, and DefaultTimeoutMs.HTTP.

.PARAMETER TimeoutMs
    HTTP request timeout in milliseconds. Defaults to $Context.DefaultTimeoutMs.HTTP (3000).

.OUTPUTS
    [PSCustomObject] -- base layer result fields plus:
        HTTPTitle  [string]
        HTTPServer [string]
        HTTPPort   [int]
        HTTPScheme [string]   http | https

.EXAMPLE
    $ctx = Get-VBEnrichmentContext
    Get-VBHTTPBanner -IPAddress '192.168.1.45' -Context $ctx

.EXAMPLE
    Get-VBHTTPBanner -IPAddress '192.168.1.45' -OpenPortsList @(80,443) -Context $ctx

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
        [int[]]$OpenPortsList = @(),

        [Parameter()]
        [PSCustomObject]$Context,

        [Parameter()]
        [int]$TimeoutMs = 3000
    )

    begin {
        $LAYER_NUM  = 6
        $LAYER_NAME = 'HTTP'

        if (-not $Context) {
            Write-Warning "[$LAYER_NAME] No context provided -- running without prerequisite validation."
        }

        if ($PSBoundParameters.ContainsKey('TimeoutMs') -eq $false -and $Context -and $Context.DefaultTimeoutMs) {
            $TimeoutMs = $Context.DefaultTimeoutMs.HTTP
        }

        $canSkipCert = $Context -and $Context.CanSkipCertCheck
        $psMajor     = if ($Context) { $Context.PSMajor } else { $PSVersionTable.PSVersion.Major }
    }

    process {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        if ($Context -and -not $Context.NetworkProbeEnabled) {
            $sw.Stop()
            return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Skipped' -ExecutionMs $sw.ElapsedMilliseconds `
                -SkipReason 'NetworkProbeDisabled' `
                -Impact 'No HTTP banner -- web-based device identification unavailable'
        }

        # Determine which ports to try based on what TCP found open
        $httpPorts = @(
            @{ Port = 80;   Scheme = 'http'  }
            @{ Port = 8080; Scheme = 'http'  }
            @{ Port = 443;  Scheme = 'https' }
            @{ Port = 8443; Scheme = 'https' }
        )

        if ($OpenPortsList.Count -gt 0) {
            $httpPorts = @($httpPorts | Where-Object { $OpenPortsList -contains $_.Port })
        }

        if ($httpPorts.Count -eq 0) {
            $sw.Stop()
            return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'NoResult' -ExecutionMs $sw.ElapsedMilliseconds
        }

        $timeoutSec = [math]::Max(1, [int][math]::Ceiling($TimeoutMs / 1000))

        foreach ($entry in $httpPorts) {
            $port   = $entry.Port
            $scheme = $entry.Scheme
            $uri    = "${scheme}://${IPAddress}:${port}/"

            $callbackSet = $false
            try {
                $response = $null

                if ($psMajor -ge 6) {
                    $response = Invoke-WebRequest -Uri $uri -TimeoutSec $timeoutSec `
                        -SkipCertificateCheck -UseBasicParsing -ErrorAction Stop
                }
                else {
                    # PS 5.1 -- apply cert bypass callback for HTTPS, always reset in finally
                    if ($scheme -eq 'https') {
                        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                        [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
                        $callbackSet = $true
                    }
                    $response = Invoke-WebRequest -Uri $uri -TimeoutSec $timeoutSec `
                        -UseBasicParsing -ErrorAction Stop
                }

                # Extract <title>
                $title = $null
                if ($response.Content -match '<title[^>]*>(.*?)</title>') {
                    $title = [System.Web.HttpUtility]::HtmlDecode($Matches[1]).Trim()
                    # Collapse whitespace
                    $title = $title -replace '\s+', ' '
                }

                # Extract Server header
                $serverHeader = $null
                if ($response.Headers.ContainsKey('Server')) {
                    $serverHeader = ($response.Headers['Server'] -join ', ')
                }

                $sw.Stop()
                return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                    -Status 'Success' -ExecutionMs $sw.ElapsedMilliseconds `
                    -ExtraFields @{
                        HTTPTitle  = $title
                        HTTPServer = $serverHeader
                        HTTPPort   = $port
                        HTTPScheme = $scheme
                    }
            }
            catch {
                Write-Verbose "[$LAYER_NAME] $uri failed: $($_.Exception.Message)"
                # Try next port
            }
            finally {
                if ($callbackSet) {
                    [Net.ServicePointManager]::ServerCertificateValidationCallback = $null
                }
            }
        }

        # All ports tried -- none responded successfully
        $sw.Stop()
        New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
            -Status 'NoResult' -ExecutionMs $sw.ElapsedMilliseconds
    }
}
