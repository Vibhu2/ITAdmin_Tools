function Get-VBDHCPLease {
<#
.SYNOPSIS
    Resolve an IP address to a DHCP lease record (Layer 2).

.DESCRIPTION
    On the first call within a session the function enumerates ALL active and
    inactive leases across every scope in $Context.DHCPScopeIds and loads them
    into a script-scope hashtable keyed by IP address. All subsequent calls
    perform a hashtable lookup -- no further DHCP server calls are made.

    If $Context.DHCPIsLocal is $true the function queries the local server
    (no -ComputerName). Otherwise it uses $Context.DHCPServer.

    Expired leases are still returned (Status = 'Success') with IsLeaseExpired
    set to $true so the orchestrator can flag DHCP churn risk.

    Prerequisites: $Context.DHCPAvailable must be $true.

.PARAMETER IPAddress
    The RFC1918 / CGNAT / link-local IP address to look up.

.PARAMETER Context
    Environment context object from Get-VBEnrichmentContext. Provides
    DHCPAvailable, DHCPIsLocal, DHCPServer, and DHCPScopeIds.

.OUTPUTS
    [PSCustomObject] -- base layer result fields plus:
        Hostname        [string]
        MACAddress      [string]
        LeaseExpiry     [datetime]
        IsLeaseExpired  [bool]
        ScopeId         [string]

.EXAMPLE
    $ctx = Get-VBEnrichmentContext -DHCPServer 'dhcp01.corp.local'
    Get-VBDHCPLease -IPAddress '192.168.1.45' -Context $ctx

.EXAMPLE
    '192.168.1.45','192.168.1.46' | Get-VBDHCPLease -Context $ctx

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
        $LAYER_NUM  = 2
        $LAYER_NAME = 'DHCP'

        if (-not $Context) {
            Write-Warning "[$LAYER_NAME] No context provided -- running without prerequisite validation."
        }

        if ($null -eq $Script:VBDhcpLeaseCache) {
            if ($Context -and -not $Context.DHCPAvailable) {
                $Script:VBDhcpLeaseCache  = @{}
                $Script:VBDhcpCacheBuilt  = $false
            }
            else {
                try {
                    Write-Verbose "[$LAYER_NAME] Building DHCP lease cache (one-shot)..."

                    $serverSplat = @{}
                    if ($Context -and -not $Context.DHCPIsLocal -and $Context.DHCPServer) {
                        $serverSplat['ComputerName'] = $Context.DHCPServer
                    }

                    $scopeIds = @()
                    if ($Context -and $Context.DHCPScopeIds.Count -gt 0) {
                        $scopeIds = $Context.DHCPScopeIds
                    }
                    else {
                        # Discover all scopes if none were supplied
                        $scopeIds = @(
                            Get-DhcpServerv4Scope @serverSplat -ErrorAction Stop |
                                ForEach-Object { $_.ScopeId.ToString() }
                        )
                    }

                    $Script:VBDhcpLeaseCache = @{}
                    foreach ($scopeId in $scopeIds) {
                        $leases = Get-DhcpServerv4Lease -ScopeId $scopeId @serverSplat `
                            -ErrorAction SilentlyContinue
                        foreach ($lease in $leases) {
                            if (-not [string]::IsNullOrWhiteSpace($lease.IPAddress)) {
                                $Script:VBDhcpLeaseCache[$lease.IPAddress.ToString()] = $lease
                            }
                        }
                    }

                    $Script:VBDhcpCacheBuilt = $true
                    Write-Verbose "[$LAYER_NAME] Cache built: $($Script:VBDhcpLeaseCache.Count) leases across $($scopeIds.Count) scope(s)"
                }
                catch {
                    Write-Warning "[$LAYER_NAME] DHCP cache build failed: $($_.Exception.Message)"
                    $Script:VBDhcpLeaseCache = @{}
                    $Script:VBDhcpCacheBuilt = $false
                }
            }
        }
    }

    process {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        if ($Context -and -not $Context.DHCPAvailable) {
            $sw.Stop()
            return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Skipped' -ExecutionMs $sw.ElapsedMilliseconds `
                -SkipReason 'DHCPUnavailable' `
                -Impact 'No DHCP-derived hostname or MAC address'
        }

        if (-not $Script:VBDhcpCacheBuilt) {
            $sw.Stop()
            return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Skipped' -ExecutionMs $sw.ElapsedMilliseconds `
                -SkipReason 'DHCPCacheBuildFailed' `
                -Impact 'DHCP cache could not be built -- check DHCP server connectivity'
        }

        try {
            $lease = $Script:VBDhcpLeaseCache[$IPAddress]

            if ($null -eq $lease) {
                $sw.Stop()
                return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                    -Status 'NoResult' -ExecutionMs $sw.ElapsedMilliseconds
            }

            $leaseExpiry   = $lease.LeaseExpiryTime
            $isExpired     = ($null -ne $leaseExpiry -and $leaseExpiry -lt (Get-Date))
            $macNormalised = ConvertTo-VBNormalisedMAC -MACAddress $lease.ClientId

            $sw.Stop()
            New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Success' -ExecutionMs $sw.ElapsedMilliseconds `
                -ExtraFields @{
                    Hostname       = $lease.HostName
                    MACAddress     = $lease.ClientId
                    MACNormalised  = $macNormalised
                    LeaseExpiry    = $leaseExpiry
                    IsLeaseExpired = $isExpired
                    ScopeId        = $lease.ScopeId.ToString()
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
