function Get-VBADComputer {
<#
.SYNOPSIS
    Resolve an IP address to an AD computer object (Layer 1).

.DESCRIPTION
    Queries Active Directory for a computer whose IPv4Address matches the supplied IP.
    On the first call within a session the function loads ALL computer objects from AD
    into a script-scope hashtable keyed by IPv4Address. All subsequent calls hit that
    hashtable -- never one LDAP query per IP.

    DC detection is also cached on first call from Get-ADDomainController.

    Prerequisites: $Context.ADAvailable must be $true. If context is not supplied the
    function emits Write-Warning and attempts the query anyway.

.PARAMETER IPAddress
    The RFC1918 / CGNAT / link-local IP address to look up.

.PARAMETER Context
    Environment context object from Get-VBEnrichmentContext. Provides the ADAvailable
    flag. If omitted the function runs without prerequisite gating.

.OUTPUTS
    [PSCustomObject] -- base layer result fields plus:
        Hostname          [string]
        OperatingSystem   [string]
        OSClass           [string]   Workstation | Server | DomainController
        OU                [string]
        DistinguishedName [string]

.EXAMPLE
    $ctx = Get-VBEnrichmentContext
    Get-VBADComputer -IPAddress '192.168.1.45' -Context $ctx

.EXAMPLE
    '192.168.1.45','192.168.1.46' | Get-VBADComputer -Context $ctx

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
        $LAYER_NUM  = 1
        $LAYER_NAME = 'AD'

        if (-not $Context) {
            Write-Warning "[$LAYER_NAME] No context provided -- running without prerequisite validation."
        }

        # One-shot cache: build on first call, reused for the lifetime of the session
        if ($null -eq $Script:VBAdComputerCache) {
            if ($Context -and -not $Context.ADAvailable) {
                # AD unavailable -- leave cache as sentinel so we skip cleanly each call
                $Script:VBAdComputerCache = @{}
                $Script:VBAdCacheBuilt    = $false
            }
            else {
                try {
                    Write-Verbose "[$LAYER_NAME] Building AD computer cache (one-shot)..."
                    $allComputers = Get-ADComputer -Filter * -Properties IPv4Address, OperatingSystem,
                        OperatingSystemVersion, DistinguishedName -ErrorAction Stop

                    $Script:VBAdComputerCache = @{}
                    foreach ($comp in $allComputers) {
                        if (-not [string]::IsNullOrWhiteSpace($comp.IPv4Address)) {
                            $Script:VBAdComputerCache[$comp.IPv4Address] = $comp
                        }
                    }

                    # Cache DC IPs separately
                    $Script:VBAdDCIPs = @{}
                    $dcList = Get-ADDomainController -Filter * -ErrorAction Stop
                    foreach ($dc in $dcList) {
                        if (-not [string]::IsNullOrWhiteSpace($dc.IPv4Address)) {
                            $Script:VBAdDCIPs[$dc.IPv4Address] = $true
                        }
                    }

                    $Script:VBAdCacheBuilt = $true
                    Write-Verbose "[$LAYER_NAME] Cache built: $($Script:VBAdComputerCache.Count) computers, $($Script:VBAdDCIPs.Count) DCs"
                }
                catch {
                    Write-Warning "[$LAYER_NAME] AD cache build failed: $($_.Exception.Message)"
                    $Script:VBAdComputerCache = @{}
                    $Script:VBAdCacheBuilt    = $false
                }
            }
        }
    }

    process {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        # Skip if AD unavailable and cache build was skipped
        if ($Context -and -not $Context.ADAvailable) {
            $sw.Stop()
            return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Skipped' -ExecutionMs $sw.ElapsedMilliseconds `
                -SkipReason 'ADUnavailable' `
                -Impact 'Cannot resolve domain-joined hosts via AD'
        }

        if (-not $Script:VBAdCacheBuilt) {
            $sw.Stop()
            return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Skipped' -ExecutionMs $sw.ElapsedMilliseconds `
                -SkipReason 'ADCacheBuildFailed' `
                -Impact 'AD cache could not be built -- check AD connectivity'
        }

        try {
            $comp = $Script:VBAdComputerCache[$IPAddress]

            if ($null -eq $comp) {
                $sw.Stop()
                return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                    -Status 'NoResult' -ExecutionMs $sw.ElapsedMilliseconds
            }

            # Determine OSClass
            $osClass = 'Workstation'
            if ($Script:VBAdDCIPs.ContainsKey($IPAddress)) {
                $osClass = 'DomainController'
            }
            elseif ($comp.OperatingSystem -match 'Server') {
                $osClass = 'Server'
            }

            # Extract OU from DistinguishedName -- everything after the first CN= component
            $ou = $null
            if ($comp.DistinguishedName -match '^CN=[^,]+,(.+)$') {
                $ou = $Matches[1]
            }

            $sw.Stop()
            New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Success' -ExecutionMs $sw.ElapsedMilliseconds `
                -ExtraFields @{
                    Hostname          = $comp.DNSHostName
                    OperatingSystem   = $comp.OperatingSystem
                    OSClass           = $osClass
                    OU                = $ou
                    DistinguishedName = $comp.DistinguishedName
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
