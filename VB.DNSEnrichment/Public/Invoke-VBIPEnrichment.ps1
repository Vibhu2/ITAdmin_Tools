function Invoke-VBIPEnrichment {
<#
.SYNOPSIS
    Orchestrate the full enrichment pipeline against a list of private IP addresses.

.DESCRIPTION
    Runs all enabled layers in priority order for each IP, merges results, classifies
    each device, persists to SQLite, and returns one enrichment object per IP.

    Round 2 ships passive-only mode (layers 1-4 + classification + storage).
    Active probe layers (5-10) are called stubs -- they return Skipped until Round 3.

    Execution flow:
        1.  Validate context; warn if missing.
        2.  Validate IP list -- skip public addresses with Write-Warning.
        3.  Load existing rows from SQLite for supplied IPs.
        4.  Decide which IPs to probe (missing, stale, unresolved, or -ForceRefresh).
        5.  For each IP to probe (always sequential in Round 2):
              Step 1  Get-VBADComputer    -> may set Hostname, OSClass
              Step 2  Get-VBDHCPLease    -> may set Hostname, MAC (always runs for MAC)
              Step 3  Get-VBPTRRecord    -> may set Hostname if not already resolved
              Step 4  Get-VBARPEntry     -> may set MAC if not already known
              Steps 5-11  Skipped (active layers -- Round 3)
              Step 12 Resolve-VBDeviceClass
        6.  Compare result to existing SQLite row; write EnrichmentHistory on change.
        7.  Upsert row into SQLite.
        8.  Emit object (stream immediately if -PassThru; collect otherwise).
        9.  Emit summary via Write-Verbose.

.PARAMETER IPAddress
    One or more private IP addresses to enrich. Accepts pipeline input.

.PARAMETER Context
    Environment context from Get-VBEnrichmentContext. Mandatory for full operation;
    omitting emits a warning and uses degraded defaults.

.PARAMETER SkipActiveProbes
    Force-skip layers 5-10 even in Round 3+. Useful for scheduled passive-only runs.

.PARAMETER ForceRefresh
    Ignore SQLite cache -- re-probe every IP even if recently enriched.

.PARAMETER StaleThresholdHours
    Re-probe rows whose UpdatedAt is older than this many hours. Default 168 (7 days).

.PARAMETER PassThru
    Emit each result immediately as it completes rather than collecting and emitting at end.

.PARAMETER ProgressUpdateInterval
    Minimum seconds between Write-Progress updates. Default 1.

.OUTPUTS
    [PSCustomObject[]] -- one enrichment object per IP (see design spec section 14).

.EXAMPLE
    $ctx = Get-VBEnrichmentContext
    '192.168.1.45','192.168.1.46' | Invoke-VBIPEnrichment -Context $ctx

.EXAMPLE
    Import-Csv ips.csv | Select-Object -ExpandProperty IPAddress |
        Invoke-VBIPEnrichment -Context $ctx -ForceRefresh -PassThru

.NOTES
    Version:      1.0.0
    MinPSVersion: 5.1
    Author:       VB
    ChangeLog:
        1.0.0 -- 2026-05-11 -- Round 2: passive layers only (1-4)
#>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string[]]$IPAddress,

        [Parameter()]
        [PSCustomObject]$Context,

        [Parameter()]
        [switch]$SkipActiveProbes,

        [Parameter()]
        [switch]$ForceRefresh,

        [Parameter()]
        [int]$StaleThresholdHours = 168,

        [Parameter()]
        [switch]$PassThru,

        [Parameter()]
        [int]$ProgressUpdateInterval = 1
    )

    begin {
        $ErrorActionPreference = 'Stop'

        if (-not $Context) {
            Write-Warning "[Orchestrator] No context provided -- running without prerequisite gating. Results may be incomplete."
        }

        $dbPath     = if ($Context) { $Context.DatabasePath } else { Join-Path $env:LOCALAPPDATA 'VB.DNSEnrichment\enrichment.db' }
        $dbEnabled  = $Context -and $Context.DatabaseInitialized
        $allIPs     = New-Object System.Collections.Generic.List[string]
        $results    = New-Object System.Collections.Generic.List[PSCustomObject]
        $runStart   = Get-Date
        $sw         = [System.Diagnostics.Stopwatch]::StartNew()
    }

    process {
        foreach ($ip in $IPAddress) {
            $allIPs.Add($ip)
        }
    }

    end {
        # --- Step 1: Validate and filter IPs ---
        $validIPs = New-Object System.Collections.Generic.List[string]
        foreach ($ip in $allIPs) {
            if (Test-VBPrivateIP -IPAddress $ip) {
                $validIPs.Add($ip)
            }
            else {
                Write-Warning "[Orchestrator] Skipping public IP: $ip"
            }
        }

        if ($validIPs.Count -eq 0) {
            Write-Verbose "[Orchestrator] No valid private IPs to process."
            return
        }

        # --- Step 2: Load existing SQLite rows ---
        $cachedRows = @{}
        if ($dbEnabled) {
            try {
                $placeholders = ($validIPs | ForEach-Object { "'$_'" }) -join ','
                $rows = Invoke-VBSqliteCommand -DatabasePath $dbPath `
                    -Query "SELECT * FROM Enrichment WHERE IPAddress IN ($placeholders)"
                foreach ($row in $rows) {
                    $cachedRows[$row.IPAddress] = $row
                }
                Write-Verbose "[Orchestrator] Loaded $($cachedRows.Count) existing rows from SQLite"
            }
            catch {
                Write-Warning "[Orchestrator] SQLite read failed: $($_.Exception.Message)"
            }
        }

        # --- Step 3: Decide which IPs to probe ---
        $toProbe = New-Object System.Collections.Generic.List[string]
        $fromCache = New-Object System.Collections.Generic.List[PSCustomObject]
        $staleThreshold = (Get-Date).AddHours(-$StaleThresholdHours)

        foreach ($ip in $validIPs) {
            if ($ForceRefresh) {
                $toProbe.Add($ip)
                continue
            }
            $existing = $cachedRows[$ip]
            if ($null -eq $existing) {
                $toProbe.Add($ip)
            }
            elseif ($existing.IsResolved -eq 0) {
                $toProbe.Add($ip)
            }
            elseif ($existing.UpdatedAt -and ([datetime]$existing.UpdatedAt) -lt $staleThreshold) {
                $toProbe.Add($ip)
            }
            else {
                # Return from cache
                $fromCache.Add((Invoke-VBBuildEnrichmentObject -Row $existing -FromCache $true))
            }
        }

        Write-Verbose "[Orchestrator] $($validIPs.Count) IPs total: $($toProbe.Count) to probe, $($fromCache.Count) from cache"

        # Emit cached results
        foreach ($cached in $fromCache) {
            if ($PassThru) { $cached } else { $results.Add($cached) }
        }

        # --- Step 4: Probe each IP ---
        $total   = $toProbe.Count
        $current = 0

        foreach ($ip in $toProbe) {
            $current++
            $ipSw = [System.Diagnostics.Stopwatch]::StartNew()
            $layerTrace = New-Object System.Collections.Generic.List[PSCustomObject]

            # State accumulator for this IP
            $state = @{
                IPAddress       = $ip
                Hostname        = $null
                HostnameSource  = $null
                IsResolved      = $false
                MACAddress      = $null
                MACNormalised   = $null
                OSClass         = $null
                OperatingSystem = $null
                OU              = $null
                OpenPorts       = $null
                HTTPTitle       = $null
                HTTPServer      = $null
                SNMPDescr       = $null
                RTSPBanner      = $null
                MDNSServiceType = $null
                Location        = $null
                LeaseExpiry     = $null
                VendorDeviceClass = $null
                OUIVendor       = $null
            }

            # ---- Step 1: AD ----
            Write-VBEnrichmentProgress -Current $current -Total $total -IPAddress $ip `
                -StepNumber 1 -LayerName 'AD' -ElapsedMs $sw.ElapsedMilliseconds
            Write-Verbose "[$ip] Step 1 AD"

            $adResult = Get-VBADComputer -IPAddress $ip -Context $Context
            $layerTrace.Add([PSCustomObject]@{
                Step       = 1
                Name       = 'AD'
                Status     = $adResult.Status
                DurationMs = $adResult.ExecutionMs
                Detail     = if ($adResult.Status -eq 'Success') { "$($adResult.OSClass) | $($adResult.OU)" } else { $adResult.SkipReason + $adResult.ErrorDetail }
            })
            if ($adResult.Status -eq 'Success') {
                $state.Hostname        = $adResult.Hostname
                $state.HostnameSource  = 'AD'
                $state.IsResolved      = $true
                $state.OSClass         = $adResult.OSClass
                $state.OperatingSystem = $adResult.OperatingSystem
                $state.OU              = $adResult.OU
            }
            Write-Verbose "[$ip] Step 1 AD -> $($adResult.Status)"

            # ---- Step 2: DHCP (always runs -- provides MAC even if already resolved) ----
            Write-VBEnrichmentProgress -Current $current -Total $total -IPAddress $ip `
                -StepNumber 2 -LayerName 'DHCP' -ElapsedMs $sw.ElapsedMilliseconds
            Write-Verbose "[$ip] Step 2 DHCP"

            $dhcpResult = Get-VBDHCPLease -IPAddress $ip -Context $Context
            $layerTrace.Add([PSCustomObject]@{
                Step       = 2
                Name       = 'DHCP'
                Status     = $dhcpResult.Status
                DurationMs = $dhcpResult.ExecutionMs
                Detail     = if ($dhcpResult.Status -eq 'Success') { "$($dhcpResult.Hostname) MAC:$($dhcpResult.MACAddress)" } else { $dhcpResult.SkipReason + $dhcpResult.ErrorDetail }
            })
            if ($dhcpResult.Status -eq 'Success') {
                if (-not [string]::IsNullOrWhiteSpace($dhcpResult.MACAddress)) {
                    $state.MACAddress   = $dhcpResult.MACAddress
                    $state.MACNormalised = $dhcpResult.MACNormalised
                }
                if (-not $state.IsResolved -and -not [string]::IsNullOrWhiteSpace($dhcpResult.Hostname)) {
                    $state.Hostname       = $dhcpResult.Hostname
                    $state.HostnameSource = 'DHCP'
                    $state.IsResolved     = $true
                    $state.LeaseExpiry    = $dhcpResult.LeaseExpiry
                }
            }
            Write-Verbose "[$ip] Step 2 DHCP -> $($dhcpResult.Status)"

            # ---- Step 3: PTR (only if not already resolved) ----
            Write-VBEnrichmentProgress -Current $current -Total $total -IPAddress $ip `
                -StepNumber 3 -LayerName 'PTR' -ElapsedMs $sw.ElapsedMilliseconds

            if (-not $state.IsResolved) {
                Write-Verbose "[$ip] Step 3 PTR"
                $ptrResult = Get-VBPTRRecord -IPAddress $ip -Context $Context
                $layerTrace.Add([PSCustomObject]@{
                    Step       = 3
                    Name       = 'PTR'
                    Status     = $ptrResult.Status
                    DurationMs = $ptrResult.ExecutionMs
                    Detail     = if ($ptrResult.Status -eq 'Success') { "$($ptrResult.Hostname) (fwd:$($ptrResult.ForwardConfirmed))" } else { $ptrResult.SkipReason + $ptrResult.ErrorDetail }
                })
                if ($ptrResult.Status -eq 'Success' -and $ptrResult.ForwardConfirmed) {
                    $state.Hostname       = $ptrResult.Hostname
                    $state.HostnameSource = 'PTR'
                    $state.IsResolved     = $true
                }
                Write-Verbose "[$ip] Step 3 PTR -> $($ptrResult.Status)"
            }
            else {
                $layerTrace.Add([PSCustomObject]@{
                    Step = 3; Name = 'PTR'; Status = 'Skipped'; DurationMs = 0
                    Detail = 'Already resolved'
                })
                Write-Verbose "[$ip] Step 3 PTR -> Skipped (already resolved)"
            }

            # ---- Step 4: ARP (always runs -- provides MAC) ----
            Write-VBEnrichmentProgress -Current $current -Total $total -IPAddress $ip `
                -StepNumber 4 -LayerName 'ARP' -ElapsedMs $sw.ElapsedMilliseconds
            Write-Verbose "[$ip] Step 4 ARP"

            $arpResult = Get-VBARPEntry -IPAddress $ip -Context $Context
            $layerTrace.Add([PSCustomObject]@{
                Step       = 4
                Name       = 'ARP'
                Status     = $arpResult.Status
                DurationMs = $arpResult.ExecutionMs
                Detail     = if ($arpResult.Status -eq 'Success') { "MAC:$($arpResult.MACAddress) ($($arpResult.ARPType))" } else { $arpResult.SkipReason + $arpResult.ErrorDetail }
            })
            if ($arpResult.Status -eq 'Success' -and [string]::IsNullOrWhiteSpace($state.MACAddress)) {
                $state.MACAddress    = $arpResult.MACAddress
                $state.MACNormalised = $arpResult.MACNormalised
            }
            Write-Verbose "[$ip] Step 4 ARP -> $($arpResult.Status)"

            # ---- Steps 5-11: Active probes ----
            $openPortsList = @()

            if (-not $SkipActiveProbes) {

                # ---- Step 5: TCP fingerprint (always runs -- enrichment only) ----
                Write-VBEnrichmentProgress -Current $current -Total $total -IPAddress $ip `
                    -StepNumber 5 -LayerName 'TCP' -ElapsedMs $sw.ElapsedMilliseconds
                Write-Verbose "[$ip] Step 5 TCP"

                $tcpResult = Get-VBTCPFingerprint -IPAddress $ip -Context $Context
                $layerTrace.Add([PSCustomObject]@{
                    Step       = 5
                    Name       = 'TCP'
                    Status     = $tcpResult.Status
                    DurationMs = $tcpResult.ExecutionMs
                    Detail     = if ($tcpResult.Status -eq 'Success') { $tcpResult.OpenPorts } else { $tcpResult.SkipReason + $tcpResult.ErrorDetail }
                })
                if ($tcpResult.Status -eq 'Success') {
                    $state.OpenPorts = $tcpResult.OpenPorts
                    $openPortsList   = $tcpResult.OpenPortsList
                }
                Write-Verbose "[$ip] Step 5 TCP -> $($tcpResult.Status) ports:$($tcpResult.OpenPorts)"

                # ---- Step 6: HTTP banner (gated on 80/443/8080/8443 open) ----
                Write-VBEnrichmentProgress -Current $current -Total $total -IPAddress $ip `
                    -StepNumber 6 -LayerName 'HTTP' -ElapsedMs $sw.ElapsedMilliseconds

                $httpGatePorts = @(80, 443, 8080, 8443)
                $httpOpen = @($openPortsList | Where-Object { $httpGatePorts -contains $_ })

                if ($httpOpen.Count -gt 0) {
                    Write-Verbose "[$ip] Step 6 HTTP"
                    $httpResult = Get-VBHTTPBanner -IPAddress $ip -OpenPortsList $openPortsList -Context $Context
                    $layerTrace.Add([PSCustomObject]@{
                        Step       = 6
                        Name       = 'HTTP'
                        Status     = $httpResult.Status
                        DurationMs = $httpResult.ExecutionMs
                        Detail     = if ($httpResult.Status -eq 'Success') { "$($httpResult.HTTPTitle) [$($httpResult.HTTPServer)]" } else { $httpResult.SkipReason + $httpResult.ErrorDetail }
                    })
                    if ($httpResult.Status -eq 'Success') {
                        $state.HTTPTitle  = $httpResult.HTTPTitle
                        $state.HTTPServer = $httpResult.HTTPServer
                        # HTTP hostname resolution (conclusive title only) handled in Resolve-VBDeviceClass
                    }
                    Write-Verbose "[$ip] Step 6 HTTP -> $($httpResult.Status)"
                }
                else {
                    $layerTrace.Add([PSCustomObject]@{
                        Step = 6; Name = 'HTTP'; Status = 'Skipped'; DurationMs = 0
                        Detail = 'No HTTP ports open (80/443/8080/8443)'
                    })
                    Write-Verbose "[$ip] Step 6 HTTP -> Skipped (no HTTP ports open)"
                }

                # ---- Step 7: SNMP (gated on port 161 open OR SNMP probing not blocked) ----
                Write-VBEnrichmentProgress -Current $current -Total $total -IPAddress $ip `
                    -StepNumber 7 -LayerName 'SNMP' -ElapsedMs $sw.ElapsedMilliseconds

                # SNMP is UDP so TCP scan won't find it -- always attempt if SNMP is available
                if ($Context -and $Context.SNMPAvailable) {
                    Write-Verbose "[$ip] Step 7 SNMP"
                    $snmpResult = Get-VBSNMPIdentity -IPAddress $ip -Context $Context
                    $layerTrace.Add([PSCustomObject]@{
                        Step       = 7
                        Name       = 'SNMP'
                        Status     = $snmpResult.Status
                        DurationMs = $snmpResult.ExecutionMs
                        Detail     = if ($snmpResult.Status -eq 'Success') { "$($snmpResult.SNMPDescr) loc:$($snmpResult.Location)" } else { $snmpResult.SkipReason + $snmpResult.ErrorDetail }
                    })
                    if ($snmpResult.Status -eq 'Success') {
                        $state.SNMPDescr = $snmpResult.SNMPDescr
                        $state.Location  = $snmpResult.Location
                        if (-not $state.IsResolved -and -not [string]::IsNullOrWhiteSpace($snmpResult.Hostname)) {
                            $state.Hostname       = $snmpResult.Hostname
                            $state.HostnameSource = 'SNMP'
                            $state.IsResolved     = $true
                        }
                    }
                    Write-Verbose "[$ip] Step 7 SNMP -> $($snmpResult.Status)"
                }
                else {
                    $layerTrace.Add([PSCustomObject]@{
                        Step = 7; Name = 'SNMP'; Status = 'Skipped'; DurationMs = 0
                        Detail = 'SNMP unavailable (olePrn COM not present)'
                    })
                    Write-Verbose "[$ip] Step 7 SNMP -> Skipped (unavailable)"
                }

                # ---- Steps 8-10: RTSP, mDNS, Switch -- Round 4 ----
                foreach ($stub in @(
                    @{ Step=8;  Name='RTSP';   Detail='Round 4' }
                    @{ Step=9;  Name='mDNS';   Detail='Round 4' }
                    @{ Step=10; Name='Switch';  Detail='Round 4' }
                )) {
                    $layerTrace.Add([PSCustomObject]@{
                        Step = $stub.Step; Name = $stub.Name
                        Status = 'Skipped'; DurationMs = 0
                        Detail = $stub.Detail
                    })
                }

            }
            else {
                # -SkipActiveProbes -- stub out 5-10
                foreach ($stub in @(
                    @{ Step=5;  Name='TCP'    }
                    @{ Step=6;  Name='HTTP'   }
                    @{ Step=7;  Name='SNMP'   }
                    @{ Step=8;  Name='RTSP'   }
                    @{ Step=9;  Name='mDNS'   }
                    @{ Step=10; Name='Switch' }
                )) {
                    $layerTrace.Add([PSCustomObject]@{
                        Step = $stub.Step; Name = $stub.Name
                        Status = 'Skipped'; DurationMs = 0
                        Detail = '-SkipActiveProbes set'
                    })
                }
            }

            # ---- Step 11: OUI vendor (always runs if MAC known) ----
            Write-VBEnrichmentProgress -Current $current -Total $total -IPAddress $ip `
                -StepNumber 11 -LayerName 'OUI' -ElapsedMs $sw.ElapsedMilliseconds
            Write-Verbose "[$ip] Step 11 OUI"

            $ouiResult = Get-VBOUIVendor -MACAddress $state.MACAddress -IPAddress $ip -Context $Context
            $layerTrace.Add([PSCustomObject]@{
                Step       = 11
                Name       = 'OUI'
                Status     = $ouiResult.Status
                DurationMs = $ouiResult.ExecutionMs
                Detail     = if ($ouiResult.Status -eq 'Success') { $ouiResult.Vendor } else { $ouiResult.SkipReason + $ouiResult.ErrorDetail }
            })
            if ($ouiResult.Status -eq 'Success') {
                $state.OUIVendor        = $ouiResult.Vendor
                $state.VendorDeviceClass = $ouiResult.VendorDeviceClass
            }
            Write-Verbose "[$ip] Step 11 OUI -> $($ouiResult.Status) vendor:$($ouiResult.Vendor)"

            # ---- Classification ----
            Write-VBEnrichmentProgress -Current $current -Total $total -IPAddress $ip `
                -StepNumber 12 -LayerName 'Classify' -ElapsedMs $sw.ElapsedMilliseconds

            $classResult = Resolve-VBDeviceClass `
                -OSClass           $state.OSClass `
                -OpenPorts         $state.OpenPorts `
                -HTTPTitle         $state.HTTPTitle `
                -HTTPServer        $state.HTTPServer `
                -SNMPDescr         $state.SNMPDescr `
                -RTSPBanner        $state.RTSPBanner `
                -OUIVendor         $state.OUIVendor `
                -VendorDeviceClass $state.VendorDeviceClass `
                -MDNSServiceType   $state.MDNSServiceType

            Write-Verbose "[$ip] Classify -> $($classResult.DeviceClass) ($($classResult.Confidence)) via $($classResult.DeviceClassSource)"

            # ---- Tally steps ----
            $attempted  = @($layerTrace | Where-Object { $_.Status -ne 'Skipped' }).Count
            $succeeded  = @($layerTrace | Where-Object { $_.Status -eq 'Success'  }).Count
            $noResult   = @($layerTrace | Where-Object { $_.Status -eq 'NoResult' }).Count
            $skipped    = @($layerTrace | Where-Object { $_.Status -eq 'Skipped'  }).Count
            $failed     = @($layerTrace | Where-Object { $_.Status -eq 'Failed'   }).Count

            $layerTraceJson = $layerTrace | ConvertTo-Json -Compress -Depth 3

            $ipSw.Stop()
            $now       = Get-Date
            $nowIso    = $now.ToString('o')

            # ---- Compare to existing row for change detection ----
            $existing     = $cachedRows[$ip]
            $changeReason = 'NewIP'
            if ($existing) {
                if ($existing.Hostname -ne $state.Hostname -and -not [string]::IsNullOrWhiteSpace($state.Hostname)) {
                    $changeReason = 'HostnameChanged'
                }
                elseif ($existing.MACAddressNormalised -ne $state.MACNormalised -and -not [string]::IsNullOrWhiteSpace($state.MACNormalised)) {
                    $changeReason = 'MACChanged'
                }
                elseif ($existing.DeviceClass -ne $classResult.DeviceClass) {
                    $changeReason = 'ClassChanged'
                }
                else {
                    $changeReason = 'StaleRefresh'
                }
            }

            $firstSeen = if ($existing -and $existing.FirstSeenAt) { $existing.FirstSeenAt } else { $nowIso }

            # ---- Upsert into SQLite ----
            if ($dbEnabled) {
                try {
                    $upsertSql = @'
INSERT INTO Enrichment (
    IPAddress, Hostname, HostnameSource, MACAddress, MACAddressNormalised,
    DeviceClass, DeviceClassSource, Confidence, OSClass, OperatingSystem,
    OU, OpenPorts, LeaseExpiry, StepsAttempted, StepsSucceeded,
    StepsNoResult, StepsSkipped, StepsFailed, LayerTraceJson,
    IsResolved, IsUnresolved, EnrichedAt, EnrichmentDurationMs,
    FirstSeenAt, UpdatedAt
) VALUES (
    @ip, @hostname, @hostnameSource, @mac, @macNorm,
    @deviceClass, @deviceClassSource, @confidence, @osClass, @os,
    @ou, @openPorts, @leaseExpiry, @attempted, @succeeded,
    @noResult, @skipped, @failed, @traceJson,
    @isResolved, @isUnresolved, @enrichedAt, @durationMs,
    @firstSeen, @updatedAt
)
ON CONFLICT(IPAddress) DO UPDATE SET
    Hostname              = excluded.Hostname,
    HostnameSource        = excluded.HostnameSource,
    MACAddress            = excluded.MACAddress,
    MACAddressNormalised  = excluded.MACAddressNormalised,
    DeviceClass           = excluded.DeviceClass,
    DeviceClassSource     = excluded.DeviceClassSource,
    Confidence            = excluded.Confidence,
    OSClass               = excluded.OSClass,
    OperatingSystem       = excluded.OperatingSystem,
    OU                    = excluded.OU,
    OpenPorts             = excluded.OpenPorts,
    LeaseExpiry           = excluded.LeaseExpiry,
    StepsAttempted        = excluded.StepsAttempted,
    StepsSucceeded        = excluded.StepsSucceeded,
    StepsNoResult         = excluded.StepsNoResult,
    StepsSkipped          = excluded.StepsSkipped,
    StepsFailed           = excluded.StepsFailed,
    LayerTraceJson        = excluded.LayerTraceJson,
    IsResolved            = excluded.IsResolved,
    IsUnresolved          = excluded.IsUnresolved,
    EnrichedAt            = excluded.EnrichedAt,
    EnrichmentDurationMs  = excluded.EnrichmentDurationMs,
    UpdatedAt             = excluded.UpdatedAt
'@
                    $isResolved   = if ($state.IsResolved) { 1 } else { 0 }
                    $isUnresolved = if ($classResult.DeviceClass -eq 'Unknown') { 1 } else { 0 }

                    Invoke-VBSqliteCommand -DatabasePath $dbPath -NonQuery -Query $upsertSql `
                        -SqlParameters @{
                            ip               = $ip
                            hostname         = $state.Hostname
                            hostnameSource   = $state.HostnameSource
                            mac              = $state.MACAddress
                            macNorm          = $state.MACNormalised
                            deviceClass      = $classResult.DeviceClass
                            deviceClassSource = $classResult.DeviceClassSource
                            confidence       = $classResult.Confidence
                            osClass          = $state.OSClass
                            os               = $state.OperatingSystem
                            ou               = $state.OU
                            openPorts        = $state.OpenPorts
                            leaseExpiry      = if ($state.LeaseExpiry) { $state.LeaseExpiry.ToString('o') } else { $null }
                            attempted        = $attempted
                            succeeded        = $succeeded
                            noResult         = $noResult
                            skipped          = $skipped
                            failed           = $failed
                            traceJson        = $layerTraceJson
                            isResolved       = $isResolved
                            isUnresolved     = $isUnresolved
                            enrichedAt       = $nowIso
                            durationMs       = [int]$ipSw.ElapsedMilliseconds
                            firstSeen        = $firstSeen
                            updatedAt        = $nowIso
                        } | Out-Null

                    # Write history row for meaningful changes
                    if ($changeReason -ne 'StaleRefresh' -and $changeReason -ne 'NewIP') {
                        $historySql = @'
INSERT INTO EnrichmentHistory (
    IPAddress, OldHostname, NewHostname, OldMACAddress, NewMACAddress,
    OldDeviceClass, NewDeviceClass, ChangeReason, ChangedAt
) VALUES (
    @ip, @oldHostname, @newHostname, @oldMAC, @newMAC,
    @oldClass, @newClass, @reason, @changedAt
)
'@
                        Invoke-VBSqliteCommand -DatabasePath $dbPath -NonQuery -Query $historySql `
                            -SqlParameters @{
                                ip          = $ip
                                oldHostname = $existing.Hostname
                                newHostname = $state.Hostname
                                oldMAC      = $existing.MACAddress
                                newMAC      = $state.MACAddress
                                oldClass    = $existing.DeviceClass
                                newClass    = $classResult.DeviceClass
                                reason      = $changeReason
                                changedAt   = $nowIso
                            } | Out-Null
                    }

                    Write-Verbose "[$ip] DB upsert -> $changeReason"
                }
                catch {
                    Write-Warning "[$ip] SQLite upsert failed: $($_.Exception.Message)"
                }
            }

            # ---- Build final enrichment object ----
            $enriched = [PSCustomObject]@{
                IPAddress            = $ip
                Hostname             = $state.Hostname
                HostnameSource       = $state.HostnameSource
                MACAddress           = $state.MACAddress
                MACAddressNormalised = $state.MACNormalised
                Vendor               = $state.OUIVendor
                DeviceClass          = $classResult.DeviceClass
                DeviceClassSource    = $classResult.DeviceClassSource
                Confidence           = $classResult.Confidence
                OSClass              = $state.OSClass
                OperatingSystem      = $state.OperatingSystem
                Model                = $null
                Location             = $state.Location
                OU                   = $state.OU
                OpenPorts            = $state.OpenPorts
                HTTPTitle            = $state.HTTPTitle
                HTTPServer           = $state.HTTPServer
                SNMPDescr            = $state.SNMPDescr
                RTSPBanner           = $state.RTSPBanner
                MDNSServiceType      = $state.MDNSServiceType
                StepsAttempted       = $attempted
                StepsSucceeded       = $succeeded
                StepsNoResult        = $noResult
                StepsSkipped         = $skipped
                StepsFailed          = $failed
                LayerTrace           = $layerTrace.ToArray()
                IsResolved           = $state.IsResolved
                IsUnresolved         = ($classResult.DeviceClass -eq 'Unknown')
                EnrichedAt           = $now
                EnrichmentDurationMs = [int]$ipSw.ElapsedMilliseconds
                FirstSeenAt          = [datetime]$firstSeen
                UpdatedAt            = $now
                ChangeReason         = $changeReason
                FromCache            = $false
            }

            if ($PassThru) { $enriched } else { $results.Add($enriched) }
        }

        Write-VBEnrichmentProgress -Completed

        # --- Summary ---
        $fromCacheCount = $fromCache.Count
        $resolvedCount  = @($results | Where-Object { $_.IsResolved }).Count + @($fromCache | Where-Object { $_.IsResolved }).Count
        $unknownCount   = @($results | Where-Object { $_.IsUnresolved }).Count + @($fromCache | Where-Object { $_.IsUnresolved }).Count
        Write-Verbose "[Orchestrator] Processed $($validIPs.Count). From cache: $fromCacheCount. Newly resolved: $resolvedCount. Still unknown: $unknownCount. Duration: $($sw.ElapsedMilliseconds) ms"

        if (-not $PassThru) {
            foreach ($r in $results) { $r }
        }
    }
}

function Invoke-VBBuildEnrichmentObject {
<#
.SYNOPSIS
    Reconstruct an enrichment PSCustomObject from a SQLite row.
    Private helper used by Invoke-VBIPEnrichment.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Row,

        [Parameter()]
        [bool]$FromCache = $false
    )

    $layerTrace = @()
    if ($Row.LayerTraceJson) {
        try {
            $layerTrace = $Row.LayerTraceJson | ConvertFrom-Json
        }
        catch { }
    }

    [PSCustomObject]@{
        IPAddress            = $Row.IPAddress
        Hostname             = $Row.Hostname
        HostnameSource       = $Row.HostnameSource
        MACAddress           = $Row.MACAddress
        MACAddressNormalised = $Row.MACAddressNormalised
        Vendor               = $Row.Vendor
        DeviceClass          = $Row.DeviceClass
        DeviceClassSource    = $Row.DeviceClassSource
        Confidence           = $Row.Confidence
        OSClass              = $Row.OSClass
        OperatingSystem      = $Row.OperatingSystem
        Model                = $Row.Model
        Location             = $Row.Location
        OU                   = $Row.OU
        OpenPorts            = $Row.OpenPorts
        HTTPTitle            = $Row.HTTPTitle
        HTTPServer           = $Row.HTTPServer
        SNMPDescr            = $Row.SNMPDescr
        RTSPBanner           = $Row.RTSPBanner
        MDNSServiceType      = $Row.MDNSServiceType
        StepsAttempted       = [int]$Row.StepsAttempted
        StepsSucceeded       = [int]$Row.StepsSucceeded
        StepsNoResult        = [int]$Row.StepsNoResult
        StepsSkipped         = [int]$Row.StepsSkipped
        StepsFailed          = [int]$Row.StepsFailed
        LayerTrace           = $layerTrace
        IsResolved           = [bool]$Row.IsResolved
        IsUnresolved         = [bool]$Row.IsUnresolved
        EnrichedAt           = if ($Row.EnrichedAt) { [datetime]$Row.EnrichedAt } else { $null }
        EnrichmentDurationMs = [int]$Row.EnrichmentDurationMs
        FirstSeenAt          = if ($Row.FirstSeenAt) { [datetime]$Row.FirstSeenAt } else { $null }
        UpdatedAt            = if ($Row.UpdatedAt) { [datetime]$Row.UpdatedAt } else { $null }
        ChangeReason         = 'NoChange'
        FromCache            = $FromCache
    }
}
