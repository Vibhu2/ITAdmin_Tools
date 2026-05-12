function Invoke-VBIPEnrichment {
<#
.SYNOPSIS
    Orchestrate the full enrichment pipeline against a list of private IP addresses.

.DESCRIPTION
    Runs all enabled layers in priority order for each IP, merges results, classifies
    each device, persists to SQLite, and returns one enrichment object per IP.

    Full 11-layer pipeline: passive (AD, DHCP, PTR, ARP) then active (TCP, HTTP,
    SNMP, RTSP, mDNS, Switch ARP, OUI). Layers that are unavailable or whose
    prerequisites are not met return Skipped automatically.

    On PS 7 with $Context.CanUseParallel = $true, active probes (steps 5-10) run
    in parallel across IPs using ForEach-Object -Parallel. Passive layers always
    run sequentially. Classification and SQLite writes always run sequentially.

    Execution flow:
        1.  Validate context; warn if missing.
        2.  Validate IP list -- skip public addresses with Write-Warning.
        3.  Load existing rows from SQLite for supplied IPs.
        4.  Decide which IPs to probe (missing, stale, unresolved, or -ForceRefresh).
        5a. Passive layers (sequential for all IPs):
              Step 1  Get-VBADComputer    -> may set Hostname, OSClass
              Step 2  Get-VBDHCPLease    -> may set Hostname, MAC
              Step 3  Get-VBPTRRecord    -> may set Hostname if not resolved
              Step 4  Get-VBARPEntry     -> may set MAC if not known
        5b. Active layers per IP (parallel on PS7, sequential on PS5.1):
              Step 5  Get-VBTCPFingerprint -> OpenPorts (always)
              Step 6  Get-VBHTTPBanner     -> gated on 80/443/8080/8443 open
              Step 7  Get-VBSNMPIdentity   -> gated on SNMPAvailable
              Step 8  Get-VBRTSPBanner     -> gated on port 554 open
              Step 9  Get-VBmDNSRecord     -> gated on mDNSAvailable
              Step 10 Get-VBSwitchARP      -> gated on SwitchTargets configured
              Step 11 Get-VBOUIVendor      -> always runs if MAC known
        6.  Resolve-VBDeviceClass on merged signals.
        7.  Compare result to existing SQLite row; write EnrichmentHistory on change.
        8.  Upsert row into SQLite.
        9.  Emit object (stream immediately if -PassThru; collect otherwise).
        10. Emit summary via Write-Verbose.

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
        2.0.0 -- 2026-05-11 -- Round 4: layers 8-10 wired; PS7 parallel mode for active probes
#>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('IP_Address', 'IP Address', 'IP')]
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
        if (-not $Context) {
            Write-Verbose "[Orchestrator] No context provided -- auto-building context."
            $Context = Get-VBEnrichmentContext -Quiet
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
                $paramHash = @{}
                $placeholders = for ($i = 0; $i -lt $validIPs.Count; $i++) {
                    $paramHash["ip$i"] = $validIPs[$i]
                    "@ip$i"
                }
                $inClause = $placeholders -join ','
                $rows = Invoke-VBSqliteCommand -DatabasePath $dbPath `
                    -Query "SELECT * FROM Enrichment WHERE IPAddress IN ($inClause)" `
                    -SqlParameters $paramHash
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
                $fromCache.Add((ConvertFrom-VBSqliteEnrichmentRow -Row $existing -FromCache $true))
            }
        }

        Write-Verbose "[Orchestrator] $($validIPs.Count) IPs total: $($toProbe.Count) to probe, $($fromCache.Count) from cache"

        # Emit cached results
        foreach ($cached in $fromCache) {
            if ($PassThru) { $cached } else { $results.Add($cached) }
        }

        # --- Step 4: Passive probes (sequential -- build one-shot caches) ---
        $total   = $toProbe.Count
        $current = 0

        # Collect state objects for all IPs (passive pass)
        $stateMap = [ordered]@{}

        foreach ($ip in $toProbe) {
            $current++
            $layerTrace = New-Object System.Collections.Generic.List[PSCustomObject]

            # State accumulator for this IP
            $state = @{
                IPAddress         = $ip
                Hostname          = $null
                HostnameSource    = $null
                IsResolved        = $false
                MACAddress        = $null
                MACNormalised     = $null
                OSClass           = $null
                OperatingSystem   = $null
                OU                = $null
                OpenPorts         = $null
                HTTPTitle         = $null
                HTTPServer        = $null
                SNMPDescr         = $null
                RTSPBanner        = $null
                MDNSServiceType   = $null
                Location          = $null
                LeaseExpiry       = $null
                VendorDeviceClass = $null
                OUIVendor         = $null
                PassiveTrace      = $layerTrace  # carry through to active phase
            }

            # ---- Step 1: AD ----
            Write-VBEnrichmentProgress -Current $current -Total $total -IPAddress $ip `
                -StepNumber 1 -LayerName 'AD' -ElapsedMs $sw.ElapsedMilliseconds
            Write-Verbose "[$ip] Step 1 AD"

            $adResult = Get-VBADComputer -IPAddress $ip -Context $Context
            $adDetail = if ($adResult.Status -eq 'Success') { "$($adResult.OSClass) | $($adResult.OU)" } else { $adResult.SkipReason + $adResult.ErrorDetail }
            $layerTrace.Add([PSCustomObject]@{
                Step       = 1
                Name       = 'AD'
                Status     = $adResult.Status
                DurationMs = $adResult.ExecutionMs
                Detail     = $adDetail
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
            $dhcpDetail = if ($dhcpResult.Status -eq 'Success') { "$($dhcpResult.Hostname) MAC:$($dhcpResult.MACAddress)" } else { $dhcpResult.SkipReason + $dhcpResult.ErrorDetail }
            $layerTrace.Add([PSCustomObject]@{
                Step       = 2
                Name       = 'DHCP'
                Status     = $dhcpResult.Status
                DurationMs = $dhcpResult.ExecutionMs
                Detail     = $dhcpDetail
            })
            if ($dhcpResult.Status -eq 'Success') {
                if (-not [string]::IsNullOrWhiteSpace($dhcpResult.MACAddress)) {
                    $state.MACAddress    = $dhcpResult.MACAddress
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
                $ptrDetail = if ($ptrResult.Status -eq 'Success') { "$($ptrResult.Hostname) (fwd:$($ptrResult.ForwardConfirmed))" } else { $ptrResult.SkipReason + $ptrResult.ErrorDetail }
                $layerTrace.Add([PSCustomObject]@{
                    Step       = 3
                    Name       = 'PTR'
                    Status     = $ptrResult.Status
                    DurationMs = $ptrResult.ExecutionMs
                    Detail     = $ptrDetail
                })
                if ($ptrResult.Status -eq 'Success' -and -not [string]::IsNullOrWhiteSpace($ptrResult.Hostname)) {
                    $state.Hostname       = $ptrResult.Hostname
                    $state.HostnameSource = if ($ptrResult.ForwardConfirmed) { 'PTR' } else { 'PTR-Unconfirmed' }
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
            $arpDetail = if ($arpResult.Status -eq 'Success') { "MAC:$($arpResult.MACAddress) ($($arpResult.ARPType))" } else { $arpResult.SkipReason + $arpResult.ErrorDetail }
            $layerTrace.Add([PSCustomObject]@{
                Step       = 4
                Name       = 'ARP'
                Status     = $arpResult.Status
                DurationMs = $arpResult.ExecutionMs
                Detail     = $arpDetail
            })
            if ($arpResult.Status -eq 'Success' -and [string]::IsNullOrWhiteSpace($state.MACAddress)) {
                $state.MACAddress    = $arpResult.MACAddress
                $state.MACNormalised = $arpResult.MACNormalised
            }
            Write-Verbose "[$ip] Step 4 ARP -> $($arpResult.Status)"

            $stateMap[$ip] = $state
        }

        # --- Steps 5-11: Active probes (parallel on PS7 when context allows, sequential otherwise) ---

        $useParallel = ($PSVersionTable.PSVersion.Major -ge 7) -and
                       ($Context -and $Context.CanUseParallel) -and
                       (-not $SkipActiveProbes) -and
                       ($stateMap.Count -gt 1)

        if ($useParallel) {
            Write-Verbose "[Orchestrator] PS7 parallel active probes -- $($stateMap.Count) IPs, throttle $($Context.ParallelThrottleLimit)"

            # Build a serialisable list of per-IP inputs for the parallel block
            $parallelInputs = foreach ($ip in $stateMap.Keys) {
                $s = $stateMap[$ip]
                [PSCustomObject]@{
                    IPAddress         = $ip
                    IsResolved        = $s.IsResolved
                    MACAddress        = $s.MACAddress
                    MACNormalised     = $s.MACNormalised
                    SNMPAvailable     = ($Context -and $Context.SNMPAvailable)
                    mDNSAvailable     = ($Context -and $Context.mDNSAvailable)
                    RTSPProbeEnabled  = ($Context -and $Context.RTSPProbeEnabled)
                    SwitchTargetCount = if ($Context -and $Context.SwitchTargets) { [int]$Context.SwitchTargets.Count } else { 0 }
                }
            }

            $throttle   = $Context.ParallelThrottleLimit
            $modulePath = (Get-Module -Name 'VB.DNSEnrichment' | Select-Object -ExpandProperty Path -First 1)
            if (-not $modulePath) {
                throw "[Orchestrator] Cannot resolve VB.DNSEnrichment module path for parallel block."
            }

            # Pre-warm OUI table in main scope so $using: can pass it into runspaces
            $null = Get-VBOUIVendor -MACAddress '000000000000' -Context $Context
            $ouiTableSnapshot = $Script:VBOUITable

            $parallelResults = $parallelInputs |
                ForEach-Object -ThrottleLimit $throttle -Parallel {
                    $inp     = $_
                    $ctx     = $using:Context
                    $modPath = $using:modulePath
                    if (-not (Get-Module -Name 'VB.DNSEnrichment')) {
                        Import-Module $modPath -ErrorAction Stop
                    }
                    $Script:VBOUITable = $using:ouiTableSnapshot

                    $ip            = $inp.IPAddress
                    $activeTrace   = New-Object System.Collections.Generic.List[PSCustomObject]
                    $openPortsList = @()
                    $fields        = @{
                        OpenPorts         = $null
                        HTTPTitle         = $null
                        HTTPServer        = $null
                        SNMPDescr         = $null
                        RTSPBanner        = $null
                        MDNSServiceType   = $null
                        Location          = $null
                        OUIVendor         = $null
                        VendorDeviceClass = $null
                        MACAddress        = $inp.MACAddress
                        MACNormalised     = $inp.MACNormalised
                        IsResolved        = $inp.IsResolved
                        Hostname          = $null
                        HostnameSource    = $null
                    }

                    # Step 5 TCP
                    $tcpResult = Get-VBTCPFingerprint -IPAddress $ip -Context $ctx
                    $activeTrace.Add([PSCustomObject]@{ Step=5; Name='TCP'; Status=$tcpResult.Status; DurationMs=$tcpResult.ExecutionMs; Detail=if($tcpResult.Status -eq 'Success'){$tcpResult.OpenPorts}else{$tcpResult.SkipReason+$tcpResult.ErrorDetail} })
                    if ($tcpResult.Status -eq 'Success') { $fields.OpenPorts=$tcpResult.OpenPorts; $openPortsList=$tcpResult.OpenPortsList }

                    # Step 6 HTTP
                    $httpGate = @(80,443,8080,8443)
                    if (($openPortsList | Where-Object { $httpGate -contains $_ }).Count -gt 0) {
                        $httpResult = Get-VBHTTPBanner -IPAddress $ip -OpenPortsList $openPortsList -Context $ctx
                        $activeTrace.Add([PSCustomObject]@{ Step=6; Name='HTTP'; Status=$httpResult.Status; DurationMs=$httpResult.ExecutionMs; Detail=if($httpResult.Status -eq 'Success'){"$($httpResult.HTTPTitle) [$($httpResult.HTTPServer)]"}else{$httpResult.SkipReason+$httpResult.ErrorDetail} })
                        if ($httpResult.Status -eq 'Success') { $fields.HTTPTitle=$httpResult.HTTPTitle; $fields.HTTPServer=$httpResult.HTTPServer }
                    } else {
                        $activeTrace.Add([PSCustomObject]@{ Step=6; Name='HTTP'; Status='Skipped'; DurationMs=0; Detail='No HTTP ports open (80/443/8080/8443)' })
                    }

                    # Step 7 SNMP
                    if ($inp.SNMPAvailable) {
                        $snmpResult = Get-VBSNMPIdentity -IPAddress $ip -Context $ctx
                        $activeTrace.Add([PSCustomObject]@{ Step=7; Name='SNMP'; Status=$snmpResult.Status; DurationMs=$snmpResult.ExecutionMs; Detail=if($snmpResult.Status -eq 'Success'){"$($snmpResult.SNMPDescr) loc:$($snmpResult.Location)"}else{$snmpResult.SkipReason+$snmpResult.ErrorDetail} })
                        if ($snmpResult.Status -eq 'Success') {
                            $fields.SNMPDescr=$snmpResult.SNMPDescr; $fields.Location=$snmpResult.Location
                            if (-not $fields.IsResolved -and -not [string]::IsNullOrWhiteSpace($snmpResult.Hostname)) { $fields.Hostname=$snmpResult.Hostname; $fields.HostnameSource='SNMP'; $fields.IsResolved=$true }
                        }
                    } else {
                        $activeTrace.Add([PSCustomObject]@{ Step=7; Name='SNMP'; Status='Skipped'; DurationMs=0; Detail='SNMP unavailable (olePrn COM not present)' })
                    }

                    # Step 8 RTSP
                    if ($openPortsList -contains 554 -and $inp.RTSPProbeEnabled) {
                        $rtspResult = Get-VBRTSPBanner -IPAddress $ip -Context $ctx
                        $activeTrace.Add([PSCustomObject]@{ Step=8; Name='RTSP'; Status=$rtspResult.Status; DurationMs=$rtspResult.ExecutionMs; Detail=if($rtspResult.Status -eq 'Success'){$rtspResult.RTSPBanner.Substring(0,[math]::Min(80,$rtspResult.RTSPBanner.Length))}else{$rtspResult.SkipReason+$rtspResult.ErrorDetail} })
                        if ($rtspResult.Status -eq 'Success') { $fields.RTSPBanner=$rtspResult.RTSPBanner }
                    } else {
                        $activeTrace.Add([PSCustomObject]@{ Step=8; Name='RTSP'; Status='Skipped'; DurationMs=0; Detail=if($openPortsList -notcontains 554){'Port 554 closed'}else{'RTSPProbeDisabled'} })
                    }

                    # Step 9 mDNS
                    if ($inp.mDNSAvailable) {
                        $mdnsResult = Get-VBmDNSRecord -IPAddress $ip -Context $ctx
                        $activeTrace.Add([PSCustomObject]@{ Step=9; Name='mDNS'; Status=$mdnsResult.Status; DurationMs=$mdnsResult.ExecutionMs; Detail=if($mdnsResult.Status -eq 'Success'){"$($mdnsResult.MDNSServiceType) $($mdnsResult.MDNSServiceName)"}else{$mdnsResult.SkipReason+$mdnsResult.ErrorDetail} })
                        if ($mdnsResult.Status -eq 'Success') {
                            $fields.MDNSServiceType=$mdnsResult.MDNSServiceType
                            if (-not $fields.IsResolved -and -not [string]::IsNullOrWhiteSpace($mdnsResult.MDNSServiceName)) { $fields.Hostname=$mdnsResult.MDNSServiceName; $fields.HostnameSource='mDNS'; $fields.IsResolved=$true }
                        }
                    } else {
                        $activeTrace.Add([PSCustomObject]@{ Step=9; Name='mDNS'; Status='Skipped'; DurationMs=0; Detail='dns-sd.exe not on PATH' })
                    }

                    # Step 10 Switch
                    if ($inp.SwitchTargetCount -gt 0 -and $inp.SNMPAvailable) {
                        $switchResult = Get-VBSwitchARP -IPAddress $ip -Context $ctx
                        $activeTrace.Add([PSCustomObject]@{ Step=10; Name='Switch'; Status=$switchResult.Status; DurationMs=$switchResult.ExecutionMs; Detail=if($switchResult.Status -eq 'Success'){"SW:$($switchResult.SwitchIP) Port:$($switchResult.SwitchPort) $($switchResult.PortDescription)"}else{$switchResult.SkipReason+$switchResult.ErrorDetail} })
                        if ($switchResult.Status -eq 'Success') {
                            if ([string]::IsNullOrWhiteSpace($fields.MACAddress)) { $fields.MACAddress=$switchResult.MACAddress; $fields.MACNormalised=($switchResult.MACAddress -replace '[:\-\.]','').ToUpperInvariant() }
                            if ([string]::IsNullOrWhiteSpace($fields.Location)) { $fields.Location=$switchResult.PortDescription }
                        }
                    } else {
                        $activeTrace.Add([PSCustomObject]@{ Step=10; Name='Switch'; Status='Skipped'; DurationMs=0; Detail=if($inp.SwitchTargetCount -eq 0){'No SwitchTargets configured'}else{'SNMPUnavailable'} })
                    }

                    # Step 11 OUI
                    $ouiResult = Get-VBOUIVendor -MACAddress $fields.MACAddress -IPAddress $ip -Context $ctx
                    $activeTrace.Add([PSCustomObject]@{ Step=11; Name='OUI'; Status=$ouiResult.Status; DurationMs=$ouiResult.ExecutionMs; Detail=if($ouiResult.Status -eq 'Success'){$ouiResult.Vendor}else{$ouiResult.SkipReason+$ouiResult.ErrorDetail} })
                    if ($ouiResult.Status -eq 'Success') { $fields.OUIVendor=$ouiResult.Vendor; $fields.VendorDeviceClass=$ouiResult.VendorDeviceClass }

                    [PSCustomObject]@{ IPAddress=$ip; Fields=$fields; ActiveTrace=$activeTrace }
                }

            # Merge parallel results back into stateMap
            foreach ($pr in $parallelResults) {
                $s = $stateMap[$pr.IPAddress]
                $f = $pr.Fields
                $s.OpenPorts        = $f.OpenPorts
                $s.HTTPTitle        = $f.HTTPTitle
                $s.HTTPServer       = $f.HTTPServer
                $s.SNMPDescr        = $f.SNMPDescr
                $s.RTSPBanner       = $f.RTSPBanner
                $s.MDNSServiceType  = $f.MDNSServiceType
                $s.Location         = $f.Location
                $s.OUIVendor        = $f.OUIVendor
                $s.VendorDeviceClass = $f.VendorDeviceClass
                if (-not [string]::IsNullOrWhiteSpace($f.MACAddress) -and [string]::IsNullOrWhiteSpace($s.MACAddress)) {
                    $s.MACAddress    = $f.MACAddress
                    $s.MACNormalised = $f.MACNormalised
                }
                if ($f.IsResolved -and -not $s.IsResolved) {
                    $s.IsResolved     = $true
                    $s.Hostname       = $f.Hostname
                    $s.HostnameSource = $f.HostnameSource
                }
                foreach ($entry in $pr.ActiveTrace) { $s.PassiveTrace.Add($entry) }
            }
        }
        else {
            # Sequential active probes (PS5.1 or single IP or SkipActiveProbes)
            foreach ($ip in $stateMap.Keys) {
                $state     = $stateMap[$ip]
                $layerTrace = $state.PassiveTrace
                $openPortsList = @()

                if (-not $SkipActiveProbes) {
                    # ---- Step 5: TCP ----
                    $tcpResult = Get-VBTCPFingerprint -IPAddress $ip -Context $Context
                    $tcpDetail = if ($tcpResult.Status -eq 'Success') { $tcpResult.OpenPorts } else { $tcpResult.SkipReason + $tcpResult.ErrorDetail }
                    $layerTrace.Add([PSCustomObject]@{
                        Step       = 5; Name = 'TCP'; Status = $tcpResult.Status; DurationMs = $tcpResult.ExecutionMs
                        Detail     = $tcpDetail
                    })
                    if ($tcpResult.Status -eq 'Success') { $state.OpenPorts=$tcpResult.OpenPorts; $openPortsList=$tcpResult.OpenPortsList }
                    Write-Verbose "[$ip] Step 5 TCP -> $($tcpResult.Status) ports:$($tcpResult.OpenPorts)"

                    # ---- Step 6: HTTP ----
                    $httpGatePorts = @(80, 443, 8080, 8443)
                    $httpOpen = @($openPortsList | Where-Object { $httpGatePorts -contains $_ })
                    if ($httpOpen.Count -gt 0) {
                        $httpResult = Get-VBHTTPBanner -IPAddress $ip -OpenPortsList $openPortsList -Context $Context
                        $layerTrace.Add([PSCustomObject]@{
                            Step=6; Name='HTTP'; Status=$httpResult.Status; DurationMs=$httpResult.ExecutionMs
                            Detail=if($httpResult.Status -eq 'Success'){"$($httpResult.HTTPTitle) [$($httpResult.HTTPServer)]"}else{$httpResult.SkipReason+$httpResult.ErrorDetail}
                        })
                        if ($httpResult.Status -eq 'Success') { $state.HTTPTitle=$httpResult.HTTPTitle; $state.HTTPServer=$httpResult.HTTPServer }
                        Write-Verbose "[$ip] Step 6 HTTP -> $($httpResult.Status)"
                    }
                    else {
                        $layerTrace.Add([PSCustomObject]@{ Step=6; Name='HTTP'; Status='Skipped'; DurationMs=0; Detail='No HTTP ports open (80/443/8080/8443)' })
                        Write-Verbose "[$ip] Step 6 HTTP -> Skipped"
                    }

                    # ---- Step 7: SNMP ----
                    if ($Context -and $Context.SNMPAvailable) {
                        $snmpResult = Get-VBSNMPIdentity -IPAddress $ip -Context $Context
                        $layerTrace.Add([PSCustomObject]@{
                            Step=7; Name='SNMP'; Status=$snmpResult.Status; DurationMs=$snmpResult.ExecutionMs
                            Detail=if($snmpResult.Status -eq 'Success'){"$($snmpResult.SNMPDescr) loc:$($snmpResult.Location)"}else{$snmpResult.SkipReason+$snmpResult.ErrorDetail}
                        })
                        if ($snmpResult.Status -eq 'Success') {
                            $state.SNMPDescr=$snmpResult.SNMPDescr; $state.Location=$snmpResult.Location
                            if (-not $state.IsResolved -and -not [string]::IsNullOrWhiteSpace($snmpResult.Hostname)) { $state.Hostname=$snmpResult.Hostname; $state.HostnameSource='SNMP'; $state.IsResolved=$true }
                        }
                        Write-Verbose "[$ip] Step 7 SNMP -> $($snmpResult.Status)"
                    }
                    else {
                        $layerTrace.Add([PSCustomObject]@{ Step=7; Name='SNMP'; Status='Skipped'; DurationMs=0; Detail='SNMP unavailable (olePrn COM not present)' })
                        Write-Verbose "[$ip] Step 7 SNMP -> Skipped"
                    }

                    # ---- Step 8: RTSP ----
                    if ($openPortsList -contains 554 -and ($Context -and $Context.RTSPProbeEnabled)) {
                        $rtspResult = Get-VBRTSPBanner -IPAddress $ip -Context $Context
                        $layerTrace.Add([PSCustomObject]@{
                            Step=8; Name='RTSP'; Status=$rtspResult.Status; DurationMs=$rtspResult.ExecutionMs
                            Detail=if($rtspResult.Status -eq 'Success'){$rtspResult.RTSPBanner.Substring(0,[math]::Min(80,$rtspResult.RTSPBanner.Length))}else{$rtspResult.SkipReason+$rtspResult.ErrorDetail}
                        })
                        if ($rtspResult.Status -eq 'Success') { $state.RTSPBanner=$rtspResult.RTSPBanner }
                        Write-Verbose "[$ip] Step 8 RTSP -> $($rtspResult.Status)"
                    }
                    else {
                        $layerTrace.Add([PSCustomObject]@{ Step=8; Name='RTSP'; Status='Skipped'; DurationMs=0; Detail=if($openPortsList -notcontains 554){'Port 554 closed'}else{'RTSPProbeDisabled'} })
                        Write-Verbose "[$ip] Step 8 RTSP -> Skipped"
                    }

                    # ---- Step 9: mDNS ----
                    if ($Context -and $Context.mDNSAvailable) {
                        $mdnsResult = Get-VBmDNSRecord -IPAddress $ip -Context $Context
                        $layerTrace.Add([PSCustomObject]@{
                            Step=9; Name='mDNS'; Status=$mdnsResult.Status; DurationMs=$mdnsResult.ExecutionMs
                            Detail=if($mdnsResult.Status -eq 'Success'){"$($mdnsResult.MDNSServiceType) $($mdnsResult.MDNSServiceName)"}else{$mdnsResult.SkipReason+$mdnsResult.ErrorDetail}
                        })
                        if ($mdnsResult.Status -eq 'Success') {
                            $state.MDNSServiceType=$mdnsResult.MDNSServiceType
                            if (-not $state.IsResolved -and -not [string]::IsNullOrWhiteSpace($mdnsResult.MDNSServiceName)) { $state.Hostname=$mdnsResult.MDNSServiceName; $state.HostnameSource='mDNS'; $state.IsResolved=$true }
                        }
                        Write-Verbose "[$ip] Step 9 mDNS -> $($mdnsResult.Status)"
                    }
                    else {
                        $layerTrace.Add([PSCustomObject]@{ Step=9; Name='mDNS'; Status='Skipped'; DurationMs=0; Detail='dns-sd.exe not on PATH' })
                        Write-Verbose "[$ip] Step 9 mDNS -> Skipped"
                    }

                    # ---- Step 10: Switch ----
                    if ($Context -and $Context.SwitchTargets.Count -gt 0 -and $Context.SNMPAvailable) {
                        $switchResult = Get-VBSwitchARP -IPAddress $ip -Context $Context
                        $layerTrace.Add([PSCustomObject]@{
                            Step=10; Name='Switch'; Status=$switchResult.Status; DurationMs=$switchResult.ExecutionMs
                            Detail=if($switchResult.Status -eq 'Success'){"SW:$($switchResult.SwitchIP) Port:$($switchResult.SwitchPort) $($switchResult.PortDescription)"}else{$switchResult.SkipReason+$switchResult.ErrorDetail}
                        })
                        if ($switchResult.Status -eq 'Success') {
                            if ([string]::IsNullOrWhiteSpace($state.MACAddress)) { $state.MACAddress=$switchResult.MACAddress; $state.MACNormalised=ConvertTo-VBNormalisedMAC -MACAddress $switchResult.MACAddress }
                            $state.Location=$switchResult.PortDescription
                        }
                        Write-Verbose "[$ip] Step 10 Switch -> $($switchResult.Status)"
                    }
                    else {
                        $layerTrace.Add([PSCustomObject]@{ Step=10; Name='Switch'; Status='Skipped'; DurationMs=0; Detail=if(-not($Context -and $Context.SwitchTargets.Count -gt 0)){'No SwitchTargets configured'}else{'SNMPUnavailable'} })
                        Write-Verbose "[$ip] Step 10 Switch -> Skipped"
                    }

                    # ---- Step 11: OUI ----
                    $ouiResult = Get-VBOUIVendor -MACAddress $state.MACAddress -IPAddress $ip -Context $Context
                    $layerTrace.Add([PSCustomObject]@{
                        Step=11; Name='OUI'; Status=$ouiResult.Status; DurationMs=$ouiResult.ExecutionMs
                        Detail=if($ouiResult.Status -eq 'Success'){$ouiResult.Vendor}else{$ouiResult.SkipReason+$ouiResult.ErrorDetail}
                    })
                    if ($ouiResult.Status -eq 'Success') { $state.OUIVendor=$ouiResult.Vendor; $state.VendorDeviceClass=$ouiResult.VendorDeviceClass }
                    Write-Verbose "[$ip] Step 11 OUI -> $($ouiResult.Status) vendor:$($ouiResult.Vendor)"
                }
                else {
                    foreach ($stub in @(
                        @{ Step=5; Name='TCP' }; @{ Step=6; Name='HTTP' }; @{ Step=7; Name='SNMP' }
                        @{ Step=8; Name='RTSP' }; @{ Step=9; Name='mDNS' }; @{ Step=10; Name='Switch' }
                        @{ Step=11; Name='OUI' }
                    )) {
                        $layerTrace.Add([PSCustomObject]@{ Step=$stub.Step; Name=$stub.Name; Status='Skipped'; DurationMs=0; Detail='-SkipActiveProbes set' })
                    }
                }
            }
        }

        # --- Classification + SQLite writes (always sequential) ---
        $current = 0
        foreach ($ip in $stateMap.Keys) {
            $current++
            $state      = $stateMap[$ip]
            $layerTrace = $state.PassiveTrace
            $ipSw       = [System.Diagnostics.Stopwatch]::StartNew()

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
    IPAddress, Hostname, HostnameSource, MACAddress, MACAddressNormalised, Vendor,
    DeviceClass, DeviceClassSource, Confidence, OSClass, OperatingSystem,
    OU, OpenPorts, LeaseExpiry, StepsAttempted, StepsSucceeded,
    StepsNoResult, StepsSkipped, StepsFailed, LayerTraceJson,
    IsResolved, IsUnresolved, EnrichedAt, EnrichmentDurationMs,
    FirstSeenAt, UpdatedAt
) VALUES (
    @ip, @hostname, @hostnameSource, @mac, @macNorm, @vendor,
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
    Vendor                = excluded.Vendor,
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
                    $isResolved      = if ($state.IsResolved) { 1 } else { 0 }
                    $isUnresolved    = if ($classResult.DeviceClass -eq 'Unknown') { 1 } else { 0 }
                    $leaseExpiryVal  = if ($state.LeaseExpiry) { $state.LeaseExpiry.ToString('o') } else { $null }

                    Invoke-VBSqliteCommand -DatabasePath $dbPath-Query $upsertSql `
                        -SqlParameters @{
                            ip               = $ip
                            hostname         = $state.Hostname
                            hostnameSource   = $state.HostnameSource
                            mac              = $state.MACAddress
                            macNorm          = $state.MACNormalised
                            vendor           = $state.OUIVendor
                            deviceClass      = $classResult.DeviceClass
                            deviceClassSource = $classResult.DeviceClassSource
                            confidence       = $classResult.Confidence
                            osClass          = $state.OSClass
                            os               = $state.OperatingSystem
                            ou               = $state.OU
                            openPorts        = $state.OpenPorts
                            leaseExpiry      = $leaseExpiryVal
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
                        Invoke-VBSqliteCommand -DatabasePath $dbPath-Query $historySql `
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
