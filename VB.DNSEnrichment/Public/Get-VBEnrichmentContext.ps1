function Get-VBEnrichmentContext {
<#
.SYNOPSIS
    Build the environment context object that drives every other VB.DNSEnrichment function.

.DESCRIPTION
    Probes the local environment to determine which enrichment layers are usable
    in the current runtime. Returns a single PSCustomObject containing PowerShell
    version flags, machine identity, per-layer availability flags, storage paths,
    performance defaults, and a PrerequisiteReport array (one row per check).

    Unless -Quiet is supplied, prints a formatted console report so the operator
    knows exactly which layers will run, which will skip, and why.

    Checks performed (in order):
        1.  PowerShell version       -> PSMajor, CanUseParallel, CanSkipCertCheck
        2.  Domain join              -> IsDomainJoined, DomainName
        3.  DC role                  -> IsDomainController
        4.  DNS resolution works     -> DNSAvailable
        5.  Local DNS role           -> DNSIsLocal
        6.  AD module + Get-ADDomain -> ADAvailable, ADIsLocal
        7.  DhcpServer module        -> DHCPAvailable
        8.  DHCP server reachable    -> DHCPAvailable confirmed
        9.  Local DHCP role          -> DHCPIsLocal
        10. SNMP olePrn COM          -> SNMPAvailable
        11. OUI file                 -> OUIFileAvailable, OUIFileAge
        12. mDNS dns-sd.exe          -> mDNSAvailable
        13. PSSQLite module          -> CanUsePSSQLite
        14. SQLite database init     -> DatabaseInitialized

    This function MUST be called before any other VB.DNSEnrichment function.

.PARAMETER DHCPServer
    FQDN or IP of the DHCP server to query in subsequent runs. Optional --
    omit to skip DHCP probing entirely.

.PARAMETER DHCPScopeIds
    Array of DHCP scope IDs (e.g. '192.168.1.0') the cache should pre-load.
    Optional -- omit to enumerate all scopes the server exposes.

.PARAMETER SwitchTargets
    Array of managed-switch IPs for the optional Switch ARP layer (10).
    Empty = layer 10 skipped.

.PARAMETER SNMPCommunityStrings
    Array of SNMP community strings to try in order. First match wins.
    Defaults to @('public').

.PARAMETER OUIFilePath
    Path to the IEEE OUI CSV. Defaults to "<module>\Data\oui.csv".

.PARAMETER DatabasePath
    Path to the SQLite database. Defaults to "$env:LOCALAPPDATA\VB.DNSEnrichment\enrichment.db".

.PARAMETER ParallelThrottleLimit
    Maximum parallel runspaces for active probes on PS 7. Defaults to 10.

.PARAMETER DisableNetworkProbe
    Switch -- disables active layers (4-10) entirely. Use when network probing is
    not permitted (e.g. running from a hardened jumpbox).

.PARAMETER Quiet
    Switch -- suppress the console prerequisite report. The PrerequisiteReport
    property on the returned object is still populated.

.OUTPUTS
    [PSCustomObject] -- the full context object. See module design v2 section 7
    for the complete schema.

.EXAMPLE
    $ctx = Get-VBEnrichmentContext

.EXAMPLE
    $ctx = Get-VBEnrichmentContext -DHCPServer 'dhcp01.corp.local' `
        -SNMPCommunityStrings 'public','readonly' -Quiet

.NOTES
    Version:      1.0.0
    MinPSVersion: 5.1
    Author:       VB
    ChangeLog:
        1.0.0 -- 2026-05-10 -- Initial release
#>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter()]
        [string]$DHCPServer,

        [Parameter()]
        [string[]]$DHCPScopeIds = @(),

        [Parameter()]
        [string[]]$SwitchTargets = @(),

        [Parameter()]
        [string[]]$SNMPCommunityStrings = @('public'),

        [Parameter()]
        [string]$OUIFilePath,

        [Parameter()]
        [string]$DatabasePath,

        [Parameter()]
        [int]$ParallelThrottleLimit = 10,

        [Parameter()]
        [switch]$DisableNetworkProbe,

        [Parameter()]
        [switch]$Quiet
    )

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $report = New-Object System.Collections.Generic.List[PSCustomObject]

    # Resolve default paths -- $PSScriptRoot is Public\, module root is one level up
    if (-not $OUIFilePath) {
        $OUIFilePath = Join-Path $PSScriptRoot '..\Data\oui.csv'
    }
    if (-not $DatabasePath) {
        $DatabasePath = Join-Path $env:LOCALAPPDATA 'VB.DNSEnrichment\enrichment.db'
    }

    # --- Helper: append a PrerequisiteReport row ---
    $addReport = {
        param($Layer, $Prerequisite, $Status, $Detail, $LayersAffected, $SkippedLayers, $Impact, $Remediation)
        $report.Add([PSCustomObject]@{
            Layer          = $Layer
            Prerequisite   = $Prerequisite
            Status         = $Status
            Detail         = $Detail
            LayersAffected = $LayersAffected
            SkippedLayers  = $SkippedLayers
            Impact         = $Impact
            Remediation    = $Remediation
        })
    }

    # ============================================================
    # CHECK 1 -- PowerShell version
    # ============================================================
    $psVersion = $PSVersionTable.PSVersion
    $psMajor   = $psVersion.Major
    $psEdition = $PSVersionTable.PSEdition

    $canUseParallel   = ($psMajor -ge 7)
    $canSkipCertCheck = ($psMajor -ge 6)

    & $addReport 0 'PowerShell Version' 'Available' "$psVersion ($psEdition)" `
        'All layers' '' '' ''

    # ============================================================
    # CHECK 2 -- Domain join
    # ============================================================
    $isDomainJoined = $false
    $domainName     = $null
    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -OperationTimeoutSec 5 -ErrorAction Stop
        $isDomainJoined = [bool]$cs.PartOfDomain
        if ($isDomainJoined) { $domainName = $cs.Domain }
    }
    catch {
        Write-Verbose "[Context] Win32_ComputerSystem query failed: $($_.Exception.Message)"
    }

    # ============================================================
    # CHECK 3 -- Domain Controller role
    # ============================================================
    $isDomainController = $false
    try {
        $productOptions = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\ProductOptions' `
            -ErrorAction Stop
        if ($productOptions.ProductType -eq 'LanmanNT') {
            $isDomainController = $true
        }
    }
    catch {
        Write-Verbose "[Context] ProductOptions registry read failed: $($_.Exception.Message)"
    }

    # ============================================================
    # CHECK 4 -- DNS resolution works (Layer 3 -- PTR)
    # ============================================================
    $dnsAvailable = $false
    $dnsServer    = $null
    try {
        if (Get-Command -Name Resolve-DnsName -ErrorAction SilentlyContinue) {
            $null = Resolve-DnsName -Name '127.0.0.1' -Type PTR -ErrorAction Stop
            $dnsAvailable = $true
        }
        else {
            $null = [System.Net.Dns]::GetHostEntry('127.0.0.1')
            $dnsAvailable = $true
        }
    }
    catch {
        Write-Verbose "[Context] DNS resolution test failed: $($_.Exception.Message)"
    }

    try {
        $dnsClient = Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
            Where-Object { $_.ServerAddresses.Count -gt 0 } |
            Select-Object -First 1
        if ($dnsClient) {
            $dnsServer = $dnsClient.ServerAddresses[0]
        }
    }
    catch {
        Write-Verbose "[Context] DNS client server enumeration failed: $($_.Exception.Message)"
    }

    if ($dnsAvailable) {
        & $addReport 3 'DNS (PTR)' 'Available' 'Resolve-DnsName / GetHostEntry works' `
            'Layer 3 (PTR)' '' '' ''
    }
    else {
        & $addReport 3 'DNS (PTR)' 'Unavailable' 'Resolve-DnsName failed' `
            'Layer 3 (PTR)' 'Layer 3 (PTR)' `
            'PTR-based hostname resolution unavailable' `
            'Verify DNS client configuration; try Test-NetConnection <ip> -InformationLevel Detailed'
    }

    # ============================================================
    # CHECK 5 -- Local DNS server role
    # ============================================================
    $dnsIsLocal = $false
    $dnsService = Get-Service -Name 'DNS' -ErrorAction SilentlyContinue
    if ($dnsService -and $dnsService.Status -eq 'Running') {
        $dnsIsLocal = $true
    }

    # ============================================================
    # CHECK 6 -- AD module + Get-ADDomain
    # ============================================================
    $adAvailable = $false
    $adIsLocal   = $isDomainController
    try {
        Import-Module -Name ActiveDirectory -ErrorAction Stop
        $null = Get-ADDomain -ErrorAction Stop
        $adAvailable = $true
    }
    catch {
        Write-Verbose "[Context] AD module/Get-ADDomain failed: $($_.Exception.Message)"
    }

    if ($adAvailable) {
        $detail = if ($adIsLocal) { 'Local DC -- auth-free queries' } else { 'Remote AD via RSAT' }
        & $addReport 1 'Active Directory' 'Available' $detail `
            'Layer 1 (AD)' '' '' ''
    }
    else {
        & $addReport 1 'Active Directory' 'Unavailable' 'AD module not loaded or Get-ADDomain failed' `
            'Layer 1 (AD)' 'Layer 1 (AD)' `
            'Cannot resolve domain-joined hosts via AD -- relies on DHCP/PTR/probe layers' `
            'Install RSAT-AD-PowerShell: Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"'
    }

    # ============================================================
    # CHECK 7 + 8 -- DhcpServer module + reachable
    # ============================================================
    $dhcpAvailable = $false
    $dhcpScopeIdsResolved = @()
    try {
        Import-Module -Name DhcpServer -ErrorAction Stop

        if ($DHCPServer) {
            $scopes = Get-DhcpServerv4Scope -ComputerName $DHCPServer -ErrorAction Stop
            $dhcpAvailable = $true
            if ($DHCPScopeIds.Count -gt 0) {
                $dhcpScopeIdsResolved = $DHCPScopeIds
            }
            else {
                $dhcpScopeIdsResolved = @($scopes | ForEach-Object { $_.ScopeId.ToString() })
            }
        }
        elseif ($isDomainController -or (Get-Service -Name 'DHCPServer' -ErrorAction SilentlyContinue)) {
            # Local DHCP role available -- try localhost
            $scopes = Get-DhcpServerv4Scope -ErrorAction Stop
            $dhcpAvailable = $true
            $dhcpScopeIdsResolved = @($scopes | ForEach-Object { $_.ScopeId.ToString() })
        }
    }
    catch {
        Write-Verbose "[Context] DhcpServer module/scope query failed: $($_.Exception.Message)"
    }

    if ($dhcpAvailable) {
        $detail = if ($DHCPServer) { "Server: $DHCPServer ($($dhcpScopeIdsResolved.Count) scopes)" } else { "Local DHCP role ($($dhcpScopeIdsResolved.Count) scopes)" }
        & $addReport 2 'DHCP Leases' 'Available' $detail `
            'Layer 2 (DHCP)' '' '' ''
    }
    elseif ($DHCPServer) {
        & $addReport 2 'DHCP Leases' 'Unavailable' "DhcpServer module load or scope query against '$DHCPServer' failed" `
            'Layer 2 (DHCP)' 'Layer 2 (DHCP)' `
            'No DHCP-derived hostname/MAC for leased IPs' `
            'Install RSAT-DHCP: Add-WindowsCapability -Online -Name "Rsat.DHCP.Tools~~~~0.0.1.0"'
    }
    else {
        & $addReport 2 'DHCP Leases' 'NotConfigured' 'No -DHCPServer supplied and no local DHCP role detected' `
            'Layer 2 (DHCP)' 'Layer 2 (DHCP)' `
            'No DHCP-derived hostname/MAC for leased IPs' `
            'Re-run Get-VBEnrichmentContext with -DHCPServer <fqdn>'
    }

    # ============================================================
    # CHECK 9 -- Local DHCP server role
    # ============================================================
    $dhcpIsLocal = $false
    $dhcpService = Get-Service -Name 'DHCPServer' -ErrorAction SilentlyContinue
    if ($dhcpService -and $dhcpService.Status -eq 'Running') {
        $dhcpIsLocal = $true
    }

    # ============================================================
    # CHECK 10 -- SNMP olePrn COM
    # ============================================================
    $snmpAvailable = $false
    $snmp = $null
    try {
        $snmp = New-Object -ComObject olePrn.OleSNMP -ErrorAction Stop
        $snmpAvailable = $true
    }
    catch {
        Write-Verbose "[Context] olePrn.OleSNMP COM creation failed: $($_.Exception.Message)"
    }
    finally {
        if ($snmp) {
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($snmp) | Out-Null } catch { }
            $snmp = $null
        }
    }

    if ($snmpAvailable) {
        & $addReport 7 'SNMP (olePrn)' 'Available' 'COM object created OK' `
            'Layer 7 (SNMP), Layer 10 (Switch ARP)' '' '' ''
    }
    else {
        & $addReport 7 'SNMP (olePrn)' 'Unavailable' 'olePrn.OleSNMP COM not creatable' `
            'Layer 7 (SNMP), Layer 10 (Switch ARP)' 'Layer 7 (SNMP), Layer 10 (Switch)' `
            'No SNMP sysName/sysLocation; no switch ARP/port lookup' `
            'Install Print and Document Services -> LPD Service feature; reboot if required'
    }

    # ============================================================
    # CHECK 11 -- OUI file
    # ============================================================
    $ouiFileAvailable = $false
    $ouiFileAge = $null
    if (Test-Path -LiteralPath $OUIFilePath) {
        $ouiItem = Get-Item -LiteralPath $OUIFilePath
        # Per design: "if file > 1MB" treat as valid; otherwise treat as missing/corrupt
        if ($ouiItem.Length -gt 1MB) {
            $ouiFileAvailable = $true
            $ouiFileAge = (Get-Date) - $ouiItem.CreationTime
        }
    }

    if ($ouiFileAvailable) {
        $sizeMB = [math]::Round($ouiItem.Length / 1MB, 1)
        $ageDays = [int]$ouiFileAge.TotalDays
        & $addReport 11 'OUI Vendor Lookup' 'Available' "oui.csv $sizeMB MB, age $ageDays days" `
            'Layer 11 (OUI)' '' '' ''
    }
    else {
        & $addReport 11 'OUI Vendor Lookup' 'NotConfigured' "oui.csv missing or < 1MB at $OUIFilePath" `
            'Layer 11 (OUI)' '' `
            'No vendor lookup until first run -- Get-VBOUIVendor downloads on first call' `
            'Will be auto-downloaded on first run; or pre-stage with Invoke-WebRequest -Uri https://standards-oui.ieee.org/oui/oui.csv'
    }

    # ============================================================
    # CHECK 12 -- mDNS dns-sd.exe
    # ============================================================
    $mDNSAvailable = $false
    if (Get-Command -Name 'dns-sd.exe' -ErrorAction SilentlyContinue) {
        $mDNSAvailable = $true
    }

    if ($mDNSAvailable) {
        & $addReport 9 'mDNS Discovery' 'Available' 'dns-sd.exe found on PATH' `
            'Layer 9 (mDNS)' '' '' ''
    }
    else {
        & $addReport 9 'mDNS Discovery' 'Unavailable' 'dns-sd.exe not on PATH' `
            'Layer 9 (mDNS)' 'Layer 9 (mDNS)' `
            'Printers/scanners using mDNS only may be missed' `
            'Install Bonjour Print Services from Apple, or skip if mDNS is not in use'
    }

    # ============================================================
    # CHECK 13 -- PSSQLite module
    # ============================================================
    $canUsePSSQLite = [bool](Get-Module -Name PSSQLite -ListAvailable)
    if ($canUsePSSQLite) {
        & $addReport 0 'PSSQLite Module' 'Available' 'Module loadable' `
            'Storage layer' '' '' ''
    }
    else {
        & $addReport 0 'PSSQLite Module' 'Unavailable' 'PSSQLite not installed' `
            'Storage layer' 'All caching/persistence' `
            'No SQLite cache -- every run re-probes every IP' `
            'Install-Module PSSQLite -Scope CurrentUser'
    }

    # ============================================================
    # CHECK 14 -- SQLite database init
    # ============================================================
    $databaseInitialized = $false
    if ($canUsePSSQLite) {
        try {
            $initResult = Initialize-VBEnrichmentDatabase -DatabasePath $DatabasePath
            if ($initResult.Status -eq 'Success') {
                $databaseInitialized = $true
                & $addReport 0 'SQLite Database' 'Available' "$DatabasePath (schema v$($initResult.Version))" `
                    'Storage layer' '' '' ''
            }
            else {
                & $addReport 0 'SQLite Database' 'Unavailable' $initResult.Error `
                    'Storage layer' 'All caching/persistence' `
                    'Cannot persist enrichment results' `
                    "Verify write access to $DatabasePath"
            }
        }
        catch {
            & $addReport 0 'SQLite Database' 'Unavailable' $_.Exception.Message `
                'Storage layer' 'All caching/persistence' `
                'Cannot persist enrichment results' `
                "Verify write access to $DatabasePath"
        }
    }

    # ============================================================
    # NetworkProbeEnabled / RTSPProbeEnabled / report rows for layers 4-8, 10
    # ============================================================
    $networkProbeEnabled = -not $DisableNetworkProbe
    $rtspProbeEnabled    = $networkProbeEnabled

    if ($networkProbeEnabled) {
        & $addReport 4 'ARP Cache'      'Available' 'Native -- arp -a'                 'Layer 4 (ARP)'  '' '' ''
        & $addReport 5 'TCP Port Scan'  'Available' 'Network probe enabled'            'Layer 5 (TCP)'  '' '' ''
        & $addReport 6 'HTTP Banner'    'Available' (
            if ($canSkipCertCheck) { 'PS 6+ native cert bypass' } else { 'PS 5.1 -- cert callback workaround' }
        ) 'Layer 6 (HTTP)' '' '' ''
        & $addReport 8 'RTSP Probe'     'Available' 'Network probe enabled'            'Layer 8 (RTSP)' '' '' ''
    }
    else {
        & $addReport 4 'ARP Cache'      'Skipped' 'Network probe disabled' 'Layer 4 (ARP)'  'Layer 4'  'No MAC for unleased IPs' 'Re-run without -DisableNetworkProbe'
        & $addReport 5 'TCP Port Scan'  'Skipped' 'Network probe disabled' 'Layer 5 (TCP)'  'Layer 5'  'No port-based device class signals' 'Re-run without -DisableNetworkProbe'
        & $addReport 6 'HTTP Banner'    'Skipped' 'Network probe disabled' 'Layer 6 (HTTP)' 'Layer 6'  'No HTTP banner -- many printers/cameras unidentified' 'Re-run without -DisableNetworkProbe'
        & $addReport 8 'RTSP Probe'     'Skipped' 'Network probe disabled' 'Layer 8 (RTSP)' 'Layer 8'  'No RTSP banner for cameras' 'Re-run without -DisableNetworkProbe'
    }

    if ($SwitchTargets.Count -gt 0 -and $snmpAvailable) {
        & $addReport 10 'Switch ARP' 'Available' "$($SwitchTargets.Count) switch(es) configured" 'Layer 10 (Switch)' '' '' ''
    }
    elseif ($SwitchTargets.Count -gt 0 -and -not $snmpAvailable) {
        & $addReport 10 'Switch ARP' 'Unavailable' 'SNMP unavailable' 'Layer 10 (Switch)' 'Layer 10' 'No switch port location lookup' 'Resolve SNMP COM availability'
    }
    else {
        & $addReport 10 'Switch ARP' 'NotConfigured' 'No SwitchTargets supplied' 'Layer 10 (Switch)' 'Layer 10' "Unresolved IPs won't get switch port location" 'Re-run with -SwitchTargets <ip[,ip...]>'
    }

    $sw.Stop()

    # ============================================================
    # Build the context object
    # ============================================================
    $context = [PSCustomObject]@{
        # PowerShell environment
        PSVersion             = $psVersion
        PSMajor               = $psMajor
        PSEdition             = $psEdition
        CanUseParallel        = $canUseParallel
        CanSkipCertCheck      = $canSkipCertCheck
        CanUsePSSQLite        = $canUsePSSQLite

        # Machine identity
        ComputerName          = $env:COMPUTERNAME
        IsDomainJoined        = $isDomainJoined
        IsDomainController    = $isDomainController
        DomainName            = $domainName

        # Layer availability
        DNSAvailable          = $dnsAvailable
        DNSIsLocal            = $dnsIsLocal
        DNSServer             = $dnsServer
        DHCPAvailable         = $dhcpAvailable
        DHCPIsLocal           = $dhcpIsLocal
        DHCPServer            = $DHCPServer
        DHCPScopeIds          = $dhcpScopeIdsResolved
        ADAvailable           = $adAvailable
        ADIsLocal             = $adIsLocal
        NetworkProbeEnabled   = $networkProbeEnabled
        SNMPAvailable         = $snmpAvailable
        SNMPCommunityStrings  = $SNMPCommunityStrings
        OUIFileAvailable      = $ouiFileAvailable
        OUIFilePath           = $OUIFilePath
        OUIFileAge            = $ouiFileAge
        mDNSAvailable         = $mDNSAvailable
        RTSPProbeEnabled      = $rtspProbeEnabled
        SwitchTargets         = $SwitchTargets

        # Storage
        DatabasePath          = $DatabasePath
        DatabaseInitialized   = $databaseInitialized

        # Performance
        ParallelThrottleLimit = $ParallelThrottleLimit
        DefaultTimeoutMs      = @{
            TCP   = 300
            HTTP  = 3000
            SNMP  = 2000
            RTSP  = 2000
        }

        # Reporting
        PrerequisiteReport    = $report.ToArray()
        ContextBuiltAt        = (Get-Date)
        ContextDurationMs     = [int]$sw.ElapsedMilliseconds
    }

    if (-not $Quiet) {
        Format-VBEnrichmentContext -Context $context
    }

    $context
}

function Format-VBEnrichmentContext {
<#
.SYNOPSIS
    Render the context object as a formatted console report.

.DESCRIPTION
    Internal display helper invoked by Get-VBEnrichmentContext when -Quiet is not
    supplied. Writes the report via Write-Host using the VB colour palette.
    Returns nothing to the pipeline.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Context
    )

    $line = ('=' * 71)
    $sep  = ('-' * 71)

    Write-Host ''
    Write-Host $line -ForegroundColor Cyan
    Write-Host '  VB.DNSEnrichment -- Environment Prerequisites' -ForegroundColor Cyan
    Write-Host $line -ForegroundColor Cyan

    $machineLine = "  Computer    : $($Context.ComputerName)"
    if ($Context.IsDomainController) { $machineLine += ' (Domain Controller)' }
    elseif ($Context.IsDomainJoined) { $machineLine += " (Domain: $($Context.DomainName))" }
    else                             { $machineLine += ' (Workgroup)' }
    Write-Host $machineLine -ForegroundColor White

    $parallelText = if ($Context.CanUseParallel) { 'ENABLED' } else { 'DISABLED (PS < 7)' }
    Write-Host "  PowerShell  : $($Context.PSVersion) $($Context.PSEdition)  -> Parallel execution: $parallelText" -ForegroundColor White

    $dbText = if ($Context.DatabaseInitialized) { "$($Context.DatabasePath) (initialized)" } else { "$($Context.DatabasePath) (NOT INITIALIZED)" }
    Write-Host "  Database    : $dbText" -ForegroundColor White

    $probeText = if ($Context.NetworkProbeEnabled) { 'Active probes ENABLED' } else { 'Active probes DISABLED' }
    Write-Host "  Probes      : $probeText" -ForegroundColor White

    Write-Host $sep -ForegroundColor Gray
    Write-Host ''
    Write-Host '  Layer  Prerequisite          Status         Detail' -ForegroundColor White
    Write-Host '  -----  -------------------   ------------   -------------------------------' -ForegroundColor Gray

    $available    = 0
    $unavailable  = 0
    $notConfigured = 0
    $skipped      = 0
    $skippedLayers = New-Object System.Collections.Generic.List[string]

    foreach ($row in $Context.PrerequisiteReport) {
        $statusColour = switch ($row.Status) {
            'Available'     { 'Green';   $available++;     break }
            'Unavailable'   { 'Red';     $unavailable++;   break }
            'NotConfigured' { 'Yellow';  $notConfigured++; break }
            'Skipped'       { 'Yellow';  $skipped++;       break }
            'Degraded'      { 'Yellow';  break }
            default         { 'White' }
        }

        $layerText = if ($row.Layer -gt 0) { ([string]$row.Layer).PadLeft(2) } else { '--' }
        $preq      = $row.Prerequisite.PadRight(20)
        $statusPad = $row.Status.PadRight(13)
        $detail    = if ($row.Detail) { $row.Detail } else { '' }

        Write-Host ('   {0}    {1}  ' -f $layerText, $preq) -NoNewline -ForegroundColor White
        Write-Host $statusPad -NoNewline -ForegroundColor $statusColour
        Write-Host "  $detail" -ForegroundColor White

        if ($row.Status -ne 'Available' -and $row.SkippedLayers) {
            $skippedLayers.Add($row.SkippedLayers)
            if ($row.Impact) {
                Write-Host ("                                              -> Impact: $($row.Impact)") -ForegroundColor Gray
            }
            if ($row.Remediation) {
                Write-Host ("                                              -> Fix: $($row.Remediation)") -ForegroundColor Gray
            }
        }
    }

    Write-Host ''
    Write-Host $line -ForegroundColor Cyan
    Write-Host "  Summary: $available available, $unavailable unavailable, $notConfigured not configured" -ForegroundColor White
    if ($skippedLayers.Count -gt 0) {
        Write-Host "  Skipped: $($skippedLayers -join ', ')" -ForegroundColor Yellow
    }
    $modeText = if ($Context.CanUseParallel) { "Parallel x $($Context.ParallelThrottleLimit) on active probes" } else { 'Sequential (PS 5.1)' }
    Write-Host "  Mode   : $modeText" -ForegroundColor White
    Write-Host "  Built  : $($Context.ContextDurationMs) ms" -ForegroundColor Gray
    Write-Host $line -ForegroundColor Cyan
    Write-Host ''
}
