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

    # Resolve default paths -- derive module root from the .psm1 path, not $PSScriptRoot
    if (-not $OUIFilePath) {
        $moduleRoot  = Split-Path -Path (Get-Module -Name 'VB.DNSEnrichment').Path -Parent
        $OUIFilePath = Join-Path $moduleRoot 'Data\oui.csv'
        Write-Verbose "[Context] OUI path: $OUIFilePath"
    }
    if (-not $DatabasePath) {
        $DatabasePath = Join-Path $env:LOCALAPPDATA 'VB.DNSEnrichment\enrichment.db'
    }

    # --- Helper: append a PrerequisiteReport row ---
    $addReport = {
        param($Layer, $Prerequisite, $Status, $Detail, $LayersAffected, $SkippedLayers, $Impact, $Remediation, $Severity = 'Info')
        $report.Add([PSCustomObject]@{
            Layer          = $Layer
            Prerequisite   = $Prerequisite
            Status         = $Status
            Severity       = $Severity
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
    $psVersion    = $PSVersionTable.PSVersion
    $psMajor      = $psVersion.Major
    $psEditionStr = $PSVersionTable.PSEdition

    $canUseParallel = $false
    if ($psMajor -ge 7) {
        try {
            $null = [System.Management.Automation.Runspaces.Runspace]::DefaultRunspace
            $canUseParallel = $true
        }
        catch {
            Write-Verbose "[Context] ForEach-Object -Parallel not available in this runspace: $($_.Exception.Message)"
        }
    }
    $canSkipCertCheck = ($psMajor -ge 6)

    & $addReport 0 'PowerShell Version' 'Available' "$psVersion ($psEditionStr)" `
        'All layers' '' '' '' 'Info'

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
            'Layer 3 (PTR)' '' '' '' 'Info'
    }
    else {
        & $addReport 3 'DNS (PTR)' 'Unavailable' 'Resolve-DnsName failed' `
            'Layer 3 (PTR)' 'Layer 3 (PTR)' `
            'PTR-based hostname resolution unavailable' `
            'Verify DNS client configuration; try Test-NetConnection <ip> -InformationLevel Detailed' `
            'Critical'
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
            'Layer 1 (AD)' '' '' '' 'Info'
    }
    else {
        & $addReport 1 'Active Directory' 'Unavailable' 'AD module not loaded or Get-ADDomain failed' `
            'Layer 1 (AD)' 'Layer 1 (AD)' `
            'Cannot resolve domain-joined hosts via AD -- relies on DHCP/PTR/probe layers' `
            'Install RSAT-AD-PowerShell: Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"' `
            'Critical'
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
            'Layer 2 (DHCP)' '' '' '' 'Info'
    }
    elseif ($DHCPServer) {
        & $addReport 2 'DHCP Leases' 'Unavailable' "DhcpServer module load or scope query against '$DHCPServer' failed" `
            'Layer 2 (DHCP)' 'Layer 2 (DHCP)' `
            'No DHCP-derived hostname/MAC for leased IPs' `
            'Install RSAT-DHCP: Add-WindowsCapability -Online -Name "Rsat.DHCP.Tools~~~~0.0.1.0"' `
            'Warning'
    }
    else {
        & $addReport 2 'DHCP Leases' 'NotConfigured' 'No -DHCPServer supplied and no local DHCP role detected' `
            'Layer 2 (DHCP)' 'Layer 2 (DHCP)' `
            'No DHCP-derived hostname/MAC for leased IPs' `
            'Re-run Get-VBEnrichmentContext with -DHCPServer <fqdn>' `
            'Warning'
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
            'Layer 7 (SNMP), Layer 10 (Switch ARP)' '' '' '' 'Info'
    }
    else {
        & $addReport 7 'SNMP (olePrn)' 'Unavailable' 'olePrn.OleSNMP COM not creatable' `
            'Layer 7 (SNMP), Layer 10 (Switch ARP)' 'Layer 7 (SNMP), Layer 10 (Switch)' `
            'No SNMP sysName/sysLocation; no switch ARP/port lookup' `
            'Install Print and Document Services -> LPD Service feature; reboot if required' `
            'Warning'
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
            $ouiFileAge = (Get-Date) - $ouiItem.LastWriteTime
        }
    }

    if ($ouiFileAvailable) {
        $sizeMB = [math]::Round($ouiItem.Length / 1MB, 1)
        $ageDays = [int]$ouiFileAge.TotalDays
        & $addReport 11 'OUI Vendor Lookup' 'Available' "oui.csv $sizeMB MB, age $ageDays days" `
            'Layer 11 (OUI)' '' '' '' 'Info'
    }
    else {
        & $addReport 11 'OUI Vendor Lookup' 'NotConfigured' "oui.csv missing or < 1MB at $OUIFilePath" `
            'Layer 11 (OUI)' '' `
            'No vendor lookup until first run -- Get-VBOUIVendor downloads on first call' `
            'Will be auto-downloaded on first run; or pre-stage with Invoke-WebRequest -Uri https://standards-oui.ieee.org/oui/oui.csv' `
            'Warning'
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
            'Layer 9 (mDNS)' '' '' '' 'Info'
    }
    else {
        & $addReport 9 'mDNS Discovery' 'Unavailable' 'dns-sd.exe not on PATH' `
            'Layer 9 (mDNS)' 'Layer 9 (mDNS)' `
            'Printers/scanners using mDNS only may be missed' `
            'Install Bonjour Print Services from Apple, or skip if mDNS is not in use' `
            'Info'
    }

    # ============================================================
    # CHECK 13 -- PSSQLite module
    # ============================================================
    $canUsePSSQLite = [bool](Get-Module -Name PSSQLite -ListAvailable)
    if ($canUsePSSQLite) {
        & $addReport 0 'PSSQLite Module' 'Available' 'Module loadable' `
            'Storage layer' '' '' '' 'Info'
    }
    else {
        & $addReport 0 'PSSQLite Module' 'Unavailable' 'PSSQLite not installed' `
            'Storage layer' 'All caching/persistence' `
            'No SQLite cache -- every run re-probes every IP' `
            'Install-Module PSSQLite -Scope CurrentUser' `
            'Warning'
    }

    # ============================================================
    # CHECK 14 -- SQLite database init
    # ============================================================
    $databaseInitialized = $false
    if ($canUsePSSQLite) {
        try {
            $sqlFolder = Join-Path (Split-Path (Get-Module 'VB.DNSEnrichment').Path -Parent) 'Sql'
            if (-not (Test-Path -LiteralPath $sqlFolder)) {
                & $addReport 0 'SQLite Database' 'Unavailable' `
                    "Sql migration folder not found at: $sqlFolder" `
                    'Storage layer' 'All caching/persistence' `
                    'Cannot persist enrichment results' `
                    "Verify the module is imported from its installed location (Sql\ folder must exist beside .psm1)" `
                    'Critical'
            }
            else {
                $initResult = Initialize-VBEnrichmentDatabase -DatabasePath $DatabasePath
                if ($initResult.Status -eq 'Success') {
                    $databaseInitialized = $true
                    & $addReport 0 'SQLite Database' 'Available' "$DatabasePath (schema v$($initResult.Version))" `
                        'Storage layer' '' '' '' 'Info'
                }
                else {
                    & $addReport 0 'SQLite Database' 'Unavailable' "$($initResult.Error) | Expected SQL folder: $sqlFolder" `
                        'Storage layer' 'All caching/persistence' `
                        'Cannot persist enrichment results' `
                        "Verify write access to $DatabasePath and that Sql\001_init.sql exists" `
                        'Critical'
                }
            }
        }
        catch {
            $sqlFolderHint = Join-Path (Split-Path (Get-Module 'VB.DNSEnrichment').Path -Parent) 'Sql'
            & $addReport 0 'SQLite Database' 'Unavailable' "$($_.Exception.Message) | Expected SQL folder: $sqlFolderHint" `
                'Storage layer' 'All caching/persistence' `
                'Cannot persist enrichment results' `
                "Verify write access to $DatabasePath and that Sql\001_init.sql exists" `
                'Critical'
        }
    }

    # ============================================================
    # NetworkProbeEnabled / RTSPProbeEnabled / report rows for layers 4-8, 10
    # ============================================================
    $networkProbeEnabled = -not $DisableNetworkProbe
    $rtspProbeEnabled    = $networkProbeEnabled

    if ($networkProbeEnabled) {
        & $addReport 4 'ARP Cache'      'Available' 'Native -- arp -a'                 'Layer 4 (ARP)'  '' '' '' 'Info'
        & $addReport 5 'TCP Port Scan'  'Available' 'Network probe enabled'            'Layer 5 (TCP)'  '' '' '' 'Info'
        $httpBannerDetail = if ($canSkipCertCheck) { 'PS 6+ native cert bypass' } else { 'PS 5.1 -- cert callback workaround' }
        & $addReport 6 'HTTP Banner'    'Available' $httpBannerDetail 'Layer 6 (HTTP)' '' '' '' 'Info'
        & $addReport 8 'RTSP Probe'     'Available' 'Network probe enabled'            'Layer 8 (RTSP)' '' '' '' 'Info'
    }
    else {
        & $addReport 4 'ARP Cache'      'Skipped' 'Network probe disabled' 'Layer 4 (ARP)'  'Layer 4'  'No MAC for unleased IPs' 'Re-run without -DisableNetworkProbe' 'Info'
        & $addReport 5 'TCP Port Scan'  'Skipped' 'Network probe disabled' 'Layer 5 (TCP)'  'Layer 5'  'No port-based device class signals' 'Re-run without -DisableNetworkProbe' 'Info'
        & $addReport 6 'HTTP Banner'    'Skipped' 'Network probe disabled' 'Layer 6 (HTTP)' 'Layer 6'  'No HTTP banner -- many printers/cameras unidentified' 'Re-run without -DisableNetworkProbe' 'Info'
        & $addReport 8 'RTSP Probe'     'Skipped' 'Network probe disabled' 'Layer 8 (RTSP)' 'Layer 8'  'No RTSP banner for cameras' 'Re-run without -DisableNetworkProbe' 'Info'
    }

    if ($SwitchTargets.Count -gt 0 -and $snmpAvailable) {
        & $addReport 10 'Switch ARP' 'Available' "$($SwitchTargets.Count) switch(es) configured" 'Layer 10 (Switch)' '' '' '' 'Info'
    }
    elseif ($SwitchTargets.Count -gt 0 -and -not $snmpAvailable) {
        & $addReport 10 'Switch ARP' 'Unavailable' 'SNMP unavailable' 'Layer 10 (Switch)' 'Layer 10' 'No switch port location lookup' 'Resolve SNMP COM availability' 'Warning'
    }
    else {
        & $addReport 10 'Switch ARP' 'NotConfigured' 'No SwitchTargets supplied' 'Layer 10 (Switch)' 'Layer 10' "Unresolved IPs won't get switch port location" 'Re-run with -SwitchTargets <ip[,ip...]>' 'Info'
    }

    $sw.Stop()

    # ============================================================
    # Build the context object
    # ============================================================
    $context = [PSCustomObject]@{
        # PowerShell environment
        PSVersion             = $psVersion
        PSMajor               = $psMajor
        PSEdition             = $psEditionStr
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
        SNMPCommunityStrings  = @($SNMPCommunityStrings | ForEach-Object { ConvertTo-SecureString $_ -AsPlainText -Force })
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
            TCP   = 1000
            HTTP  = 5000
            SNMP  = 3000
            RTSP  = 3000
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
