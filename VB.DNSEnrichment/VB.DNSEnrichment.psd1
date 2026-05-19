#
# Module manifest for module 'VB.DNSEnrichment'
# Generated: 2026-05-10
# Author   : Vibhu Bhatnagar
#

@{

    # -- Module identity
    ModuleVersion     = '0.5.0'
    GUID              = '4bbf090c-c895-4a7f-babd-fa826a1787c1'
    Author            = 'Vibhu Bhatnagar'
    CompanyName       = 'pwsh.in'
    Copyright         = '(c) 2026 Vibhu. All rights reserved.'
    Description       = 'Resolve private IP addresses to hostname, MAC, vendor, and device class through an 11-step enrichment pipeline. Persists results in SQLite for cross-run caching and DHCP churn analysis.'

    # -- Runtime requirements
    PowerShellVersion = '5.1'

    # -- Module entry point
    RootModule        = 'VB.DNSEnrichment.psm1'

    # -- Required modules (must be installed before importing VB.DNSEnrichment)
    RequiredModules   = @(
        @{ ModuleName = 'PSSQLite'; ModuleVersion = '1.0.0' }
    )

    # -- Exported public functions (Private functions are NOT listed here)
    # Round 1 ships only the context/storage layer; subsequent rounds add layers.
    FunctionsToExport = @(
        'Get-VBEnrichmentContext',
        'Initialize-VBEnrichmentDatabase',
        'Get-VBADComputer',
        'Get-VBDHCPLease',
        'Get-VBPTRRecord',
        'Get-VBARPEntry',
        'Resolve-VBDeviceClass',
        'Invoke-VBIPEnrichment',
        'Get-VBEnrichmentResult',
        'Export-VBEnrichmentResult',
        'Get-VBTCPFingerprint',
        'Get-VBHTTPBanner',
        'Get-VBSNMPIdentity',
        'Get-VBOUIVendor',
        'Get-VBRTSPBanner',
        'Get-VBmDNSRecord',
        'Get-VBSwitchARP'
    )

    # -- Not exporting cmdlets, variables, or aliases
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    # -- Module metadata
    PrivateData = @{
        PSData = @{
            Tags         = @('DNS', 'Enrichment', 'IP', 'SQLite', 'AD', 'DHCP', 'SNMP', 'Networking', 'RTSP', 'mDNS', 'PowerShell', 'Sysadmin', 'Windows', 'MSP', 'Automation', 'VBTools')
            ProjectUri   = 'https://github.com/Vibhu2/ITAdmin_Tools/tree/main/VB.DNSEnrichment'
            LicenseUri   = 'https://github.com/Vibhu2/ITAdmin_Tools/blob/main/LICENSE'
            HelpInfoUri  = 'https://pwsh.in/posts/'
            ReleaseNotes = @"
v0.4.0 (2026-05-11) -- Round 4: Sensor layers (RTSP, mDNS, Switch ARP) + PS7 parallel active probes
- Get-VBRTSPBanner (Layer 8): TCP 554 banner grab; high-confidence camera/NVR signal
- Get-VBmDNSRecord (Layer 9): dns-sd.exe browse+resolve for 5 service types; one-shot session cache
- Get-VBSwitchARP (Layer 10): SNMP walk of ARP/FDB/ifAlias on managed switches; one-shot session cache
- Invoke-VBIPEnrichment: PS7 ForEach-Object -Parallel for active probes (steps 5-11) when Context.CanUseParallel; passive layers always sequential; classification + SQLite always sequential

v0.3.0 (2026-05-11) -- Round 3: Active probe layers (TCP, HTTP, SNMP, OUI)
- Get-VBTCPFingerprint: 24-port async concurrent scan using BeginConnect/WaitOne (~300 ms)
- Get-VBHTTPBanner: title + Server header, PS 5.1 cert callback workaround, 4-port fallback chain
- Get-VBSNMPIdentity: olePrn COM sysDescr/sysName/sysLocation, community string iteration
- Get-VBOUIVendor: IEEE OUI CSV auto-download + 30-day refresh, session hashtable cache, vendor->class map
- Invoke-VBIPEnrichment: active probe stubs replaced; SNMP gates on context flag; OUI always runs if MAC known
- Orchestrator now wires HTTP/SNMP hostname resolution into IsResolved gate

v0.2.0 (2026-05-11) -- Round 2: Passive layers + orchestrator (Phase 6 ship gate)
- Get-VBADComputer (Layer 1): one-shot AD cache, OSClass/OU detection
- Get-VBDHCPLease (Layer 2): one-shot DHCP lease cache across all scopes
- Get-VBPTRRecord (Layer 3): PTR lookup with forward-confirmation, PS5.1 fallback
- Get-VBARPEntry (Layer 4): arp -a cache with optional ping-to-populate
- Resolve-VBDeviceClass: 13-tier classification logic (pure signal -> class)
- Invoke-VBIPEnrichment: orchestrator with SQLite cache, stale-row detection, DHCP churn history
- Get-VBEnrichmentResult: SQLite query with IP/class/date/unresolved filters
- Export-VBEnrichmentResult: CSV (UTF-8 BOM) / JSON / Object pipeline pass-through

v0.1.0 (2026-05-10) -- Round 1: Skeleton + Context + Storage
- Module skeleton (manifest, loader, .editorconfig)
- Sql/001_init.sql schema (Enrichment, EnrichmentHistory, SchemaVersion)
- Private helpers: Test-VBPrivateIP, ConvertTo-VBNormalisedMAC, New-VBLayerResult, Invoke-VBSqliteCommand, Write-VBEnrichmentProgress
- Get-VBEnrichmentContext (14 prerequisite checks + console report)
- Initialize-VBEnrichmentDatabase (idempotent schema migration)
"@
        }
    }

}
