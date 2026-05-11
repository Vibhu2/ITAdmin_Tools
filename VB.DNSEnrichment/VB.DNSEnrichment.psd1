#
# Module manifest for module 'VB.DNSEnrichment'
# Generated: 2026-05-10
# Author   : Vibhu Bhatnagar
#

@{

    # -- Module identity
    ModuleVersion     = '0.2.0'
    GUID              = '4bbf090c-c895-4a7f-babd-fa826a1787c1'
    Author            = 'Vibhu Bhatnagar'
    CompanyName       = 'Internal IT'
    Copyright         = '(c) 2026 Vibhu Bhatnagar. All rights reserved.'
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
        'Export-VBEnrichmentResult'
    )

    # -- Not exporting cmdlets, variables, or aliases
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    # -- Module metadata
    PrivateData = @{
        PSData = @{
            Tags         = @('DNS', 'Enrichment', 'IP', 'SQLite', 'AD', 'DHCP', 'SNMP', 'Networking')
            ProjectUri   = 'https://github.com/VibhuBhatnagar/ITAdmin_Tools'
            ReleaseNotes = @"
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
