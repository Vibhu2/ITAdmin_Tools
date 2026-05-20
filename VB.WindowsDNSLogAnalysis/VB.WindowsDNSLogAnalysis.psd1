#
# Module manifest for module 'VB.WindowsDNSLogAnalysis'
# Generated: 2026-05-07
# Author   : Vibhu Bhatnagar
#

@{

    # -- Module identity
    ModuleVersion     = '1.1.1'
    GUID              = 'a4c2e8f1-3b7d-4e91-82a0-5f6c9d1e4b23'
    Author            = 'Vibhu Bhatnagar'
    CompanyName       = 'Internal IT'
    Copyright         = '(c) 2026 Vibhu Bhatnagar. All rights reserved.'
    Description       = 'Parse Windows DNS debug logs once, store all records in SQLite, and expose the data through SQL-backed query functions. Designed for internal DNS traffic analysis, troubleshooting, and security investigation.'

    # -- Runtime requirements
    PowerShellVersion = '7.0'

    # -- Module entry point
    RootModule        = 'VB.WindowsDNSLogAnalysis.psm1'

    # -- Required modules (must be installed before importing VB.WindowsDNSLogAnalysis)
    RequiredModules   = @(
        @{ ModuleName = 'PSSQLite'; ModuleVersion = '1.0.0' }
    )

    # -- Exported public functions (Private functions are NOT listed here)
    FunctionsToExport = @(
        'Initialize-VBDNSLogDatabase',
        'Import-VBDNSLog',
        'Get-VBDNSLog',
        'Invoke-VBDNSLogQuery',
        'Get-VBDNSLogStatistics',
        'Export-VBDNSLogReport'
    )

    # -- Not exporting cmdlets, variables, or aliases
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    # -- Module metadata
    PrivateData = @{
        PSData = @{
            Tags         = @('DNS', 'SQLite', 'Logging', 'Analysis', 'WindowsDNS', 'Diagnostics', 'Security')
            ProjectUri   = 'https://github.com/VibhuBhatnagar/ITAdmin_Tools'
            ReleaseNotes = @"
v1.1.1 (2026-05-20)
- Patch release to trigger PSGallery publish (initial publish failed due to missing PSSQLite dependency on CI runner)

v1.1.0 (2026-05-07)
- Renamed module from DNSLogDB to VB.WindowsDNSLogAnalysis
- Added TalkerDetail report type (per-IP query type and protocol breakdown)
- TopTalkers/TopDomains: -Top now optional (omit for all results)
- TopTalkers/TopDomains: summary always shown, no switch needed
- Fixed ResponseCode scoping to response packets only (PacketKind=R)
- Added parse-phase and insert-phase progress bars with real percentage

v1.0.0 (2026-05-07)
- Initial release
- StreamReader + compiled regex parsing engine (2-5 min for 450MB file)
- SHA256 dedup via ImportLog table
- ForEach-Object -Parallel multi-file import (PS7)
- 9 pre-built statistical reports (Get-VBDNSLogStatistics)
- CSV, HTML, XLSX export (Export-VBDNSLogReport)
"@
        }
    }

}
