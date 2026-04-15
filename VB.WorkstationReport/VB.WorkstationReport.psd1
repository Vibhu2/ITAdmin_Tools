#
# VB.WorkstationReport.psd1 -- Module manifest
# Version : 1.2.0
# Author  : Vibhu Bhatnagar
#

@{
    # Module identity
    ModuleVersion     = '1.2.0'
    GUID              = 'a3f9d2b1-4c7e-4f8a-9b2d-1e5f6c3a7d0e'
    Author            = 'Vibhu Bhatnagar'
    CompanyName       = 'Realtime-IT'
    Description       = 'Workstation reporting and Nextcloud upload utilities for Windows sysadmin environments.'
    PowerShellVersion = '5.1'

    # Root module
    RootModule        = 'VB.WorkstationReport.psm1'

    # Dependencies
    RequiredModules   = @('VB.NextCloud')

    # Exported functions -- loader controls exports dynamically, never list individually
    FunctionsToExport = '*'

    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    # Module metadata
    PrivateData = @{
        PSData = @{
            Tags         = @('Workstation', 'Reporting', 'Nextcloud', 'Printers', 'FolderRedirection', 'SyncCenter')
            ProjectUri   = 'https://github.com/Vibhu2/ITAdmin_Tools'
            ReleaseNotes = @'
v1.2.0 -- 15-04-2026 -- Function rename for clarity
- Renamed: Set-VBNextcloudFiles -> Start-VBNextcloudUpload (batch upload orchestrator)
- Set-VBNextcloudFile (single file PUT) remains unchanged
- Updated: New-VBNextcloudFolder caller reference

v1.1.1 -- 14-04-2026 -- Standards compliance fixes
- Added mandatory function file headers to all .ps1 files
- Fixed: FunctionsToExport = '*' (was listing individual functions)
- Fixed: Author changed to 'Vibhu' (was 'VB Admin Tools')
- Fixed: DarkGray colour replaced with Gray across all functions
- Fixed: Non-ASCII box-drawing characters removed from source files
- Fixed: ComputerName added as first property to Nextcloud function outputs
- Fixed: CollectionTime added to Get-VBUserFolderRedirections and Get-VBNextcloudFiles outputs
- Fixed: .psm1 uses inline Export-ModuleMember pattern (no explicit list)

v1.1.0 -- Initial modular release
- Split monolithic script into individual .ps1 files (one function per file)
- DRY refactor: eliminated local/remote code duplication across all collection functions
- Fixed: Set-VBNextcloudFile process block missing opening brace
- Fixed: Set-VBNextcloudFiles replaced plain-text Username/Password with PSCredential
- Fixed: Invoke-VBWorkstationReport variable reuse bug ($VBUserFolderRedirections used for NI output)
- New:   New-VBNextcloudFolder (Private) — was referenced but never defined
- New:   Invoke-VBWorkstationReport — wraps report generation; removed hardcoded credentials
- Improved: switch-based registry value translation in Get-VBSyncCenterStatus
- Improved: Generic.List used instead of += array concatenation throughout
- Improved: Full comment-based help added to all functions
'@
        }
    }
}
