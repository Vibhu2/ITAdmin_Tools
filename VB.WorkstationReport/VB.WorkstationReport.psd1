#
# VB.WorkstationReport.psd1 -- Module manifest
# Version : 1.5.1
# Author  : Vibhu Bhatnagar
#

@{
    # Module identity
    ModuleVersion     = '1.11.1'
    GUID              = 'a3f9d2b1-4c7e-4f8a-9b2d-1e5f6c3a7d0e'
    Author            = 'Vibhu Bhatnagar'
    CompanyName       = 'Realtime-IT'
    Description       = 'Workstation reporting and Nextcloud upload utilities for Windows sysadmin environments.'
    PowerShellVersion = '5.1'

    # Root module
    RootModule        = 'VB.WorkstationReport.psm1'

    # Note: VB.NextCloud is a soft dependency (required at runtime by Invoke-VBWorkstationReport).
    # It is not listed in RequiredModules to avoid Test-ModuleManifest failures on systems
    # where VB.NextCloud is not yet installed. Install it separately: Install-Module VB.NextCloud

    # Exported functions -- loader controls exports dynamically, never list individually
    FunctionsToExport = '*'

    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    # Module metadata
    PrivateData = @{
        PSData = @{
            Tags         = @('PowerShell', 'Sysadmin', 'Windows', 'AD', 'MSP', 'Automation', 'VBTools', 'Workstation', 'Reporting', 'Nextcloud', 'Printers', 'FolderRedirection', 'SyncCenter')
            ProjectUri   = 'https://github.com/Vibhu2/ITAdmin_Tools/tree/main/VB.WorkstationReport'
            LicenseUri   = 'https://github.com/Vibhu2/ITAdmin_Tools/blob/main/LICENSE'
            HelpInfoUri  = 'https://pwsh.in/posts/printer-mapping-powershell-vb-workstationreport/'
            LinkedinUri  = 'https://www.linkedin.com/in/vibhu-bhatnagar-02622798/'
            ReleaseNotes = @'
v1.11.1 -- 04-06-2026 -- Fixed: Invoke-VBWorkstationReport report generation logic for user-applied GPOs; corrected function call issue causing script self-invocation; fixed spelling error in function name reference.
v1.11.0 -- 03-06-2026 -- New: Get-VBLoggedOnUser -- returns interactive logon sessions with quser enrichment (user, domain, SID, session name/ID, state, idle time, logon time).
                         New: Get-VBSystemGpResult -- parses gpresult /r /scope computer into structured PSCustomObjects (SystemInfo, AppliedGPOs, SecurityGroups).
                         New: Get-VBLoggedOnUserGpResult -- runs user-scoped gpresult for every logged-on interactive user; returns per-user policy data.
                         Updated: Invoke-VBWorkstationReport now includes all three new reports in its orchestrated output (7 new CSVs: LoggedOnUsers, sysGpResult x3, UserGpResult x3).
v1.10.4 -- 12-05-2026 -- Structural fix: Add-VBUserPrinter moved to Printer_Management subfolder; Mount-VBUserHive and Dismount-VBUserHive moved to Private (no longer exported).
v1.10.3 -- 04-05-2026 -- Set-VBUserPrinterMigration: Existing printer port corrected via Set-Printer; fixes UNC path in printer properties.
v1.10.2 -- 04-05-2026 -- Set-VBUserPrinterMigration: Username filter now matches short name, DOMAIN\user, and profile folder name.
v1.10.1 -- 04-05-2026 -- Set-VBUserPrinterMigration: Add-Printer "already exists" treated as skip; machine-level errors no longer abort per-user loop.
v1.10.0 -- 04-05-2026 -- Set-VBUserPrinterMigration: -TargetUser replaced with -Username[] and -SID[] arrays for multi-user targeting.
v1.9.0 -- 23-04-2026 -- New: Invoke-VBasCurrentUser -- run a scriptblock as the logged-on user from SYSTEM context. Fully self-contained -- C# type compiled inline, no external module required.
v1.8.0 -- 23-04-2026 -- New: Add-VBUserPrinter -- adds net-new UNC or IP printer to all or targeted user profiles.
v1.7.2 -- 23-04-2026 -- Set-VBUserPrinterMigration: added CSV format sample and step-by-step creation example.
v1.7.1 -- 23-04-2026 -- Finalized: full help blocks and examples on all printer migration functions.
v1.7.0 -- 23-04-2026 -- New: Set-VBUserPrinterMigration (public), Update-VBUserPrinterRegistry (private) -- printer mapping migration UNC <-> IP.
v1.6.0 -- 23-04-2026 -- New: Dismount-VBUserHive -- safe hive unload, pipeline-compatible with Mount-VBUserHive.
v1.5.0 -- 16-04-2026 -- Invoke-VBWorkstationReport: all parameters mandatory, defaults removed.
                         OutputPath validated upfront -- warns and stops if path does not exist.
                         OutputPath and CSV count validated before upload attempt.
v1.4.0 -- 16-04-2026 -- Invoke-VBWorkstationReport updated: added ODFB, UP, USF reports.
                         Updated defaults: OutputPath -> C:\Realtime\Reports, NextcloudDestination -> Realtime-IT/Reports.
v1.3.1 -- 15-04-2026 -- Moved Get-VBNextcloudFiles to VB.NextCloud module where it belongs.
v1.3.0 -- 15-04-2026 -- Module renamed to VB.WorkstationReport. VB.NextCloud declared as soft runtime dependency. Author standardised to Vibhu Bhatnagar.
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

v1.1.0 -- Initial modular release
- fixed and tested all functions locally and remotely
- added full comment-based help to all functions
- added version notes to all functions and module manifest
- added error handling and status reporting to all functions
- standardized output objects across all functions (ComputerName, Status, CollectionTime, Error)
'@
        }
    }
}
