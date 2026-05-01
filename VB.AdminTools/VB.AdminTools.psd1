#
# VB.AdminTools.psd1 -- Module manifest
# Author  : Vibhu Bhatnagar
#

@{
    # -- Identity --
    ModuleVersion     = '1.0.1'
    GUID              = 'f7a2c3e1-9d4b-4a8f-b6e0-2c5d7f1a3e9b'
    Author            = 'Vibhu Bhatnagar'
    CompanyName       = 'Realtime-IT'
    Description       = 'VB.AdminTools -- Miscellaneous sysadmin utilities for day-to-day Windows administration. Covers reboot detection, environment checks, and general-purpose tooling not specific to a dedicated module.'
    Copyright         = '(c) 2026 Vibhu Bhatnagar. All rights reserved.'

    # -- Requirements --
    PowerShellVersion = '5.1'

    # -- Root module --
    RootModule        = 'VB.AdminTools.psm1'

    # -- Exports -- loader controls what gets defined, never list individual functions --
    FunctionsToExport = '*'
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    # -- Metadata --
    PrivateData = @{
        PSData = @{
            Tags         = @('PowerShell', 'Sysadmin', 'Windows', 'AD', 'MSP', 'Automation', 'VBTools', 'AdminTools', 'Utilities')
            ProjectUri   = 'https://github.com/Vibhu2/ITAdmin_Tools/tree/main/VB.AdminTools'
            LicenseUri   = 'https://github.com/Vibhu2/ITAdmin_Tools/blob/main/LICENSE'
            HelpInfoUri  = 'https://pwsh.in/posts/'
            ReleaseNotes = ReleaseNotes = 'v1.0.1 -- 02-05-2026 -- Workflow trigger bump. No functional changes.
            v1.0.0 -- 30-04-2026 -- Initial release. Get-VBPendingReboot: checks CBS, Windows Update, pending file rename operations, Server Feature Manager, computer rename, and domain join reboot flags. Supports local and remote execution, optional SCCM detection via -IncludeSCCM.'
        }
    }
}
