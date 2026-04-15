@{
    # -- Identity --
    ModuleVersion     = '1.0.0'
    GUID              = 'c2e4f1a3-8b7d-4e9f-a1c5-3d6e8f2b0a4c'
    Author            = 'Vibhu Bhatnagar'
    CompanyName       = 'Realtime-IT'
    Description       = 'VB.NextCloud -- WebDAV upload utilities for Nextcloud. Provides Set-VBNextcloudFile (single file PUT) and Start-VBNextcloudUpload (batch orchestrator).'
    Copyright         = '(c) 2026 Vibhu. All rights reserved.'

    # -- Requirements --
    PowerShellVersion = '5.1'

    # -- Root module --
    RootModule        = 'VB.NextCloud.psm1'

    # -- Exports -- loader controls what gets defined, never list individual functions --
    FunctionsToExport = '*'
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    # -- Metadata --
    PrivateData = @{
        PSData = @{
            Tags         = @('Nextcloud', 'WebDAV', 'Upload', 'CloudStorage', 'Sysadmin', 'Realtime', 'VBTools')
            ProjectUri   = 'https://github.com/Vibhu2/ITAdmin_Tools'
            ReleaseNotes = 'v1.0.0 -- 15-04-2026 -- Initial release. Set-VBNextcloudFile (single file WebDAV PUT) and Start-VBNextcloudUpload (batch wrapper with auto folder creation).'
        }
    }
}
