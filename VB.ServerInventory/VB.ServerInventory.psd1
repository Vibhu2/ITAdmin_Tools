@{
    # -- Identity --
    ModuleVersion     = '1.0.1'
    GUID              = '88DE9814-273C-4FC8-852C-7B6D76D79338'
    Author            = 'Vibhu Bhatnagar'
    CompanyName       = 'Realtime-IT'
    Description       = 'VB Server Inventory Module -- collects system, AD, GPO, security, apps, and service data from Windows servers'
    Copyright         = '(c) 2026 Vibhu. All rights reserved.'

    # -- Requirements --
    PowerShellVersion = '5.1'

    # -- Root module --
    RootModule        = 'VB.ServerInventory.psm1'

    # -- Exports -- loader controls what gets defined, never list individual functions --
    FunctionsToExport = '*'
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    # -- Metadata --
    PrivateData = @{
        PSData = @{
            Tags         = @('Windows', 'Server', 'Inventory', 'AD', 'GPO', 'Security', 'Sysadmin', 'Realtime', 'VBTools')
            ProjectUri   = 'https://github.com/Vibhu2/ITAdmin_Tools'
            ReleaseNotes = 'v1.0.1 -- 10-04-2026 -- Initial release. 34 functions covering System, Disk, Network, AD, GPO, Security, Apps, Printing, Services, and Scheduled Tasks.'
        }
    }
}
