@{
    # -- Identity --
    ModuleVersion     = '1.0.4'
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
            Tags         = @('PowerShell', 'Sysadmin', 'Windows', 'AD', 'MSP', 'Automation', 'VBTools', 'Server', 'Inventory', 'GPO', 'Security')
            ProjectUri   = 'https://github.com/Vibhu2/ITAdmin_Tools/tree/main/VB.ServerInventory'
            LicenseUri   = 'https://github.com/Vibhu2/ITAdmin_Tools/blob/main/LICENSE'
            HelpInfoUri  = 'https://pwsh.in/posts/'
            ReleaseNotes = 'v1.0.4 -- 03-06-2026 -- Removed VB.ServerInventory.xlsx and Server_invetory_Fix scratch folder from module root; fixes PS5.1 install failure (PSGet 1.0.0.1 rejected nupkg with non-module files).
v1.0.3 -- 24-05-2026 -- Fixed DiskInformation call: Get-VBDiskInformation -> Get-VBDiskInventory.
v1.0.2 -- 15-04-2026 -- Author standardised to Vibhu Bhatnagar across all 34 functions.
v1.0.1 -- 10-04-2026 -- Initial release. 34 functions covering System, Disk, Network, AD, GPO, Security, Apps, Printing, Services, and Scheduled Tasks.'
        }
    }
}
