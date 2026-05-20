# ============================================================
# FUNCTION : Repair-VBUserShellFolderRedirects
# MODULE   : VB.WorkstationReport
# VERSION  : 1.0.0
# CHANGED  : 20-05-2026 -- Initial release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Replaces UNC Shell Folder registry paths with local paths for all loaded user hives
# ENCODING : UTF-8 with BOM
# ============================================================
function Repair-VBUserShellFolderRedirects {
    <#
    .SYNOPSIS
        Replaces UNC Shell Folder registry paths with local paths for all loaded user hives.

    .DESCRIPTION
        Repair-VBUserShellFolder enumerates all standard user hives (S-1-5-21-*) currently
        loaded in HKEY_USERS and checks both Shell Folders and User Shell Folders registry
        keys for values pointing to UNC paths (\\server\Users\...). When found:

          Shell Folders      -- replaced with an expanded local path (C:\Users\username\folder)
          User Shell Folders -- replaced with an unexpanded env-var path (%USERPROFILE%\folder)

        Supports -WhatIf to preview changes without modifying the registry.
        Call Get-VBUserProfile | Mount-VBUserHive first to ensure offline profiles are loaded.

    .PARAMETER None
        No parameters. Operates on the local machine only.

    .EXAMPLE
        Repair-VBUserShellFolder -WhatIf
        Previews all UNC paths that would be replaced without making any changes.

    .EXAMPLE
        Repair-VBUserShellFolder | Where-Object Changed
        Runs the repair and returns only entries where a change was made.

    .EXAMPLE
        Repair-VBUserShellFolder | Export-Csv C:\Realtime\repairs.csv -NoTypeInformation -Encoding UTF8
        Runs the repair and exports the full change log to CSV.

    .OUTPUTS
        PSCustomObject
        Returns one object per matched registry value with:
          - ComputerName   : Local computer name
          - Status         : 'Success' or 'Failed'
          - SID            : User SID
          - Username       : Username extracted from the UNC path
          - RegistryKey    : 'Shell Folders' or 'User Shell Folders'
          - PropertyName   : Registry value name (e.g. 'Desktop', 'Personal')
          - OldValue       : Original UNC path
          - NewValue       : Replacement local path
          - Changed        : True if the value was written; False on -WhatIf or if unchanged
          - Error          : Error message ($null on success)
          - CollectionTime : Timestamp of data collection

    .NOTES
        Version      : 1.0.0
        Author       : Vibhu Bhatnagar
        Category     : Workstation Fixes
        Requirements :
          - PowerShell 5.1 or higher
          - Must run as Administrator (registry write access to HKEY_USERS)
          - Target user hives must be loaded in HKEY_USERS
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([PSCustomObject])]
    param()

    process {
        $collectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
        $found          = $false

        try {
            $userHives = Get-ChildItem -Path 'Registry::HKEY_USERS' |
                         Where-Object { $_.Name -match 'S-1-5-21-\d+-\d+-\d+-\d+$' }

            foreach ($hive in $userHives) {
                $sid                  = $hive.PSChildName
                $shellFoldersPath     = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
                $userShellFoldersPath = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"

                Write-Verbose "[$sid] Checking Shell Folders"

                if (Test-Path $shellFoldersPath) {
                    $shellFolders = Get-ItemProperty -Path $shellFoldersPath -ErrorAction SilentlyContinue

                    foreach ($property in $shellFolders.PSObject.Properties) {
                        if ($property.Value -like '\\*\Users\*') {
                            $found      = $true
                            $uncPath    = $property.Value
                            $username   = $uncPath -replace '^.*\\Users\\([^\\]+)\\.*$', '$1'
                            $folderName = Split-Path -Path $uncPath -Leaf
                            $newPath    = Join-Path "$env:SystemDrive\Users\$username" $folderName

                            $changesMade = $false
                            if ($PSCmdlet.ShouldProcess(
                                    "$shellFoldersPath [$($property.Name)]",
                                    "Replace '$uncPath' with '$newPath'")) {
                                Set-ItemProperty -Path $shellFoldersPath -Name $property.Name -Value $newPath -ErrorAction Stop
                                $changesMade = $true
                                Write-Verbose "[$sid] Shell Folders: $($property.Name) updated"
                            }

                            [PSCustomObject]@{
                                ComputerName   = $env:COMPUTERNAME
                                Status         = 'Success'
                                SID            = $sid
                                Username       = $username
                                RegistryKey    = 'Shell Folders'
                                PropertyName   = $property.Name
                                OldValue       = $uncPath
                                NewValue       = $newPath
                                Changed        = $changesMade
                                Error          = $null
                                CollectionTime = $collectionTime
                            }
                        }
                    }
                }

                Write-Verbose "[$sid] Checking User Shell Folders"

                if (Test-Path $userShellFoldersPath) {
                    $userShellFolders = Get-ItemProperty -Path $userShellFoldersPath -ErrorAction SilentlyContinue

                    foreach ($property in $userShellFolders.PSObject.Properties) {
                        if ($property.Value -like '\\*\Users\*') {
                            $found      = $true
                            $uncPath    = $property.Value
                            $folderName = Split-Path -Path $uncPath -Leaf
                            $newPath    = "%USERPROFILE%\$folderName"
                            $username   = $uncPath -replace '^.*\\Users\\([^\\]+)\\.*$', '$1'

                            $changesMade = $false
                            if ($PSCmdlet.ShouldProcess(
                                    "$userShellFoldersPath [$($property.Name)]",
                                    "Replace '$uncPath' with '$newPath'")) {
                                Set-ItemProperty -Path $userShellFoldersPath -Name $property.Name -Value $newPath -ErrorAction Stop
                                $changesMade = $true
                                Write-Verbose "[$sid] User Shell Folders: $($property.Name) updated"
                            }

                            [PSCustomObject]@{
                                ComputerName   = $env:COMPUTERNAME
                                Status         = 'Success'
                                SID            = $sid
                                Username       = $username
                                RegistryKey    = 'User Shell Folders'
                                PropertyName   = $property.Name
                                OldValue       = $uncPath
                                NewValue       = $newPath
                                Changed        = $changesMade
                                Error          = $null
                                CollectionTime = $collectionTime
                            }
                        }
                    }
                }
            }

            if (-not $found) {
                [PSCustomObject]@{
                    ComputerName   = $env:COMPUTERNAME
                    Status         = 'Success'
                    SID            = 'N/A'
                    Username       = 'N/A'
                    RegistryKey    = 'N/A'
                    PropertyName   = 'N/A'
                    OldValue       = 'N/A'
                    NewValue       = 'N/A'
                    Changed        = $false
                    Error          = $null
                    CollectionTime = $collectionTime
                }
            }
        }
        catch {
            [PSCustomObject]@{
                ComputerName   = $env:COMPUTERNAME
                Status         = 'Failed'
                SID            = 'N/A'
                Username       = 'N/A'
                RegistryKey    = 'N/A'
                PropertyName   = 'N/A'
                OldValue       = 'N/A'
                NewValue       = 'N/A'
                Changed        = $false
                Error          = $_.Exception.Message
                CollectionTime = $collectionTime
            }
        }
    }
}
