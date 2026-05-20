# ============================================================
# FUNCTION : Repair-VBUserShellFolderRedirects
# MODULE   : VB.WorkstationReport
# VERSION  : 2.0.1
# CHANGED  : 20-05-2026 -- v2.0.1: Null guard added after Get-ItemProperty on Shell Folders
#                          and User Shell Folders -- prevents null-reference throw when
#                          a key exists but the read returns nothing (ACL/locked hive)
#                          v2.0.0: Full rewrite: canonical folder mapping tables, broader fix
#                          scope, -Username filter, -Force mode, per-entry error objects,
#                          ShouldProcess on registry write and folder creation,
#                          O(1) reverse lookup, unknown key warnings
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Repairs broken Shell Folder registry paths for all loaded user hives
# ENCODING : UTF-8 with BOM
# ============================================================
function Repair-VBUserShellFolderRedirects {
    <#
    .SYNOPSIS
        Repairs broken Shell Folder registry paths for all loaded user hives.

    .DESCRIPTION
        Repair-VBUserShellFolderRedirects enumerates all standard user hives
        (S-1-5-21-*) currently loaded in HKEY_USERS and checks both Shell Folders
        and User Shell Folders registry keys for broken path values.

        Without -Force, fixes only paths that are unambiguously broken:
          -- UNC paths (\\server\...)
          -- Paths under the Windows directory
          -- Paths under ProgramData
          -- Absolute paths outside the user's own profile path

        With -Force, resets every key to its canonical Windows default, regardless
        of the current value. Use this when a profile has been partially corrupted
        or intentionally redirected and needs a full reset to local defaults.

        Shell Folders values are written as expanded local paths:
          e.g.  C:\Users\jsmith\Documents

        User Shell Folders values are written as unexpanded env-var paths:
          e.g.  %USERPROFILE%\Documents

        Registry keys not present in the known folder mapping table are skipped
        with a warning rather than receiving a guessed path.

        Supports -WhatIf to preview all changes without touching the registry.
        Call Get-VBUserProfile | Mount-VBUserHive first to load offline profiles.

    .PARAMETER Username
        One or more usernames to target. Accepts plain username or DOMAIN\username.
        If omitted, all loaded standard user hives are processed.

    .PARAMETER Force
        Resets all Shell Folder keys to canonical Windows defaults, not just
        those that are clearly broken. Does not affect -WhatIf behaviour.

    .EXAMPLE
        Repair-VBUserShellFolderRedirects -WhatIf
        Previews all broken paths that would be fixed without making any changes.

    .EXAMPLE
        Repair-VBUserShellFolderRedirects | Where-Object Changed
        Runs the repair and returns only entries where a change was made.

    .EXAMPLE
        Repair-VBUserShellFolderRedirects -Username 'jsmith' -Force -WhatIf
        Previews a full reset to Windows defaults for user jsmith only.

    .EXAMPLE
        Get-VBUserProfile | Mount-VBUserHive
        Repair-VBUserShellFolderRedirects | Export-Csv C:\Realtime\repairs.csv -NoTypeInformation -Encoding UTF8
        Loads all offline profiles first, runs the repair, and exports the change log.

    .OUTPUTS
        PSCustomObject
        Returns one object per matched registry value with:
          - ComputerName   : Local computer name
          - Status         : 'Success' or 'Failed'
          - SID            : User SID
          - Username       : Resolved username
          - RegistryKey    : 'Shell Folders' or 'User Shell Folders'
          - PropertyName   : Registry value name (e.g. 'Personal', 'Desktop')
          - FolderName     : Human-readable folder name (e.g. 'Documents', 'Desktop')
          - OldValue       : Value before repair
          - NewValue       : Value after repair (or intended value on -WhatIf)
          - FolderCreated  : True if the local folder did not exist and was created
          - Changed        : True if the registry value was written
          - Error          : Error message ($null on success)
          - CollectionTime : Timestamp of data collection

    .NOTES
        Version      : 2.0.1
        Author       : Vibhu Bhatnagar
        Category     : Workstation Fixes
        Requirements :
          - PowerShell 5.1 or higher
          - Must run as Administrator (registry write access to HKEY_USERS)
          - Target user hives must be loaded in HKEY_USERS
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([PSCustomObject])]
    param(
        [Parameter()]
        [string[]]$Username,

        [Parameter()]
        [switch]$Force
    )

    begin {
        # Maps human-readable folder name -> registry key name used in Shell Folders
        $folderMappings = @{
            'Desktop'        = 'Desktop'
            'Documents'      = 'Personal'
            'Pictures'       = 'My Pictures'
            'Music'          = 'My Music'
            'Videos'         = 'My Video'
            'Favorites'      = 'Favorites'
            'Downloads'      = '{374DE290-123F-4565-9164-39C4925E467B}'
            'Start Menu'     = 'Start Menu'
            'Programs'       = 'Programs'
            'Startup'        = 'Startup'
            'Recent'         = 'Recent'
            'SendTo'         = 'SendTo'
            'Templates'      = 'Templates'
            'NetHood'        = 'NetHood'
            'PrintHood'      = 'PrintHood'
            'History'        = 'History'
            'Cookies'        = 'Cookies'
            'Cache'          = 'Cache'
            'AppData'        = 'AppData'
            'Local AppData'  = 'Local AppData'
            'Contacts'       = '{56784854-C6CB-462B-8169-88E350ACB882}'
            'Links'          = '{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}'
            'Saved Games'    = '{4C5C32FF-BB9D-43B0-B5B4-2D72E54EAAA4}'
            'Searches'       = '{7D1D3A04-DEBB-4115-95CF-2F29DA2920DA}'
        }

        # Maps human-readable folder name -> relative path under the user profile
        $defaultPaths = @{
            'Desktop'        = 'Desktop'
            'Documents'      = 'Documents'
            'Pictures'       = 'Pictures'
            'Music'          = 'Music'
            'Videos'         = 'Videos'
            'Favorites'      = 'Favorites'
            'Downloads'      = 'Downloads'
            'Start Menu'     = 'AppData\Roaming\Microsoft\Windows\Start Menu'
            'Programs'       = 'AppData\Roaming\Microsoft\Windows\Start Menu\Programs'
            'Startup'        = 'AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup'
            'Recent'         = 'AppData\Roaming\Microsoft\Windows\Recent'
            'SendTo'         = 'AppData\Roaming\Microsoft\Windows\SendTo'
            'Templates'      = 'AppData\Roaming\Microsoft\Windows\Templates'
            'NetHood'        = 'AppData\Roaming\Microsoft\Windows\Network Shortcuts'
            'PrintHood'      = 'AppData\Roaming\Microsoft\Windows\Printer Shortcuts'
            'History'        = 'AppData\Local\Microsoft\Windows\History'
            'Cookies'        = 'AppData\Local\Microsoft\Windows\INetCookies'
            'Cache'          = 'AppData\Local\Microsoft\Windows\INetCache'
            'AppData'        = 'AppData\Roaming'
            'Local AppData'  = 'AppData\Local'
            'Contacts'       = 'Contacts'
            'Links'          = 'Links'
            'Saved Games'    = 'Saved Games'
            'Searches'       = 'Searches'
        }

        # Build reverse lookup O(1): registry key name -> human-readable folder name
        $keyToFolder = @{}
        foreach ($entry in $folderMappings.GetEnumerator()) {
            $keyToFolder[$entry.Value] = $entry.Key
        }

        # PS* properties present on every Get-ItemProperty result -- always excluded
        $psProps = @('PSPath', 'PSParentPath', 'PSChildName', 'PSDrive', 'PSProvider')

        # Normalise Username filter: strip domain prefix so both 'jsmith' and 'DOMAIN\jsmith' match
        $normalizedFilter = @()
        if ($Username) {
            $normalizedFilter = $Username | ForEach-Object {
                if ($_ -match '\\') { ($_ -split '\\')[1] } else { $_ }
            }
        }
    }

    process {
        $collectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
        $found          = $false

        try {
            $userHives = Get-ChildItem -Path 'Registry::HKEY_USERS' |
                         Where-Object { $_.Name -match 'S-1-5-21-\d+-\d+-\d+-\d+$' }

            foreach ($hive in $userHives) {
                $sid = $hive.PSChildName

                # Step 1 -- Resolve profile path from ProfileList
                $profilePath = $null
                try {
                    $profilePath = (Get-ItemProperty `
                        -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid" `
                        -ErrorAction Stop).ProfileImagePath
                }
                catch {
                    Write-Warning "[$sid] Could not read ProfileList entry -- skipping"
                    continue
                }

                if (-not $profilePath) {
                    Write-Warning "[$sid] ProfileImagePath is empty -- skipping"
                    continue
                }

                if (-not (Test-Path $profilePath)) {
                    Write-Warning "[$sid] Profile path '$profilePath' does not exist on disk -- skipping"
                    continue
                }

                # Step 2 -- Resolve username (SID translation, fall back to profile folder leaf)
                $resolvedUsername = $null
                try {
                    $ntAccount       = (New-Object System.Security.Principal.SecurityIdentifier($sid)).Translate([System.Security.Principal.NTAccount]).Value
                    $resolvedUsername = if ($ntAccount -match '\\') { ($ntAccount -split '\\')[1] } else { $ntAccount }
                }
                catch {
                    $resolvedUsername = Split-Path -Path $profilePath -Leaf
                }

                # Step 3 -- Apply Username filter
                if ($normalizedFilter.Count -gt 0 -and $resolvedUsername -notin $normalizedFilter) {
                    continue
                }

                $shellFoldersPath     = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
                $userShellFoldersPath = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"

                # Step 4 -- Process Shell Folders (expanded local paths)
                Write-Verbose "[$resolvedUsername] Checking Shell Folders"

                if (Test-Path $shellFoldersPath) {
                    $shellFolders = Get-ItemProperty -Path $shellFoldersPath -ErrorAction SilentlyContinue

                    if ($shellFolders) {
                    foreach ($property in ($shellFolders.PSObject.Properties | Where-Object { $_.Name -notin $psProps })) {
                        $currentValue = $property.Value

                        # Reverse lookup -- skip unknown keys with a warning
                        $folderDisplayName = $keyToFolder[$property.Name]
                        if (-not $folderDisplayName) {
                            Write-Warning "[$resolvedUsername] Shell Folders: Unknown key '$($property.Name)' -- skipping"
                            continue
                        }

                        $expectedPath = Join-Path $profilePath $defaultPaths[$folderDisplayName]

                        # Determine whether this entry needs fixing
                        $needsFix = if ($Force) {
                            $currentValue -ne $expectedPath
                        }
                        else {
                            $currentValue -like '\\*' -or
                            $currentValue -like "$env:SystemRoot\*" -or
                            $currentValue -like "$env:ProgramData\*" -or
                            ($currentValue -match '^[A-Za-z]:\\' -and $currentValue -notlike "$profilePath\*")
                        }

                        if (-not $needsFix) { continue }

                        $found = $true

                        try {
                            $folderCreated = $false
                            $changesMade   = $false

                            if ($PSCmdlet.ShouldProcess(
                                    "$shellFoldersPath [$($property.Name)]",
                                    "Replace '$currentValue' with '$expectedPath'")) {

                                if (-not (Test-Path $expectedPath)) {
                                    New-Item -Path $expectedPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
                                    $folderCreated = $true
                                    Write-Verbose "[$resolvedUsername] Created folder: $expectedPath"
                                }

                                Set-ItemProperty -Path $shellFoldersPath -Name $property.Name -Value $expectedPath -ErrorAction Stop
                                $changesMade = $true
                                Write-Verbose "[$resolvedUsername] Shell Folders: $($property.Name) -> $expectedPath"
                            }

                            [PSCustomObject]@{
                                ComputerName   = $env:COMPUTERNAME
                                Status         = 'Success'
                                SID            = $sid
                                Username       = $resolvedUsername
                                RegistryKey    = 'Shell Folders'
                                PropertyName   = $property.Name
                                FolderName     = $folderDisplayName
                                OldValue       = $currentValue
                                NewValue       = $expectedPath
                                FolderCreated  = $folderCreated
                                Changed        = $changesMade
                                Error          = $null
                                CollectionTime = $collectionTime
                            }
                        }
                        catch {
                            [PSCustomObject]@{
                                ComputerName   = $env:COMPUTERNAME
                                Status         = 'Failed'
                                SID            = $sid
                                Username       = $resolvedUsername
                                RegistryKey    = 'Shell Folders'
                                PropertyName   = $property.Name
                                FolderName     = $folderDisplayName
                                OldValue       = $currentValue
                                NewValue       = $expectedPath
                                FolderCreated  = $false
                                Changed        = $false
                                Error          = $_.Exception.Message
                                CollectionTime = $collectionTime
                            }
                        }
                    }
                    } # end if ($shellFolders)
                }

                # Step 5 -- Process User Shell Folders (unexpanded env-var paths)
                Write-Verbose "[$resolvedUsername] Checking User Shell Folders"

                if (Test-Path $userShellFoldersPath) {
                    $userShellFolders = Get-ItemProperty -Path $userShellFoldersPath -ErrorAction SilentlyContinue

                    if ($userShellFolders) {
                    foreach ($property in ($userShellFolders.PSObject.Properties | Where-Object { $_.Name -notin $psProps })) {
                        $currentValue = $property.Value

                        # Reverse lookup -- skip unknown keys with a warning
                        $folderDisplayName = $keyToFolder[$property.Name]
                        if (-not $folderDisplayName) {
                            Write-Warning "[$resolvedUsername] User Shell Folders: Unknown key '$($property.Name)' -- skipping"
                            continue
                        }

                        $expectedPath    = "%USERPROFILE%\$($defaultPaths[$folderDisplayName])"
                        $expandedForDisk = Join-Path $profilePath $defaultPaths[$folderDisplayName]

                        # Determine whether this entry needs fixing
                        $needsFix = if ($Force) {
                            $currentValue -ne $expectedPath
                        }
                        else {
                            $currentValue -like '\\*' -or
                            $currentValue -like "$env:SystemRoot\*" -or
                            $currentValue -like "$env:ProgramData\*" -or
                            ($currentValue -match '^[A-Za-z]:\\' -and $currentValue -notlike "$profilePath\*")
                        }

                        if (-not $needsFix) { continue }

                        $found = $true

                        try {
                            $folderCreated = $false
                            $changesMade   = $false

                            if ($PSCmdlet.ShouldProcess(
                                    "$userShellFoldersPath [$($property.Name)]",
                                    "Replace '$currentValue' with '$expectedPath'")) {

                                if (-not (Test-Path $expandedForDisk)) {
                                    New-Item -Path $expandedForDisk -ItemType Directory -Force -ErrorAction Stop | Out-Null
                                    $folderCreated = $true
                                    Write-Verbose "[$resolvedUsername] Created folder: $expandedForDisk"
                                }

                                Set-ItemProperty -Path $userShellFoldersPath -Name $property.Name -Value $expectedPath -ErrorAction Stop
                                $changesMade = $true
                                Write-Verbose "[$resolvedUsername] User Shell Folders: $($property.Name) -> $expectedPath"
                            }

                            [PSCustomObject]@{
                                ComputerName   = $env:COMPUTERNAME
                                Status         = 'Success'
                                SID            = $sid
                                Username       = $resolvedUsername
                                RegistryKey    = 'User Shell Folders'
                                PropertyName   = $property.Name
                                FolderName     = $folderDisplayName
                                OldValue       = $currentValue
                                NewValue       = $expectedPath
                                FolderCreated  = $folderCreated
                                Changed        = $changesMade
                                Error          = $null
                                CollectionTime = $collectionTime
                            }
                        }
                        catch {
                            [PSCustomObject]@{
                                ComputerName   = $env:COMPUTERNAME
                                Status         = 'Failed'
                                SID            = $sid
                                Username       = $resolvedUsername
                                RegistryKey    = 'User Shell Folders'
                                PropertyName   = $property.Name
                                FolderName     = $folderDisplayName
                                OldValue       = $currentValue
                                NewValue       = $expectedPath
                                FolderCreated  = $false
                                Changed        = $false
                                Error          = $_.Exception.Message
                                CollectionTime = $collectionTime
                            }
                        }
                    }
                    } # end if ($userShellFolders)
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
                    FolderName     = 'N/A'
                    OldValue       = 'N/A'
                    NewValue       = 'N/A'
                    FolderCreated  = $false
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
                FolderName     = 'N/A'
                OldValue       = 'N/A'
                NewValue       = 'N/A'
                FolderCreated  = $false
                Changed        = $false
                Error          = $_.Exception.Message
                CollectionTime = $collectionTime
            }
        }
    }
}
