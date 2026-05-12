# ============================================================
# FUNCTION : Get-VBUserShellFolders
# MODULE   : VB.WorkstationReport
# VERSION  : 1.0.1
# CHANGED  : 16-04-2026 -- Initial release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Retrieves Shell Folders and User Shell Folders registry values for user profiles
# ENCODING : UTF-8 with BOM
# ============================================================
function Get-VBUserShellFolders {
    <#
    .SYNOPSIS
        Retrieves Shell Folders and User Shell Folders registry values for user profiles.

    .DESCRIPTION
        Get-VBUserShellFolders enumerates loaded user profiles and queries the registry
        for Shell Folders and User Shell Folders values. These registry keys define
        the paths to standard Windows folders (Desktop, Documents, Downloads, Favorites,
        Music, Pictures, Videos, etc).

        Shell Folders contains expanded paths (all environment variables resolved).
        User Shell Folders contains unexpanded paths (may contain %USERPROFILE%, etc).

        Can filter by UserName or SID. If neither is provided, returns data for all
        currently loaded profiles. Only scans profiles whose hives are loaded in
        HKEY_USERS (i.e., logged-in users). Use Mount-VBUserHive to scan offline
        profiles.

    .PARAMETER ComputerName
        Computer names to query. Accepts pipeline input. Defaults to the local computer.

    .PARAMETER UserName
        Username(s) to query. Optional filter. Accepts pipeline input.
        If not provided, all loaded profiles are returned.

    .PARAMETER SID
        SID(s) to query. Optional filter.
        If not provided, all loaded profiles are returned.

    .PARAMETER Credential
        Credentials for remote computer access. Not required for local execution.

    .EXAMPLE
        Get-VBUserShellFolders
        Returns all shell folder registry values for all loaded user profiles on the
        local computer.

    .EXAMPLE
        Get-VBUserShellFolders -UserName 'vibhu.bhatnagar' -ComputerName 'WS001'
        Returns shell folder values for a specific user on a remote workstation.

    .EXAMPLE
        Get-VBUserShellFolders -SID 'S-1-5-21-3652201067-2579442110-4045031335-1001' |
            Select-Object UserName, ValueName, ValueData
        Returns Desktop, Documents, Downloads paths and others for a specific user SID.

    .EXAMPLE
        Get-VBUserShellFolders -UserName 'admin' | Where-Object ValueName -eq 'Desktop'
        Returns the Desktop folder path for a user named 'admin'.

    .OUTPUTS
        PSCustomObject
        Returns one object per registry value with:
          - ComputerName  : Target computer
          - UserSID       : User SID
          - UserName      : Username (profile folder name)
          - FolderType    : 'Shell Folders' (expanded) or 'User Shell Folders' (unexpanded)
          - ValueName     : Registry value name (e.g. 'Desktop', 'Documents', 'Favorites')
          - ValueData     : Registry value data (folder path)
          - CollectionTime: Timestamp of data collection
          - Status        : 'Success' or 'Failed'
          - Error         : Error message (only present on failure)

    .NOTES
        Version      : 1.0.1
        Author       : Vibhu Bhatnagar
        Category     : User Profile Management
        Requirements :
          - PowerShell 5.1 or higher
          - User profile hives must be loaded in HKEY_USERS (users must be logged in)
          - PowerShell Remoting enabled for remote targets

        Known Limitations :
          - Only scans currently loaded user profiles (logged-in users)
          - To scan unloaded profiles, use Mount-VBUserHive first
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('Name', 'Server', 'Host')]
        [string[]]$ComputerName = $env:COMPUTERNAME,

        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('User', 'Identity', 'SamAccountName')]
        [string[]]$UserName,

        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('SecurityIdentifier')]
        [string[]]$SID,

        [PSCredential]$Credential
    )

    begin {
        # Shared scriptblock for both local and remote execution.
        # Avoids code duplication and ensures consistent behavior.
        $scriptBlock = {
            param($FilterUserName, $FilterSID)

            $results        = [System.Collections.Generic.List[object]]::new()
            $collectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')

            try {
                # Get all currently loaded user profiles
                $loadedProfiles = Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop |
                                  Where-Object { $_.Loaded -eq $true }

                # Filter by UserName and/or SID if provided
                if ($FilterUserName -or $FilterSID) {
                    $loadedProfiles = $loadedProfiles | Where-Object {
                        $profileUserName = Split-Path $_.LocalPath -Leaf
                        $profileSID      = $_.SID

                        ($FilterUserName -and $profileUserName -in $FilterUserName) -or
                        ($FilterSID -and $profileSID -in $FilterSID)
                    }
                }

                # Scan each profile's Shell Folders registry values
                foreach ($profile in $loadedProfiles) {
                    $sid            = $profile.SID
                    $userName       = Split-Path $profile.LocalPath -Leaf
                    $shellPath      = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
                    $userShellPath  = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"

                    # Get Shell Folders (expanded paths -- env vars resolved)
                    if (Test-Path $shellPath -ErrorAction SilentlyContinue) {
                        $shellFolders = Get-ItemProperty -Path $shellPath -ErrorAction SilentlyContinue
                        if ($shellFolders) {
                            $shellFolders.PSObject.Properties |
                                Where-Object { $_.Name -notlike 'PS*' } |
                                ForEach-Object {
                                    $results.Add([PSCustomObject]@{
                                        UserSID       = $sid
                                        UserName      = $userName
                                        FolderType    = 'Shell Folders'
                                        ValueName     = $_.Name
                                        ValueData     = $_.Value
                                        CollectionTime = $collectionTime
                                        Status        = 'Success'
                                    })
                                }
                        }
                    }

                    # Get User Shell Folders (unexpanded paths -- may contain env variables)
                    if (Test-Path $userShellPath -ErrorAction SilentlyContinue) {
                        $userShellFolders = Get-ItemProperty -Path $userShellPath -ErrorAction SilentlyContinue
                        if ($userShellFolders) {
                            $userShellFolders.PSObject.Properties |
                                Where-Object { $_.Name -notlike 'PS*' } |
                                ForEach-Object {
                                    $results.Add([PSCustomObject]@{
                                        UserSID       = $sid
                                        UserName      = $userName
                                        FolderType    = 'User Shell Folders'
                                        ValueName     = $_.Name
                                        ValueData     = $_.Value
                                        CollectionTime = $collectionTime
                                        Status        = 'Success'
                                    })
                                }
                        }
                    }
                }

                return $results
            }
            catch {
                [PSCustomObject]@{
                    Error          = $_.Exception.Message
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    Status         = 'Failed'
                }
            }
        }
    }

    process {
        foreach ($computer in $ComputerName) {
            try {
                if ($computer -eq $env:COMPUTERNAME) {
                    # Local execution
                    $results = & $scriptBlock $UserName $SID
                    foreach ($result in $results) {
                        $result | Add-Member -NotePropertyName 'ComputerName' -NotePropertyValue $computer -Force
                        $result
                    }
                }
                else {
                    # Remote execution
                    $invokeParams = @{
                        ComputerName = $computer
                        ScriptBlock  = $scriptBlock
                        ArgumentList = $UserName, $SID
                        ErrorAction  = 'Stop'
                    }
                    if ($Credential) { $invokeParams['Credential'] = $Credential }

                    $results = Invoke-Command @invokeParams
                    foreach ($result in $results) {
                        $result | Add-Member -NotePropertyName 'ComputerName' -NotePropertyValue $computer -Force
                        $result
                    }
                }
            }
            catch {
                [PSCustomObject]@{
                    ComputerName   = $computer
                    Error          = $_.Exception.Message
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    Status         = 'Failed'
                }
            }
        }
    }
}