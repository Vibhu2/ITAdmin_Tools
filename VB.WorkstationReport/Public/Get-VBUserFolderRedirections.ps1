# ============================================================
# FUNCTION : Get-VBUserFolderRedirections
# MODULE   : WorkstationReport
# VERSION  : 1.3.0
# CHANGED  : 14-04-2026 -- Standards compliance fixes
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Audits folder redirections for all user profiles on local or remote computers
# ENCODING : UTF-8 with BOM
# ============================================================

function Get-VBUserFolderRedirections {
    <#
    .SYNOPSIS
    Audits folder redirections for all user profiles on local or remote computers.

    .DESCRIPTION
    Get-VBUserFolderRedirections scans Windows systems to identify folder redirections for
    all user profiles. It detects OneDrive, network-based, and manual redirections by
    examining Shell Folders and User Shell Folders registry keys. System accounts (SYSTEM,
    LOCAL SERVICE, NETWORK SERVICE) are excluded -- only domain/local user accounts are
    audited (SIDs matching S-1-5-21-*).

    Registry hives for inactive profiles are loaded using reg.exe and safely unloaded after
    inspection. The scriptblock pattern is used so the same logic runs locally or remotely
    without duplication.

    Helper functions (Test-IsRedirected, Get-RedirectionType, Get-ProfileLastUpdate) are
    defined inside the scriptblock so they are serialised correctly when sent via
    Invoke-Command to remote targets.

    .PARAMETER ComputerName
    Computer names to audit. Accepts pipeline input. Defaults to the local computer.

    .PARAMETER Credential
    Credentials for remote computer access. Not required for local execution.

    .PARAMETER TableOutput
    When specified, suppresses Write-Host console output and returns only structured objects.

    .EXAMPLE
    Get-VBUserFolderRedirections

    Audits folder redirections for all users on the local computer with console output.

    .EXAMPLE
    Get-VBUserFolderRedirections -ComputerName 'WS001','WS002' -Credential (Get-Credential)

    Audits folder redirections on two remote workstations.

    .EXAMPLE
    Get-VBUserFolderRedirections -TableOutput |
        Where-Object { $_.RedirectionCount -gt 0 } | Format-Table

    Returns only users with active redirections in table format.

    .EXAMPLE
    'WS001','WS002' | Get-VBUserFolderRedirections -TableOutput |
        Export-Csv 'FolderRedirections.csv' -NoTypeInformation

    Processes multiple computers via pipeline and exports to CSV.

    .EXAMPLE
    Get-VBUserFolderRedirections -TableOutput |
        Where-Object { $_.OneDriveRedirections -ne 'None' } |
        Select-Object ComputerName, Username, OneDriveRedirections

    Identifies all users with OneDrive folder redirections.

    .OUTPUTS
    PSCustomObject
    Returns objects with:
    - ComputerName         : Target computer
    - Username             : User profile name
    - RedirectedFolders    : Semicolon-separated folder=path summary for all redirected folders
    - RedirectionCount     : Number of redirected folders
    - RedirectionTypes     : Detected type(s) -- 'OneDrive', 'Network', 'Manual', or 'Mixed (...)'
    - Desktop, Documents, Pictures, Downloads, Music, Videos, Favorites, AppDataRoaming,
      SavedGames, Searches, StartMenu, Contacts, Links : Individual folder path [type] or 'Local'
    - OneDriveRedirections : Semicolon-separated list of OneDrive-redirected folders
    - NetworkRedirections  : Semicolon-separated list of network-redirected folders
    - ManualRedirections   : Semicolon-separated list of manually-redirected folders
    - LastProfileUpdate    : Last write time of the user profile directory
    - Status               : 'Success', 'Failed', 'Profile Missing', 'Registry Missing',
                             'SID Not Found', 'No Profiles', or 'Registry Not Accessible'
    - Error                : Error message (only present on failure)

    .NOTES
    Version : 1.3.0
    Author  : Vibhu Bhatnagar
    Category: User Profile Management

    Requirements:
    - PowerShell 5.1 or later
    - Administrative privileges (required to load user registry hives)
    - PowerShell Remoting enabled for remote targets
    - Registry access permissions for HKEY_USERS and HKEY_LOCAL_MACHINE
    #>

    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Name', 'Server')]
        [string[]]$ComputerName = $env:COMPUTERNAME,

        [PSCredential]$Credential,

        [switch]$TableOutput
    )

    process {
        # Core logic -- defined once, runs locally or remotely
        $scriptBlock = {
            param([bool]$ShowConsole)

            $computerName   = $env:COMPUTERNAME
            $collectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
            $results        = [System.Collections.Generic.List[object]]::new()

            # -- Helper: determine if a path is genuinely redirected -------------------
            function Test-IsRedirected {
                param([string]$Path, [string]$UserProfilePath)
                if (-not $Path) { return $false }
                $normalPath    = $Path.TrimEnd('\')
                $normalProfile = $UserProfilePath.TrimEnd('\')
                if ($normalPath -like "$normalProfile*")  { return $false }
                if ($normalPath -like '\\*')              { return $true  }
                if ($normalPath -like '*OneDrive*')       { return $true  }
                if ($normalPath -notlike '*\Users\*')     { return $true  }
                return $false
            }

            # -- Helper: classify redirection type -------------------------------------
            function Get-RedirectionType {
                param([string]$Path)
                if (-not $Path)                                             { return 'Local'          }
                if ($Path -like 'C:\WINDOWS\system32\config\systemprofile*') { return 'System Profile' }
                if ($Path -like '*OneDrive*')                               { return 'OneDrive'       }
                if ($Path -like '\\*')                                      { return 'Network'        }
                if ($Path -like 'C:\Users\*')                               { return 'Local'          }
                return 'Manual'
            }

            # -- Helper: safe last-write timestamp ------------------------------------
            function Get-ProfileLastUpdate {
                param([string]$FilePath)
                try {
                    return (Get-Item -Path $FilePath -ErrorAction Stop).LastWriteTime.ToString(
                        'dd-MM-yyyy HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
                }
                catch [System.IO.FileNotFoundException]       { return 'File Not Found' }
                catch [System.UnauthorizedAccessException]    { return 'Access Denied'  }
                catch                                          { return 'Error'          }
            }

            # -- Helper: build a standard 'no data' result object ---------------------
            function New-ProfileErrorResult {
                param([string]$Computer, [string]$User, [string]$Message, [string]$StatusValue, [string]$ProfilePath, [string]$CollectionTime)
                $ts = if ($ProfilePath) { Get-ProfileLastUpdate -FilePath $ProfilePath } else { 'N/A' }
                [PSCustomObject]@{
                    ComputerName         = $Computer
                    Username             = $User
                    RedirectedFolders    = $Message
                    RedirectionCount     = 0
                    RedirectionTypes     = $Message
                    Desktop              = $Message
                    Documents            = $Message
                    Pictures             = $Message
                    Downloads            = $Message
                    Music                = $Message
                    Videos               = $Message
                    Favorites            = $Message
                    AppDataRoaming       = $Message
                    SavedGames           = $Message
                    Searches             = $Message
                    StartMenu            = $Message
                    Contacts             = $Message
                    Links                = $Message
                    OneDriveRedirections = $Message
                    NetworkRedirections  = $Message
                    ManualRedirections   = $Message
                    LastProfileUpdate    = $ts
                    CollectionTime       = $CollectionTime
                    Status               = $StatusValue
                }
            }

            # -- Folder key-to-display-name mapping ------------------------------------
            $folderMappings = [ordered]@{
                'AppData'                                = 'AppData(Roaming)'
                'Desktop'                                = 'Desktop'
                'Personal'                               = 'Documents'
                'My Pictures'                            = 'Pictures'
                'My Music'                               = 'Music'
                'My Video'                               = 'Videos'
                'Favorites'                              = 'Favorites'
                '{374de290-123f-4565-9164-39c4925e467b}' = 'Downloads'
                '{4C5C32FF-BB9D-43b0-B5B4-2D72E54EAAA4}' = 'Saved Games'
                '{7d1d3a04-debb-4115-95cf-2f29da2920da}' = 'Searches'
                'Start Menu'                             = 'Start Menu'
                'Contacts'                               = 'Contacts'
                'Links'                                  = 'Links'
            }

            if ($ShowConsole) {
                Write-Host "`nStarting folder redirection audit on: $computerName" -ForegroundColor Green
                Write-Host ('=' * 50) -ForegroundColor Green
            }

            # -- Load user profiles (domain/local accounts only) -----------------------
            $userProfiles = [System.Collections.Generic.List[string]]::new()
            try {
                Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' -ErrorAction Stop |
                    ForEach-Object {
                        $sid = $_.PSChildName
                        if ($sid -match '^S-1-5-21-' -and
                            $sid -notin @('S-1-5-18', 'S-1-5-19', 'S-1-5-20')) {
                            $path = $_.GetValue('ProfileImagePath')
                            if ($path -and (Test-Path $path -ErrorAction SilentlyContinue)) {
                                $userProfiles.Add($path)
                            }
                        }
                    }
            }
            catch {
                if ($ShowConsole) { Write-Host "Error reading profile list: $_" -ForegroundColor Red }
                $results.Add([PSCustomObject]@{
                    ComputerName   = $computerName
                    Username       = 'System Error'
                    Error          = "Cannot access profile registry: $($_.Exception.Message)"
                    CollectionTime = $collectionTime
                    Status         = 'Failed'
                })
                return $results
            }

            if ($userProfiles.Count -eq 0) {
                if ($ShowConsole) { Write-Host 'No user profiles found.' -ForegroundColor Yellow }
                $results.Add((New-ProfileErrorResult -Computer $computerName -User 'No Profiles' `
                    -Message 'No user profiles found' -StatusValue 'No Profiles' -ProfilePath $null -CollectionTime $collectionTime))
                return $results
            }

            # -- Process each profile --------------------------------------------------
            foreach ($profilePath in $userProfiles) {
                $username   = Split-Path $profilePath -Leaf
                $hiveLoaded = $false

                if ($ShowConsole) {
                    Write-Host "`nChecking: $username" -ForegroundColor Cyan
                    Write-Host ('-' * 46) -ForegroundColor Gray
                }

                # Resolve SID
                $userSID = $null
                try {
                    Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' -ErrorAction Stop |
                        ForEach-Object {
                            if ($_.GetValue('ProfileImagePath') -eq $profilePath) {
                                $userSID = $_.PSChildName
                            }
                        }
                }
                catch {
                    if ($ShowConsole) { Write-Host "  Error resolving SID for $username : $_" -ForegroundColor Red }
                }

                if (-not $userSID) {
                    if ($ShowConsole) { Write-Host "  SID not found for: $username" -ForegroundColor Yellow }
                    $results.Add((New-ProfileErrorResult -Computer $computerName -User $username `
                        -Message 'SID Not Found' -StatusValue 'SID Not Found' -ProfilePath $profilePath -CollectionTime $collectionTime))
                    continue
                }

                try {
                    # Profile directory must exist
                    if (-not (Test-Path $profilePath)) {
                        if ($ShowConsole) { Write-Host "  Profile directory missing: $profilePath" -ForegroundColor Yellow }
                        $results.Add((New-ProfileErrorResult -Computer $computerName -User $username `
                            -Message 'Profile Not Found' -StatusValue 'Profile Missing' -ProfilePath $profilePath -CollectionTime $collectionTime))
                        continue
                    }

                    # Load hive if not already mounted
                    if (-not (Test-Path "Registry::HKEY_USERS\$userSID")) {
                        $ntUserPath = Join-Path $profilePath 'NTUSER.DAT'
                        if (Test-Path $ntUserPath) {
                            $loadResult = Start-Process -FilePath 'reg.exe' `
                                -ArgumentList 'load', "HKU\$userSID", $ntUserPath `
                                -Wait -PassThru -WindowStyle Hidden
                            if ($loadResult.ExitCode -eq 0) {
                                $hiveLoaded = $true
                                Start-Sleep -Milliseconds 500
                            }
                            else {
                                if ($ShowConsole) { Write-Host "  Failed to load hive for: $username" -ForegroundColor Yellow }
                            }
                        }
                        else {
                            if ($ShowConsole) { Write-Host "  NTUSER.DAT missing for: $username" -ForegroundColor Yellow }
                            $results.Add((New-ProfileErrorResult -Computer $computerName -User $username `
                                -Message 'NTUSER.DAT Missing' -StatusValue 'Registry Missing' -ProfilePath $profilePath -CollectionTime $collectionTime))
                            continue
                        }
                    }

                    # Verify the hive is accessible after load attempt
                    if (-not (Test-Path "Registry::HKEY_USERS\$userSID")) {
                        if ($ShowConsole) { Write-Host "  Registry hive not accessible for: $username" -ForegroundColor Yellow }
                        $results.Add((New-ProfileErrorResult -Computer $computerName -User $username `
                            -Message 'Registry Not Accessible' -StatusValue 'Registry Not Accessible' -ProfilePath $profilePath -CollectionTime $collectionTime))
                        continue
                    }

                    # Read Shell Folders registry keys
                    $sfPath  = "Registry::HKEY_USERS\$userSID\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
                    $usfPath = "Registry::HKEY_USERS\$userSID\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
                    $sf  = if (Test-Path $sfPath)  { Get-ItemProperty $sfPath  -ErrorAction SilentlyContinue } else { $null }
                    $usf = if (Test-Path $usfPath) { Get-ItemProperty $usfPath -ErrorAction SilentlyContinue } else { $null }

                    $redirectedFolders   = @{}
                    $redirectionTypes    = @{}
                    $oneDriveRedirects   = [System.Collections.Generic.List[string]]::new()
                    $networkRedirects    = [System.Collections.Generic.List[string]]::new()
                    $manualRedirects     = [System.Collections.Generic.List[string]]::new()

                    foreach ($regKey in $folderMappings.Keys) {
                        $folderName  = $folderMappings[$regKey]
                        $redirectPath = $null

                        if ($sf -and ($sf | Get-Member -Name $regKey -MemberType Properties -ErrorAction SilentlyContinue)) {
                            $redirectPath = $sf.$regKey
                        }
                        elseif ($usf -and ($usf | Get-Member -Name $regKey -MemberType Properties -ErrorAction SilentlyContinue)) {
                            $raw = $usf.$regKey
                            if ($raw -and $raw -like '*%*') {
                                $expanded = $raw -replace '%USERPROFILE%', $profilePath `
                                                 -replace '%USERNAME%',    $username
                                $redirectPath = [Environment]::ExpandEnvironmentVariables($expanded)
                            }
                            else { $redirectPath = $raw }
                        }

                        if ($redirectPath -and (Test-IsRedirected -Path $redirectPath -UserProfilePath $profilePath)) {
                            $rType = Get-RedirectionType -Path $redirectPath
                            $redirectedFolders[$folderName]  = $redirectPath
                            $redirectionTypes[$folderName]   = $rType

                            switch ($rType) {
                                'OneDrive' { $oneDriveRedirects.Add("$folderName=$redirectPath") }
                                'Network'  { $networkRedirects.Add("$folderName=$redirectPath")  }
                                'Manual'   { $manualRedirects.Add("$folderName=$redirectPath")   }
                            }

                            if ($ShowConsole) {
                                $color = switch ($rType) {
                                    'OneDrive' { 'Cyan'    }
                                    'Network'  { 'Yellow'  }
                                    'Manual'   { 'Magenta' }
                                    default    { 'White'   }
                                }
                                Write-Host "  $folderName -> $redirectPath [$rType]" -ForegroundColor $color
                            }
                        }
                    }

                    if ($ShowConsole) {
                        if ($redirectedFolders.Count -eq 0) {
                            Write-Host '  No folder redirections found.' -ForegroundColor Gray
                        }
                        else {
                            Write-Host "  Total redirected: $($redirectedFolders.Count)" -ForegroundColor Green
                        }
                    }

                    # Build redirection type summary
                    $allTypes = @($redirectionTypes.Values | Sort-Object -Unique)
                    $typeSummary = if ($allTypes.Count -gt 1) {
                        'Mixed (' + ($allTypes -join ', ') + ')'
                    }
                    elseif ($allTypes.Count -eq 1) { $allTypes[0] }
                    else                           { 'None'        }

                    # Helper to build per-folder field value
                    $fv = {
                        param($name)
                        if ($redirectedFolders[$name]) {
                            "$($redirectedFolders[$name]) [$($redirectionTypes[$name])]"
                        }
                        else { 'Local' }
                    }

                    $results.Add([PSCustomObject]@{
                        ComputerName         = $computerName
                        Username             = $username
                        RedirectedFolders    = if ($redirectedFolders.Count -gt 0) {
                            ($redirectedFolders.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
                        } else { 'None' }
                        RedirectionCount     = $redirectedFolders.Count
                        RedirectionTypes     = $typeSummary
                        Desktop              = & $fv 'Desktop'
                        Documents            = & $fv 'Documents'
                        Pictures             = & $fv 'Pictures'
                        Downloads            = & $fv 'Downloads'
                        Music                = & $fv 'Music'
                        Videos               = & $fv 'Videos'
                        Favorites            = & $fv 'Favorites'
                        AppDataRoaming       = & $fv 'AppData(Roaming)'
                        SavedGames           = & $fv 'Saved Games'
                        Searches             = & $fv 'Searches'
                        StartMenu            = & $fv 'Start Menu'
                        Contacts             = & $fv 'Contacts'
                        Links                = & $fv 'Links'
                        OneDriveRedirections = if ($oneDriveRedirects.Count) { $oneDriveRedirects -join '; ' } else { 'None' }
                        NetworkRedirections  = if ($networkRedirects.Count)  { $networkRedirects  -join '; ' } else { 'None' }
                        ManualRedirections   = if ($manualRedirects.Count)   { $manualRedirects   -join '; ' } else { 'None' }
                        LastProfileUpdate    = Get-ProfileLastUpdate -FilePath $profilePath
                        CollectionTime       = $collectionTime
                        Status               = 'Success'
                    })
                }
                catch {
                    if ($ShowConsole) { Write-Host "  Error processing $username : $_" -ForegroundColor Red }
                    $results.Add([PSCustomObject]@{
                        ComputerName      = $computerName
                        Username          = $username
                        Error             = $_.Exception.Message
                        LastProfileUpdate = Get-ProfileLastUpdate -FilePath $profilePath
                        CollectionTime    = $collectionTime
                        Status            = 'Failed'
                    })
                }
                finally {
                    if ($hiveLoaded) {
                        [System.GC]::Collect()
                        Start-Sleep -Milliseconds 500
                        try {
                            $unload = Start-Process -FilePath 'reg.exe' `
                                -ArgumentList 'unload', "HKU\$userSID" -Wait -PassThru -WindowStyle Hidden
                            if ($unload.ExitCode -ne 0 -and $ShowConsole) {
                                Write-Host "  Warning: Could not unload registry hive for $username" -ForegroundColor Yellow
                            }
                        }
                        catch {
                            if ($ShowConsole) {
                                Write-Host "  Warning: Registry cleanup failed for $username" -ForegroundColor Yellow
                            }
                        }
                    }
                }
            }

            return $results
        }

        foreach ($computer in $ComputerName) {
            try {
                Write-Verbose "Auditing folder redirections on: $computer"

                $showConsole = -not $TableOutput.IsPresent

                if ($computer -eq $env:COMPUTERNAME) {
                    & $scriptBlock $showConsole
                }
                else {
                    $invokeParams = @{
                        ComputerName = $computer
                        ScriptBlock  = $scriptBlock
                        ArgumentList = $showConsole
                        ErrorAction  = 'Stop'
                    }
                    if ($Credential) { $invokeParams['Credential'] = $Credential }
                    Invoke-Command @invokeParams
                }
            }
            catch {
                Write-Error -Message "Failed to connect to '$computer': $($_.Exception.Message)" -ErrorAction Continue
                [PSCustomObject]@{
                    ComputerName = $computer
                    Username     = 'Connection Error'
                    Error        = $_.Exception.Message
                    Status       = 'Failed'
                }
            }
        }
    }
}
