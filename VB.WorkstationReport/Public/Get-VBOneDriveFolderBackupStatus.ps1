# ============================================================
# FUNCTION : Get-VBOneDriveFolderBackupStatus
# MODULE   : VB.WorkstationReport
# VERSION  : 1.1.0
# CHANGED  : 16-04-2026 -- Fix hashtable Int64 type coercion; fix Invoke-Command credential
#                          and ErrorAction; move scriptblock to begin block
#            16-04-2026 -- Fix OneDrive type detection: folder name check is now primary
#                          signal; custom-domain Business accounts (e.g. @itbd.net) now
#                          correctly identified as Business
#            16-04-2026 -- Rename Get-VBOneDriveKFMStatus -> Get-VBOneDriveFolderBackupStatus
#            16-04-2026 -- Fix ghost registry stubs (no email, generic/unexpanded folder path)
#                          misclassified as Personal; now correctly reported as Not Configured
#            16-04-2026 -- Fix SID regex: added S-1-12-1-* pattern for Azure AD / Entra ID
#                          joined devices (S-1-5-21-* only covers traditional domain/local accounts)
#            16-04-2026 -- Replaced HKEY_USERS enumeration with Win32_UserProfile; embedded
#                          hive mount/unmount logic (from Mount-VBUserHive) so all user profiles
#                          are scanned regardless of whether their hive is currently loaded
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Reports OneDrive folder backup (KFM) status for all user profiles
# ENCODING : UTF-8 with BOM
# ============================================================
function Get-VBOneDriveFolderBackupStatus {
    <#
    .SYNOPSIS
        Reports OneDrive folder backup status (Known Folder Move) for all user profiles
        on local or remote computers.

    .DESCRIPTION
        Get-VBOneDriveFolderBackupStatus enumerates all user profiles via Win32_UserProfile,
        mounts any hives that are not currently loaded in HKEY_USERS (using reg.exe load),
        reads OneDrive Known Folder Move (KFM) registry data for each user, then unloads
        any hives it mounted. This ensures all profiles are scanned regardless of whether
        the user is currently logged in.

        KFM is the OneDrive feature that automatically backs up Desktop, Documents, and
        Pictures to the cloud. For each user the function reports the OneDrive account type
        (Business, Personal, or Not Configured), email address, local sync folder, and a
        per-folder breakdown of KFM migration state and scan eligibility.

        Supported account types:
          - S-1-5-21-*  : Traditional domain and local accounts
          - S-1-12-1-*  : Azure AD / Entra ID joined devices

        Hive mounting requires local admin rights on the target computer.
        The same scriptblock runs locally or via Invoke-Command without code duplication.

    .PARAMETER ComputerName
        Computer names to query. Accepts pipeline input. Defaults to the local computer.

    .PARAMETER Credential
        Credentials for remote computer access. Not required for local execution.

    .EXAMPLE
        Get-VBOneDriveFolderBackupStatus
        Returns backup status for all OneDrive users on the local computer, including
        profiles whose hives are not currently loaded.

    .EXAMPLE
        Get-VBOneDriveFolderBackupStatus -ComputerName 'WS001','WS002' -Credential (Get-Credential)
        Queries two remote workstations using alternate credentials.

    .EXAMPLE
        'WS001','WS002' | Get-VBOneDriveFolderBackupStatus -Credential $cred |
            Where-Object { $_.KFMStatus -ne '3 of 3 folders backed up' }
        Finds workstations where folder backup is not fully configured.

    .EXAMPLE
        Get-VBOneDriveFolderBackupStatus | Select-Object UserName, OneDriveType, KFMStatus, KFMFolders
        Returns a summary with per-folder detail for all local users.

    .OUTPUTS
        PSCustomObject
        Returns one object per OneDrive user with:
          - ComputerName  : Target computer
          - UserName      : Windows username
          - SID           : User SID
          - OneDriveType  : 'Business', 'Personal', 'Not Configured', or 'Unknown'
          - UserEmail     : OneDrive account email address
          - UserFolder    : Local OneDrive sync folder path
          - KFMStatus     : Summary string, e.g. '3 of 3 folders backed up'
          - KFMFolders    : Array of per-folder objects (Folder, Status, ScanStatus, AccountType)
          - CollectionTime: Timestamp of data collection
          - Status        : 'Success' or 'Failed'
          - Error         : Error message (only present on failure)

    .NOTES
        Version      : 1.1.0
        Author       : Vibhu Bhatnagar
        Category     : User Profile Management
        Requirements :
          - PowerShell 5.1 or higher
          - Local admin rights on the target computer (required for hive mounting)
          - PowerShell Remoting enabled for remote targets
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Name', 'Server')]
        [string[]]$ComputerName = $env:COMPUTERNAME,

        [PSCredential]$Credential
    )

    begin {
        $stateLookup = @{
            0 = 'Not attempted'
            1 = 'Opted out / Not applicable'
            2 = 'Failed'
            3 = 'Succeeded but disabled later'
            4 = 'Already redirected'
            5 = 'Backed up to OneDrive'
        }

        $scanLookup = @{
            0 = 'Not scanned'
            1 = 'Eligible'
            2 = 'Scan failed'
        }

        # Scriptblock defined once in begin.
        # Hive mount/unmount logic is embedded (not a call to Mount-VBUserHive) so it
        # serialises cleanly over Invoke-Command without requiring the function to exist
        # on the remote machine.
        $scriptBlock = {
            param($stateLookup, $scanLookup)

            $computerName   = $env:COMPUTERNAME
            $collectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
            $results        = [System.Collections.Generic.List[object]]::new()
            $mountedHives   = [System.Collections.Generic.List[string]]::new()

            try {
                # ── 1. Enumerate all non-special user profiles via WMI ─────────────
                # This replaces HKEY_USERS enumeration -- Win32_UserProfile returns every
                # profile regardless of whether its hive is currently loaded, and it
                # handles S-1-5-21-* (domain/local) and S-1-12-1-* (Entra ID) alike.
                try {
                    $allProfiles = Get-CimInstance -ClassName Win32_UserProfile `
                                       -Filter "Special = 'False'" `
                                       -ErrorAction Stop
                }
                catch {
                    $results.Add([PSCustomObject]@{
                        ComputerName   = $computerName
                        Error          = "Failed to enumerate user profiles: $($_.Exception.Message)"
                        Status         = 'Failed'
                        CollectionTime = $collectionTime
                    })
                    return $results
                }

                if (-not $allProfiles) {
                    $results.Add([PSCustomObject]@{
                        ComputerName   = $computerName
                        UserName       = 'No Users Found'
                        KFMStatus      = 'No user profiles detected'
                        Status         = 'Success'
                        CollectionTime = $collectionTime
                    })
                    return $results
                }

                # ── 2. Identify which hives are already loaded ─────────────────────
                $loadedSIDs = Get-ChildItem 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue |
                              Select-Object -ExpandProperty PSChildName

                # ── 3. Mount any unloaded hives ────────────────────────────────────
                # Requires local admin. Failures are non-fatal -- that profile is simply
                # skipped when its OneDrive path is checked later.
                foreach ($profile in $allProfiles) {
                    if ($loadedSIDs -notcontains $profile.SID) {
                        $ntuserPath = Join-Path $profile.LocalPath 'NTUSER.DAT'
                        if (Test-Path $ntuserPath -ErrorAction SilentlyContinue) {
                            $regResult = reg.exe load "HKU\$($profile.SID)" "$ntuserPath" 2>&1
                            if ($LASTEXITCODE -eq 0) {
                                $mountedHives.Add($profile.SID)
                                Write-Verbose "Mounted hive: $($profile.SID)"
                            }
                            else {
                                Write-Verbose "Could not mount hive for $($profile.SID): $regResult"
                            }
                        }
                    }
                }

                # ── 4. Scan OneDrive data for every profile ────────────────────────
                foreach ($profile in $allProfiles) {
                    $sid = $profile.SID

                    # Resolve username -- SID.Translate() fails for Entra ID (S-1-12-1-*)
                    # accounts; fall back to the profile folder name in that case
                    $userName = try {
                        (New-Object System.Security.Principal.SecurityIdentifier($sid)).
                            Translate([System.Security.Principal.NTAccount]).Value.Split('\')[-1]
                    }
                    catch {
                        Split-Path $profile.LocalPath -Leaf
                    }

                    $basePath = "Registry::HKEY_USERS\$sid\Software\Microsoft\OneDrive\Accounts"

                    # Skip profiles with no OneDrive installation
                    if (-not (Test-Path $basePath -ErrorAction SilentlyContinue)) { continue }

                    try {
                        $accounts      = Get-ChildItem -Path $basePath -ErrorAction Stop
                        $allKfmFolders = [System.Collections.Generic.List[object]]::new()
                        $accountTypes  = [System.Collections.Generic.List[string]]::new()
                        $userEmails    = [System.Collections.Generic.List[string]]::new()
                        $userFolders   = [System.Collections.Generic.List[string]]::new()

                        foreach ($account in $accounts) {
                            $props = Get-ItemProperty -Path $account.PSPath -ErrorAction Stop

                            # ── Type detection (priority order) ──────────────────────
                            # 0. No email + generic folder         → Not Configured (stub)
                            # 1. Folder "OneDrive - CompanyName"   → Business (definitive)
                            # 2. SharePoint URL / onmicrosoft.com  → Business
                            # 3. live.com URL / personal email     → Personal
                            # 4. Plain "\OneDrive" folder          → Personal
                            # 5. Any other custom email domain     → Business (fallback)
                            $oneDriveType = 'Unknown'

                            if (-not $props.UserEmail -and
                                ($props.UserFolder -like '*%UserProfile%*' -or
                                 ($props.UserFolder -like '*\OneDrive' -and
                                  $props.UserFolder -notlike '*OneDrive - *'))) {
                                $oneDriveType = 'Not Configured'
                            }
                            elseif ($props.UserFolder -like '*OneDrive - *') {
                                $oneDriveType = 'Business'
                            }
                            elseif ($props.WebServiceUrl -like '*-my.sharepoint.com*' -or
                                    $props.UserEmail   -like '*@*.onmicrosoft.com') {
                                $oneDriveType = 'Business'
                            }
                            elseif ($props.WebServiceUrl -like '*onedrive.live.com*' -or
                                    $props.UserEmail   -like '*@outlook.com' -or
                                    $props.UserEmail   -like '*@hotmail.com' -or
                                    $props.UserEmail   -like '*@live.com') {
                                $oneDriveType = 'Personal'
                            }
                            elseif ($props.UserFolder -like '*\OneDrive' -and
                                    $props.UserFolder -notlike '*OneDrive - *') {
                                $oneDriveType = 'Personal'
                            }
                            elseif ($props.UserEmail -match '@[^@]+\.[^@]+$' -and
                                    $props.UserEmail -notlike '*@outlook.com' -and
                                    $props.UserEmail -notlike '*@hotmail.com' -and
                                    $props.UserEmail -notlike '*@live.com') {
                                $oneDriveType = 'Business'
                            }

                            if ($props.UserEmail)  { $userEmails.Add($props.UserEmail) }
                            if ($props.UserFolder) { $userFolders.Add($props.UserFolder) }
                            $accountTypes.Add($oneDriveType)

                            # ── Parse KFM migration JSON ──────────────────────────────
                            if ($props.LastKnownFolderMigrationState -and
                                $props.LastPerFolderMigrationScanResult) {
                                try {
                                    $migration = $props.LastKnownFolderMigrationState   | ConvertFrom-Json
                                    $scan      = $props.LastPerFolderMigrationScanResult | ConvertFrom-Json

                                    foreach ($folder in $migration.PSObject.Properties.Name) {
                                        # [int] cast required: ConvertFrom-Json returns Int64 in PS 5.1;
                                        # hashtable keys are Int32 -- ContainsKey() fails without cast
                                        $stateCode = [int]$migration.$folder
                                        $scanCode  = [int]$scan.$folder

                                        $allKfmFolders.Add([PSCustomObject]@{
                                            Folder      = $folder
                                            StatusCode  = $stateCode
                                            Status      = if ($stateLookup.ContainsKey($stateCode)) { $stateLookup[$stateCode] } else { "Unknown ($stateCode)" }
                                            ScanCode    = $scanCode
                                            ScanStatus  = if ($scanLookup.ContainsKey($scanCode))  { $scanLookup[$scanCode]  } else { "Unknown ($scanCode)"  }
                                            AccountType = $oneDriveType
                                        })
                                    }
                                }
                                catch { <# JSON parse failed for this account -- continue #> }
                            }
                        }

                        # Business takes precedence when a user has both account types
                        $finalType = if ($accountTypes -contains 'Business')        { 'Business' }
                                     elseif ($accountTypes -contains 'Personal')    { 'Personal' }
                                     elseif ($accountTypes -contains 'Not Configured') { 'Not Configured' }
                                     else                                            { 'Unknown' }

                        $finalEmail  = $userEmails  | Where-Object { $_ } | Select-Object -First 1
                        $finalFolder = $userFolders | Where-Object { $_ } | Select-Object -First 1

                        $kfmSummary = if ($allKfmFolders.Count -gt 0) {
                            $backed = ($allKfmFolders | Where-Object { $_.StatusCode -eq 5 }).Count
                            "$backed of $($allKfmFolders.Count) folders backed up"
                        }
                        else { 'No KFM data available' }

                        $results.Add([PSCustomObject]@{
                            ComputerName   = $computerName
                            UserName       = $userName
                            SID            = $sid
                            OneDriveType   = $finalType
                            UserEmail      = $finalEmail
                            UserFolder     = $finalFolder
                            KFMStatus      = $kfmSummary
                            KFMFolders     = $allKfmFolders
                            CollectionTime = $collectionTime
                            Status         = 'Success'
                        })
                    }
                    catch {
                        $results.Add([PSCustomObject]@{
                            ComputerName   = $computerName
                            UserName       = $userName
                            SID            = $sid
                            Error          = "Failed to read OneDrive registry data: $($_.Exception.Message)"
                            CollectionTime = $collectionTime
                            Status         = 'Failed'
                        })
                    }
                }

                if ($results.Count -eq 0) {
                    $results.Add([PSCustomObject]@{
                        ComputerName   = $computerName
                        UserName       = 'No OneDrive Users'
                        KFMStatus      = 'No OneDrive installations found'
                        CollectionTime = $collectionTime
                        Status         = 'Success'
                    })
                }

                return $results
            }
            finally {
                # ── 5. Unload any hives we mounted ─────────────────────────────────
                # GC collect first to release any open registry handles from PowerShell;
                # reg.exe unload fails if any handle to the hive is still open.
                if ($mountedHives.Count -gt 0) {
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                    foreach ($hiveSID in $mountedHives) {
                        reg.exe unload "HKU\$hiveSID" 2>&1 | Out-Null
                        Write-Verbose "Unloaded hive: $hiveSID"
                    }
                }
            }
        }
    }

    process {
        foreach ($computer in $ComputerName) {
            try {
                Write-Verbose "Querying OneDrive folder backup status on: $computer"

                if ($computer -eq $env:COMPUTERNAME) {
                    & $scriptBlock $stateLookup $scanLookup
                }
                else {
                    $invokeParams = @{
                        ComputerName = $computer
                        ScriptBlock  = $scriptBlock
                        ArgumentList = $stateLookup, $scanLookup
                        ErrorAction  = 'Stop'
                    }
                    if ($Credential) { $invokeParams['Credential'] = $Credential }

                    Invoke-Command @invokeParams
                }
            }
            catch {
                Write-Error -Message "Failed to connect to '$computer': $($_.Exception.Message)" -ErrorAction Continue
                [PSCustomObject]@{
                    ComputerName   = $computer
                    Error          = $_.Exception.Message
                    Status         = 'Failed'
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}