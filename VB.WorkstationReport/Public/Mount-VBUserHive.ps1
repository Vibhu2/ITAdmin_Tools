function Mount-VBUserHive {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipelineByPropertyName)]
        [string]$SID,

        [Parameter(ValueFromPipelineByPropertyName)]
        [string]$Username,

        [Parameter(ValueFromPipelineByPropertyName)]
        [string]$ProfilePath,

        [Parameter(ValueFromPipelineByPropertyName)]
        [string]$ComputerName = $env:COMPUTERNAME,

        [PSCredential]$Credential
    )

    process {
        try {
            # Get all user profiles first to resolve SID/Username/ProfilePath
            if ($ComputerName -eq $env:COMPUTERNAME) {
                $allProfiles = Get-CimInstance -ClassName Win32_UserProfile -Filter "Special = 'False'" -ErrorAction Stop
            }
            else {
                $allProfiles = Get-CimInstance -ClassName Win32_UserProfile -Filter "Special = 'False'" -ComputerName $ComputerName -Credential $Credential -ErrorAction Stop
            }

            # Resolve profile information based on input
            $targetProfile = $null

            if ($SID) {
                $targetProfile = $allProfiles | Where-Object { $_.SID -eq $SID }
            }
            elseif ($Username) {
                foreach ($profile in $allProfiles) {
                    try {
                        $fullAccount = (New-Object System.Security.Principal.SecurityIdentifier($profile.SID)).Translate([System.Security.Principal.NTAccount]).Value
                        if ($fullAccount -match '\\') {
                            $accountParts = $fullAccount -split '\\'
                            $profileUsername = $accountParts[1]
                        }
                        else {
                            $profileUsername = $fullAccount
                        }

                        if ($profileUsername -eq $Username) {
                            $targetProfile = $profile
                            break
                        }
                    }
                    catch {
                        continue
                    }
                }
            }

            if (-not $targetProfile) {
                return [PSCustomObject]@{
                    ComputerName  = $ComputerName
                    SID           = $SID
                    Username      = $Username
                    HiveMounted   = $false
                    AlreadyLoaded = $false
                    Error         = "Profile not found for $(if($SID){"SID: $SID"}else{"Username: $Username"})"
                    Status        = 'Failed'
                }
            }

            # Use resolved profile information
            $resolvedSID = $targetProfile.SID
            $resolvedProfilePath = if ($ProfilePath) { $ProfilePath } else { $targetProfile.LocalPath }

            # Get username for output
            $resolvedUsername = $Username
            if (-not $resolvedUsername) {
                try {
                    $fullAccount = (New-Object System.Security.Principal.SecurityIdentifier($resolvedSID)).Translate([System.Security.Principal.NTAccount]).Value
                    if ($fullAccount -match '\\') {
                        $resolvedUsername = ($fullAccount -split '\\')[1]
                    }
                    else {
                        $resolvedUsername = $fullAccount
                    }
                }
                catch {
                    $resolvedUsername = "Unknown"
                }
            }

            # Mount the hive
            if ($ComputerName -eq $env:COMPUTERNAME) {
                $ntuserPath = Join-Path $resolvedProfilePath 'NTUSER.DAT'

                if (-not (Test-Path $ntuserPath)) {
                    return [PSCustomObject]@{
                        ComputerName  = $ComputerName
                        SID           = $resolvedSID
                        Username      = $resolvedUsername
                        ProfilePath   = $resolvedProfilePath
                        HiveMounted   = $false
                        AlreadyLoaded = $false
                        Error         = "NTUSER.DAT not found at $ntuserPath"
                        Status        = 'Failed'
                    }
                }

                $loadedSIDs = Get-ChildItem "Registry::HKEY_USERS" | Select-Object -ExpandProperty PSChildName

                if ($loadedSIDs -contains $resolvedSID) {
                    return [PSCustomObject]@{
                        ComputerName  = $ComputerName
                        SID           = $resolvedSID
                        Username      = $resolvedUsername
                        ProfilePath   = $resolvedProfilePath
                        HiveMounted   = $false
                        AlreadyLoaded = $true
                        Status        = 'Success'
                    }
                }

                $result = reg.exe load "HKU\$resolvedSID" "$ntuserPath" 2>&1

                if ($LASTEXITCODE -eq 0) {
                    [PSCustomObject]@{
                        ComputerName  = $ComputerName
                        SID           = $resolvedSID
                        Username      = $resolvedUsername
                        ProfilePath   = $resolvedProfilePath
                        HiveMounted   = $true
                        AlreadyLoaded = $false
                        Status        = 'Success'
                    }
                }
                else {
                    [PSCustomObject]@{
                        ComputerName  = $ComputerName
                        SID           = $resolvedSID
                        Username      = $resolvedUsername
                        ProfilePath   = $resolvedProfilePath
                        HiveMounted   = $false
                        AlreadyLoaded = $false
                        Error         = "Failed to mount hive: $result"
                        Status        = 'Failed'
                    }
                }
            }
            else {
                $result = Invoke-Command -ComputerName $ComputerName -Credential $Credential -ScriptBlock {
                    param($resolvedSID, $resolvedProfilePath)

                    $ntuserPath = Join-Path $resolvedProfilePath 'NTUSER.DAT'

                    if (-not (Test-Path $ntuserPath)) {
                        return @{
                            HiveMounted   = $false
                            AlreadyLoaded = $false
                            Error         = "NTUSER.DAT not found at $ntuserPath"
                            Status        = 'Failed'
                        }
                    }

                    $loadedSIDs = Get-ChildItem "Registry::HKEY_USERS" | Select-Object -ExpandProperty PSChildName

                    if ($loadedSIDs -contains $resolvedSID) {
                        return @{
                            HiveMounted   = $false
                            AlreadyLoaded = $true
                            Status        = 'Success'
                        }
                    }

                    $regResult = reg.exe load "HKU\$resolvedSID" "$ntuserPath" 2>&1

                    if ($LASTEXITCODE -eq 0) {
                        return @{
                            HiveMounted   = $true
                            AlreadyLoaded = $false
                            Status        = 'Success'
                        }
                    }
                    else {
                        return @{
                            HiveMounted   = $false
                            AlreadyLoaded = $false
                            Error         = "Failed to mount hive: $regResult"
                            Status        = 'Failed'
                        }
                    }
                } -ArgumentList $resolvedSID, $resolvedProfilePath

                [PSCustomObject]@{
                    ComputerName  = $ComputerName
                    SID           = $resolvedSID
                    Username      = $resolvedUsername
                    ProfilePath   = $resolvedProfilePath
                    HiveMounted   = $result.HiveMounted
                    AlreadyLoaded = $result.AlreadyLoaded
                    Error         = $result.Error
                    Status        = $result.Status
                }
            }
        }
        catch {
            [PSCustomObject]@{
                ComputerName  = $ComputerName
                SID           = $SID
                Username      = $Username
                HiveMounted   = $false
                AlreadyLoaded = $false
                Error         = $_.Exception.Message
                Status        = 'Failed'
            }
        }
    }
}
