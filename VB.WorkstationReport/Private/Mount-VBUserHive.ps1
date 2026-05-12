function Mount-VBUserHive {
    <#
    .SYNOPSIS
        Mount a user's NTUSER.DAT hive to the registry for offline editing.
    .DESCRIPTION
        Resolves a user profile by SID, Username, or ProfilePath and mounts the associated
        NTUSER.DAT hive to HKEY_USERS. Works locally or remotely via Invoke-Command.
        Returns detailed status including whether the hive was newly mounted or already loaded.
    .PARAMETER SID
        User security identifier (e.g., S-1-5-21-...).
    .PARAMETER Username
        Username to resolve (domain\user or local user format).
    .PARAMETER ProfilePath
        Direct path to the user profile (overrides auto-resolution).
    .PARAMETER ComputerName
        Target computer name. Default is local machine.
    .PARAMETER Credential
        Credentials for remote operations on $ComputerName.
    .EXAMPLE
        Mount-VBUserHive -Username 'contoso\jdoe'
    .EXAMPLE
        Get-ADUser jdoe | Mount-VBUserHive -ComputerName SERVER01
    .VERSION
        v1.0 - 2026-04-16
    #>
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
            #region --- Resolve Profile ---
            
            $allProfiles = @{
                Local  = { Get-CimInstance -ClassName Win32_UserProfile -Filter "Special = 'False'" -ErrorAction Stop }
                Remote = { Get-CimInstance -ClassName Win32_UserProfile -Filter "Special = 'False'" -ComputerName $ComputerName -Credential $Credential -ErrorAction Stop }
            }

            $profiles = if ($ComputerName -eq $env:COMPUTERNAME) { & $allProfiles.Local } else { & $allProfiles.Remote }
            $targetProfile = $null

            if ($SID) {
                $targetProfile = $profiles | Where-Object { $_.SID -eq $SID }
            }
            elseif ($Username) {
                foreach ($profile in $profiles) {
                    try {
                        $fullAccount = (New-Object System.Security.Principal.SecurityIdentifier($profile.SID)).Translate([System.Security.Principal.NTAccount]).Value
                        $profileUsername = if ($fullAccount -match '\\') { ($fullAccount -split '\\')[1] } else { $fullAccount }
                        
                        if ($profileUsername -eq $Username) {
                            $targetProfile = $profile
                            break
                        }
                    }
                    catch { continue }
                }
            }

            if (-not $targetProfile) {
                $notFoundMsg = if ($SID) { "SID: $SID" } else { "Username: $Username" }
                throw "Profile not found for $notFoundMsg"
            }

            $resolvedSID = $targetProfile.SID
            $resolvedProfilePath = if ($ProfilePath) { $ProfilePath } else { $targetProfile.LocalPath }

            # Resolve username if not provided
            $resolvedUsername = $Username
            if (-not $resolvedUsername) {
                try {
                    $fullAccount = (New-Object System.Security.Principal.SecurityIdentifier($resolvedSID)).Translate([System.Security.Principal.NTAccount]).Value
                    $resolvedUsername = if ($fullAccount -match '\\') { ($fullAccount -split '\\')[1] } else { $fullAccount }
                }
                catch { $resolvedUsername = "Unknown" }
            }

            #endregion

            #region --- Mount Hive (Local or Remote) ---

            $scriptBlock = {
                param($SID, $ProfilePath)
                
                $ntuserPath = Join-Path $ProfilePath 'NTUSER.DAT'
                
                if (-not (Test-Path $ntuserPath)) {
                    throw "NTUSER.DAT not found at $ntuserPath"
                }

                $loadedSIDs = Get-ChildItem "Registry::HKEY_USERS" | Select-Object -ExpandProperty PSChildName
                
                if ($loadedSIDs -contains $SID) {
                    return @{ AlreadyLoaded = $true; HiveMounted = $false }
                }

                $regResult = reg.exe load "HKU\$SID" "$ntuserPath" 2>&1
                
                if ($LASTEXITCODE -ne 0) {
                    throw "Registry load failed: $regResult"
                }

                return @{ AlreadyLoaded = $false; HiveMounted = $true }
            }

            $mountResult = if ($ComputerName -eq $env:COMPUTERNAME) {
                & $scriptBlock $resolvedSID $resolvedProfilePath
            }
            else {
                Invoke-Command -ComputerName $ComputerName -Credential $Credential -ScriptBlock $scriptBlock -ArgumentList $resolvedSID, $resolvedProfilePath
            }

            #endregion

            # Return success
            [PSCustomObject]@{
                ComputerName  = $ComputerName
                SID           = $resolvedSID
                Username      = $resolvedUsername
                ProfilePath   = $resolvedProfilePath
                HiveMounted   = $mountResult.HiveMounted
                AlreadyLoaded = $mountResult.AlreadyLoaded
                Status        = 'Success'
            }
        }
        catch {
            [PSCustomObject]@{
                ComputerName  = $ComputerName
                SID           = $SID
                Username      = $Username
                ProfilePath   = $ProfilePath
                HiveMounted   = $false
                AlreadyLoaded = $false
                Error         = $_.Exception.Message
                Status        = 'Failed'
            }
        }
    }
}