# ============================================================
# FUNCTION : Get-VBUserProfile
# MODULE   : VB.WorkstationReport
# VERSION  : 1.0.0
# CHANGED  : 16-04-2026 -- Initial release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Enumerates all non-system user profiles on local or remote computers
# ENCODING : UTF-8 with BOM
# ============================================================
function Get-VBUserProfile {
    <#
    .SYNOPSIS
        Enumerates all non-system user profiles on local or remote computers.

    .DESCRIPTION
        Get-VBUserProfile queries Win32_UserProfile to return every non-special user
        profile on the target computer. For each profile it resolves the SID to a
        domain and username via NTAccount translation.

        Entra ID / Azure AD accounts (S-1-12-1-* SIDs) cannot be resolved via
        NTAccount translation. These are detected by SID prefix and returned with
        Domain = 'AzureAD' and Username derived from the profile folder name.

        Supports local and remote execution. Credentials are only passed to
        Get-CimInstance when explicitly supplied.

    .PARAMETER ComputerName
        Computer names to query. Accepts pipeline input. Defaults to the local computer.

    .PARAMETER Credential
        Credentials for remote computer access. Not required for local execution.

    .EXAMPLE
        Get-VBUserProfile
        Returns all user profiles on the local computer.

    .EXAMPLE
        Get-VBUserProfile -ComputerName 'WS001','WS002' -Credential (Get-Credential)
        Returns all user profiles on two remote workstations.

    .EXAMPLE
        'WS001','WS002' | Get-VBUserProfile -Credential $cred | Where-Object Loaded
        Returns only profiles that are currently loaded (user is logged in).

    .EXAMPLE
        Get-VBUserProfile | Select-Object Username, Domain, ProfilePath, LastUseTime
        Returns a summary of all local profiles with last login time.

    .OUTPUTS
        PSCustomObject
        Returns one object per user profile with:
          - ComputerName   : Target computer
          - SID            : User SID
          - Domain         : Domain or machine name ('AzureAD' for Entra ID accounts)
          - Username       : Resolved username (profile folder name for Entra ID)
          - ProfilePath    : Local profile path
          - LastUseTime    : Last time the profile was used
          - Loaded         : True if the profile hive is currently loaded in HKEY_USERS
          - CollectionTime : Timestamp of data collection
          - Status         : 'Success' or 'Failed'
          - Error          : Error message (only present on failure)

    .NOTES
        Version      : 1.0.0
        Author       : Vibhu Bhatnagar
        Category     : User Profile Management
        Requirements :
          - PowerShell 5.1 or higher
          - PowerShell Remoting enabled for remote targets
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Name', 'Server')]
        [string[]]$ComputerName = $env:COMPUTERNAME,

        [PSCredential]$Credential
    )

    process {
        foreach ($computer in $ComputerName) {
            try {
                # Only pass ComputerName and Credential for remote queries
                $cimParams = @{
                    ClassName   = 'Win32_UserProfile'
                    Filter      = "Special = 'False'"
                    ErrorAction = 'Stop'
                }
                if ($computer -ne $env:COMPUTERNAME) {
                    $cimParams['ComputerName'] = $computer
                    if ($Credential) { $cimParams['Credential'] = $Credential }
                }

                $profiles       = Get-CimInstance @cimParams
                $collectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')

                foreach ($profile in $profiles) {
                    $domain   = $null
                    $username = $null

                    try {
                        # Translate SID to NTAccount (DOMAIN\Username)
                        # Fails for Entra ID (S-1-12-1-*) SIDs -- caught below
                        $fullAccount = (New-Object System.Security.Principal.SecurityIdentifier($profile.SID)).
                                           Translate([System.Security.Principal.NTAccount]).Value

                        if ($fullAccount -match '\\') {
                            $parts    = $fullAccount -split '\\'
                            $domain   = $parts[0]
                            $username = $parts[1]
                        }
                        else {
                            $domain   = $null
                            $username = $fullAccount
                        }
                    }
                    catch {
                        # S-1-12-1-* = Entra ID / Azure AD joined device account.
                        # These SIDs have no NTAccount mapping -- use profile path for username.
                        if ($profile.SID -match '^S-1-12-1-') {
                            $domain   = 'AzureAD'
                            $username = Split-Path $profile.LocalPath -Leaf
                        }
                        else {
                            $domain   = 'Unknown'
                            $username = Split-Path $profile.LocalPath -Leaf
                        }
                    }

                    [PSCustomObject]@{
                        ComputerName   = $computer
                        SID            = $profile.SID
                        Domain         = $domain
                        Username       = $username
                        ProfilePath    = $profile.LocalPath
                        LastUseTime    = $profile.LastUseTime
                        Loaded         = $profile.Loaded
                        CollectionTime = $collectionTime
                        Status         = 'Success'
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