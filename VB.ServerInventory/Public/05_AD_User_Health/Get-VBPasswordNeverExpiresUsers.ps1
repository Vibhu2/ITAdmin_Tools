# ============================================================
# FUNCTION : Get-VBPasswordNeverExpiresUsers
# VERSION  : 1.0.2
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Get Active Directory users with password never expires
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Get Active Directory users with passwords set to never expire.
.DESCRIPTION
    Queries Active Directory for enabled users who have the PasswordNeverExpires
    attribute set to true. These accounts represent a security risk and should
    be reviewed and configured with password expiration policies.
.PARAMETER ComputerName
    Domain Controller to query. Defaults to local machine. Accepts pipeline input.
.PARAMETER Credential
    Alternate credentials for the AD query.
.EXAMPLE
    Get-VBPasswordNeverExpiresUsers
.EXAMPLE
    Get-VBPasswordNeverExpiresUsers -ComputerName DC01
.EXAMPLE
    'DC01' | Get-VBPasswordNeverExpiresUsers -Credential (Get-Credential)
.OUTPUTS
    [PSCustomObject]: ComputerName, Name, SamAccountName, Enabled, whenCreated, LastLogon, Status, CollectionTime
.NOTES
    Version  : 1.0.2
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : AD User Health
#>

function Get-VBPasswordNeverExpiresUsers {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'Server', 'Host')]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [PSCredential]$Credential
    )
    begin {
        Import-Module ActiveDirectory -ErrorAction Stop
    }


    process {
        foreach ($computer in $ComputerName) {
            try {
                # Step 1 -- Build AD query parameters
                $AdParams = @{
                    Filter     = "passwordNeverExpires -eq `$true -and enabled -eq `$true"
                    Properties = 'PasswordNeverExpires', 'whenCreated', 'lastLogonTimestamp'
                }
                if ($computer -ne $env:COMPUTERNAME) {
                    $AdParams['Server'] = $computer
                }
                if ($Credential) {
                    $AdParams['Credential'] = $Credential
                }

                # Step 2 -- Get users with password never expires
                $users = Get-ADUser @AdParams

                # Step 3 -- Emit results
                foreach ($user in $users) {
                    $lastLogon = if ($user.lastLogonTimestamp) {
                        [DateTime]::FromFileTime($user.lastLogonTimestamp)
                    }
                    else {
                        'Never Logged On'
                    }

                    [PSCustomObject]@{
                        ComputerName   = $computer
                        Name           = $user.Name
                        SamAccountName = $user.SamAccountName
                        Enabled        = $user.Enabled
                        whenCreated    = $user.whenCreated
                        LastLogon      = $lastLogon
                        Status         = 'Success'
                        CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    }
                }
            }
            catch {
                [PSCustomObject]@{
                    ComputerName   = $computer
                    Name           = $null
                    SamAccountName = $null
                    Enabled        = $null
                    whenCreated    = $null
                    LastLogon      = $null
                    Error          = $_.Exception.Message
                    Status         = 'Failed'
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
