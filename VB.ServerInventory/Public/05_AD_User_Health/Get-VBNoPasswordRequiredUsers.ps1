# ============================================================
# FUNCTION : Get-VBNoPasswordRequiredUsers
# VERSION  : 1.0.0
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Get Active Directory users with password not required
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Get Active Directory users with password not required enabled.
.DESCRIPTION
    Queries Active Directory for enabled users who have the PasswordNotRequired
    attribute set to true. These accounts represent a security risk and should
    be reviewed.
.PARAMETER ComputerName
    Domain Controller to query. Defaults to local machine. Accepts pipeline input.
.PARAMETER Credential
    Alternate credentials for the AD query.
.EXAMPLE
    Get-VBNoPasswordRequiredUsers
.EXAMPLE
    Get-VBNoPasswordRequiredUsers -ComputerName DC01
.EXAMPLE
    'DC01' | Get-VBNoPasswordRequiredUsers -Credential (Get-Credential)
.OUTPUTS
    [PSCustomObject]: ComputerName, Name, SamAccountName, Enabled, whenCreated, Status, CollectionTime
.NOTES
    Version  : 1.0.0
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : AD User Health
#>

function Get-VBNoPasswordRequiredUsers {
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
                    Filter     = "passwordNotRequired -eq `$true -and enabled -eq `$true"
                    Properties = 'PasswordNotRequired', 'whenCreated'
                }
                if ($computer -ne $env:COMPUTERNAME) {
                    $AdParams['Server'] = $computer
                }
                if ($Credential) {
                    $AdParams['Credential'] = $Credential
                }

                # Step 2 -- Get users with password not required
                $users = Get-ADUser @AdParams

                # Step 3 -- Emit results
                foreach ($user in $users) {
                    [PSCustomObject]@{
                        ComputerName   = $computer
                        Name           = $user.Name
                        SamAccountName = $user.SamAccountName
                        Enabled        = $user.Enabled
                        whenCreated    = $user.whenCreated
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
                    Error          = $_.Exception.Message
                    Status         = 'Failed'
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
