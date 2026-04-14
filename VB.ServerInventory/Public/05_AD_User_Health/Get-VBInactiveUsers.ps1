# ============================================================
# FUNCTION : Get-VBInactiveUsers
# VERSION  : 1.0.0
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu
# PURPOSE  : Get Active Directory users inactive for 90+ days
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Get Active Directory users inactive for 90 days or more.
.DESCRIPTION
    Queries Active Directory for enabled users whose lastLogonTimestamp is
    older than 90 days. Returns user details including last logon date.
    Supports querying alternate domain controllers with optional credentials.
.PARAMETER ComputerName
    Domain Controller to query. Defaults to local machine. Accepts pipeline input.
.PARAMETER Credential
    Alternate credentials for the AD query.
.EXAMPLE
    Get-VBInactiveUsers
.EXAMPLE
    Get-VBInactiveUsers -ComputerName DC01
.EXAMPLE
    'DC01' | Get-VBInactiveUsers -Credential (Get-Credential)
.OUTPUTS
    [PSCustomObject]: ComputerName, Name, SamAccountName, LastLogonDate, Status, CollectionTime
.NOTES
    Version  : 1.0.0
    Author   : Vibhu
    Modified : 10-04-2026
    Category : AD User Health
#>

function Get-VBInactiveUsers {
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
                # Step 1 -- Define the inactivity threshold (90 days ago)
                $thresholdDate = (Get-Date).AddDays(-90)

                # Step 2 -- Build AD query parameters
                $AdParams = @{
                    Filter     = "enabled -eq `$true -and lastLogonTimestamp -lt `$($thresholdDate.FileTime)"
                    Properties = 'lastLogonTimestamp'
                }
                if ($computer -ne $env:COMPUTERNAME) {
                    $AdParams['Server'] = $computer
                }
                if ($Credential) {
                    $AdParams['Credential'] = $Credential
                }

                # Step 3 -- Get inactive users
                $inactiveUsers = Get-ADUser @AdParams

                # Step 4 -- Emit results
                foreach ($user in $inactiveUsers) {
                    $lastLogonDate = if ($user.lastLogonTimestamp) {
                        [DateTime]::FromFileTime($user.lastLogonTimestamp)
                    }
                    else {
                        $null
                    }

                    [PSCustomObject]@{
                        ComputerName   = $computer
                        Name           = $user.Name
                        SamAccountName = $user.SamAccountName
                        LastLogonDate  = $lastLogonDate
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
                    LastLogonDate  = $null
                    Error          = $_.Exception.Message
                    Status         = 'Failed'
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
