# ============================================================
# FUNCTION : Get-VBOldAdminPasswords
# VERSION  : 1.0.0
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu
# PURPOSE  : Get Domain Admins with passwords older than threshold
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Get Domain Admins with passwords older than configured threshold.
.DESCRIPTION
    Queries Active Directory for members of the Domain Admins group
    whose PasswordLastSet is older than the specified threshold (default 365 days).
    Returns admin details including password age for audit and compliance purposes.
.PARAMETER ComputerName
    Domain Controller to query. Defaults to local machine. Accepts pipeline input.
.PARAMETER Credential
    Alternate credentials for the AD query.
.PARAMETER DayThreshold
    Number of days for password age threshold. Default is 365 days.
.EXAMPLE
    Get-VBOldAdminPasswords
.EXAMPLE
    Get-VBOldAdminPasswords -ComputerName DC01
.EXAMPLE
    Get-VBOldAdminPasswords -DayThreshold 180
.EXAMPLE
    'DC01' | Get-VBOldAdminPasswords -Credential (Get-Credential)
.OUTPUTS
    [PSCustomObject]: ComputerName, Name, SamAccountName, Enabled, PasswordLastSet, Status, CollectionTime
.NOTES
    Version  : 1.0.0
    Author   : Vibhu
    Modified : 10-04-2026
    Category : AD User Health
#>

function Get-VBOldAdminPasswords {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'Server', 'Host')]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [PSCredential]$Credential,
        [int]$DayThreshold = 365
    )
    begin {
        Import-Module ActiveDirectory -ErrorAction Stop
    }


    process {
        foreach ($computer in $ComputerName) {
            try {
                # Step 1 -- Define the password age threshold
                $thresholdDate = (Get-Date).AddDays(-$DayThreshold)

                # Step 2 -- Build AD query parameters for Domain Admins group
                $GroupParams = @{
                    Identity  = 'Domain Admins'
                    Recursive = $true
                }
                if ($computer -ne $env:COMPUTERNAME) {
                    $GroupParams['Server'] = $computer
                }
                if ($Credential) {
                    $GroupParams['Credential'] = $Credential
                }

                # Step 3 -- Get Domain Admins members
                $admins = Get-ADGroupMember @GroupParams | Where-Object { $_.objectClass -eq 'user' }

                # Step 4 -- Build user query parameters
                $UserParams = @{
                    Properties = 'PasswordLastSet', 'Enabled'
                }
                if ($computer -ne $env:COMPUTERNAME) {
                    $UserParams['Server'] = $computer
                }
                if ($Credential) {
                    $UserParams['Credential'] = $Credential
                }

                # Step 5 -- Check each admin's password age
                foreach ($admin in $admins) {
                    try {
                        $user = Get-ADUser -Identity $admin.SamAccountName @UserParams

                        if ($user.PasswordLastSet -lt $thresholdDate) {
                            [PSCustomObject]@{
                                ComputerName   = $computer
                                Name           = $user.Name
                                SamAccountName = $user.SamAccountName
                                Enabled        = $user.Enabled
                                PasswordLastSet = $user.PasswordLastSet
                                Status         = 'Success'
                                CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                            }
                        }
                    }
                    catch {
                        [PSCustomObject]@{
                            ComputerName   = $computer
                            Name           = $admin.Name
                            SamAccountName = $admin.SamAccountName
                            Enabled        = $null
                            PasswordLastSet = $null
                            Error          = $_.Exception.Message
                            Status         = 'Failed'
                            CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                        }
                    }
                }
            }
            catch {
                [PSCustomObject]@{
                    ComputerName   = $computer
                    Name           = $null
                    SamAccountName = $null
                    Enabled        = $null
                    PasswordLastSet = $null
                    Error          = $_.Exception.Message
                    Status         = 'Failed'
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
