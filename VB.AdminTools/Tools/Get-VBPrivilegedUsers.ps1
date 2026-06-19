<#
.SYNOPSIS
Gets privileged users from Active Directory administrative groups

.DESCRIPTION
Retrieves users who are members of administrative groups in Active Directory, including their account status, password expiration dates, and group memberships. The function scans common administrative groups and returns detailed information for security auditing and compliance reporting.

.PARAMETER ComputerName
Specifies the name of the domain controller to query. Defaults to the local computer. Accepts pipeline input and supports multiple computer names.

.PARAMETER Credential
Specifies a PSCredential object for authentication when connecting to remote computers. Required for remote connections outside current domain context.

.PARAMETER IncludeGroups
Specifies which administrative groups to scan for privileged users. Defaults to common administrative groups including Administrators, Domain Admins, Enterprise Admins, Schema Admins, and various operator groups.

.EXAMPLE
Get-VBPrivilegedUsers

Retrieves all privileged users from the local domain controller, scanning default administrative groups.

.EXAMPLE
Get-VBPrivilegedUsers | Export-Csv -Path "PrivilegedUsers.csv" -NoTypeInformation

Exports privileged user information to a CSV file for reporting and compliance purposes.

.EXAMPLE
Get-VBPrivilegedUsers -ComputerName "DC01.contoso.com" -Credential (Get-Credential)

Retrieves privileged users from a specific domain controller using alternate credentials.

.EXAMPLE
@('DC01', 'DC02') | Get-VBPrivilegedUsers | Where-Object Status -eq 'Success'

Scans multiple domain controllers via pipeline and filters for successful results only.

.EXAMPLE
Get-VBPrivilegedUsers -IncludeGroups @('Domain Admins', 'Enterprise Admins') | Where-Object AccountStatus -eq 'Enabled'

Scans only Domain Admins and Enterprise Admins groups and returns only enabled accounts.

.OUTPUTS
PSCustomObject with properties:
- ComputerName: Source domain controller
- Account: User's SamAccountName
- Group: Administrative group name
- AccountStatus: Enabled or Disabled
- PasswordExpiration: Date when password expires or status
- LastLogonDate: When user last logged in
- Status: Success or Failed
- Error: Error message if Status is Failed

.NOTES
Version: 1.0
Author: Admin
Category: Windows Server Administration
Requires: ActiveDirectory PowerShell module
Requires: Domain Admin or equivalent permissions for full group enumeration
#>
function Get-VBPrivilegedUsers {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'Server', 'Host')]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [PSCredential]$Credential,
        [string[]]$IncludeGroups = @('Administrators', 'Domain Admins', 'Enterprise Admins', 'Schema Admins', 'Backup Operators', 'Server Operators', 'Account Operators', 'Print Operators')
    )
    
    process {
        foreach ($computer in $ComputerName) {
            try {
                if ($computer -eq $env:COMPUTERNAME) {
                    if (-not (Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue)) {
                        Import-Module ActiveDirectory -ErrorAction Stop
                    }
                    
                    $results = @()
                    
                    foreach ($groupName in $IncludeGroups) {
                        try {
                            $groupMembers = Get-ADGroupMember -Identity $groupName -Recursive -ErrorAction SilentlyContinue | 
                                Where-Object { $_.objectClass -eq 'user' }
                            
                            foreach ($member in $groupMembers) {
                                $userDetails = Get-ADUser -Identity $member.SID -Properties PasswordExpired, PasswordNeverExpires, AccountExpirationDate, Enabled, LastLogonDate, PasswordLastSet -ErrorAction SilentlyContinue
                                
                                $passwordExpiration = if ($userDetails.PasswordNeverExpires) {
                                    "Never expires"
                                } elseif ($userDetails.PasswordExpired) {
                                    "Expired"
                                } elseif ($userDetails.PasswordLastSet) {
                                    $maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
                                    if ($maxPasswordAge.Days -gt 0) {
                                        ($userDetails.PasswordLastSet).AddDays($maxPasswordAge.Days).ToString('dd-MM-yyyy')
                                    } else {
                                        "Never expires"
                                    }
                                } else {
                                    "Unknown"
                                }
                                
                                $results += [PSCustomObject]@{
                                    ComputerName = $computer
                                    Account = $userDetails.SamAccountName
                                    Group = $groupName
                                    AccountStatus = if ($userDetails.Enabled) { "Enabled" } else { "Disabled" }
                                    PasswordExpiration = $passwordExpiration
                                    LastLogonDate = if ($userDetails.LastLogonDate) { $userDetails.LastLogonDate.ToString('dd-MM-yyyy HH:mm:ss') } else { "Never" }
                                    Status = 'Success'
                                }
                            }
                        }
                        catch {
                            $results += [PSCustomObject]@{
                                ComputerName = $computer
                                Account = "N/A"
                                Group = $groupName
                                AccountStatus = "N/A"
                                PasswordExpiration = "N/A"
                                LastLogonDate = "N/A"
                                Error = "Group not found or access denied: $($_.Exception.Message)"
                                Status = 'Failed'
                            }
                        }
                    }
                    
                    $results
                    
                } else {
                    $results = Invoke-Command -ComputerName $computer -Credential $Credential -ScriptBlock {
                        param($Groups)
                        
                        if (-not (Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue)) {
                            Import-Module ActiveDirectory -ErrorAction Stop
                        }
                        
                        $results = @()
                        
                        foreach ($groupName in $Groups) {
                            try {
                                $groupMembers = Get-ADGroupMember -Identity $groupName -Recursive -ErrorAction SilentlyContinue | 
                                    Where-Object { $_.objectClass -eq 'user' }
                                
                                foreach ($member in $groupMembers) {
                                    $userDetails = Get-ADUser -Identity $member.SID -Properties PasswordExpired, PasswordNeverExpires, AccountExpirationDate, Enabled, LastLogonDate, PasswordLastSet -ErrorAction SilentlyContinue
                                    
                                    $passwordExpiration = if ($userDetails.PasswordNeverExpires) {
                                        "Never expires"
                                    } elseif ($userDetails.PasswordExpired) {
                                        "Expired"
                                    } elseif ($userDetails.PasswordLastSet) {
                                        $maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
                                        if ($maxPasswordAge.Days -gt 0) {
                                            ($userDetails.PasswordLastSet).AddDays($maxPasswordAge.Days).ToString('dd-MM-yyyy')
                                        } else {
                                            "Never expires"
                                        }
                                    } else {
                                        "Unknown"
                                    }
                                    
                                    $results += [PSCustomObject]@{
                                        Account = $userDetails.SamAccountName
                                        Group = $groupName
                                        AccountStatus = if ($userDetails.Enabled) { "Enabled" } else { "Disabled" }
                                        PasswordExpiration = $passwordExpiration
                                        LastLogonDate = if ($userDetails.LastLogonDate) { $userDetails.LastLogonDate.ToString('dd-MM-yyyy HH:mm:ss') } else { "Never" }
                                        Status = 'Success'
                                    }
                                }
                            }
                            catch {
                                $results += [PSCustomObject]@{
                                    Account = "N/A"
                                    Group = $groupName
                                    AccountStatus = "N/A"
                                    PasswordExpiration = "N/A"
                                    LastLogonDate = "N/A"
                                    Error = "Group not found or access denied: $($_.Exception.Message)"
                                    Status = 'Failed'
                                }
                            }
                        }
                        
                        $results
                        
                    } -ArgumentList (,$IncludeGroups)
                    
                    $results | ForEach-Object {
                        $_ | Add-Member -NotePropertyName 'ComputerName' -NotePropertyValue $computer -PassThru
                    }
                }
            }
            catch {
                [PSCustomObject]@{
                    ComputerName = $computer
                    Account = "N/A"
                    Group = "N/A"
                    AccountStatus = "N/A"
                    PasswordExpiration = "N/A"
                    LastLogonDate = "N/A"
                    Error = $_.Exception.Message
                    Status = 'Failed'
                }
            }
        }
    }
}