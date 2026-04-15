# ============================================================
# FUNCTION : Get-VBADGroupsWithMemberCount
# VERSION  : 1.0.2
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Get Active Directory groups with member count
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Get Active Directory groups with member count excluding defaults.
.DESCRIPTION
    Queries Active Directory for all groups and returns their member counts,
    excluding default system groups and built-in groups. Results are sorted by
    member count for easy identification of empty or large groups.
.PARAMETER ComputerName
    Domain Controller to query. Defaults to local machine. Accepts pipeline input.
.PARAMETER Credential
    Alternate credentials for the AD query.
.EXAMPLE
    Get-VBADGroupsWithMemberCount
.EXAMPLE
    Get-VBADGroupsWithMemberCount -ComputerName DC01
.EXAMPLE
    'DC01' | Get-VBADGroupsWithMemberCount -Credential (Get-Credential)
.OUTPUTS
    [PSCustomObject]: ComputerName, SamAccountName, MemberCount, Created, Modified, Status, CollectionTime
.NOTES
    Version  : 1.0.2
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : AD User Health
#>

function Get-VBADGroupsWithMemberCount {
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
                # Step 1 -- Define excluded default groups
                $excludeGroups = @(
                    'Domain Admins', 'Domain Users', 'Domain Guests', 'Enterprise Admins', 'Schema Admins',
                    'Administrators', 'Users', 'Guests', 'Account Operators', 'Backup Operators',
                    'Print Operators', 'Server Operators', 'Replicator', 'DnsAdmins', 'DnsUpdateProxy',
                    'Cert Publishers', 'Read-only Domain Controllers', 'Group Policy Creator Owners',
                    'Access Control Assistance Operators', 'ADSyncBrowse', 'ADSyncOperators', 'ADSyncPasswordSet',
                    'Allowed RODC Password Replication Group', 'Certificate Service DCOM Access', 'Cloneable Domain Controllers',
                    'Cryptographic Operators', 'DHCP Administrators', 'DHCP Users', 'Distributed COM Users',
                    'Enterprise Key Admins', 'Enterprise Read-only Domain Controllers', 'Event Log Readers', 'Hyper-V Administrators',
                    'Incoming Forest Trust Builders', 'Key Admins', 'Network Configuration Operators', 'Office 365 Public Folder Administration',
                    'Performance Log Users', 'Performance Monitor Users', 'Protected Users', 'RAS and IAS Servers',
                    'RDS Endpoint Servers', 'RDS Management Servers', 'RDS Remote Access Servers', 'Remote Management Users',
                    'Storage Replica Administrators'
                )

                # Step 2 -- Build AD query parameters
                $AdParams = @{
                    Filter     = '*'
                    Properties = 'whenCreated', 'whenChanged'
                }
                if ($computer -ne $env:COMPUTERNAME) {
                    $AdParams['Server'] = $computer
                }
                if ($Credential) {
                    $AdParams['Credential'] = $Credential
                }

                # Step 3 -- Get all non-default groups
                $allGroups = Get-ADGroup @AdParams | Where-Object { $excludeGroups -notcontains $_.Name }

                # Step 4 -- Count members in each group
                foreach ($group in $allGroups) {
                    try {
                        $GroupMemberParams = @{
                            Identity      = $group.DistinguishedName
                            ErrorAction   = 'SilentlyContinue'
                        }
                        if ($computer -ne $env:COMPUTERNAME) {
                            $GroupMemberParams['Server'] = $computer
                        }
                        if ($Credential) {
                            $GroupMemberParams['Credential'] = $Credential
                        }

                        $members = Get-ADGroupMember @GroupMemberParams
                        $memberCount = if ($members) { $members.Count } else { 0 }

                        [PSCustomObject]@{
                            ComputerName   = $computer
                            SamAccountName = $group.SamAccountName
                            MemberCount    = $memberCount
                            Created        = $group.whenCreated
                            Modified       = $group.whenChanged
                            Status         = 'Success'
                            CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                        }
                    }
                    catch {
                        [PSCustomObject]@{
                            ComputerName   = $computer
                            SamAccountName = $group.SamAccountName
                            MemberCount    = $null
                            Created        = $group.whenCreated
                            Modified       = $group.whenChanged
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
                    SamAccountName = $null
                    MemberCount    = $null
                    Created        = $null
                    Modified       = $null
                    Error          = $_.Exception.Message
                    Status         = 'Failed'
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
