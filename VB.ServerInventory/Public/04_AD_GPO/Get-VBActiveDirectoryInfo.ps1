# ============================================================
# FUNCTION : Get-VBActiveDirectoryInfo
# VERSION  : 1.0.0
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Collects comprehensive Active Directory information including domain controllers, servers, FSMO roles, users, and security status
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Collects comprehensive Active Directory information.

.DESCRIPTION
    Retrieves detailed information about the Active Directory environment including domain controllers,
    all Windows servers, FSMO roles, user folder configurations, SYSVOL scripts, privileged user accounts,
    functional levels, and Recycle Bin status.

.PARAMETER ComputerName
    Domain Controller to query. Defaults to local machine. Accepts pipeline input.

.PARAMETER Credential
    Alternate credentials for the AD query.

.EXAMPLE
    Get-VBActiveDirectoryInfo

.EXAMPLE
    Get-VBActiveDirectoryInfo -ComputerName DC01

.EXAMPLE
    'DC01','DC02' | Get-VBActiveDirectoryInfo -Credential (Get-Credential)

.OUTPUTS
    [PSCustomObject]: ComputerName, DomainControllers, AllServers, FSMORoles, UserFolderReport,
    SysvolScripts, PrivilegedUsers, DomainFunctionalLevel, ForestFunctionalLevel, RecycleBinEnabled,
    TombstoneLifetime, TotalADUsers, ADRecyclebin, Status, CollectionTime

.NOTES
    Version  : 1.0.0
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : AD / GPO
#>

function Get-VBActiveDirectoryInfo {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'Server', 'Host')]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [PSCredential]$Credential
    )
    begin {
        Import-Module ActiveDirectory -ErrorAction Stop
        Import-Module GroupPolicy -ErrorAction Stop
    }



    process {
        foreach ($computer in $ComputerName) {
            try {
                # Step 1 -- Prepare AD query parameters
                $AdParams = @{}
                if ($computer -ne $env:COMPUTERNAME) {
                    $AdParams['Server'] = $computer
                }
                if ($Credential) {
                    $AdParams['Credential'] = $Credential
                }

                # Step 2 -- Retrieve domain controllers
                $DomainControllers = Get-ADDomainController -Filter * @AdParams |
                    Select-Object Name, Domain, Forest, OperationMasterRoles, IsReadOnly

                # Step 3 -- Retrieve all Windows servers
                $AllServers = Get-ADComputer -Filter { OperatingSystem -Like "Windows Server*" } -Property * @AdParams |
                    Select-Object Name, IPv4Address, OperatingSystem, OperatingSystemVersion, Enabled, LastLogonDate, WhenCreated |
                    Sort-Object OperatingSystemVersion

                # Step 4 -- Retrieve FSMO roles
                $FsmoRoles = [PSCustomObject]@{
                    InfrastructureMaster = (Get-ADDomain @AdParams).InfrastructureMaster
                    PDCEmulator          = (Get-ADDomain @AdParams).PDCEmulator
                    RIDMaster            = (Get-ADDomain @AdParams).RIDMaster
                    DomainNamingMaster   = (Get-ADForest @AdParams).DomainNamingMaster
                    SchemaMaster         = (Get-ADForest @AdParams).SchemaMaster
                }

                # Step 5 -- Get functional levels
                $DomainFunctionalLevel = (Get-ADDomain @AdParams).DomainMode
                $ForestFunctionalLevel = (Get-ADForest @AdParams).ForestMode

                # Step 6 -- Check Recycle Bin status
                $RecycleBinEnabled = (Get-ADOptionalFeature -Filter { Name -eq 'Recycle Bin Feature' } @AdParams).EnabledScopes.Count -gt 0

                # Step 7 -- Get tombstone lifetime
                $RootDse = Get-ADRootDSE @AdParams
                $ConfigContext = $RootDse.configurationNamingContext
                $TombstoneLifetime = (Get-ADObject -Identity "CN=Directory Service,CN=Windows NT,CN=Services,$ConfigContext" -Properties tombstoneLifetime @AdParams).tombstoneLifetime

                # Step 8 -- Retrieve all users and build folder report
                $Users = Get-ADUser -Filter * -Properties SamAccountName, ProfilePath, ScriptPath, homeDrive, homeDirectory @AdParams

                $UserFolderReport = foreach ($user in $Users) {
                    [PSCustomObject]@{
                        SamAccountName = $user.SamAccountName
                        ProfilePath    = if ([string]::IsNullOrEmpty($user.ProfilePath)) { "N/A" } else { $user.ProfilePath }
                        LogonScript    = if ([string]::IsNullOrEmpty($user.ScriptPath)) { "N/A" } else { $user.ScriptPath }
                        HomeDrive      = if ([string]::IsNullOrEmpty($user.homeDrive)) { "N/A" } else { $user.homeDrive }
                        HomeDirectory  = if ([string]::IsNullOrEmpty($user.homeDirectory)) { "N/A" } else { $user.homeDirectory }
                    }
                }

                $TotalUsers = $Users.Count

                # Step 9 -- Retrieve SYSVOL scripts
                $ScriptBlock = {
                    if (Test-Path -Path "C:\Windows\SYSVOL\sysvol") {
                        $FolderPath = (Get-ChildItem "C:\Windows\SYSVOL\sysvol" | Where-Object { $_.PSIsContainer } | Select-Object -First 1).FullName
                        Get-ChildItem -Recurse -Path "$FolderPath" -ErrorAction SilentlyContinue |
                            Where-Object { $_.Extension -in ".bat", ".cmd", ".ps1", ".vbs", ".exe", ".msi" } |
                            Select-Object FullName, Length, LastWriteTime
                    }
                }

                if ($computer -eq $env:COMPUTERNAME) {
                    $SysvolScripts = & $ScriptBlock
                } else {
                    $SysvolScripts = Invoke-Command -ComputerName $computer -ScriptBlock $ScriptBlock -Credential $Credential -ErrorAction SilentlyContinue
                }

                # Step 10 -- Retrieve privileged users
                $PrivilegedUsers = @()

                $GroupNames = 'Enterprise Admins', 'Domain Admins', 'Schema Admins'
                foreach ($groupName in $GroupNames) {
                    try {
                        $Members = Get-ADGroupMember -Identity $groupName -Recursive -ErrorAction SilentlyContinue @AdParams | Sort-Object Name
                        $PrivilegedUsers += foreach ($member in $Members) {
                            Get-ADUser -Identity $member.SID -Properties * @AdParams |
                                Select-Object Name, @{Name = 'Group'; Expression = { $groupName } }, WhenCreated, LastLogonDate, SamAccountName
                        }
                    } catch {
                        # Group not found or inaccessible
                    }
                }

                # Step 11 -- Check Recycle Bin status with emoji replacement
                $RecycleBinStatus = if ($RecycleBinEnabled) { "ENABLED" } else { "DISABLED" }

                # Step 12 -- Return comprehensive object
                [PSCustomObject]@{
                    ComputerName          = $computer
                    DomainControllers     = $DomainControllers
                    AllServers            = $AllServers
                    FSMORoles             = $FsmoRoles
                    UserFolderReport      = $UserFolderReport
                    SysvolScripts         = $SysvolScripts
                    PrivilegedUsers       = $PrivilegedUsers
                    DomainFunctionalLevel = $DomainFunctionalLevel
                    ForestFunctionalLevel = $ForestFunctionalLevel
                    RecycleBinEnabled     = $RecycleBinEnabled
                    TombstoneLifetime     = $TombstoneLifetime
                    TotalADUsers          = $TotalUsers
                    ADRecyclebin          = $RecycleBinStatus
                    Status                = 'Success'
                    CollectionTime        = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            } catch {
                [PSCustomObject]@{
                    ComputerName = $computer
                    Error        = $_.Exception.Message
                    Status       = 'Failed'
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
