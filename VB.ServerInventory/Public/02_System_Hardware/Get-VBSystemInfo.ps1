# ============================================================
# FUNCTION : Get-VBSystemInfo
# VERSION  : 1.0.0
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Retrieves comprehensive system information from local or remote computers
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Retrieves comprehensive system information from local or remote computers.

.DESCRIPTION
    The Get-VBSystemInfo function collects detailed system information including OS details,
    hardware specifications, network configuration, and domain information. It supports both
    local and remote computer queries with pipeline input and credential authentication.

.PARAMETER ComputerName
    Target computer(s). Defaults to local machine. Accepts pipeline input.
    Aliases: Name, Server, Host

.PARAMETER Credential
    Alternate credentials for remote execution.

.EXAMPLE
    Get-VBSystemInfo

    Retrieves system information from the local computer.

.EXAMPLE
    Get-VBSystemInfo -ComputerName "SERVER01", "SERVER02"

    Retrieves system information from multiple remote servers.

.EXAMPLE
    "DC01", "WEB01" | Get-VBSystemInfo -Credential (Get-Credential)

    Uses pipeline input to query multiple servers with alternate credentials.

.OUTPUTS
    [PSCustomObject]: ComputerName, DNS Hostname, Primary IP, IP Assignment, OS Name, OS Version,
    OS Build, OS Display Version, OS Install Date, Windows Version, Hardware Abstraction Layer,
    BIOS Version, BIOS Manufacturer, BIOS Serial Number, BIOS Release Date, System Manufacturer,
    System Model, BIOS Type, Domain, Domain Role, Processor, Cores, Logical Processors,
    Total Memory (GB), Free Memory (GB), OS Server Level, Logon Server, Time Zone, Last Boot Time,
    Status, CollectionTime

.NOTES
    Version  : 1.0.0
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : System Hardware
#>

function Get-VBSystemInfo {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'Server', 'Host')]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [PSCredential]$Credential
    )

    process {
        foreach ($computer in $ComputerName) {
            try {
                # Step 1 -- Local vs remote determination and data collection
                if ($computer -eq $env:COMPUTERNAME) {
                    # Local collection
                    $OS = Get-CimInstance -ClassName Win32_OperatingSystem
                    $CS = Get-CimInstance -ClassName Win32_ComputerSystem
                    $BIOS = Get-CimInstance -ClassName Win32_BIOS
                    $Processor = Get-CimInstance -ClassName Win32_Processor | Select-Object -First 1
                    $TimeZone = Get-CimInstance -ClassName Win32_TimeZone
                    $Registry = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction SilentlyContinue
                    $NetAdapter = Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notmatch 'Loopback|vEthernet' } | Select-Object -First 1
                    $IPAddress = $NetAdapter.IPAddress
                    $NetConfig = Get-NetIPConfiguration -InterfaceIndex $NetAdapter.InterfaceIndex
                    $DHCPStatus = if ($NetConfig.NetIPv4Interface.Dhcp -eq 'Enabled') { 'DHCP' } else { 'Static' }
                } else {
                    # Step 2 -- Remote collection setup
                    $CimParams = @{
                        ComputerName = $computer
                        ErrorAction  = 'Stop'
                    }
                    if ($Credential) { $CimParams.Credential = $Credential }

                    $Session = New-CimSession @CimParams
                    $OS = Get-CimInstance -CimSession $Session -ClassName Win32_OperatingSystem
                    $CS = Get-CimInstance -CimSession $Session -ClassName Win32_ComputerSystem
                    $BIOS = Get-CimInstance -CimSession $Session -ClassName Win32_BIOS
                    $Processor = Get-CimInstance -CimSession $Session -ClassName Win32_Processor | Select-Object -First 1
                    $TimeZone = Get-CimInstance -CimSession $Session -ClassName Win32_TimeZone

                    # Step 3 -- Remote script block execution
                    $InvokeParams = @{
                        ComputerName = $computer
                        ScriptBlock  = {
                            $Registry = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction SilentlyContinue
                            $NetAdapter = Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notmatch 'Loopback|vEthernet' } | Select-Object -First 1
                            $IPAddress = $NetAdapter.IPAddress
                            $NetConfig = Get-NetIPConfiguration -InterfaceIndex $NetAdapter.InterfaceIndex
                            $DHCPStatus = if ($NetConfig.NetIPv4Interface.Dhcp -eq 'Enabled') { 'DHCP' } else { 'Static' }
                            return @{
                                IPAddress  = $IPAddress
                                DHCPStatus = $DHCPStatus
                                Registry   = $Registry
                            }
                        }
                    }
                    if ($Credential) { $InvokeParams.Credential = $Credential }
                    $RemoteData = Invoke-Command @InvokeParams
                    $IPAddress = $RemoteData.IPAddress
                    $DHCPStatus = $RemoteData.DHCPStatus
                    $Registry = $RemoteData.Registry

                    Remove-CimSession -CimSession $Session
                }

                # Step 4 -- Output object construction
                $SystemInfo = [PSCustomObject]@{
                    ComputerName                 = $computer
                    DNSHostname                  = $CS.DNSHostName
                    PrimaryIP                    = $IPAddress
                    IPAssignment                 = $DHCPStatus
                    OSName                       = $OS.Caption
                    OSVersion                    = $OS.Version
                    OSBuild                      = $OS.BuildNumber
                    OSDisplayVersion             = if ($Registry.DisplayVersion) { $Registry.DisplayVersion } else { $OS.BuildNumber }
                    OSInstallDate                = $OS.InstallDate
                    WindowsVersion               = if ($Registry.ReleaseId) { $Registry.ReleaseId } elseif ($Registry.DisplayVersion) { $Registry.DisplayVersion } else { $OS.Version }
                    HardwareAbstractionLayer     = $OS.Version
                    BIOSVersion                  = $BIOS.SMBIOSBIOSVersion
                    BIOSManufacturer             = $BIOS.Manufacturer
                    BIOSSerialNumber             = $BIOS.SerialNumber
                    BIOSReleaseDate              = $BIOS.ReleaseDate
                    SystemManufacturer           = $CS.Manufacturer
                    SystemModel                  = $CS.Model
                    BIOSType                     = switch ($BIOS.BiosCharacteristics) {
                        { $_ -contains 3 } { 'UEFI' }
                        { $_ -contains 4 } { 'Legacy' }
                        default { 'Unknown' }
                    }
                    Domain                       = $CS.Domain
                    DomainRole                   = switch ($CS.DomainRole) {
                        0 { 'Standalone Workstation' }
                        1 { 'Member Workstation' }
                        2 { 'Standalone Server' }
                        3 { 'Member Server' }
                        4 { 'Backup Domain Controller' }
                        5 { 'Primary Domain Controller' }
                        default { $CS.DomainRole }
                    }
                    Processor                    = $Processor.Name
                    Cores                        = $Processor.NumberOfCores
                    LogicalProcessors            = $Processor.NumberOfLogicalProcessors
                    TotalMemoryGB                = [math]::Round($CS.TotalPhysicalMemory / 1GB, 2)
                    FreeMemoryGB                 = [math]::Round($OS.FreePhysicalMemory / 1MB, 2)
                    OSServerLevel                = switch ($OS.ProductType) {
                        1 { 'Workstation' }
                        2 { 'Domain Controller' }
                        3 { 'Server' }
                        default { $OS.ProductType }
                    }
                    LogonServer                  = $env:LOGONSERVER
                    TimeZone                     = $TimeZone.StandardName
                    LastBootTime                 = $OS.LastBootUpTime
                    Status                       = 'Success'
                    CollectionTime               = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }

                # Step 5 -- Return object
                $SystemInfo
            }
            catch {
                # Step 6 -- Error handling
                [PSCustomObject]@{
                    ComputerName   = $computer
                    Error          = $_.Exception.Message
                    Status         = 'Failed'
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
