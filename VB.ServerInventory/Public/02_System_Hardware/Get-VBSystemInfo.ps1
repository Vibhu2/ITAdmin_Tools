# ============================================================
# FUNCTION : Get-VBSystemInfo
# VERSION  : 1.0.3
# CHANGED  : 16-04-2026 -- Added native API firmware type detection
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Retrieves comprehensive system information from local or remote computers
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Retrieves comprehensive system information from local or remote computers.

.DESCRIPTION
    The Get-VBSystemInfo function collects detailed system information including OS details,
    hardware specifications, network configuration, domain information, and BIOS/firmware type.
    It supports both local and remote computer queries with pipeline input and credential authentication.
    Uses native Windows API for accurate BIOS/firmware type detection (UEFI vs Legacy).

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
    Version  : 1.0.3
    Author   : Vibhu Bhatnagar
    Modified : 16-04-2026
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

    begin {
        if (-not ([System.Management.Automation.PSTypeName]'FirmwareType').Type) {
            Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class FirmwareType {
    [DllImport("kernel32.dll")]
    public static extern bool GetFirmwareType(out uint firmwareType);
}
"@
        }
    }

    process {
        foreach ($computer in $ComputerName) {
            try {
                if ($computer -eq $env:COMPUTERNAME) {
                    # Local collection
                    $OS = Get-CimInstance -ClassName Win32_OperatingSystem
                    $CS = Get-CimInstance -ClassName Win32_ComputerSystem
                    $BIOS = Get-CimInstance -ClassName Win32_BIOS
                    $Processor = Get-CimInstance -ClassName Win32_Processor | Select-Object -First 1
                    $TimeZone = Get-CimInstance -ClassName Win32_TimeZone
                    $Registry = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction SilentlyContinue
                    
                    # Network info
                    $NetAdapter = Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notmatch 'Loopback|vEthernet' } | Select-Object -First 1
                    $IPAddress = $NetAdapter.IPAddress
                    $NetConfig = Get-NetIPConfiguration -InterfaceIndex $NetAdapter.InterfaceIndex
                    $DHCPStatus = if ($NetConfig.NetIPv4Interface.Dhcp -eq 'Enabled') { 'Enabled' } else { 'Disabled' }
                    $DNSServers = ($NetConfig.DnsServer.ServerAddresses -join ', ')
                    
                    # Get DNS suffix from multiple sources
                    $DNSSuffix = $NetConfig.DnsSuffix
                    if (-not $DNSSuffix) {
                        $DNSSuffixSearch = Get-DnsClientGlobalSetting -ErrorAction SilentlyContinue
                        $DNSSuffix = $DNSSuffixSearch.SuffixSearchList[0]
                    }
                    
                    # DHCP lease info from registry (more reliable)
                    $DHCPLeaseInfo = Get-ChildItem "HKLM:\SYSTEM\CurrentControlSet\Services\Dhcp\Parameters\Interfaces" -ErrorAction SilentlyContinue | ForEach-Object {
                        Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue | Select-Object -First 1
                    }
                    $DHCPServer = $DHCPLeaseInfo.DhcpServer
                    $DHCPLeaseObtained = if ($DHCPLeaseInfo.LeaseObtainedTime) { 
                        [DateTime]::FromFileTime($DHCPLeaseInfo.LeaseObtainedTime).ToString('g') 
                    }
                    else { 
                        'N/A' 
                    }
                    $DHCPLeaseExpires = if ($DHCPLeaseInfo.T1) { 
                        [DateTime]::FromFileTime($DHCPLeaseInfo.T1).ToString('g') 
                    }
                    else { 
                        'N/A' 
                    }

                    # Firmware type
                    $fw = 0
                    [FirmwareType]::GetFirmwareType([ref]$fw) | Out-Null
                    $BIOSType = switch ($fw) {
                        1 { "BIOS" }
                        2 { "UEFI" }
                        default { "Unknown" }
                    }

                    $LogonServer = $env:LOGONSERVER
                }
                else {
                    # Remote collection
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

                    # Remote script block with registry-based DHCP info
                    $InvokeParams = @{
                        ComputerName = $computer
                        ScriptBlock  = {
                            $Registry = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction SilentlyContinue
                            $NetAdapter = Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notmatch 'Loopback|vEthernet' } | Select-Object -First 1
                            $IPAddress = $NetAdapter.IPAddress
                            $NetConfig = Get-NetIPConfiguration -InterfaceIndex $NetAdapter.InterfaceIndex
                            $DHCPStatus = if ($NetConfig.NetIPv4Interface.Dhcp -eq 'Enabled') { 'Enabled' } else { 'Disabled' }
                            $DNSServers = ($NetConfig.DnsServer.ServerAddresses -join ', ')
                            
                            # DNS suffix
                            $DNSSuffix = $NetConfig.DnsSuffix
                            if (-not $DNSSuffix) {
                                $DNSSuffixSearch = Get-DnsClientGlobalSetting -ErrorAction SilentlyContinue
                                $DNSSuffix = $DNSSuffixSearch.SuffixSearchList[0]
                            }
                            
                            # DHCP lease from registry
                            $DHCPLeaseInfo = Get-ChildItem "HKLM:\SYSTEM\CurrentControlSet\Services\Dhcp\Parameters\Interfaces" -ErrorAction SilentlyContinue | ForEach-Object {
                                Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue | Select-Object -First 1
                            }
                            $DHCPServer = $DHCPLeaseInfo.DhcpServer
                            $DHCPLeaseObtained = if ($DHCPLeaseInfo.LeaseObtainedTime) { 
                                [DateTime]::FromFileTime($DHCPLeaseInfo.LeaseObtainedTime).ToString('g') 
                            }
                            else { 
                                'N/A' 
                            }
                            $DHCPLeaseExpires = if ($DHCPLeaseInfo.T1) { 
                                [DateTime]::FromFileTime($DHCPLeaseInfo.T1).ToString('g') 
                            }
                            else { 
                                'N/A' 
                            }

                            # Firmware type
                            if (-not ([System.Management.Automation.PSTypeName]'FirmwareType').Type) {
                                Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class FirmwareType {
    [DllImport("kernel32.dll")]
    public static extern bool GetFirmwareType(out uint firmwareType);
}
"@
                            }
                            $fw = 0
                            [FirmwareType]::GetFirmwareType([ref]$fw) | Out-Null
                            $BIOSType = switch ($fw) {
                                1 { "BIOS" }
                                2 { "UEFI" }
                                default { "Unknown" }
                            }

                            $LogonServer = $env:LOGONSERVER

                            return @{
                                IPAddress         = $IPAddress
                                DHCPStatus        = $DHCPStatus
                                DNSServers        = $DNSServers
                                DNSSuffix         = $DNSSuffix
                                DHCPServer        = $DHCPServer
                                DHCPLeaseObtained = $DHCPLeaseObtained
                                DHCPLeaseExpires  = $DHCPLeaseExpires
                                BIOSType          = $BIOSType
                                LogonServer       = $LogonServer
                                Registry          = $Registry
                            }
                        }
                    }
                    if ($Credential) { $InvokeParams.Credential = $Credential }
                    $RemoteData = Invoke-Command @InvokeParams
                    $IPAddress = $RemoteData.IPAddress
                    $DHCPStatus = $RemoteData.DHCPStatus
                    $DNSServers = $RemoteData.DNSServers
                    $DNSSuffix = $RemoteData.DNSSuffix
                    $DHCPServer = $RemoteData.DHCPServer
                    $DHCPLeaseObtained = $RemoteData.DHCPLeaseObtained
                    $DHCPLeaseExpires = $RemoteData.DHCPLeaseExpires
                    $BIOSType = $RemoteData.BIOSType
                    $LogonServer = $RemoteData.LogonServer
                    $Registry = $RemoteData.Registry

                    Remove-CimSession -CimSession $Session
                }

                # Output object
                $SystemInfo = [PSCustomObject]@{
                    ComputerName       = $computer
                    DNSHostname        = $CS.DNSHostName
                    Domain             = $CS.Domain
                    DomainRole         = switch ($CS.DomainRole) {
                        0 { 'Standalone Workstation' }
                        1 { 'Member Workstation' }
                        2 { 'Standalone Server' }
                        3 { 'Member Server' }
                        4 { 'Backup Domain Controller' }
                        5 { 'Primary Domain Controller' }
                        default { $CS.DomainRole }
                    }
                    LogonServer        = $LogonServer
                    PrimaryIP          = $IPAddress
                    DHCPStatus         = $DHCPStatus
                    DHCPServer         = $DHCPServer
                    DHCPLeaseObtained  = $DHCPLeaseObtained
                    DHCPLeaseExpires   = $DHCPLeaseExpires
                    DNSServers         = $DNSServers
                    DNSSuffix          = $DNSSuffix
                    OSName             = $OS.Caption
                    OSVersion          = $OS.Version
                    OSBuild            = $OS.BuildNumber
                    OSDisplayVersion   = if ($Registry.DisplayVersion) { $Registry.DisplayVersion } else { $OS.BuildNumber }
                    OSInstallDate      = $OS.InstallDate
                    SystemManufacturer = $CS.Manufacturer
                    SystemModel        = $CS.Model
                    Processor          = $Processor.Name
                    Cores              = $Processor.NumberOfCores
                    LogicalProcessors  = $Processor.NumberOfLogicalProcessors
                    TotalMemoryGB      = [math]::Round($CS.TotalPhysicalMemory / 1GB, 2)
                    FreeMemoryGB       = [math]::Round($OS.FreePhysicalMemory / 1MB, 2)
                    BIOSType           = $BIOSType
                    BIOSVersion        = $BIOS.SMBIOSBIOSVersion
                    BIOSManufacturer   = $BIOS.Manufacturer
                    BIOSSerialNumber   = $BIOS.SerialNumber
                    BIOSReleaseDate    = $BIOS.ReleaseDate
                    TimeZone           = $TimeZone.StandardName
                    LastBootTime       = $OS.LastBootUpTime
                    Status             = 'Success'
                    CollectionTime     = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }

                $SystemInfo
            }
            catch {
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