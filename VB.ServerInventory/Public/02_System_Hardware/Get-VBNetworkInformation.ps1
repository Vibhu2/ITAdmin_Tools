# ============================================================
# FUNCTION : Get-VBNetworkInformation
# VERSION  : 1.0.0
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Retrieves network configuration information from local or remote computers
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Retrieves network configuration information from local or remote computers.

.DESCRIPTION
    The Get-VBNetworkInformation function collects detailed network configuration information
    including IPv4/IPv6 addresses, subnet prefixes, default gateways, DNS servers, DHCP status,
    MAC addresses, and interface status for all network adapters. It supports both local and
    remote computer queries with pipeline input and credential authentication.

.PARAMETER ComputerName
    Target computer(s). Defaults to local machine. Accepts pipeline input.
    Aliases: Name, Server, Host

.PARAMETER Credential
    Alternate credentials for remote execution.

.EXAMPLE
    Get-VBNetworkInformation

    Retrieves network configuration from the local computer.

.EXAMPLE
    Get-VBNetworkInformation -ComputerName "SERVER01"

    Retrieves network configuration from a remote server.

.EXAMPLE
    'SRV01', 'SRV02' | Get-VBNetworkInformation -Credential (Get-Credential)

    Uses pipeline input to query multiple servers with alternate credentials.

.OUTPUTS
    [PSCustomObject]: ComputerName, InterfaceAlias, InterfaceDescription, IPv4Address, IPv6Address,
    SubnetPrefix, DefaultGateway, DNSServers, DHCPEnabled, MACAddress, InterfaceIndex, Status,
    CollectionTime

.NOTES
    Version  : 1.0.0
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : System Hardware
#>

function Get-VBNetworkInformation {
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
                # Step 1 -- Local vs remote determination
                $IsLocalComputer = ($computer -eq $env:COMPUTERNAME) -or ($computer -eq 'localhost') -or ($computer -eq '.')

                # Step 2 -- Retrieve network adapters
                if ($IsLocalComputer) {
                    # Local collection
                    $Adapters = Get-NetIPConfiguration -ErrorAction Stop
                    $AllNetAdapters = Get-NetAdapter -ErrorAction Stop
                    $AllDnsServers = Get-DnsClientServerAddress -ErrorAction Stop
                } else {
                    # Step 3 -- Remote collection setup
                    $Adapters = Get-NetIPConfiguration -ComputerName $computer -ErrorAction Stop
                    $AllNetAdapters = Get-CimInstance -ComputerName $computer -ClassName Win32_NetworkAdapter -ErrorAction Stop
                    $AllNetAdapterConfigs = Get-CimInstance -ComputerName $computer -ClassName Win32_NetworkAdapterConfiguration -ErrorAction Stop
                }

                # Step 4 -- Process each adapter
                foreach ($adapter in $Adapters) {
                    if ($IsLocalComputer) {
                        # Local adapter processing
                        $NetAdapter = $AllNetAdapters | Where-Object { $_.InterfaceIndex -eq $adapter.InterfaceIndex }
                        $DNS = $AllDnsServers | Where-Object { $_.InterfaceIndex -eq $adapter.InterfaceIndex }

                        [PSCustomObject]@{
                            ComputerName          = $computer
                            InterfaceAlias        = $adapter.InterfaceAlias
                            InterfaceDescription = $adapter.InterfaceDescription
                            IPv4Address           = $adapter.IPv4Address.IPAddress
                            IPv6Address           = $adapter.IPv6Address.IPAddress
                            SubnetPrefix          = $adapter.IPv4Address.PrefixLength
                            DefaultGateway        = $adapter.IPv4DefaultGateway.NextHop
                            DNSServers            = $DNS.ServerAddresses -join ', '
                            DHCPEnabled           = $adapter.DhcpEnabled
                            MACAddress            = $NetAdapter.MacAddress
                            InterfaceIndex        = $adapter.InterfaceIndex
                            AdapterStatus         = $NetAdapter.Status
                            Status                = 'Success'
                            CollectionTime        = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                        }
                    } else {
                        # Step 5 -- Remote adapter processing
                        $NetAdapter = $AllNetAdapters | Where-Object { $_.DeviceID -eq $adapter.InterfaceIndex }
                        $NetAdapterConfig = $AllNetAdapterConfigs | Where-Object { $_.InterfaceIndex -eq $adapter.InterfaceIndex }

                        [PSCustomObject]@{
                            ComputerName          = $computer
                            InterfaceAlias        = $adapter.InterfaceAlias
                            InterfaceDescription = $adapter.InterfaceDescription
                            IPv4Address           = $adapter.IPv4Address.IPAddress
                            IPv6Address           = $adapter.IPv6Address.IPAddress
                            SubnetPrefix          = $adapter.IPv4Address.PrefixLength
                            DefaultGateway        = $adapter.IPv4DefaultGateway.NextHop
                            DNSServers            = $NetAdapterConfig.DNSServerSearchOrder -join ', '
                            DHCPEnabled           = $NetAdapterConfig.DHCPEnabled
                            MACAddress            = $NetAdapterConfig.MACAddress
                            InterfaceIndex        = $adapter.InterfaceIndex
                            AdapterStatus         = $NetAdapter.NetConnectionStatus
                            Status                = 'Success'
                            CollectionTime        = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                        }
                    }
                }
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
