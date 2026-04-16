# ============================================================
# FUNCTION : Get-VBNetworkInformation
# VERSION  : 1.1.0
# CHANGED  : 16-04-2026 -- Refactored for reliable local/remote collection
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Retrieves network configuration information from local or remote computers
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Retrieves network configuration information from local or remote computers.

.DESCRIPTION
    Collects detailed network adapter configuration including IPv4/IPv6 addresses, subnet
    prefixes, default gateways, DNS servers, DHCP status, MAC addresses, and connection
    status. Uses reliable WMI/CIM methods for both local and remote execution.

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
    [PSCustomObject]: ComputerName, InterfaceAlias, InterfaceDescription, IPv4Address,
    IPv6Address, SubnetMask, DefaultGateway, DNSServers, DHCPEnabled, MACAddress,
    InterfaceIndex, ConnectionStatus, Status, CollectionTime

.NOTES
    Version  : 1.1.0
    Author   : Vibhu Bhatnagar
    Modified : 16-04-2026
    Category : System Network
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
                $scriptBlock = {
                    # Step 1 -- Collect all network configuration data
                    $collectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    $adapters       = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -ErrorAction Stop
                    $netAdapters    = Get-CimInstance -ClassName Win32_NetworkAdapter -ErrorAction Stop
                    $ipAddresses    = Get-CimInstance -ClassName Win32_NetworkAdapterSetting -ErrorAction Stop

                    $networkInfo    = @()

                    # Step 2 -- Process each network adapter
                    foreach ($adapter in $adapters) {
                        $netAdapter = $netAdapters | Where-Object { $_.InterfaceIndex -eq $adapter.InterfaceIndex }

                        if ($netAdapter) {
                            # Step 3 -- Extract IPv4 address and subnet
                            $ipv4Address = $null
                            $subnetMask  = $null

                            if ($adapter.IPAddress -and $adapter.IPAddress.Count -gt 0) {
                                # Find first IPv4 address (exclude IPv6)
                                $ipv4Array = $adapter.IPAddress | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' }
                                if ($ipv4Array) {
                                    $ipv4Address = $ipv4Array[0]
                                    $subnetMask  = $adapter.IPSubnet[0]
                                }
                            }

                            # Step 4 -- Extract IPv6 address
                            $ipv6Address = $null
                            if ($adapter.IPAddress -and $adapter.IPAddress.Count -gt 0) {
                                $ipv6Array = $adapter.IPAddress | Where-Object { $_ -match '^[0-9a-f:]+$' }
                                if ($ipv6Array) {
                                    $ipv6Address = $ipv6Array[0]
                                }
                            }

                            # Step 5 -- Extract default gateway
                            $defaultGateway = $null
                            if ($adapter.DefaultIPGateway -and $adapter.DefaultIPGateway.Count -gt 0) {
                                $defaultGateway = $adapter.DefaultIPGateway[0]
                            }

                            # Step 6 -- Extract DNS servers
                            $dnsServers = $null
                            if ($adapter.DNSServerSearchOrder -and $adapter.DNSServerSearchOrder.Count -gt 0) {
                                $dnsServers = $adapter.DNSServerSearchOrder -join ', '
                            }

                            # Step 7 -- Map connection status
                            $connectionStatus = switch ($netAdapter.NetConnectionStatus) {
                                0 { 'Disconnected' }
                                1 { 'Connecting' }
                                2 { 'Connected' }
                                3 { 'Disconnecting' }
                                4 { 'Hardware Not Present' }
                                5 { 'Hardware Disabled' }
                                6 { 'Hardware Malfunction' }
                                7 { 'Media Disconnected' }
                                8 { 'Authenticating' }
                                9 { 'Authentication Succeeded' }
                                10 { 'Authentication Failed' }
                                default { 'Unknown' }
                            }

                            # Step 8 -- Build output object
                            $networkInfo += [PSCustomObject]@{
                                ComputerName         = $env:COMPUTERNAME
                                InterfaceAlias       = $netAdapter.NetConnectionID
                                InterfaceDescription = $netAdapter.Description
                                IPv4Address          = $ipv4Address
                                IPv6Address          = $ipv6Address
                                SubnetMask           = $subnetMask
                                DefaultGateway       = $defaultGateway
                                DNSServers           = $dnsServers
                                DHCPEnabled          = $adapter.DHCPEnabled
                                MACAddress           = $adapter.MACAddress
                                InterfaceIndex       = $adapter.InterfaceIndex
                                ConnectionStatus     = $connectionStatus
                                Status               = 'Success'
                                CollectionTime       = $collectionTime
                            }
                        }
                    }

                    return $networkInfo
                }

                # Step 9 -- Local vs remote execution
                if ($computer -eq $env:COMPUTERNAME) {
                    $result = & $scriptBlock
                } else {
                    $params = @{
                        ComputerName = $computer
                        ScriptBlock  = $scriptBlock
                        ErrorAction  = 'Stop'
                    }
                    if ($Credential) { $params.Credential = $Credential }
                    $result = Invoke-Command @params
                }

                $result
            }
            catch {
                # Step 10 -- Error handling
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
