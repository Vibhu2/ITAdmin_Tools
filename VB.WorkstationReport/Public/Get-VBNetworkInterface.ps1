# ============================================================
# FUNCTION : Get-VBNetworkInterface
# MODULE   : WorkstationReport
# VERSION  : 1.1.1
# CHANGED  : 14-04-2026 -- Standards compliance fixes
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Retrieves active physical NIC configuration from local or remote computers
# ENCODING : UTF-8 with BOM
# ============================================================

function Get-VBNetworkInterface {
    <#
    .SYNOPSIS
    Retrieves active physical network interface configuration from local or remote computers.

    .DESCRIPTION
    Get-VBNetworkInterface queries active network adapters on local or remote computers,
    filtering out virtual and pseudo adapters. For each active interface it returns the IP
    address, subnet prefix length, default gateway, DNS servers, and whether the address
    was assigned by DHCP or is statically configured.

    A single scriptblock handles both local and remote execution to avoid code duplication.
    ComputerName is passed into the scriptblock so output objects are consistently populated
    regardless of whether Invoke-Command or direct invocation is used.

    .PARAMETER ComputerName
    Computer names to query. Accepts pipeline input. Defaults to the local computer.

    .PARAMETER Credential
    Credentials for remote computer access. Not required for local execution.

    .EXAMPLE
    Get-VBNetworkInterface

    Returns active network interface details for the local computer.

    .EXAMPLE
    Get-VBNetworkInterface -ComputerName 'WS001','WS002' -Credential (Get-Credential)

    Returns network interface details for two remote workstations.

    .EXAMPLE
    'WS001','WS002' | Get-VBNetworkInterface |
        Where-Object { $_.DHCPEnabled -eq $false } |
        Select-Object ComputerName, InterfaceName, IPAddress, DefaultGateway

    Pipeline input -- finds all statically-configured interfaces across multiple machines.

    .EXAMPLE
    Get-VBNetworkInterface -ComputerName $computers -Credential $cred |
        Export-Csv 'NetworkInterfaces.csv' -NoTypeInformation

    Exports network interface data for a list of computers to CSV.

    .OUTPUTS
    PSCustomObject
    Returns one object per active interface with:
    - ComputerName         : Target computer
    - InterfaceName        : Adapter short name (e.g. 'Ethernet')
    - InterfaceDescription : Full adapter description string
    - IPAddress            : IPv4 address, or 'None' if not assigned
    - SubnetMask           : CIDR prefix length (e.g. 24), or 'None'
    - DefaultGateway       : Default gateway IP, or 'None'
    - DNSServers           : Comma-separated DNS server IPs, or 'None'
    - IPType               : 'Dynamic' (DHCP) or 'Static'
    - DHCPEnabled          : Boolean -- $true if address is DHCP-assigned
    - CollectionTime       : Timestamp of data collection
    - Status               : 'Success' or 'Failed'
    - Error                : Error message (only present on failure)

    .NOTES
    Version : 1.1.1
    Author  : Vibhu Bhatnagar
    Category: Windows Workstation Administration

    Requirements:
    - PowerShell 5.1 or higher
    - NetAdapter and NetTCPIP modules (included in Windows 8.1 / Server 2012 R2 and later)
    - PowerShell Remoting enabled for remote targets
    #>

    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Name', 'Server', 'Host')]
        [string[]]$ComputerName = $env:COMPUTERNAME,

        [PSCredential]$Credential
    )

    process {
        # Core logic defined once -- executed locally or via Invoke-Command
        $scriptBlock = {
            param([string]$TargetName)

            $interfaces = Get-NetAdapter |
                Where-Object {
                    $_.Status -eq 'Up' -and
                    $_.InterfaceDescription -notlike '*Virtual*' -and
                    $_.InterfaceDescription -notlike '*Pseudo*'
                }

            $collectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')

            foreach ($interface in $interfaces) {
                $ipConfig  = Get-NetIPConfiguration -InterfaceIndex $interface.InterfaceIndex -ErrorAction SilentlyContinue
                $ipAddress = Get-NetIPAddress -InterfaceIndex $interface.InterfaceIndex `
                    -AddressFamily IPv4 -ErrorAction SilentlyContinue

                $isDHCP     = if ($ipAddress) { $ipAddress.PrefixOrigin -eq 'Dhcp' } else { $false }
                $dnsServers = $ipConfig.DNSServer |
                    Where-Object { $_.AddressFamily -eq 2 } |
                    Select-Object -ExpandProperty ServerAddresses

                [PSCustomObject]@{
                    ComputerName         = $TargetName
                    InterfaceName        = $interface.Name
                    InterfaceDescription = $interface.InterfaceDescription
                    IPAddress            = if ($ipAddress) { $ipAddress.IPAddress }                      else { 'None' }
                    SubnetMask           = if ($ipAddress) { $ipAddress.PrefixLength }                   else { 'None' }
                    DefaultGateway       = if ($ipConfig.IPv4DefaultGateway) { $ipConfig.IPv4DefaultGateway.NextHop } else { 'None' }
                    DNSServers           = if ($dnsServers) { $dnsServers -join ', ' }                   else { 'None' }
                    IPType               = if ($isDHCP) { 'Dynamic' } else { 'Static' }
                    DHCPEnabled          = $isDHCP
                    CollectionTime       = $collectionTime
                    Status               = 'Success'
                }
            }
        }

        foreach ($computer in $ComputerName) {
            try {
                Write-Verbose "Querying network interfaces on: $computer"

                if ($computer -eq $env:COMPUTERNAME) {
                    & $scriptBlock $computer
                }
                else {
                    $invokeParams = @{
                        ComputerName = $computer
                        ScriptBlock  = $scriptBlock
                        ArgumentList = $computer
                        ErrorAction  = 'Stop'
                    }
                    if ($Credential) { $invokeParams['Credential'] = $Credential }
                    Invoke-Command @invokeParams
                }
            }
            catch {
                Write-Error -Message "Failed to query '$computer': $($_.Exception.Message)" -ErrorAction Continue
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
