# ============================================================
# FUNCTION : Get-VBDNSServerInfo
# VERSION  : 1.0.2
# CHANGED  : 10-04-2026 -- VB-compliant refactor and cleanup
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Collect comprehensive DNS server configuration
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Collects comprehensive DNS server configuration including zones, forwarders, and network settings.

.DESCRIPTION
    Get-VBDNSServerInfo gathers detailed DNS server configuration from local or remote servers.
    Returns structured objects containing zone information, forwarders, DNS service status,
    and network adapter DNS configuration.

.PARAMETER ComputerName
    Target DNS server(s). Defaults to local machine. Accepts pipeline input.
    Supports multiple servers and fully-qualified domain names.

.PARAMETER Credential
    Alternate credentials for remote DNS server access.

.EXAMPLE
    Get-VBDNSServerInfo
    Retrieves DNS configuration from the local server.

.EXAMPLE
    Get-VBDNSServerInfo -ComputerName DNS-SRV01
    Retrieves DNS configuration from remote server DNS-SRV01.

.EXAMPLE
    'DNS-SRV01','DNS-SRV02' | Get-VBDNSServerInfo -Credential (Get-Credential)
    Retrieves DNS configuration from multiple servers using pipeline input.

.OUTPUTS
    [PSCustomObject]: ComputerName, ServiceStatus, ForwardZones, ReverseZones, TotalZones,
    Zones, Forwarders, NetworkAdapters, Status, CollectionTime

.NOTES
    Version  : 1.0.2
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : Services
#>

function Get-VBDNSServerInfo {
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
                # Step 1 -- Build parameter hashtable
                $Params = @{
                    ComputerName = $computer
                    ErrorAction  = 'SilentlyContinue'
                }
                if ($Credential) {
                    $Params.Credential = $Credential
                }

                # Step 2 -- Determine if local or remote execution
                if ($computer -eq $env:COMPUTERNAME) {
                    # Step 3 -- Collect local DNS information
                    $Service = Get-Service -Name 'DNS' -ErrorAction SilentlyContinue
                    $Zones = Get-DnsServerZone @Params
                    $Forwarders = (Get-DnsServerForwarder @Params -ErrorAction SilentlyContinue).IPAddress
                    $NetworkConfig = Get-CimInstance -ClassName 'Win32_NetworkAdapterConfiguration' @Params -Filter 'IPEnabled=True' -ErrorAction SilentlyContinue
                } else {
                    # Step 4 -- Collect remote DNS information
                    $RemoteResult = Invoke-Command -ComputerName $computer -Credential $Credential -ScriptBlock {
                        $Service = Get-Service -Name 'DNS' -ErrorAction SilentlyContinue
                        $Zones = Get-DnsServerZone -ErrorAction SilentlyContinue
                        $Forwarders = (Get-DnsServerForwarder -ErrorAction SilentlyContinue).IPAddress
                        $NetworkConfig = Get-CimInstance -ClassName 'Win32_NetworkAdapterConfiguration' -Filter 'IPEnabled=True' -ErrorAction SilentlyContinue

                        [PSCustomObject]@{
                            Service       = $Service
                            Zones         = $Zones
                            Forwarders    = $Forwarders
                            NetworkConfig = $NetworkConfig
                        }
                    }
                    $Service = $RemoteResult.Service
                    $Zones = $RemoteResult.Zones
                    $Forwarders = $RemoteResult.Forwarders
                    $NetworkConfig = $RemoteResult.NetworkConfig
                }

                # Step 5 -- Process zone information
                $ZoneInfo = @()
                foreach ($zone in $Zones) {
                    if ($zone.ZoneName -ne 'TrustAnchors') {
                        $RecordCount = (Get-DnsServerResourceRecord -ZoneName $zone.ZoneName @Params -ErrorAction SilentlyContinue | Measure-Object).Count
                        $ZoneInfo += [PSCustomObject]@{
                            ZoneName      = $zone.ZoneName
                            ZoneType      = $zone.ZoneType
                            DynamicUpdate = $zone.DynamicUpdate
                            RecordCount   = $RecordCount
                            IsReverse     = $zone.IsReverseLookupZone
                        }
                    }
                }

                # Step 6 -- Process network adapter information
                $NetworkAdapterInfo = $NetworkConfig | Select-Object -Property Description, @{Name = 'IPAddress'; Expression = { $_.IPAddress -join ',' } }, @{Name = 'DNSServerSearchOrder'; Expression = { $_.DNSServerSearchOrder -join ',' } }

                # Step 7 -- Return success object
                [PSCustomObject]@{
                    ComputerName   = $computer
                    ServiceStatus  = $Service.Status
                    ForwardZones   = ($ZoneInfo | Where-Object { -not $_.IsReverse }).Count
                    ReverseZones   = ($ZoneInfo | Where-Object { $_.IsReverse }).Count
                    TotalZones     = $ZoneInfo.Count
                    Zones          = $ZoneInfo
                    Forwarders     = $Forwarders
                    NetworkAdapters = $NetworkAdapterInfo
                    Status         = 'Success'
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
            catch {
                # Step 8 -- Return error object
                [PSCustomObject]@{
                    ComputerName = $computer
                    ServiceStatus = $null
                    ForwardZones = $null
                    ReverseZones = $null
                    TotalZones = $null
                    Zones = $null
                    Forwarders = $null
                    NetworkAdapters = $null
                    Error        = $_.Exception.Message
                    Status       = 'Failed'
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
