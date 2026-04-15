# ============================================================
# FUNCTION : Get-VBDHCPInformation
# VERSION  : 1.0.0
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Collect comprehensive DHCP server configuration
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Collects comprehensive DHCP server configuration including scopes, reservations, and DNS settings.

.DESCRIPTION
    Get-VBDHCPInformation gathers detailed DHCP server configuration data for both IPv4 and IPv6 from local or remote servers.
    Returns structured objects containing scope information, server options, reservations, and DNS settings.

.PARAMETER ComputerName
    Target DHCP server(s). Defaults to local machine. Accepts pipeline input.
    Supports multiple servers and fully-qualified domain names.

.PARAMETER Credential
    Alternate credentials for remote DHCP server access. Required for servers in different domains.

.EXAMPLE
    Get-VBDHCPInformation
    Retrieves DHCP configuration from the local server.

.EXAMPLE
    Get-VBDHCPInformation -ComputerName DHCP-SRV01
    Retrieves DHCP configuration from remote server DHCP-SRV01.

.EXAMPLE
    'DHCP-SRV01','DHCP-SRV02' | Get-VBDHCPInformation -Credential (Get-Credential)
    Retrieves DHCP configuration from multiple servers using pipeline input.

.OUTPUTS
    [PSCustomObject]: One object per server with properties ComputerName, IPv4Scopes, IPv4Options,
    IPv4Reservations, IPv6Scopes, IPv6Options, IPv6Reservations, IPv4DnsSettings, IPv6DnsSettings, Status, CollectionTime

.NOTES
    Version  : 1.0.0
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : Services
#>

function Get-VBDHCPInformation {
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
                # Step 1 -- Build parameter hashtable for remote commands
                $InvokeParams = @{
                    ComputerName = $computer
                    ErrorAction  = 'SilentlyContinue'
                }
                if ($Credential) {
                    $InvokeParams.Credential = $Credential
                }

                # Step 2 -- Determine if local or remote execution
                if ($computer -eq $env:COMPUTERNAME) {
                    # Step 3 -- Collect IPv4 DHCP information
                    $IPv4Scopes = Get-DhcpServerv4Scope @InvokeParams
                    $IPv4Options = Get-DhcpServerv4OptionValue @InvokeParams
                    $IPv4Reservations = @()
                    if ($IPv4Scopes) {
                        foreach ($scope in $IPv4Scopes) {
                            $IPv4Reservations += Get-DhcpServerv4Reservation @InvokeParams -ScopeId $scope.ScopeId
                        }
                    }

                    # Step 4 -- Collect IPv6 DHCP information
                    $IPv6Scopes = Get-DhcpServerv6Scope @InvokeParams
                    $IPv6Options = $null
                    $IPv6Reservations = @()
                    if ($IPv6Scopes) {
                        $IPv6Options = Get-DhcpServerv6OptionValue @InvokeParams
                        foreach ($scope in $IPv6Scopes) {
                            $IPv6Reservations += Get-DhcpServerv6Reservation @InvokeParams -ScopeId $scope.ScopeId
                        }
                    }

                    # Step 5 -- Collect DNS settings
                    $IPv4DnsSettings = Get-DhcpServerv4DnsSetting @InvokeParams
                    $IPv6DnsSettings = Get-DhcpServerv6DnsSetting @InvokeParams
                } else {
                    # Step 6 -- Remote execution via Invoke-Command
                    $RemoteResult = Invoke-Command -ComputerName $computer -Credential $Credential -ScriptBlock {
                        $IPv4Scopes = Get-DhcpServerv4Scope -ErrorAction SilentlyContinue
                        $IPv4Options = Get-DhcpServerv4OptionValue -ErrorAction SilentlyContinue
                        $IPv4Reservations = @()
                        if ($IPv4Scopes) {
                            foreach ($scope in $IPv4Scopes) {
                                $IPv4Reservations += Get-DhcpServerv4Reservation -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue
                            }
                        }

                        $IPv6Scopes = Get-DhcpServerv6Scope -ErrorAction SilentlyContinue
                        $IPv6Options = $null
                        $IPv6Reservations = @()
                        if ($IPv6Scopes) {
                            $IPv6Options = Get-DhcpServerv6OptionValue -ErrorAction SilentlyContinue
                            foreach ($scope in $IPv6Scopes) {
                                $IPv6Reservations += Get-DhcpServerv6Reservation -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue
                            }
                        }

                        $IPv4DnsSettings = Get-DhcpServerv4DnsSetting -ErrorAction SilentlyContinue
                        $IPv6DnsSettings = Get-DhcpServerv6DnsSetting -ErrorAction SilentlyContinue

                        [PSCustomObject]@{
                            IPv4Scopes      = $IPv4Scopes
                            IPv4Options     = $IPv4Options
                            IPv4Reservations = $IPv4Reservations
                            IPv6Scopes      = $IPv6Scopes
                            IPv6Options     = $IPv6Options
                            IPv6Reservations = $IPv6Reservations
                            IPv4DnsSettings = $IPv4DnsSettings
                            IPv6DnsSettings = $IPv6DnsSettings
                        }
                    }
                    $IPv4Scopes = $RemoteResult.IPv4Scopes
                    $IPv4Options = $RemoteResult.IPv4Options
                    $IPv4Reservations = $RemoteResult.IPv4Reservations
                    $IPv6Scopes = $RemoteResult.IPv6Scopes
                    $IPv6Options = $RemoteResult.IPv6Options
                    $IPv6Reservations = $RemoteResult.IPv6Reservations
                    $IPv4DnsSettings = $RemoteResult.IPv4DnsSettings
                    $IPv6DnsSettings = $RemoteResult.IPv6DnsSettings
                }

                # Step 7 -- Return success object
                [PSCustomObject]@{
                    ComputerName      = $computer
                    IPv4Scopes        = $IPv4Scopes
                    IPv4Options       = $IPv4Options
                    IPv4Reservations  = $IPv4Reservations
                    IPv6Scopes        = $IPv6Scopes
                    IPv6Options       = $IPv6Options
                    IPv6Reservations  = $IPv6Reservations
                    IPv4DnsSettings   = $IPv4DnsSettings
                    IPv6DnsSettings   = $IPv6DnsSettings
                    Status            = 'Success'
                    CollectionTime    = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
            catch {
                # Step 8 -- Return error object
                [PSCustomObject]@{
                    ComputerName   = $computer
                    IPv4Scopes     = $null
                    IPv4Options    = $null
                    IPv4Reservations = $null
                    IPv6Scopes     = $null
                    IPv6Options    = $null
                    IPv6Reservations = $null
                    IPv4DnsSettings = $null
                    IPv6DnsSettings = $null
                    Error          = $_.Exception.Message
                    Status         = 'Failed'
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
