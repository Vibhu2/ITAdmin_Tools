# ============================================================
# FUNCTION : Get-VBDhcpInfo
# VERSION  : 1.0.2
# CHANGED  : 10-04-2026 -- VB-compliant refactor and cleanup
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Comprehensive DHCP server analysis and reporting
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Comprehensive DHCP server analysis and reporting tool for IPv4 and IPv6 configurations.

.DESCRIPTION
    Get-VBDhcpInfo provides detailed analysis of DHCP server configurations including scopes,
    reservations, exclusions, server options, and utilization statistics. Supports both local
    and remote DHCP servers with structured object output for further processing.

.PARAMETER ComputerName
    Target DHCP server(s). Defaults to local machine. Accepts pipeline input.
    Supports multiple servers and fully-qualified domain names.

.PARAMETER Credential
    Alternate credentials for remote DHCP server access.

.EXAMPLE
    Get-VBDhcpInfo
    Analyzes the local DHCP server configuration.

.EXAMPLE
    Get-VBDhcpInfo -ComputerName "DHCP-SRV01"
    Analyzes remote DHCP server DHCP-SRV01.

.EXAMPLE
    'DHCP-SRV01', 'DHCP-SRV02' | Get-VBDhcpInfo -Credential (Get-Credential)
    Analyzes multiple DHCP servers using pipeline input.

.OUTPUTS
    [PSCustomObject]: ComputerName, ScanDateTime, IPv4ScopeCount, IPv4Scopes, IPv4Options,
    IPv4Reservations, IPv4Exclusions, TotalIPPool, UsedIPs, UtilizationPercent, IPv6ScopeCount,
    IPv6Scopes, IPv6Options, IPv6Reservations, IPv4DnsSettings, IPv6DnsSettings, Status

.NOTES
    Version  : 1.0.2
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : Services
#>

function Get-VBDhcpInfo {
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

                # Step 2 -- Collect IPv4 scope information
                $IPv4Scopes = @(Get-DhcpServerv4Scope @Params)
                $IPv4Options = Get-DhcpServerv4OptionValue @Params

                $AllReservations = @()
                $TotalAddresses = 0
                $UsedAddresses = 0

                if ($IPv4Scopes.Count -gt 0) {
                    foreach ($scope in $IPv4Scopes) {
                        $ScopeReservations = Get-DhcpServerv4Reservation @Params -ScopeId $scope.ScopeId
                        if ($ScopeReservations) {
                            $AllReservations += $ScopeReservations
                        }

                        # Step 3 -- Calculate scope statistics
                        $ScopeStats = Get-DhcpServerv4ScopeStatistics @Params -ScopeId $scope.ScopeId
                        if ($ScopeStats) {
                            $TotalAddresses += $ScopeStats.AddressesFree + $ScopeStats.AddressesInUse
                            $UsedAddresses += $ScopeStats.AddressesInUse
                        }
                    }
                }

                # Step 4 -- Collect IPv6 scope information
                $IPv6Scopes = @(Get-DhcpServerv6Scope @Params)
                $IPv6Options = if ($IPv6Scopes.Count -gt 0) {
                    Get-DhcpServerv6OptionValue @Params
                }

                $AllIPv6Reservations = @()
                if ($IPv6Scopes.Count -gt 0) {
                    foreach ($scope in $IPv6Scopes) {
                        $ScopeReservations = Get-DhcpServerv6Reservation @Params -ScopeId $scope.ScopeId
                        if ($ScopeReservations) {
                            $AllIPv6Reservations += $ScopeReservations
                        }
                    }
                }

                # Step 5 -- Collect DNS settings
                $IPv4DnsSettings = Get-DhcpServerv4DnsSetting @Params
                $IPv6DnsSettings = Get-DhcpServerv6DnsSetting @Params

                # Step 6 -- Collect exclusion ranges
                $AllExclusions = @()
                foreach ($scope in $IPv4Scopes) {
                    $Exclusions = Get-DhcpServerv4ExclusionRange @Params -ScopeId $scope.ScopeId
                    if ($Exclusions) {
                        $AllExclusions += $Exclusions | ForEach-Object {
                            [PSCustomObject]@{
                                ScopeId    = $scope.ScopeId
                                StartRange = $_.StartRange
                                EndRange   = $_.EndRange
                            }
                        }
                    }
                }

                # Step 7 -- Calculate utilization percentage
                $UtilizationPercent = if ($TotalAddresses -gt 0) {
                    [Math]::Round(($UsedAddresses / $TotalAddresses) * 100, 1)
                } else {
                    0
                }

                # Step 8 -- Return success object
                [PSCustomObject]@{
                    ComputerName         = $computer
                    ScanDateTime         = Get-Date
                    IPv4ScopeCount       = $IPv4Scopes.Count
                    IPv4Scopes           = $IPv4Scopes
                    IPv4Options          = $IPv4Options
                    IPv4Reservations     = $AllReservations
                    IPv4ReservationCount = $AllReservations.Count
                    IPv4Exclusions       = $AllExclusions
                    IPv4ExclusionCount   = $AllExclusions.Count
                    TotalIPPool          = $TotalAddresses
                    UsedIPs              = $UsedAddresses
                    UtilizationPercent   = $UtilizationPercent
                    IPv6ScopeCount       = $IPv6Scopes.Count
                    IPv6Scopes           = $IPv6Scopes
                    IPv6Options          = $IPv6Options
                    IPv6Reservations     = $AllIPv6Reservations
                    IPv6ReservationCount = $AllIPv6Reservations.Count
                    IPv4DnsSettings      = $IPv4DnsSettings
                    IPv6DnsSettings      = $IPv6DnsSettings
                    Status               = 'Success'
                    CollectionTime       = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
            catch {
                # Step 9 -- Return error object
                [PSCustomObject]@{
                    ComputerName = $computer
                    ScanDateTime = Get-Date
                    Error        = $_.Exception.Message
                    Status       = 'Failed'
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
