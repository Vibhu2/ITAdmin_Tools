# ============================================================
# FUNCTION : Get-VBFirewallPortRules
# VERSION  : 1.0.0
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu
# PURPOSE  : Retrieve Windows Firewall port rules from target computer
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Retrieve Windows Firewall inbound port rules from target computer(s).

.DESCRIPTION
    Queries firewall rules with port filters. Filters out Microsoft default rules
    by default. Supports filtering by protocol, port, status. Can export to CSV.
    Returns detailed rule information including status and configuration.

.PARAMETER ComputerName
    Target computer(s). Defaults to local machine. Accepts pipeline input.
    Aliases: Name, Server, Host

.PARAMETER Credential
    Alternate credentials for remote execution.

.PARAMETER Protocol
    Filter rules by protocol (TCP, UDP, Any). Case-insensitive.

.PARAMETER Port
    Filter rules by specific port number.

.PARAMETER IncludeBlocked
    Include rules where Action is Block. Default excludes blocked rules.

.PARAMETER IncludeDisabled
    Include disabled rules. Default shows enabled rules only.

.PARAMETER IncludeDefaultRules
    Include Windows/Microsoft default rules. Default shows custom rules only.

.PARAMETER Detailed
    Show additional properties including description, addresses, and authentication.

.PARAMETER ExportCSV
    Export results to CSV file at specified path.

.EXAMPLE
    Get-VBFirewallPortRules
    Retrieves custom inbound firewall rules on local computer.

.EXAMPLE
    Get-VBFirewallPortRules -ComputerName SERVER01 -Protocol TCP
    Retrieves custom TCP firewall rules from SERVER01.

.EXAMPLE
    Get-VBFirewallPortRules -Port 443 -Detailed
    Retrieves rules for port 443 with detailed information.

.EXAMPLE
    'SERVER01', 'SERVER02' | Get-VBFirewallPortRules -ExportCSV 'C:\rules.csv'
    Exports firewall rules to CSV.

.OUTPUTS
    [PSCustomObject]: ComputerName, RuleID, Name, Enabled, Direction, Profile, Action,
                     Protocol, LocalPort, RemotePort, Program, ProgramPath, Status,
                     CollectionTime (plus optional detailed properties)

.NOTES
    Version  : 1.0.0
    Author   : Vibhu
    Modified : 10-04-2026
    Category : Security
#>

function Get-VBFirewallPortRules {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'Server', 'Host')]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [PSCredential]$Credential,
        [ValidateSet('TCP', 'UDP', 'Any', IgnoreCase = $true)]
        [string]$Protocol,
        [string]$Port,
        [switch]$IncludeBlocked,
        [switch]$IncludeDisabled,
        [switch]$IncludeDefaultRules,
        [switch]$Detailed,
        [string]$ExportCSV
    )
    process {
        foreach ($computer in $ComputerName) {
            try {
                # Step 1 -- Define remote script block
                $scriptBlock = {
                    param($protocol, $port, $includeBlocked, $includeDisabled, $includeDefaultRules, $detailed)

                    # Step 1a -- Get base firewall rules
                    $rules = Get-NetFirewallRule -PolicyStore ActiveStore | Where-Object {
                        ($_.Direction -eq 'Inbound') -and
                        ($includeBlocked -or $_.Action -eq 'Allow') -and
                        ($includeDisabled -or $_.Enabled -eq $true) -and
                        ($includeDefaultRules -or -not ($_.Owner -like '*Microsoft*' -or
                            $_.DisplayName -like '*Windows*' -or
                            $_.DisplayGroup -like '*Windows*' -or
                            $_.DisplayGroup -like '*Microsoft*' -or
                            $_.Group -like '*Windows*' -or
                            $_.Group -like '*Microsoft*' -or
                            $_.Group -like '@*' -or
                            $_.DisplayName -like '@*'))
                    }

                    $results = [System.Collections.ArrayList]::new()

                    # Step 1b -- Process each rule
                    foreach ($rule in $rules) {
                        $portFilters = Get-NetFirewallPortFilter -AssociatedNetFirewallRule $rule -ErrorAction SilentlyContinue

                        if (-not $portFilters -and -not $detailed) { continue }

                        $addressFilter = Get-NetFirewallAddressFilter -AssociatedNetFirewallRule $rule -ErrorAction SilentlyContinue
                        $appFilter = Get-NetFirewallApplicationFilter -AssociatedNetFirewallRule $rule -ErrorAction SilentlyContinue
                        $securityFilter = Get-NetFirewallSecurityFilter -AssociatedNetFirewallRule $rule -ErrorAction SilentlyContinue

                        if ($portFilters) {
                            foreach ($filter in $portFilters) {
                                if ($protocol -and $filter.Protocol -ne $protocol) { continue }

                                if ($port) {
                                    if (-not $filter.LocalPort -or
                                        ($filter.LocalPort -ne $port -and
                                         $filter.LocalPort -ne 'Any' -and
                                         $filter.LocalPort -notlike "*,$port,*" -and
                                         $filter.LocalPort -notlike "$port,*" -and
                                         $filter.LocalPort -notlike "*,$port")) {
                                        continue
                                    }
                                }

                                $resultObj = [PSCustomObject]@{
                                    RuleID = $rule.Name
                                    Name = $rule.DisplayName
                                    Enabled = if ($rule.Enabled -eq $true) { 'Yes' } else { 'No' }
                                    Direction = $rule.Direction
                                    Profile = $rule.Profile
                                    Action = $rule.Action
                                    Protocol = if ($filter.Protocol -eq 'Any') { 'Any' } else { $filter.Protocol }
                                    LocalPort = if ($filter.LocalPort -eq 'Any') { 'Any' } else { $filter.LocalPort }
                                    RemotePort = if ($filter.RemotePort -eq 'Any') { 'Any' } else { $filter.RemotePort }
                                    Program = if ($appFilter.Program -eq '*') { 'Any' } else { Split-Path $appFilter.Program -Leaf }
                                    ProgramPath = if ($appFilter.Program -eq '*') { 'Any' } else { $appFilter.Program }
                                }

                                if ($detailed) {
                                    Add-Member -InputObject $resultObj -NotePropertyName 'Description' -NotePropertyValue $rule.Description
                                    Add-Member -InputObject $resultObj -NotePropertyName 'Group' -NotePropertyValue $rule.Group
                                    Add-Member -InputObject $resultObj -NotePropertyName 'LocalAddress' -NotePropertyValue ($addressFilter.LocalAddress -join ', ')
                                    Add-Member -InputObject $resultObj -NotePropertyName 'RemoteAddress' -NotePropertyValue ($addressFilter.RemoteAddress -join ', ')
                                    Add-Member -InputObject $resultObj -NotePropertyName 'Authentication' -NotePropertyValue $securityFilter.Authentication
                                    Add-Member -InputObject $resultObj -NotePropertyName 'Encryption' -NotePropertyValue $securityFilter.Encryption
                                }

                                [void]$results.Add($resultObj)
                            }
                        }
                        elseif ($detailed) {
                            $resultObj = [PSCustomObject]@{
                                RuleID = $rule.Name
                                Name = $rule.DisplayName
                                Enabled = if ($rule.Enabled -eq $true) { 'Yes' } else { 'No' }
                                Direction = $rule.Direction
                                Profile = $rule.Profile
                                Action = $rule.Action
                                Protocol = 'N/A'
                                LocalPort = 'N/A'
                                RemotePort = 'N/A'
                                Program = if ($appFilter.Program -eq '*') { 'Any' } else { Split-Path $appFilter.Program -Leaf }
                                ProgramPath = if ($appFilter.Program -eq '*') { 'Any' } else { $appFilter.Program }
                                Description = $rule.Description
                                Group = $rule.Group
                                LocalAddress = ($addressFilter.LocalAddress -join ', ')
                                RemoteAddress = ($addressFilter.RemoteAddress -join ', ')
                                Authentication = $securityFilter.Authentication
                                Encryption = $securityFilter.Encryption
                            }
                            [void]$results.Add($resultObj)
                        }
                    }

                    return $results | Sort-Object -Property Name
                }

                # Step 2 -- Execute locally or remotely
                if ($computer -eq $env:COMPUTERNAME) {
                    $results = & $scriptBlock -protocol $Protocol -port $Port -includeBlocked $IncludeBlocked -includeDisabled $IncludeDisabled -includeDefaultRules $IncludeDefaultRules -detailed $Detailed
                } else {
                    $splat = @{
                        ComputerName = $computer
                        ScriptBlock = $scriptBlock
                        ArgumentList = @($Protocol, $Port, $IncludeBlocked, $IncludeDisabled, $IncludeDefaultRules, $Detailed)
                    }
                    if ($Credential) {
                        $splat['Credential'] = $Credential
                    }
                    $results = Invoke-Command @splat
                }

                # Step 3 -- Output results with metadata
                foreach ($item in $results) {
                    [PSCustomObject]@{
                        ComputerName = $computer
                        RuleID = $item.RuleID
                        Name = $item.Name
                        Enabled = $item.Enabled
                        Direction = $item.Direction
                        Profile = $item.Profile
                        Action = $item.Action
                        Protocol = $item.Protocol
                        LocalPort = $item.LocalPort
                        RemotePort = $item.RemotePort
                        Program = $item.Program
                        ProgramPath = $item.ProgramPath
                        Status = 'Success'
                        CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    }
                }
            }
            catch {
                [PSCustomObject]@{
                    ComputerName = $computer
                    Status = 'Failed'
                    Error = $_.Exception.Message
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
