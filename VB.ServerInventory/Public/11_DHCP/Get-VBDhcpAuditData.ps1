# ============================================================
# FUNCTION : Get-VBDhcpAuditData
# VERSION  : 1.4.0
# CHANGED  : 08-06-2026 -- Cached authorized servers (1 AD query total);
#                          cached admin group membership (1 query total);
#                          bulk scope statistics fetch (1 call per server);
#                          added per-collector Verbose timing; pre-computed
#                          ScopeId strings in loops; fixed NeedScopes logic
# AUTHOR   : Vibhu
# PURPOSE  : Query one or more DHCP servers and stream audit rows
# ENCODING : UTF-8 with BOM
# ------------------------------------------------------------
# CHANGELOG (last 3-5 only -- full history in Git)
# v1.4.0 -- 08-06-2026 -- Bulk caching: authorized servers, admin members, scope stats; timing
# v1.3.0 -- 08-06-2026 -- Non-mandatory ComputerName; single scope fetch; order fix; $RowError
# v1.2.0 -- 08-06-2026 -- Bug fixes: ServerName empty, IPAddress wrong source, null CimSession splat
# v1.1.0 -- 08-06-2026 -- VB noun prefix; output standard compliance
# v1.0.0 -- 08-06-2026 -- Initial release
# ============================================================

$ErrorActionPreference = 'Stop'

#region --- PRIVATE COLLECTORS ---

function Invoke-VBDhcpServerCollector {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [string]     $ComputerName,
        [CimSession] $CimSession,
        [object[]]   $AuthorizedServers
    )
    process {
        try {
            Write-Verbose "[$ComputerName] Collecting server info"

            $cimSplat = if ($null -ne $CimSession) { @{ CimSession = $CimSession } } else { @{ ComputerName = $ComputerName } }

            $os  = Get-CimInstance -ClassName Win32_OperatingSystem @cimSplat
            $svc = Get-CimInstance -ClassName Win32_Service -Filter "Name='DHCPServer'" @cimSplat

            $ipAddress = $null
            try {
                $nic       = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=True' @cimSplat |
                             Select-Object -First 1
                $ipAddress = if ($null -ne $nic) { ($nic.IPAddress | Where-Object { $_ -match '\.' } | Select-Object -First 1) } else { $null }
            }
            catch {
                Write-Warning "[$ComputerName] Could not query network adapter IP: $($_.Exception.Message)"
            }

            # Use pre-cached authorized server list -- no AD query here
            $authorized = ($AuthorizedServers | Where-Object { $_.DnsName -eq $ComputerName -or $_.IPAddress -eq $ComputerName }).Count -gt 0

            $dbState = $null
            try {
                $db      = Get-DhcpServerDatabase -ComputerName $ComputerName -ErrorAction SilentlyContinue
                $dbState = if ($null -ne $db) { $db.FileName } else { $null }
            }
            catch {
                Write-Warning "[$ComputerName] Could not query DHCP database info: $($_.Exception.Message)"
            }

            $auditEnabled = $false
            try {
                $audit        = Get-DhcpServerAuditLog -ComputerName $ComputerName -ErrorAction SilentlyContinue
                $auditEnabled = if ($null -ne $audit) { $audit.Enable } else { $false }
            }
            catch {
                Write-Warning "[$ComputerName] Could not query audit log settings: $($_.Exception.Message)"
            }

            New-VBDhcpRow -Type 'Server' -ServerName $ComputerName -Status 'Success' -RowError $null -Override @{
                IPAddress        = $ipAddress
                OS               = if ($null -ne $os) { $os.Caption } else { $null }
                OSVersion        = if ($null -ne $os) { $os.Version } else { $null }
                Authorized       = $authorized
                ServiceState     = if ($null -ne $svc) { $svc.State } else { $null }
                DatabaseState    = $dbState
                AuditLogsEnabled = $auditEnabled
            }
        }
        catch {
            Write-Warning "[$ComputerName] Server collector failed: $($_.Exception.Message)"
            New-VBDhcpRow -Type 'Server' -ServerName $ComputerName -Status 'Failed' -RowError $_.Exception.Message -Override @{}
        }
    }
}

function Invoke-VBDhcpScopeCollector {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [string]     $ComputerName,
        [CimSession] $CimSession,
        [object[]]   $Scopes,
        [hashtable]  $StatsLookup,
        [hashtable]  $VlanMap
    )
    process {
        foreach ($scope in $Scopes) {
            $scopeIdStr = $scope.ScopeId.ToString()
            try {
                $stats = $StatsLookup[$scopeIdStr]
                $vlan  = Resolve-VBDhcpVlan -ScopeId $scopeIdStr -ScopeName $scope.Name -Description $scope.Description -VlanMap $VlanMap

                New-VBDhcpRow -Type 'Scope' -ServerName $ComputerName -Status 'Success' -RowError $null -Override @{
                    ScopeID        = $scopeIdStr
                    ScopeName      = $scope.Name
                    SubnetMask     = $scope.SubnetMask.ToString()
                    VLAN           = $vlan
                    RangeStart     = $scope.StartRange.ToString()
                    RangeEnd       = $scope.EndRange.ToString()
                    LeaseDuration  = $scope.LeaseDuration
                    ScopeState     = $scope.State.ToString()
                    TotalAddresses = if ($null -ne $stats) { [int]$stats.TotalAddresses } else { $null }
                    InUse          = if ($null -ne $stats) { [int]$stats.InUse } else { $null }
                    Free           = if ($null -ne $stats) { [int]$stats.Free } else { $null }
                    UtilizationPct = if ($null -ne $stats) { [math]::Round([double]$stats.PercentageInUse, 2) } else { $null }
                }
            }
            catch {
                Write-Warning "[$ComputerName] Scope $scopeIdStr failed: $($_.Exception.Message)"
                New-VBDhcpRow -Type 'Scope' -ServerName $ComputerName -Status 'Failed' -RowError $_.Exception.Message -Override @{
                    ScopeID = $scopeIdStr
                }
            }
        }
    }
}

function Invoke-VBDhcpExclusionCollector {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [string]     $ComputerName,
        [CimSession] $CimSession,
        [object[]]   $Scopes
    )
    process {
        foreach ($scope in $Scopes) {
            $scopeIdStr = $scope.ScopeId.ToString()
            try {
                $exclusions = Get-DhcpServerv4ExclusionRange -ComputerName $ComputerName -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue
                foreach ($excl in $exclusions) {
                    New-VBDhcpRow -Type 'Exclusion' -ServerName $ComputerName -Status 'Success' -RowError $null -Override @{
                        ScopeID        = $scopeIdStr
                        ExclusionStart = $excl.StartRange.ToString()
                        ExclusionEnd   = $excl.EndRange.ToString()
                    }
                }
            }
            catch {
                Write-Warning "[$ComputerName] Exclusion query failed for scope $scopeIdStr`: $($_.Exception.Message)"
                New-VBDhcpRow -Type 'Exclusion' -ServerName $ComputerName -Status 'Failed' -RowError $_.Exception.Message -Override @{
                    ScopeID = $scopeIdStr
                }
            }
        }
    }
}

function Invoke-VBDhcpReservationCollector {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [string]     $ComputerName,
        [CimSession] $CimSession,
        [object[]]   $Scopes
    )
    process {
        foreach ($scope in $Scopes) {
            $scopeIdStr = $scope.ScopeId.ToString()
            try {
                $reservations = Get-DhcpServerv4Reservation -ComputerName $ComputerName -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue
                foreach ($res in $reservations) {
                    New-VBDhcpRow -Type 'Reservation' -ServerName $ComputerName -Status 'Success' -RowError $null -Override @{
                        ScopeID                = $scopeIdStr
                        ReservationName        = $res.Name
                        IPAddress              = $res.IPAddress.ToString()
                        ReservationMAC         = $res.ClientId
                        ReservationDescription = $res.Description
                    }
                }
            }
            catch {
                Write-Warning "[$ComputerName] Reservation query failed for scope $scopeIdStr`: $($_.Exception.Message)"
                New-VBDhcpRow -Type 'Reservation' -ServerName $ComputerName -Status 'Failed' -RowError $_.Exception.Message -Override @{
                    ScopeID = $scopeIdStr
                }
            }
        }
    }
}

function Invoke-VBDhcpFailoverCollector {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [string]     $ComputerName,
        [CimSession] $CimSession
    )
    process {
        try {
            Write-Verbose "[$ComputerName] Collecting failover relationships"

            $failovers = Get-DhcpServerv4Failover -ComputerName $ComputerName -ErrorAction Stop

            foreach ($fo in $failovers) {
                foreach ($scopeId in $fo.ScopeId) {
                    New-VBDhcpRow -Type 'Failover' -ServerName $ComputerName -Status 'Success' -RowError $null -Override @{
                        ScopeID       = $scopeId.ToString()
                        PartnerServer = $fo.PartnerServer
                        FailoverMode  = $fo.Mode.ToString()
                        FailoverState = $fo.State.ToString()
                    }
                }
            }
        }
        catch {
            Write-Warning "[$ComputerName] Failover collector failed: $($_.Exception.Message)"
            New-VBDhcpRow -Type 'Failover' -ServerName $ComputerName -Status 'Failed' -RowError $_.Exception.Message -Override @{}
        }
    }
}

function Invoke-VBDhcpDnsConfigCollector {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [string]     $ComputerName,
        [CimSession] $CimSession
    )
    process {
        try {
            Write-Verbose "[$ComputerName] Collecting DNS config"

            $dns = Get-DhcpServerv4DnsSetting -ComputerName $ComputerName -ErrorAction Stop

            $credConfigured = $false
            try {
                $dnsCred        = Get-DhcpServerDnsCredential -ComputerName $ComputerName -ErrorAction SilentlyContinue
                $credConfigured = $null -ne $dnsCred
            }
            catch {
                Write-Warning "[$ComputerName] Could not query DNS credential config: $($_.Exception.Message)"
            }

            New-VBDhcpRow -Type 'DNSConfig' -ServerName $ComputerName -Status 'Success' -RowError $null -Override @{
                DynamicUpdates        = $dns.DynamicUpdates.ToString()
                SecureUpdates         = [bool]($dns.DynamicUpdates -eq 'OnClientRequest' -or $dns.UpdateDnsRRForOlderClients -eq $false)
                PTRRegistration       = [bool]$dns.DeleteDnsRROnLeaseExpiry
                CredentialsConfigured = $credConfigured
            }
        }
        catch {
            Write-Warning "[$ComputerName] DNS config collector failed: $($_.Exception.Message)"
            New-VBDhcpRow -Type 'DNSConfig' -ServerName $ComputerName -Status 'Failed' -RowError $_.Exception.Message -Override @{}
        }
    }
}

function Invoke-VBDhcpOptionCollector {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [string]     $ComputerName,
        [CimSession] $CimSession,
        [object[]]   $Scopes
    )
    process {
        try {
            Write-Verbose "[$ComputerName] Collecting server-level options"

            $serverOptions = Get-DhcpServerv4OptionValue -ComputerName $ComputerName -ErrorAction SilentlyContinue
            foreach ($opt in $serverOptions) {
                New-VBDhcpRow -Type 'ServerOption' -ServerName $ComputerName -Status 'Success' -RowError $null -Override @{
                    OptionCode  = $opt.OptionId
                    OptionName  = $opt.Name
                    OptionValue = ($opt.Value -join ', ')
                    OptionScope = 'Server'
                    ScopeID     = $null
                }
            }
        }
        catch {
            Write-Warning "[$ComputerName] Server option collection failed: $($_.Exception.Message)"
            New-VBDhcpRow -Type 'ServerOption' -ServerName $ComputerName -Status 'Failed' -RowError $_.Exception.Message -Override @{}
        }

        Write-Verbose "[$ComputerName] Collecting scope-level options"

        foreach ($scope in $Scopes) {
            $scopeIdStr = $scope.ScopeId.ToString()
            try {
                $scopeOptions = Get-DhcpServerv4OptionValue -ComputerName $ComputerName -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue
                foreach ($opt in $scopeOptions) {
                    New-VBDhcpRow -Type 'ScopeOption' -ServerName $ComputerName -Status 'Success' -RowError $null -Override @{
                        OptionCode  = $opt.OptionId
                        OptionName  = $opt.Name
                        OptionValue = ($opt.Value -join ', ')
                        OptionScope = 'Scope'
                        ScopeID     = $scopeIdStr
                    }
                }
            }
            catch {
                Write-Warning "[$ComputerName] Scope option query failed for scope $scopeIdStr`: $($_.Exception.Message)"
                New-VBDhcpRow -Type 'ScopeOption' -ServerName $ComputerName -Status 'Failed' -RowError $_.Exception.Message -Override @{
                    ScopeID = $scopeIdStr
                }
            }
        }
    }
}

function Invoke-VBDhcpAdminGroupCollector {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [string]     $ComputerName,
        [CimSession] $CimSession,
        [object[]]   $AdminMembers,
        [string]     $AdminMembersError
    )
    process {
        $groupName = 'DHCP Administrators'

        if (-not [string]::IsNullOrWhiteSpace($AdminMembersError)) {
            Write-Warning "[$ComputerName] Admin group not available: $AdminMembersError"
            New-VBDhcpRow -Type 'AdminGroup' -ServerName $ComputerName -Status 'Failed' -RowError $AdminMembersError -Override @{
                AdminGroup  = $groupName
                MemberCount = $null
            }
            return
        }

        New-VBDhcpRow -Type 'AdminGroup' -ServerName $ComputerName -Status 'Success' -RowError $null -Override @{
            AdminGroup  = $groupName
            MemberCount = @($AdminMembers).Count
        }
    }
}

#endregion

#region --- PRIVATE HELPERS ---

function New-VBDhcpRow {
    # Builds a consistent flat PSCustomObject; Override hashtable sets domain columns
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [string]    $Type,
        [string]    $ServerName,
        [string]    $Status,
        [string]    $RowError,
        [hashtable] $Override = @{}
    )
    process {
        $row = [ordered]@{
            ServerName             = $ServerName
            Status                 = $Status
            Type                   = $Type
            IPAddress              = $null
            OS                     = $null
            OSVersion              = $null
            Authorized             = $null
            ServiceState           = $null
            DatabaseState          = $null
            AuditLogsEnabled       = $null
            ScopeID                = $null
            ScopeName              = $null
            SubnetMask             = $null
            VLAN                   = $null
            RangeStart             = $null
            RangeEnd               = $null
            LeaseDuration          = $null
            ScopeState             = $null
            TotalAddresses         = $null
            InUse                  = $null
            Free                   = $null
            UtilizationPct         = $null
            ExclusionStart         = $null
            ExclusionEnd           = $null
            ReservationName        = $null
            ReservationMAC         = $null
            ReservationDescription = $null
            PartnerServer          = $null
            FailoverMode           = $null
            FailoverState          = $null
            DynamicUpdates         = $null
            SecureUpdates          = $null
            PTRRegistration        = $null
            CredentialsConfigured  = $null
            OptionCode             = $null
            OptionName             = $null
            OptionValue            = $null
            OptionScope            = $null
            AdminGroup             = $null
            MemberCount            = $null
            Error                  = $RowError
            CollectionTime         = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
        }

        foreach ($key in $Override.Keys) {
            $row[$key] = $Override[$key]
        }

        [PSCustomObject]$row
    }
}

function Resolve-VBDhcpVlan {
    # Returns VLAN string from map lookup or regex; $null if unresolvable
    [CmdletBinding()]
    param(
        [string]    $ScopeId,
        [string]    $ScopeName,
        [string]    $Description,
        [hashtable] $VlanMap
    )
    process {
        # Step 1 -- map lookup
        if ($null -ne $VlanMap -and $VlanMap.ContainsKey($ScopeId)) {
            return $VlanMap[$ScopeId]
        }

        # Step 2 -- regex on Name then Description
        $candidates = @($ScopeName, $Description) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        foreach ($text in $candidates) {
            if ($text -match '(?i)vlan\s*[-=]?\s*(\d+)') { return $Matches[1] }
            if ($text -match '(?i)vl(\d+)')               { return $Matches[1] }
            if ($text -match '^(\d+)-')                    { return $Matches[1] }
        }

        return $null
    }
}

function Resolve-VBDhcpCollectorNames {
    # Returns collector function names in fixed execution order, deduplicated
    [CmdletBinding()]
    param(
        [string[]] $Type
    )
    process {
        # Ordered so output type sequence is always deterministic
        $orderedMap = [ordered]@{
            Server       = 'Invoke-VBDhcpServerCollector'
            Scope        = 'Invoke-VBDhcpScopeCollector'
            Exclusion    = 'Invoke-VBDhcpExclusionCollector'
            Reservation  = 'Invoke-VBDhcpReservationCollector'
            Failover     = 'Invoke-VBDhcpFailoverCollector'
            DNSConfig    = 'Invoke-VBDhcpDnsConfigCollector'
            ServerOption = 'Invoke-VBDhcpOptionCollector'
            ScopeOption  = 'Invoke-VBDhcpOptionCollector'
            AdminGroup   = 'Invoke-VBDhcpAdminGroupCollector'
        }

        $runAll   = $Type -contains 'All'
        $selected = [System.Collections.Generic.List[string]]::new()

        foreach ($key in $orderedMap.Keys) {
            if ($runAll -or $Type -contains $key) {
                $name = $orderedMap[$key]
                if (-not $selected.Contains($name)) { $selected.Add($name) }
            }
        }

        return $selected
    }
}

#endregion

#region --- PUBLIC ORCHESTRATOR ---

function Get-VBDhcpAuditData {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string[]] $ComputerName = $env:COMPUTERNAME,

        [Parameter()]
        [System.Management.Automation.CredentialAttribute()]
        [PSCredential] $Credential = [PSCredential]::Empty,

        [Parameter()]
        [ValidateSet('All', 'Server', 'Scope', 'Exclusion', 'Reservation',
                     'Failover', 'DNSConfig', 'ServerOption', 'ScopeOption', 'AdminGroup')]
        [string[]] $Type = 'All',

        [Parameter()]
        [string] $VlanMapPath
    )

    begin {
        # Step 1 -- verify DhcpServer module is available
        if (-not (Get-Module -Name DhcpServer -ListAvailable -ErrorAction SilentlyContinue)) {
            Write-Error 'DhcpServer module (RSAT) is not installed. Install RSAT-DHCP before running this function.'
            return
        }

        # Step 2 -- load VLAN map if path supplied
        $VlanMap = $null
        if (-not [string]::IsNullOrWhiteSpace($VlanMapPath)) {
            if (Test-Path -LiteralPath $VlanMapPath) {
                try {
                    $VlanMap = @{}
                    Import-Csv -LiteralPath $VlanMapPath | ForEach-Object {
                        $VlanMap[$_.Subnet] = $_.VLAN
                    }
                    Write-Verbose "Loaded VLAN map: $($VlanMap.Count) entries from $VlanMapPath"
                }
                catch {
                    Write-Warning "Failed to load VlanMapPath '$VlanMapPath': $($_.Exception.Message). Falling back to regex."
                    $VlanMap = $null
                }
            }
            else {
                Write-Warning "VlanMapPath '$VlanMapPath' not found. Falling back to regex VLAN resolution."
            }
        }

        # Step 3 -- resolve which collectors to run
        $CollectorNames = Resolve-VBDhcpCollectorNames -Type $Type
        Write-Verbose "Collectors selected: $($CollectorNames -join ', ')"

        $ScopeConsumers = @(
            'Invoke-VBDhcpScopeCollector'
            'Invoke-VBDhcpExclusionCollector'
            'Invoke-VBDhcpReservationCollector'
            'Invoke-VBDhcpOptionCollector'
        )
        $NeedScopes = @($CollectorNames | Where-Object { $ScopeConsumers -contains $_ }).Count -gt 0

        # Step 4 -- cache authorized DHCP server list once for all servers
        $AuthorizedServers = @()
        if ($CollectorNames -contains 'Invoke-VBDhcpServerCollector') {
            try {
                $AuthorizedServers = @(Get-DhcpServerInDC -ErrorAction SilentlyContinue)
                Write-Verbose "Cached $($AuthorizedServers.Count) authorized DHCP server(s) from AD"
            }
            catch {
                Write-Warning "Could not query authorized DHCP servers from AD: $($_.Exception.Message)"
            }
        }

        # Step 5 -- cache DHCP Administrators group membership once for all servers
        $AdminMembers      = @()
        $AdminMembersError = $null
        if ($CollectorNames -contains 'Invoke-VBDhcpAdminGroupCollector') {
            $groupName = 'DHCP Administrators'
            if (Get-Module -Name ActiveDirectory -ListAvailable -ErrorAction SilentlyContinue) {
                try {
                    $AdminMembers = @(Get-ADGroupMember -Identity $groupName -Recursive -ErrorAction Stop)
                    Write-Verbose "Cached $($AdminMembers.Count) DHCP admin group member(s) from AD"
                }
                catch {
                    Write-Warning "AD group query failed, falling back to local: $($_.Exception.Message)"
                }
            }

            if ($AdminMembers.Count -eq 0 -and $null -eq $AdminMembersError) {
                try {
                    $AdminMembers = @(Get-LocalGroupMember -Group $groupName -ErrorAction Stop)
                    Write-Verbose "Cached $($AdminMembers.Count) DHCP admin group member(s) from local group"
                }
                catch {
                    $AdminMembersError = $_.Exception.Message
                    Write-Warning "Local group query also failed: $AdminMembersError"
                }
            }
        }
    }

    process {
        foreach ($computer in $ComputerName) {
            Write-Verbose "[$computer] Starting audit"

            # Step 6 -- establish CimSession if credential supplied
            $cimSession = $null
            if ($Credential -ne [PSCredential]::Empty) {
                try {
                    $cimSession = New-CimSession -ComputerName $computer -Credential $Credential -ErrorAction Stop
                    Write-Verbose "[$computer] CimSession established"
                }
                catch {
                    Write-Warning "[$computer] CimSession failed: $($_.Exception.Message). Skipping server."
                    continue
                }
            }

            # Step 7 -- fetch scopes once for all collectors that need them
            $scopes = @()
            if ($NeedScopes) {
                try {
                    $scopes = @(Get-DhcpServerv4Scope -ComputerName $computer -ErrorAction Stop)
                    Write-Verbose "[$computer] Retrieved $($scopes.Count) scope(s)"
                }
                catch {
                    Write-Warning "[$computer] Could not retrieve scopes: $($_.Exception.Message)"
                }
            }

            # Step 8 -- fetch all scope statistics in one call and build lookup
            $StatsLookup = @{}
            if ($NeedScopes -and $scopes.Count -gt 0 -and $CollectorNames -contains 'Invoke-VBDhcpScopeCollector') {
                try {
                    $allStats = @(Get-DhcpServerv4ScopeStatistics -ComputerName $computer -ErrorAction Stop)
                    foreach ($stat in $allStats) {
                        $StatsLookup[$stat.ScopeId.ToString()] = $stat
                    }
                    Write-Verbose "[$computer] Cached statistics for $($allStats.Count) scope(s)"
                }
                catch {
                    Write-Warning "[$computer] Could not retrieve scope statistics: $($_.Exception.Message)"
                }
            }

            # Step 9 -- run each selected collector and stream results
            $baseSplat = @{ ComputerName = $computer }
            if ($null -ne $cimSession) { $baseSplat['CimSession'] = $cimSession }

            $i = 0
            foreach ($collectorName in $CollectorNames) {
                $i++
                Write-Progress -Activity "[$computer] Collecting DHCP audit data" -Status $collectorName -PercentComplete (($i / $CollectorNames.Count) * 100)

                $extraSplat = @{}
                if ($ScopeConsumers -contains $collectorName)              { $extraSplat['Scopes']            = $scopes            }
                if ($collectorName -eq 'Invoke-VBDhcpScopeCollector')      { $extraSplat['StatsLookup']       = $StatsLookup       }
                if ($collectorName -eq 'Invoke-VBDhcpScopeCollector')      { $extraSplat['VlanMap']           = $VlanMap           }
                if ($collectorName -eq 'Invoke-VBDhcpServerCollector')     { $extraSplat['AuthorizedServers'] = $AuthorizedServers  }
                if ($collectorName -eq 'Invoke-VBDhcpAdminGroupCollector') { $extraSplat['AdminMembers']      = $AdminMembers       }
                if ($collectorName -eq 'Invoke-VBDhcpAdminGroupCollector') { $extraSplat['AdminMembersError'] = $AdminMembersError  }

                try {
                    $sw = [System.Diagnostics.Stopwatch]::StartNew()
                    & $collectorName @baseSplat @extraSplat
                    $sw.Stop()
                    Write-Verbose "[$computer] $collectorName completed in $($sw.ElapsedMilliseconds) ms"
                }
                catch {
                    Write-Warning "[$computer] $collectorName threw an unexpected error: $($_.Exception.Message)"
                }
            }

            Write-Progress -Activity "[$computer] Collecting DHCP audit data" -Completed

            # Step 10 -- clean up CimSession
            if ($null -ne $cimSession) {
                Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue
            }
        }
    }
}

#endregion

<#
$DHCPData = Get-VBDhcpAuditData

$finalData = $DHCPData | Select-Object `
    ComputerName, Status, IPAddress, OS, OSVersion, Authorized,
    ServiceState, DatabaseState, AuditLogsEnabled,
    @{N='Scopes';       E={ $_.Scopes       | Where-Object { $_.ScopeState -eq 'Active' } }},
    @{N='Exclusions';   E={ $_.Exclusions   }},
    @{N='Reservations'; E={ $_.Reservations }},
    @{N='Failovers';    E={ $_.Failovers    }}

# Server-level info DHCP
$finalData | Format-List ComputerName, Status, IPAddress, OS, Authorized, ServiceState

# Active scopes DHCP
$finalData.Scopes | Format-Table

# Reservations DHCP
$finalData.Reservations | Format-Table ScopeID, ReservationName, IPAddress, ReservationMAC

# Export everything to CSV DHCP (flattens nested arrays into separate files)
$finalData | Select-Object -ExcludeProperty Scopes, Exclusions, Reservations, Failovers | Export-Csv server.csv
$finalData.Scopes        | Export-Csv scopes.csv
$finalData.Reservations  | Export-Csv reservations.csv
#>
