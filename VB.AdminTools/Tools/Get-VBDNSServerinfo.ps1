function Get-DNSServerInfo {
    <#
    .SYNOPSIS
        Collects comprehensive DNS server information including configuration, zones, and statistics.
    
    .DESCRIPTION
        This function gathers essential DNS server information including:
        - Server configuration and settings
        - Zone information (forward and reverse lookup zones)
        - Forwarders and root hints
        - Cache statistics
        - Event logs related to DNS
        - Network configuration
        - Performance counters
    
    .PARAMETER ComputerName
        Specifies the DNS server to query. Defaults to local computer.
    
    .PARAMETER IncludeEventLogs
        Include recent DNS-related event log entries.
    
    .PARAMETER IncludePerformanceCounters
        Include DNS performance counter data.
    
    .EXAMPLE
        Get-DNSServerInfo
        
    .EXAMPLE
        Get-DNSServerInfo -ComputerName "DC01.domain.com" -IncludeEventLogs -IncludePerformanceCounters
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ComputerName = $env:COMPUTERNAME,
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeEventLogs,
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludePerformanceCounters
    )
    
    begin {
        # Check if DNS Server module is available
        if (-not (Get-Module -ListAvailable -Name DnsServer)) {
            Write-Warning "DNS Server PowerShell module is not available. Some information may be limited."
            $ModuleAvailable = $false
        } else {
            Import-Module DnsServer -ErrorAction SilentlyContinue
            $ModuleAvailable = $true
        }
        
        # Initialize result object
        $DNSInfo = [PSCustomObject]@{
            ServerName = $ComputerName
            CollectionTime = Get-Date
            ServerConfiguration = $null
            ZoneInformation = @()
            Forwarders = @()
            RootHints = @()
            CacheInfo = $null
            NetworkConfiguration = $null
            EventLogs = @()
            PerformanceCounters = $null
            ServiceStatus = $null
            Errors = @()
        }
    }
    
    process {
        Write-Host "Collecting DNS Server information for: $ComputerName" -ForegroundColor Green
        
        try {
            # Get DNS Service Status
            Write-Host "- Checking DNS service status..." -ForegroundColor Yellow
            $DNSInfo.ServiceStatus = Get-Service -Name DNS -ComputerName $ComputerName -ErrorAction Stop
            
            if ($ModuleAvailable) {
                # Get DNS Server Configuration
                Write-Host "- Gathering server configuration..." -ForegroundColor Yellow
                $DNSInfo.ServerConfiguration = Get-DnsServerSetting -ComputerName $ComputerName -ErrorAction Stop
                
                # Get DNS Zones
                Write-Host "- Collecting zone information..." -ForegroundColor Yellow
                $zones = Get-DnsServerZone -ComputerName $ComputerName -ErrorAction Stop
                foreach ($zone in $zones) {
                    $zoneInfo = [PSCustomObject]@{
                        ZoneName = $zone.ZoneName
                        ZoneType = $zone.ZoneType
                        DynamicUpdate = $zone.DynamicUpdate
                        ReplicationScope = $zone.ReplicationScope
                        DirectoryPartition = $zone.DirectoryPartition
                        IsPaused = $zone.IsPaused
                        IsReadOnly = $zone.IsReadOnly
                        IsReverseLookupZone = $zone.IsReverseLookupZone
                        RecordCount = (Get-DnsServerResourceRecord -ZoneName $zone.ZoneName -ComputerName $ComputerName -ErrorAction SilentlyContinue | Measure-Object).Count
                    }
                    $DNSInfo.ZoneInformation += $zoneInfo
                }
                
                # Get Forwarders
                Write-Host "- Collecting forwarder information..." -ForegroundColor Yellow
                $forwarders = Get-DnsServerForwarder -ComputerName $ComputerName -ErrorAction Stop
                $DNSInfo.Forwarders = $forwarders.IPAddress
                
                # Get Root Hints
                Write-Host "- Collecting root hints..." -ForegroundColor Yellow
                $rootHints = Get-DnsServerRootHint -ComputerName $ComputerName -ErrorAction Stop
                $DNSInfo.RootHints = $rootHints | Select-Object NameServer, IPAddress
                
                # Get Cache Information
                Write-Host "- Gathering cache statistics..." -ForegroundColor Yellow
                $DNSInfo.CacheInfo = Get-DnsServerStatistics -ComputerName $ComputerName -ErrorAction Stop | Select-Object *Cache*, *Memory*
            }
            
            # Get Network Configuration
            Write-Host "- Collecting network configuration..." -ForegroundColor Yellow
            $networkConfig = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $ComputerName -Filter "IPEnabled=True" -ErrorAction Stop
            $DNSInfo.NetworkConfiguration = $networkConfig | Select-Object Description, IPAddress, SubnetMask, DefaultIPGateway, DNSServerSearchOrder, DNSDomain
            
            # Get Event Logs if requested
            if ($IncludeEventLogs) {
                Write-Host "- Collecting DNS event logs..." -ForegroundColor Yellow
                $events = Get-WinEvent -ComputerName $ComputerName -FilterHashtable @{LogName='DNS Server'; StartTime=(Get-Date).AddDays(-7)} -MaxEvents 100 -ErrorAction SilentlyContinue
                if ($events) {
                    $DNSInfo.EventLogs = $events | Select-Object TimeCreated, Id, LevelDisplayName, Message
                }
            }
            
            # Get Performance Counters if requested
            if ($IncludePerformanceCounters) {
                Write-Host "- Collecting performance counters..." -ForegroundColor Yellow
                try {
                    $perfCounters = @(
                        "\DNS\Total Query Received",
                        "\DNS\Total Response Sent",
                        "\DNS\Recursive Queries",
                        "\DNS\Recursive Query Failure",
                        "\DNS\TCP Query Received",
                        "\DNS\UDP Query Received",
                        "\DNS\Secure Update Received",
                        "\DNS\Dynamic Update Received"
                    )
                    
                    $counterData = @{}
                    foreach ($counter in $perfCounters) {
                        try {
                            $value = (Get-Counter -Counter $counter -ComputerName $ComputerName -ErrorAction SilentlyContinue).CounterSamples.CookedValue
                            $counterData[$counter] = $value
                        } catch {
                            $counterData[$counter] = "N/A"
                        }
                    }
                    $DNSInfo.PerformanceCounters = $counterData
                } catch {
                    $DNSInfo.Errors += "Failed to collect performance counters: $($_.Exception.Message)"
                }
            }
            
        } catch {
            $DNSInfo.Errors += "Error collecting DNS information: $($_.Exception.Message)"
            Write-Warning "Error occurred: $($_.Exception.Message)"
        }
    }
    
    end {
        Write-Host "DNS information collection completed." -ForegroundColor Green
        
        # Enhanced formatted output
        Show-DNSServerReport -DNSInfo $DNSInfo
        
        #return $DNSInfo
    }
}

function Show-DNSServerReport {
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$DNSInfo
    )
    
    # Clear screen for better presentation
    Clear-Host
    
    # Header
    #Write-Host "=" * 80 -ForegroundColor Cyan
    Write-Host "DNS SERVER INFORMATION REPORT" -ForegroundColor Cyan
    #Write-Host "=" * 80 -ForegroundColor Cyan
    #Write-Host ""
    
    # Server Overview Section
    Write-Host "📊 SERVER OVERVIEW" -ForegroundColor Green
    Write-Host "-" * 50 -ForegroundColor Gray
    Write-Host "Server Name       : " -NoNewline -ForegroundColor White
    Write-Host $DNSInfo.ServerName -ForegroundColor Yellow
    Write-Host "Collection Time   : " -NoNewline -ForegroundColor White
    Write-Host $DNSInfo.CollectionTime.ToString("yyyy-MM-dd HH:mm:ss") -ForegroundColor Yellow
    Write-Host "Service Status    : " -NoNewline -ForegroundColor White
    $statusColor = if ($DNSInfo.ServiceStatus.Status -eq "Running") { "Green" } else { "Red" }
    Write-Host $DNSInfo.ServiceStatus.Status -ForegroundColor $statusColor
    Write-Host ""
    
    # Server Configuration Section
    if ($DNSInfo.ServerConfiguration) {
        Write-Host "⚙️  SERVER CONFIGURATION" -ForegroundColor Green
        Write-Host "-" * 50 -ForegroundColor Gray
        Write-Host "Listening Port    : " -NoNewline -ForegroundColor White
        Write-Host $DNSInfo.ServerConfiguration.ListeningIPAddress -ForegroundColor Yellow
        Write-Host "Boot Method       : " -NoNewline -ForegroundColor White
        Write-Host $DNSInfo.ServerConfiguration.BootMethod -ForegroundColor Yellow
        Write-Host "Recursion         : " -NoNewline -ForegroundColor White
        Write-Host $DNSInfo.ServerConfiguration.DisableRecursion -ForegroundColor Yellow
        Write-Host ""
    }
    
    # DNS Zones Section
    if ($DNSInfo.ZoneInformation.Count -gt 0) {
        Write-Host "🌐 DNS ZONES ($($DNSInfo.ZoneInformation.Count) zones)" -ForegroundColor Green
        Write-Host "-" * 50 -ForegroundColor Gray
        
        # Forward Lookup Zones
        $forwardZones = $DNSInfo.ZoneInformation | Where-Object { -not $_.IsReverseLookupZone }
        if ($forwardZones) {
            Write-Host "Forward Lookup Zones:" -ForegroundColor Cyan
            $forwardZones | Format-Table -Property @(
                @{Name="Zone Name"; Expression={$_.ZoneName}; Width=25},
                @{Name="Type"; Expression={$_.ZoneType}; Width=10},
                @{Name="Dynamic Update"; Expression={$_.DynamicUpdate}; Width=15},
                @{Name="Records"; Expression={$_.RecordCount}; Width=8}
            ) -AutoSize | Out-String | ForEach-Object { Write-Host $_ -ForegroundColor White }
        }
        
        # Reverse Lookup Zones
        $reverseZones = $DNSInfo.ZoneInformation | Where-Object { $_.IsReverseLookupZone }
        if ($reverseZones) {
            Write-Host "Reverse Lookup Zones:" -ForegroundColor Cyan
            $reverseZones | Format-Table -Property @(
                @{Name="Zone Name"; Expression={$_.ZoneName}; Width=25},
                @{Name="Type"; Expression={$_.ZoneType}; Width=10},
                @{Name="Dynamic Update"; Expression={$_.DynamicUpdate}; Width=15},
                @{Name="Records"; Expression={$_.RecordCount}; Width=8}
            ) -AutoSize | Out-String | ForEach-Object { Write-Host $_ -ForegroundColor White }
        }
    }
    
    # Forwarders Section
    if ($DNSInfo.Forwarders.Count -gt 0) {
        Write-Host "🔄 DNS FORWARDERS" -ForegroundColor Green
        Write-Host "-" * 50 -ForegroundColor Gray
        foreach ($forwarder in $DNSInfo.Forwarders) {
            Write-Host "  • $forwarder" -ForegroundColor Yellow
        }
        Write-Host ""
    }
    
    # Network Configuration Section
    if ($DNSInfo.NetworkConfiguration) {
        Write-Host "🌐 NETWORK CONFIGURATION" -ForegroundColor Green
        Write-Host "-" * 50 -ForegroundColor Gray
        foreach ($adapter in $DNSInfo.NetworkConfiguration) {
            Write-Host "Adapter: " -NoNewline -ForegroundColor White
            Write-Host $adapter.Description -ForegroundColor Yellow
            if ($adapter.IPAddress) {
                Write-Host "  IP Address(es): " -NoNewline -ForegroundColor White
                Write-Host ($adapter.IPAddress -join ", ") -ForegroundColor Cyan
            }
            if ($adapter.DNSServerSearchOrder) {
                Write-Host "  DNS Servers   : " -NoNewline -ForegroundColor White
                Write-Host ($adapter.DNSServerSearchOrder -join ", ") -ForegroundColor Cyan
            }
            Write-Host ""
        }
    }
    
    # Cache Information Section
    if ($DNSInfo.CacheInfo) {
        Write-Host "💾 CACHE STATISTICS" -ForegroundColor Green
        Write-Host "-" * 50 -ForegroundColor Gray
        if ($DNSInfo.CacheInfo.CacheStatistics) {
            Write-Host "Cache Entries     : " -NoNewline -ForegroundColor White
            Write-Host $DNSInfo.CacheInfo.CacheStatistics.CacheNodes -ForegroundColor Yellow
            Write-Host "Cache Hit Ratio   : " -NoNewline -ForegroundColor White
            Write-Host "$([math]::Round($DNSInfo.CacheInfo.CacheStatistics.CacheHitRatio * 100, 2))%" -ForegroundColor Yellow
        }
        if ($DNSInfo.CacheInfo.MemoryStatistics) {
            Write-Host "Memory Used       : " -NoNewline -ForegroundColor White
            Write-Host "$([math]::Round($DNSInfo.CacheInfo.MemoryStatistics.MemoryUsed / 1MB, 2)) MB" -ForegroundColor Yellow
        }
        Write-Host ""
    }
    
    # Performance Counters Section
    if ($DNSInfo.PerformanceCounters) {
        Write-Host "📈 PERFORMANCE COUNTERS" -ForegroundColor Green
        Write-Host "-" * 50 -ForegroundColor Gray
        foreach ($counter in $DNSInfo.PerformanceCounters.GetEnumerator()) {
            $counterName = $counter.Key -replace "\\DNS\\", ""
            Write-Host "  $counterName" -NoNewline -ForegroundColor White
            Write-Host " : " -NoNewline -ForegroundColor Gray
            Write-Host $counter.Value -ForegroundColor Yellow
        }
        Write-Host ""
    }
    
    # Event Logs Section
    if ($DNSInfo.EventLogs.Count -gt 0) {
        Write-Host "📋 RECENT DNS EVENTS" -ForegroundColor Green
        Write-Host "-" * 50 -ForegroundColor Gray
        $DNSInfo.EventLogs | Select-Object -First 10 | Format-Table -Property @(
            @{Name="Time"; Expression={$_.TimeCreated.ToString("MM/dd HH:mm")}; Width=12},
            @{Name="Level"; Expression={$_.LevelDisplayName}; Width=10},
            @{Name="Event ID"; Expression={$_.Id}; Width=8},
            @{Name="Message"; Expression={$_.Message.Substring(0, [Math]::Min(60, $_.Message.Length)) + "..."}; Width=60}
        ) -AutoSize | Out-String | ForEach-Object { Write-Host $_ -ForegroundColor White }
    }
    
    # Errors Section
    if ($DNSInfo.Errors.Count -gt 0) {
        Write-Host "⚠️  ERRORS ENCOUNTERED" -ForegroundColor Red
        Write-Host "-" * 50 -ForegroundColor Gray
        foreach ($error in $DNSInfo.Errors) {
            Write-Host "  • $error" -ForegroundColor Red
        }
        Write-Host ""
    }
    
    # Footer
    #Write-Host "=" * 80 -ForegroundColor Cyan
    Write-Host "Report generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
    #Write-Host "=" * 80 -ForegroundColor Cyan
}

# Example usage and helper function to export results
function Export-DNSServerInfo {
    <#
    .SYNOPSIS
        Exports DNS server information to various formats.
    
    .PARAMETER DNSInfo
        The DNS information object from Get-DNSServerInfo.
    
    .PARAMETER Path
        Output path for the export files.
    
    .PARAMETER Format
        Export format: JSON, XML, or CSV.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [PSCustomObject]$DNSInfo,
        
        [Parameter(Mandatory = $false)]
        [string]$Path = "C:\DNSReports",
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("JSON", "XML", "CSV")]
        [string]$Format = "JSON"
    )
    
    if (-not (Test-Path $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $filename = "$Path\DNSServerInfo_$($DNSInfo.ServerName)_$timestamp"
    
    switch ($Format) {
        "JSON" {
            $DNSInfo | ConvertTo-Json -Depth 10 | Out-File "$filename.json"
            Write-Host "DNS information exported to: $filename.json" -ForegroundColor Green
        }
        "XML" {
            $DNSInfo | Export-Clixml "$filename.xml"
            Write-Host "DNS information exported to: $filename.xml" -ForegroundColor Green
        }
        "CSV" {
            # Export zones to CSV
            $DNSInfo.ZoneInformation | Export-Csv "$filename`_Zones.csv" -NoTypeInformation
            Write-Host "DNS zone information exported to: $filename`_Zones.csv" -ForegroundColor Green
        }
    }
}

# No HTML support needed

# Usage Examples:
<#
# Basic usage with enhanced formatting
$dnsInfo = Get-DNSServerInfo

# Comprehensive collection with all options
$dnsInfo = Get-DNSServerInfo -ComputerName "DC01.domain.com" -IncludeEventLogs -IncludePerformanceCounters

# Export the results
$dnsInfo | Export-DNSServerInfo -Path "C:\Reports" -Format "JSON"

# Display specific information in formatted tables
Write-Host "Zone Details:" -ForegroundColor Green
$dnsInfo.ZoneInformation | Format-Table -AutoSize

Write-Host "Server Configuration:" -ForegroundColor Green  
$dnsInfo.ServerConfiguration | Format-List
#>