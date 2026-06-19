function Test-VBADReplication {
    <#
    .SYNOPSIS
    Tests Active Directory replication status between all domain controllers with detailed console output.

    .DESCRIPTION
    This function provides a comprehensive test of Active Directory replication health by examining 
    replication metadata between all domain controllers in the environment. It displays real-time 
    progress information, color-coded status messages, and detailed summary statistics on the console.
    
    The function automatically discovers all domain controllers and tests bi-directional replication 
    paths between each pair. It provides immediate visual feedback during execution and returns 
    detailed replication objects for further analysis. The console output includes execution timing, 
    computer information, and formatted result tables.

    .PARAMETER WarningThresholdHours
    Specifies the number of hours after which replication delay triggers a warning status.
    Default value is 2 hours. Replication older than this threshold will be marked as Warning
    and displayed in yellow on the console output.

    .EXAMPLE
    Test-VBADReplication
    
    Tests replication between all discovered domain controllers using the default 2-hour warning 
    threshold. Displays progress information, colored status indicators, and a summary table.

    .EXAMPLE
    Test-VBADReplication -WarningThresholdHours 4
    
    Tests AD replication with a 4-hour warning threshold instead of the default 2 hours. 
    Replication delays over 4 hours will be marked as warnings.

    .EXAMPLE
    $results = Test-VBADReplication
    $failedReplications = $results | Where-Object { $_.Status -like "*Error*" -or $_.Status -eq "Failing" }
    
    Captures the detailed replication results and filters for failed replications to investigate issues.

    .OUTPUTS
    PSCustomObject
    Returns an array of replication status objects with the following properties:
    - SourceServer: Source domain controller name
    - DestinationServer: Target domain controller name  
    - LastReplicationAttempt: Timestamp of last replication attempt
    - LastReplicationSuccess: Timestamp of last successful replication
    - TimeSinceLastSuccess: Human-readable time since last success (e.g., "1d 3h 25m")
    - ConsecutiveFailures: Number of consecutive replication failures
    - PartnerType: Type of replication partner relationship
    - Status: Overall status (Healthy, Warning, Failing, Never Replicated, Error)
    - FailingDSAs: List of failing directory service agents
    - ScheduledSync: Whether replication is scheduled
    - Writable: Whether the partition is writable
    - LastChangeUSN: Last change update sequence number

    .NOTES
    Version: 1.0
    Author: System Administrator
    Category: Active Directory Administration
    
    Requirements:
    - Active Directory PowerShell module (automatically imported if available)
    - Domain administrator or equivalent permissions
    - Network connectivity to all domain controllers
    - RSAT tools if running from workstation
    
    The function provides extensive console output including:
    - Real-time progress indicators
    - Color-coded status messages (Green=Healthy, Yellow=Warning, Red=Failed)
    - Summary statistics with connection counts
    - Formatted result table with key replication metrics
    - Execution timing and computer information
    
    Console output is designed for interactive use and troubleshooting scenarios.
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [int]$WarningThresholdHours = 2
    )

    # Clear the screen for better visibility
    Clear-Host
    
    # Get computer name information
    $ComputerName = $env:COMPUTERNAME
    $CurrentUser = $env:USERNAME
    $CurrentDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    Write-Host "=== Active Directory Replication Status Tester ===" -ForegroundColor Cyan
    Write-Host "Running on computer: $ComputerName" -ForegroundColor Cyan
    Write-Host "Executed by user: $CurrentUser" -ForegroundColor Cyan
    Write-Host "Started at: $CurrentDate" -ForegroundColor Cyan
    Write-Host "--------------------------------------------------" -ForegroundColor Cyan
    
    # Ensure the Active Directory module is loaded
    if (-not (Get-Module -Name ActiveDirectory)) {
        try {
            Import-Module ActiveDirectory -ErrorAction Stop
            Write-Host "Successfully loaded Active Directory module." -ForegroundColor Green
        }
        catch {
            Write-Host "ERROR: Failed to load Active Directory module. This script requires the AD PowerShell module." -ForegroundColor Red
            Write-Host "Please install RSAT tools or run this on a domain controller." -ForegroundColor Red
            return
        }
    }
    
    try {
        # Get all domain controllers in the environment
        Write-Host "Retrieving domain controllers..." -ForegroundColor Yellow
        $DomainControllers = Get-ADDomainController -Filter * | 
                            Select-Object -ExpandProperty Name | 
                            Sort-Object
        
        if (-not $DomainControllers -or $DomainControllers.Count -eq 0) {
            Write-Host "ERROR: No domain controllers found!" -ForegroundColor Red
            return
        }
        
        # Display the list of domain controllers being evaluated
        Write-Host "Found $($DomainControllers.Count) domain controllers:" -ForegroundColor Green
        $DomainControllers | ForEach-Object { Write-Host " - $_" -ForegroundColor Green }
        Write-Host "--------------------------------------------------" -ForegroundColor Cyan
        
        # Prepare an array to hold all the output objects
        $ReplicationResults = @()
        $FailureCount = 0
        $WarningCount = 0
        $SuccessCount = 0
        
        # Loop through each domain controller to set it as the source (EnumerationServer)
        foreach ($SourceDC in $DomainControllers) {
            # Loop through each domain controller to set it as the target
            foreach ($TargetDC in $DomainControllers) {
                # Skip if SourceDC and TargetDC are the same
                if ($SourceDC -ne $TargetDC) {
                    Write-Host "Testing replication from $SourceDC to $TargetDC..." -NoNewline
                    
                    try {
                        # Get replication metadata with the specified source and target
                        $ReplicationMetadata = Get-ADReplicationPartnerMetadata -EnumerationServer $SourceDC `
                                                -Target $TargetDC -Scope Server -Partition * -ErrorAction Stop |
                                                Select-Object -First 1
                        
                        # Calculate time difference
                        $TimeDifference = $null
                        $Status = "Unknown"
                        $StatusColor = "Gray"
                        
                        if ($ReplicationMetadata.LastReplicationSuccess) {
                            $TimeDifference = (Get-Date) - $ReplicationMetadata.LastReplicationSuccess
                            
                            if ($ReplicationMetadata.ConsecutiveReplicationFailures -gt 0) {
                                $Status = "Failing"
                                $StatusColor = "Red"
                                $FailureCount++
                            }
                            elseif ($TimeDifference.TotalHours -gt $WarningThresholdHours) {
                                $Status = "Warning"
                                $StatusColor = "Yellow"
                                $WarningCount++
                            }
                            else {
                                $Status = "Healthy"
                                $StatusColor = "Green"
                                $SuccessCount++
                            }
                        }
                        else {
                            $Status = "Never Replicated"
                            $StatusColor = "Red"
                            $FailureCount++
                        }
                        Write-Host " $Status" -ForegroundColor $StatusColor
                        # Create a custom object with the replication metadata
                        $ReplicationResult = [PSCustomObject]@{
                            "SourceServer" = $SourceDC
                            "DestinationServer" = $TargetDC
                            "LastReplicationAttempt" = $ReplicationMetadata.LastReplicationAttempt
                            "LastReplicationSuccess" = $ReplicationMetadata.LastReplicationSuccess
                            "TimeSinceLastSuccess" = if ($TimeDifference) { 
                                "{0}d {1}h {2}m" -f $TimeDifference.Days, $TimeDifference.Hours, $TimeDifference.Minutes 
                            } else { "N/A" }
                            "ConsecutiveFailures" = $ReplicationMetadata.ConsecutiveReplicationFailures
                            "PartnerType" = $ReplicationMetadata.PartnerType
                            "Status" = $Status
                            "FailingDSAs" = $ReplicationMetadata.FailingSyncPartners
                            "ScheduledSync" = $ReplicationMetadata.ScheduledSync
                            "Writable" = $ReplicationMetadata.Writable
                            "LastChangeUSN" = $ReplicationMetadata.LastChangeUsn
                        }
                        
                        # Add the result to the results array
                        $ReplicationResults += $ReplicationResult
                    }
                    catch {
                        Write-Host " Error!" -ForegroundColor Red
                        Write-Host "   $_" -ForegroundColor Red
                        
                        # Add error result
                        $ReplicationResults += [PSCustomObject]@{
                            "SourceServer" = $SourceDC
                            "DestinationServer" = $TargetDC
                            "LastReplicationAttempt" = $null
                            "LastReplicationSuccess" = $null
                            "TimeSinceLastSuccess" = "N/A"
                            "ConsecutiveFailures" = "N/A"
                            "PartnerType" = "N/A"
                            "Status" = "Error: $($_.Exception.Message)"
                            "FailingDSAs" = "N/A"
                            "ScheduledSync" = $null
                            "Writable" = $null
                            "LastChangeUSN" = $null
                        }
                        $FailureCount++
                    }
                }
            }
        }
        
        # Output summary
        Write-Host "--------------------------------------------------" -ForegroundColor Cyan
        Write-Host "Replication Status Summary:" -ForegroundColor Cyan
        Write-Host "  Healthy Connections: $SuccessCount" -ForegroundColor Green
        Write-Host "  Warning Connections: $WarningCount" -ForegroundColor Yellow
        Write-Host "  Failed Connections:  $FailureCount" -ForegroundColor Red
        Write-Host "--------------------------------------------------" -ForegroundColor Cyan
        
        # Format the results as a table with most important properties
        $ReplicationResults | 
            Format-Table -Property @{
                Label="Source"; Expression={$_.SourceServer}; Width=15
            }, @{
                Label="Destination"; Expression={$_.DestinationServer}; Width=15
            }, @{
                Label="Last Success"; Expression={$_.LastReplicationSuccess}; Width=20
            }, @{
                Label="Time Since"; Expression={$_.TimeSinceLastSuccess}; Width=12
            }, @{
                Label="Failures"; Expression={$_.ConsecutiveFailures}; Width=8
            }, @{
                Label="Status"; Expression={$_.Status}; Width=15
            } -AutoSize
        
        # Display completion message with computer information
        Write-Host "--------------------------------------------------" -ForegroundColor Cyan
        Write-Host "Test completed on $ComputerName at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
        Write-Host "Total execution time: $([math]::Round(((Get-Date) - [datetime]::Parse($CurrentDate)).TotalSeconds, 2)) seconds" -ForegroundColor Cyan
        
        # Return the results for further processing if needed
        return $ReplicationResults
    }
    catch {
        Write-Host "ERROR: An unexpected error occurred:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}