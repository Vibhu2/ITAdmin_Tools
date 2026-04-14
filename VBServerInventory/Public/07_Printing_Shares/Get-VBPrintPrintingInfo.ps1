# ============================================================
# FUNCTION : Get-VBPrintPrintingInfo
# VERSION  : 1.0.0
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu
# PURPOSE  : Comprehensive print job analysis and printer monitoring
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Comprehensive print job analysis and printer monitoring on local and remote systems.

.DESCRIPTION
    Retrieves detailed print job history, printer status, and usage statistics from Windows Print Service event logs.
    Supports multiple analysis modes: Jobs, Printers, Stats, Monitor, or All-inclusive reports.
    Data can be exported to Object, Table, or CSV format.

.PARAMETER ComputerName
    Target computer(s). Defaults to local machine. Accepts pipeline input.
    Supports aliases: Name, Server.

.PARAMETER Credential
    Alternate credentials for remote execution.

.PARAMETER Mode
    Analysis mode: Jobs, Stats, Printers, Monitor, All.
    Default: Jobs.

.PARAMETER Days
    Look-back period in days. Range: 1-365.
    Default: 7.

.PARAMETER MaxEvents
    Maximum events to retrieve. Range: 100-10000.
    Default: 1000.

.PARAMETER OutputFormat
    Output format: Object, Table, CSV.
    Default: Object.

.PARAMETER OutputPath
    File path for CSV export. Auto-generated if not specified with CSV format.

.EXAMPLE
    Get-VBPrintPrintingInfo

.EXAMPLE
    Get-VBPrintPrintingInfo -ComputerName SERVER01 -Mode Jobs

.EXAMPLE
    'SRV01','SRV02' | Get-VBPrintPrintingInfo -Mode All -OutputFormat CSV -OutputPath C:\Reports\PrintReport.csv

.OUTPUTS
    [PSCustomObject]: ComputerName, TimeCreated, EventID, Source, PrinterName, UserName, ClientMachine, DocumentName, PagesPrinted, JobSize, RawMessage, Status, CollectionTime

.NOTES
    Version  : 1.0.0
    Author   : Vibhu
    Modified : 10-04-2026
    Category : Printing
#>

function Get-VBPrintPrintingInfo {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'Server')]
        [string[]]$ComputerName = $env:COMPUTERNAME,

        [PSCredential]$Credential,

        [ValidateSet('Jobs', 'Stats', 'Printers', 'Monitor', 'All')]
        [string]$Mode = 'Jobs',

        [ValidateRange(1, 365)]
        [int]$Days = 7,

        [ValidateRange(100, 10000)]
        [int]$MaxEvents = 1000,

        [ValidateSet('Object', 'Table', 'CSV')]
        [string]$OutputFormat = 'Object',

        [string]$OutputPath
    )

    begin {
        $allResults = @()

        # Step 1 -- Generate output path for CSV if needed
        if (-not $OutputPath -and ($OutputFormat -eq 'CSV')) {
            $timestamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
            $OutputPath = "PrintReport_$timestamp.csv"
        }

        # Step 2 -- Enable PrintService event log if disabled
        try {
            $logStatus = wevtutil get-log Microsoft-Windows-PrintService/Admin
            if ($logStatus -match 'enabled:\s*false') {
                Write-Verbose 'Enabling Microsoft-Windows-PrintService/Admin log'
                wevtutil set-log Microsoft-Windows-PrintService/Admin /enabled:true
            }
            $testEvents = Get-WinEvent -LogName 'Microsoft-Windows-PrintService/Admin' -MaxEvents 1 -ErrorAction SilentlyContinue
            if (-not $testEvents) {
                Write-Verbose 'No events found in Microsoft-Windows-PrintService/Admin log.'
            }
        }
        catch {
            Write-Verbose "Failed to check or enable PrintService log: $($_.Exception.Message)"
        }
    }

    process {
        foreach ($computer in $ComputerName) {
            if (-not $computer) {
                Write-Verbose 'ComputerName is null or empty. Skipping.'
                continue
            }

            try {
                Write-Verbose "Processing $computer with Mode: $Mode"

                # Step 3 -- Define nested helper functions and main logic
                $scriptBlock = {
                    param($Mode, $Days, $MaxEvents, $ComputerTarget)

                    $startTime = (Get-Date).AddDays(-$Days)

                    # Nested helper: Get print job events
                    function Get-PrintJobs {
                        $printEvents = @()

                        try {
                            $filterHash = @{
                                LogName   = 'Microsoft-Windows-PrintService/Admin'
                                StartTime = $startTime
                                ID        = 307
                            }
                            $printEvents = Get-WinEvent -FilterHashtable $filterHash -MaxEvents $MaxEvents -ErrorAction Stop
                            Write-Verbose "Found $($printEvents.Count) print job events."
                        }
                        catch {
                            Write-Verbose "PrintService log not accessible: $($_.Exception.Message)"
                        }

                        foreach ($event in $printEvents) {
                            $eventData = @{}
                            try {
                                $eventXml = [xml]$event.ToXml()
                                if ($eventXml.Event.EventData.Data) {
                                    for ($i = 0; $i -lt $eventXml.Event.EventData.Data.Count; $i++) {
                                        $eventData["Param$($i+1)"] = $eventXml.Event.EventData.Data[$i].'#text'
                                    }
                                }
                            }
                            catch {
                                Write-Verbose "XML parsing failed: $($_.Exception.Message)"
                            }

                            $message = $event.Message
                            $printer = if ($eventData.Param1 -and $eventData.Param1.Trim()) { $eventData.Param1.Trim() } else { 'Unknown' }
                            if ($printer -match "^(.+?)(?:\.|:\s+)(.+?)\.\s*$") {
                                $printer = $matches[2].Trim()
                            }
                            elseif ($message -match "printer[:\s]+([^\r\n,]+)") {
                                $printer = $matches[1].Trim()
                            }

                            $user = if ($eventData.Param2 -and $eventData.Param2.Trim()) { $eventData.Param2.Trim() } else { 'Unknown' }
                            if ($message -match "user[:\s]+([^\r\n,]+)") {
                                $user = $matches[1].Trim()
                            }

                            $document = if ($eventData.Param3 -and $eventData.Param3.Trim()) { $eventData.Param3.Trim() } else { 'Unknown' }
                            if ($message -match "document[:\s]+([^\r\n,]+)") {
                                $document = $matches[1].Trim()
                            }

                            $client = if ($eventData.Param4 -and $eventData.Param4.Trim()) { $eventData.Param4.Trim() } else { 'Unknown' }
                            if ($message -match "(?:client|computer)[:\s]+([^\r\n,.]+)") {
                                $client = $matches[1].Trim()
                            }

                            $pages = 0
                            if ($eventData.Param5 -and $eventData.Param5 -match '^\d+$') {
                                $pages = [int]$eventData.Param5
                            }
                            elseif ($message -match "(?:pages?|pages printed)[:\s]*(\d+)") {
                                $pages = [int]$matches[1]
                            }

                            $size = 0
                            if ($eventData.Param6 -and $eventData.Param6 -match '^\d+$') {
                                $size = [int]$eventData.Param6
                            }
                            elseif ($message -match "size[:\s]*(\d+)") {
                                $size = [int]$matches[1]
                            }

                            [PSCustomObject]@{
                                ComputerName  = $ComputerTarget
                                TimeCreated   = $event.TimeCreated
                                EventID       = $event.Id
                                Source        = $event.LogName
                                PrinterName   = $printer
                                UserName      = $user
                                ClientMachine = $client
                                DocumentName  = $document
                                PagesPrinted  = $pages
                                JobSize       = $size
                                RawMessage    = $message
                                Status        = 'Success'
                            }
                        }

                        try {
                            $activeJobs = Get-CimInstance -ClassName Win32_PrintJob -ErrorAction Stop |
                                Where-Object { $_.TimeSubmitted -ge $startTime }
                            foreach ($job in $activeJobs) {
                                $printerName = if ($job.Name) { ($job.Name -split ',')[0].Trim() } else { 'Unknown' }
                                $owner = if ($job.Owner) { $job.Owner } else { 'Unknown' }
                                $documentName = if ($job.Document) { $job.Document } else { 'Unknown' }
                                $totalPages = if ($job.TotalPages) { $job.TotalPages } else { 0 }
                                $clientMachine = if ($job.HostComputerName) { $job.HostComputerName -replace '.*\\', '' } else { 'Unknown' }

                                [PSCustomObject]@{
                                    ComputerName  = $ComputerTarget
                                    TimeCreated   = $job.TimeSubmitted
                                    EventID       = 0
                                    Source        = 'ActiveJob'
                                    PrinterName   = $printerName
                                    UserName      = $owner
                                    ClientMachine = $clientMachine
                                    DocumentName  = $documentName
                                    PagesPrinted  = $totalPages
                                    JobSize       = if ($job.Size) { $job.Size } else { 0 }
                                    RawMessage    = "Active job: $($job.Status)"
                                    Status        = 'Active'
                                }
                            }
                        }
                        catch {
                            Write-Verbose "Could not retrieve active print jobs: $($_.Exception.Message)"
                        }
                    }

                    # Nested helper: Get printer information
                    function Get-PrinterInfo {
                        try {
                            $printers = Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop
                            $printJobs = Get-PrintJobs

                            foreach ($printer in $printers) {
                                $printerJobs = $printJobs | Where-Object { $_.PrinterName -like "*$($printer.Name)*" }
                                $totalJobs = ($printerJobs | Measure-Object).Count
                                $totalPages = ($printerJobs | Measure-Object PagesPrinted -Sum).Sum

                                $status = switch ($printer.PrinterStatus) {
                                    1 { 'Other' }
                                    2 { 'Unknown' }
                                    3 { 'Idle' }
                                    4 { 'Printing' }
                                    5 { 'Warmup' }
                                    6 { 'Stopped Printing' }
                                    7 { 'Offline' }
                                    default { 'Unknown' }
                                }

                                [PSCustomObject]@{
                                    ComputerName = $ComputerTarget
                                    PrinterName  = $printer.Name
                                    Status       = $status
                                    Location     = if ($printer.Location) { $printer.Location } else { 'Not Set' }
                                    DriverName   = if ($printer.DriverName) { $printer.DriverName } else { 'Unknown' }
                                    PortName     = if ($printer.PortName) { $printer.PortName } else { 'Unknown' }
                                    Shared       = $printer.Shared
                                    RecentJobs   = $totalJobs
                                    RecentPages  = $totalPages
                                    QueuedJobs   = if ($printer.JobCount) { $printer.JobCount } else { 0 }
                                    LastUsed     = if ($printerJobs) { ($printerJobs | Sort-Object TimeCreated -Descending | Select-Object -First 1).TimeCreated } else { 'Never' }
                                }
                            }
                        }
                        catch {
                            [PSCustomObject]@{
                                ComputerName = $ComputerTarget
                                PrinterName  = 'Error'
                                Status       = 'Failed'
                                Error        = $_.Exception.Message
                            }
                        }
                    }

                    # Nested helper: Get print statistics
                    function Get-PrintStats {
                        $printJobs = Get-PrintJobs | Where-Object { $_.PagesPrinted -gt 0 }
                        if (-not $printJobs) {
                            Write-Verbose 'No print jobs with pages printed found for stats.'
                            return
                        }

                        $stats = $printJobs |
                            Group-Object UserName |
                            ForEach-Object {
                                $totalPages = ($_.Group | Measure-Object PagesPrinted -Sum).Sum
                                $totalJobs = $_.Count
                                $averagePages = if ($totalJobs -gt 0) { [math]::Round($totalPages / $totalJobs, 2) } else { 0 }

                                [PSCustomObject]@{
                                    ComputerName   = $ComputerTarget
                                    UserName       = $_.Name
                                    TotalJobs      = $totalJobs
                                    TotalPages     = $totalPages
                                    AveragePages   = $averagePages
                                    AnalysisPeriod = "$Days days"
                                }
                            }

                        return $stats | Sort-Object TotalPages -Descending
                    }

                    # Step 4 -- Execute mode-specific logic
                    switch ($Mode) {
                        'Jobs' { Get-PrintJobs }
                        'Printers' { Get-PrinterInfo }
                        'Stats' { Get-PrintStats }
                        'All' {
                            @{
                                Jobs     = Get-PrintJobs
                                Printers = Get-PrinterInfo
                                Stats    = Get-PrintStats
                            }
                        }
                        'Monitor' {
                            Get-PrintJobs | Where-Object { $_.TimeCreated -gt (Get-Date).AddMinutes(-5) }
                        }
                    }
                }

                # Step 5 -- Execute script block locally or remotely
                if ($computer -eq $env:COMPUTERNAME) {
                    $result = & $scriptBlock $Mode $Days $MaxEvents $computer
                }
                else {
                    $params = @{
                        ComputerName = $computer
                        ScriptBlock  = $scriptBlock
                        ArgumentList = $Mode, $Days, $MaxEvents, $computer
                    }
                    if ($Credential) { $params.Credential = $Credential }
                    $result = Invoke-Command @params
                }

                # Step 6 -- Handle Monitor mode output
                if ($Mode -eq 'Monitor') {
                    if ($result) {
                        Write-Verbose "Recent print activity on $computer"
                        $result | Format-Table TimeCreated, UserName, PrinterName, DocumentName, PagesPrinted -AutoSize
                    }
                    else {
                        Write-Verbose "No recent print activity on $computer"
                    }
                    continue
                }

                # Step 7 -- Accumulate results
                $allResults += $result
            }
            catch {
                $errorResult = [PSCustomObject]@{
                    ComputerName   = $computer
                    Error          = $_.Exception.Message
                    Status         = 'Failed'
                    TimeCreated    = Get-Date
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
                $allResults += $errorResult
            }
        }
    }

    end {
        # Step 8 -- Handle Monitor mode termination
        if ($Mode -eq 'Monitor') {
            return
        }

        # Step 9 -- Validate results
        if (-not $allResults) {
            Write-Verbose 'No data retrieved from any computer. Verify print services and event logs.'
        }

        # Step 10 -- Flatten results from All mode
        if ($Mode -eq 'All') {
            $consolidatedResults = @()
            foreach ($result in $allResults) {
                if ($result -is [Hashtable]) {
                    $consolidatedResults += $result.Jobs
                    $consolidatedResults += $result.Printers
                    $consolidatedResults += $result.Stats
                }
                else {
                    $consolidatedResults += $result
                }
            }
            $allResults = $consolidatedResults
        }

        # Step 11 -- Output in requested format
        switch ($OutputFormat) {
            'Object' {
                $allResults
            }
            'Table' {
                $allResults | Format-Table -AutoSize | Out-String
            }
            'CSV' {
                $allResults | Export-Csv -Path $OutputPath -NoTypeInformation
                Write-Verbose "Results exported to: $OutputPath"
            }
        }
    }
}
