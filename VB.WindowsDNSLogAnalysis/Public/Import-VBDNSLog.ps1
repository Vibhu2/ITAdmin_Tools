# ============================================================
# FUNCTION : Import-VBDNSLog
# VERSION  : 1.0.0
# CHANGED  : 2026-05-07 -- Initial build
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Parse DNS debug log files and populate the SQLite database
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Parses one or more Windows DNS debug log files and stores all PACKET records
    in the DNSLogDB SQLite database.

.DESCRIPTION
    The main workhorse of the DNSLogDB module. Orchestrates the full import pipeline:

    1. Resolves -InputPath to a list of .log files (single file, wildcard, or directory)
    2. For each file, calls Get-VBImportStatus to check SHA256 hash vs ImportLog
       -- files already imported are skipped unless -Force is specified
    3. Uses ForEach-Object -Parallel (PS7) to process multiple files concurrently
    4. Per thread: Invoke-VBDNSLogParser reads and parses the file into an object[] buffer
    5. Per thread: Invoke-VBBulkInsert writes the buffer to SQLite using a prepared
       statement inside a batched transaction
    6. Returns one summary object per file

    The database must be initialised with Initialize-VBDNSLogDatabase before calling
    this function. Each parallel thread creates its own SQLiteConnection (WAL mode
    supports concurrent writes).

.PARAMETER InputPath
    Path to a single .log file, a directory containing .log files, or a wildcard
    pattern (e.g. 'C:\DNS\Logs\*.log').

.PARAMETER DatabasePath
    Full path to the SQLite .db file. Must be initialised with Initialize-VBDNSLogDatabase.

.PARAMETER Recurse
    Search subdirectories for .log files when -InputPath is a directory.

.PARAMETER ThrottleLimit
    Maximum number of parallel threads for ForEach-Object -Parallel.
    Default: 8. Reduce if memory pressure is observed on large files.

.PARAMETER ExcludePrivateIPs
    Skip PACKET lines where the client IP is classified as private at parse time.
    This is a storage decision -- private IP records are never written to the database.

.PARAMETER Force
    Re-import files even if their SHA256 hash already exists in ImportLog.
    The existing ImportLog row is deleted before re-importing.

.EXAMPLE
    Import-VBDNSLog -InputPath 'C:\DNS\Logs\dns_20260501.log' -DatabasePath 'C:\DNS\dns_analysis.db'

.EXAMPLE
    Import-VBDNSLog -InputPath 'C:\DNS\Logs\' -DatabasePath 'C:\DNS\dns_analysis.db' -ThrottleLimit 4

.EXAMPLE
    Import-VBDNSLog -InputPath 'C:\DNS\Logs\' -DatabasePath 'C:\DNS\dns_analysis.db' -Recurse -ExcludePrivateIPs

.EXAMPLE
    Import-VBDNSLog -InputPath 'C:\DNS\Logs\dns.log' -DatabasePath 'C:\DNS\dns_analysis.db' -Force -Verbose

.OUTPUTS
    [PSCustomObject] One object per processed file:
        FileName      -- File name (leaf only)
        FilePath      -- Full path
        RecordCount   -- Rows inserted into DNSLog
        ErrorCount    -- Lines that failed to parse
        DurationSec   -- Total parse + insert time
        Status        -- 'Success', 'Skipped', or 'Failed'
        Message       -- Additional detail (skip reason or error message)

.NOTES
    Version  : 1.0.0
    Author   : Vibhu Bhatnagar
    Modified : 2026-05-07
    Category : Public

    Requires: PowerShell 7.0+ (ForEach-Object -Parallel)
    Requires: PSSQLite module (Install-Module PSSQLite)
    Run after: Initialize-VBDNSLogDatabase
#>

function Import-VBDNSLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [string]$DatabasePath,

        [switch]$Recurse,

        [int]$ThrottleLimit = 8,

        [switch]$ExcludePrivateIPs,

        [switch]$Force
    )

    # Step 1 -- Validate database exists
    if (-not (Test-Path $DatabasePath)) {
        throw "Import-VBDNSLog: Database not found at '$DatabasePath'. Run Initialize-VBDNSLogDatabase first."
    }

    # Step 2 -- Resolve input path to list of .log files
    $logFiles = @()

    if (Test-Path $InputPath -PathType Container) {
        $gciParams = @{ Path = $InputPath; Filter = '*.log'; File = $true }
        if ($Recurse) { $gciParams.Recurse = $true }
        $logFiles = Get-ChildItem @gciParams | Select-Object -ExpandProperty FullName
    }
    elseif ($InputPath -match '[*?]') {
        $logFiles = Get-ChildItem -Path $InputPath -File | Select-Object -ExpandProperty FullName
    }
    else {
        if (-not (Test-Path $InputPath -PathType Leaf)) {
            throw "Import-VBDNSLog: File not found: '$InputPath'"
        }
        $logFiles = @($InputPath)
    }

    if ($logFiles.Count -eq 0) {
        Write-Warning "Import-VBDNSLog: No .log files found at '$InputPath'"
        return
    }

    Write-Verbose "Import-VBDNSLog: Found $($logFiles.Count) file(s) to process"

    # Step 3 -- Pre-check import status for all files (sequential, before parallel block)
    $fileStatuses = @{}
    foreach ($file in $logFiles) {
        $status = Get-VBImportStatus -FilePath $file -DatabasePath $DatabasePath
        $fileStatuses[$file] = $status

        if ($status.AlreadyImported -and -not $Force) {
            Write-Verbose "Import-VBDNSLog: Skipping '$file' (already imported on $($status.ImportedAt))"
        }
        elseif ($status.AlreadyImported -and $Force) {
            Invoke-SqliteQuery -DataSource $DatabasePath `
                -Query "DELETE FROM ImportLog WHERE FileHash = @hash" `
                -SqlParameters @{ hash = $status.FileHash }
            Write-Verbose "Import-VBDNSLog: Force re-import -- removed existing ImportLog entry for '$file'"
        }
    }

    # Step 4 -- Filter to files that need processing
    $filesToProcess = $logFiles | Where-Object {
        -not $fileStatuses[$_].AlreadyImported -or $Force
    }

    # Emit skipped file results
    $logFiles | Where-Object {
        $fileStatuses[$_].AlreadyImported -and -not $Force
    } | ForEach-Object {
        [PSCustomObject]@{
            FileName    = Split-Path $_ -Leaf
            FilePath    = $_
            RecordCount = 0
            ErrorCount  = 0
            DurationSec = 0
            Status      = 'Skipped'
            Message     = "Already imported on $($fileStatuses[$_].ImportedAt)"
        }
    }

    if ($filesToProcess.Count -eq 0) {
        Write-Verbose "Import-VBDNSLog: No new files to import."
        return
    }

    Write-Verbose "Import-VBDNSLog: Processing $($filesToProcess.Count) file(s) with ThrottleLimit=$ThrottleLimit"

    # Step 5 -- Parallel parse and insert (one thread per file)
    # PS7 parallel runspaces do not inherit module functions -- pass the Private folder
    # path via $using: and dot-source the required functions inside each thread.
    #
    # Progress strategy: Write-Progress inside ForEach-Object -Parallel is NOT forwarded
    # to the host console. Instead each thread writes counters into a shared
    # ConcurrentDictionary, and the main thread polls that dictionary in a tight loop
    # driving all three progress bars (Read -> Parse -> Insert) itself.
    $excludePrivate = $ExcludePrivateIPs.IsPresent
    $dbPath         = $DatabasePath
    $hashLookup     = $fileStatuses
    $privateDir     = Join-Path $PSScriptRoot '..\Private'

    # One shared state dict per file being processed. Keyed by file name (leaf).
    # Parallel threads write; the polling loop below reads.
    $progressStates = [System.Collections.Concurrent.ConcurrentDictionary[string,
        System.Collections.Concurrent.ConcurrentDictionary[string,object]]]::new()
    foreach ($f in $filesToProcess) {
        $state = [System.Collections.Concurrent.ConcurrentDictionary[string,object]]::new()
        $state['Stage']              = 'Queued'   # Queued | Reading | Parsing | Inserting | Done | Failed
        $state['FileName']           = (Split-Path $f -Leaf)
        $state['FileSizeBytes']      = [long](Get-Item $f).Length
        $state['ParseTotalBytes']    = [long]0
        $state['ParseBytesRead']     = [long]0
        $state['ParseLinesRead']     = [long]0
        $state['ParsePacketCount']   = [long]0
        $state['ParseDone']          = $false
        $state['InsertTotalRows']    = [long]0
        $state['InsertRowsInserted'] = [long]0
        $state['InsertDone']         = $false
        $progressStates[(Split-Path $f -Leaf)] = $state
    }

    # Launch the parallel pipeline as a background job so the main thread is free to poll
    $parallelJob = $filesToProcess | ForEach-Object -AsJob -ThrottleLimit $ThrottleLimit -Parallel {
        $filePath       = $_
        $dbPath         = $using:dbPath
        $excludePrivate = $using:excludePrivate
        $hashLookup     = $using:hashLookup
        $privateDir     = $using:privateDir
        $progressStates = $using:progressStates

        # Dot-source Private functions into this runspace -- required for PS7 parallel threads
        Get-ChildItem -Path $privateDir -Filter '*.ps1' | ForEach-Object { . $_.FullName }

        $fileName    = Split-Path $filePath -Leaf
        $fileHash    = $hashLookup[$filePath].FileHash
        $stopwatch   = [System.Diagnostics.Stopwatch]::StartNew()
        $recordCount = 0
        $errorCount  = 0
        $state       = $progressStates[$fileName]

        try {
            if ($null -ne $state) { $state['Stage'] = 'Parsing' }

            $parseParams = @{
                FilePath      = $filePath
                ProgressState = $state
            }
            if ($excludePrivate) { $parseParams.ExcludePrivateIPs = $true }

            $parseResult = Invoke-VBDNSLogParser @parseParams
            $buffer      = $parseResult.Buffer      # List[object[]] -- never piped through PS pipeline
            $errorCount  = $parseResult.ErrorCount
            $recordCount = $buffer.Count

            if ($recordCount -gt 0) {
                if ($null -ne $state) { $state['Stage'] = 'Inserting' }

                $duration = $stopwatch.Elapsed.TotalSeconds

                Invoke-VBBulkInsert `
                    -DatabasePath    $dbPath `
                    -Buffer          $buffer `
                    -SourceFile      $filePath `
                    -RecordCount     $recordCount `
                    -ErrorCount      $errorCount `
                    -FileHash        $fileHash `
                    -DurationSeconds $duration `
                    -ProgressState   $state | Out-Null
            }

            $stopwatch.Stop()
            if ($null -ne $state) { $state['Stage'] = 'Done' }

            [PSCustomObject]@{
                FileName    = $fileName
                FilePath    = $filePath
                RecordCount = $recordCount
                ErrorCount  = $errorCount
                DurationSec = [math]::Round($stopwatch.Elapsed.TotalSeconds, 1)
                Status      = 'Success'
                Message     = ''
            }
        }
        catch {
            $stopwatch.Stop()
            if ($null -ne $state) { $state['Stage'] = 'Failed' }
            [PSCustomObject]@{
                FileName    = $fileName
                FilePath    = $filePath
                RecordCount = $recordCount
                ErrorCount  = $errorCount
                DurationSec = [math]::Round($stopwatch.Elapsed.TotalSeconds, 1)
                Status      = 'Failed'
                Message     = $_.Exception.Message
            }
        }
    }

    # Step 6 -- Progress polling loop on the main thread.
    # Runs until the background job finishes, updating 3-stage progress bars every 200ms.
    # Bar IDs: 1 = per-file parent, 2 = Parse child, 3 = Insert child.
    $totalFiles = $filesToProcess.Count
    $fileIndex  = 0

    while ($parallelJob.State -in 'NotStarted','Running') {
        $fileIndex = 0
        foreach ($kvp in $progressStates.GetEnumerator()) {
            $s         = $kvp.Value
            $fileName  = $s['FileName']
            $stage     = $s['Stage']
            $fileIndex++

            $fileLabel = "[$fileIndex/$totalFiles] $fileName"

            switch ($stage) {
                'Queued' {
                    Write-Progress -Id 1 -Activity "Import-VBDNSLog" `
                        -Status "Queued: $fileLabel" -PercentComplete 0
                }

                'Parsing' {
                    $totalBytes   = [long]$s['ParseTotalBytes']
                    $bytesRead    = [long]$s['ParseBytesRead']
                    $linesRead    = [long]$s['ParseLinesRead']
                    $packetCount  = [long]$s['ParsePacketCount']
                    $pct          = if ($totalBytes -gt 0) { [math]::Min(99, [math]::Round($bytesRead / $totalBytes * 100)) } else { 0 }

                    Write-Progress -Id 1 -Activity "Import-VBDNSLog" `
                        -Status "Parsing: $fileLabel" -PercentComplete $pct

                    Write-Progress -Id 2 -ParentId 1 -Activity "Parsing log file" `
                        -Status ("Lines: {0:N0}  |  PACKET rows: {1:N0}  |  {2}%" -f $linesRead, $packetCount, $pct) `
                        -PercentComplete $pct

                    Write-Progress -Id 3 -ParentId 1 -Activity "Populating database" `
                        -Status "Waiting for parse to complete..." -PercentComplete 0
                }

                'Inserting' {
                    $totalRows    = [long]$s['InsertTotalRows']
                    $rowsInserted = [long]$s['InsertRowsInserted']
                    $pct          = if ($totalRows -gt 0) { [math]::Min(99, [math]::Round($rowsInserted / $totalRows * 100)) } else { 0 }
                    $packetCount  = [long]$s['ParsePacketCount']

                    Write-Progress -Id 1 -Activity "Import-VBDNSLog" `
                        -Status "Inserting: $fileLabel" -PercentComplete (50 + [math]::Round($pct / 2))

                    Write-Progress -Id 2 -ParentId 1 -Activity "Parsing log file" `
                        -Status ("Complete  |  {0:N0} PACKET rows parsed" -f $packetCount) `
                        -PercentComplete 100

                    Write-Progress -Id 3 -ParentId 1 -Activity "Populating database" `
                        -Status ("{0:N0} of {1:N0} rows inserted  ({2}%)" -f $rowsInserted, $totalRows, $pct) `
                        -PercentComplete $pct
                }

                'Done' {
                    Write-Progress -Id 3 -ParentId 1 -Activity "Populating database" -Completed
                    Write-Progress -Id 2 -ParentId 1 -Activity "Parsing log file"    -Completed
                    Write-Progress -Id 1 -Activity "Import-VBDNSLog" `
                        -Status "Done: $fileLabel" -PercentComplete 100
                }

                'Failed' {
                    Write-Progress -Id 3 -ParentId 1 -Activity "Populating database" -Completed
                    Write-Progress -Id 2 -ParentId 1 -Activity "Parsing log file"    -Completed
                    Write-Progress -Id 1 -Activity "Import-VBDNSLog" `
                        -Status "Failed: $fileLabel" -PercentComplete 100
                }
            }
        }

        Start-Sleep -Milliseconds 200
    }

    # Clear all progress bars
    Write-Progress -Id 3 -Activity "Populating database" -Completed
    Write-Progress -Id 2 -Activity "Parsing log file"    -Completed
    Write-Progress -Id 1 -Activity "Import-VBDNSLog"     -Completed

    # Collect and emit results from the completed job
    Receive-Job -Job $parallelJob
    Remove-Job  -Job $parallelJob
}
