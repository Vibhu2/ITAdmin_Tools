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
    $excludePrivate = $ExcludePrivateIPs.IsPresent
    $dbPath         = $DatabasePath
    $hashLookup     = $fileStatuses
    $privateDir     = Join-Path $PSScriptRoot '..\Private'

    $filesToProcess | ForEach-Object -Parallel {
        $filePath       = $_
        $dbPath         = $using:dbPath
        $excludePrivate = $using:excludePrivate
        $hashLookup     = $using:hashLookup
        $privateDir     = $using:privateDir

        # Dot-source Private functions into this runspace -- required for PS7 parallel threads
        Get-ChildItem -Path $privateDir -Filter '*.ps1' | ForEach-Object { . $_.FullName }

        $fileName    = Split-Path $filePath -Leaf
        $fileHash    = $hashLookup[$filePath].FileHash
        $stopwatch   = [System.Diagnostics.Stopwatch]::StartNew()
        $recordCount = 0
        $errorCount  = 0

        try {
            $parseParams = @{ FilePath = $filePath }
            if ($excludePrivate) { $parseParams.ExcludePrivateIPs = $true }

            $parseResult = Invoke-VBDNSLogParser @parseParams
            $buffer      = $parseResult.Buffer      # List[object[]] -- never piped through PS pipeline
            $errorCount  = $parseResult.ErrorCount
            $recordCount = $buffer.Count

            if ($recordCount -gt 0) {
                $stopwatch.Stop()
                $duration = $stopwatch.Elapsed.TotalSeconds

                Invoke-VBBulkInsert `
                    -DatabasePath    $dbPath `
                    -Buffer          $buffer `
                    -SourceFile      $filePath `
                    -RecordCount     $recordCount `
                    -ErrorCount      $errorCount `
                    -FileHash        $fileHash `
                    -DurationSeconds $duration | Out-Null
            }

            $stopwatch.Stop()

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

    } -ThrottleLimit $ThrottleLimit
}
