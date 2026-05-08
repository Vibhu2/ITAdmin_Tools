# ============================================================
# FUNCTION : Invoke-VBBulkInsert
# VERSION  : 1.0.0
# CHANGED  : 2026-05-07 -- Initial build
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Bulk-insert parsed DNS records into SQLite using prepared statements
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Inserts a buffer of parsed DNS log rows into the DNSLog SQLite table using a
    prepared statement inside a batched transaction.

.DESCRIPTION
    Consumes the List[object[]] buffer produced by Invoke-VBDNSLogParser and writes
    every row to the DNSLog table using the following SQLite performance strategy:

    1. PRAGMA tuning applied once per connection before any inserts
       (synchronous=NORMAL, cache_size=64MB, temp_store=MEMORY, mmap_size=256MB)
    2. Single INSERT statement compiled once (prepared statement)
    3. Rows committed in batches of -BatchSize (default 50,000) to control memory
       and reduce disk sync frequency vs. one transaction per file
    4. Each PS7 parallel thread creates its own SQLiteConnection (WAL mode allows this)

    The 21-column parameter binding order MUST match the object[] index contract
    defined in Invoke-VBDNSLogParser. Never reorder without updating both functions.

    After all rows are inserted, writes one audit row to the ImportLog table.

.PARAMETER DatabasePath
    Full path to the SQLite .db file. Must be initialised by Initialize-VBDNSLogDatabase.

.PARAMETER Buffer
    The List[object[]] of parsed rows from Invoke-VBDNSLogParser.

.PARAMETER SourceFile
    Full path of the log file being imported (written to ImportLog).

.PARAMETER RecordCount
    Number of records inserted (written to ImportLog).

.PARAMETER ErrorCount
    Number of parse errors encountered (written to ImportLog).

.PARAMETER FileHash
    SHA256 hash of the source file (written to ImportLog).

.PARAMETER DurationSeconds
    Parse + insert duration for this file (written to ImportLog).

.PARAMETER BatchSize
    Number of rows per transaction commit. Default: 50,000.

.OUTPUTS
    [int] Number of rows successfully inserted into DNSLog.

.NOTES
    Version  : 1.0.0
    Author   : Vibhu Bhatnagar
    Modified : 2026-05-07
    Category : Private
    Called by: Import-VBDNSLog (inside ForEach-Object -Parallel thread)

    Requires: PSSQLite module (Install-Module PSSQLite)
    WAL mode must be set on the database (done by Initialize-VBDNSLogDatabase).
    Never pass a SQLiteConnection across parallel thread boundaries.
#>

function Invoke-VBBulkInsert {
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath,

        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[object[]]]$Buffer,

        [Parameter(Mandatory = $true)]
        [string]$SourceFile,

        [Parameter(Mandatory = $true)]
        [int]$RecordCount,

        [Parameter(Mandatory = $true)]
        [int]$ErrorCount,

        [Parameter(Mandatory = $true)]
        [string]$FileHash,

        [Parameter(Mandatory = $true)]
        [double]$DurationSeconds,

        [int]$BatchSize = 50000,

        # Shared progress state written by this function; read by the main thread polling loop.
        # Keys written: InsertRowsInserted (int64), InsertTotalRows (int64), InsertDone (bool).
        [System.Collections.Concurrent.ConcurrentDictionary[string,object]]$ProgressState
    )

    $conn         = $null
    $tx           = $null
    $cmd          = $null
    $rowsInserted = 0

    try {
        # Step 1 -- Open connection (each thread gets its own connection -- WAL allows concurrent writes)
        $conn = [System.Data.SQLite.SQLiteConnection]::new("Data Source=$DatabasePath;Version=3;")
        $conn.Open()

        # Step 2 -- PRAGMA tuning (must be set per connection, before any inserts)
        $pragmaCmd = $conn.CreateCommand()
        $pragmaCmd.CommandText = @"
PRAGMA synchronous   = NORMAL;
PRAGMA cache_size    = -65536;
PRAGMA temp_store    = MEMORY;
PRAGMA mmap_size     = 268435456;
"@
        $pragmaCmd.ExecuteNonQuery() | Out-Null
        $pragmaCmd.Dispose()

        # Step 3 -- Prepare INSERT statement once (compiled once, reused per row)
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = @"
INSERT INTO DNSLog (
    LogDateTime, LogDate, LogTime, ThreadId, PacketId,
    Protocol, Direction, IPAddress, IPVersion, IsPrivate,
    TransactionId, PacketKind, Opcode, FlagsHex, FlagsChar,
    ResponseCode, Status, Error, QueryType, QueryName,
    SourceFile, ImportedAt
) VALUES (
    @p0,  @p1,  @p2,  @p3,  @p4,
    @p5,  @p6,  @p7,  @p8,  @p9,
    @p10, @p11, @p12, @p13, @p14,
    @p15, @p16, @p17, @p18, @p19,
    @p20, @p21
)
"@

        # Add parameters once -- values are updated per row in the loop
        $importedAt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
        for ($i = 0; $i -le 20; $i++) {
            $null = $cmd.Parameters.AddWithValue("@p$i", $null)
        }
        $null = $cmd.Parameters.AddWithValue('@p21', $importedAt)

        # Step 4 -- Begin first transaction
        $tx = $conn.BeginTransaction()
        $cmd.Transaction = $tx

        # Total rows for progress percentage
        $totalRows = $Buffer.Count

        # Seed shared state so the polling loop transitions to the insert stage immediately
        if ($null -ne $ProgressState) {
            $ProgressState['InsertTotalRows']   = [long]$totalRows
            $ProgressState['InsertRowsInserted'] = [long]0
            $ProgressState['InsertDone']        = $false
        }

        # Step 5 -- Iterate buffer and bind each object[] row by index
        foreach ($row in $Buffer) {
            for ($i = 0; $i -le 20; $i++) {
                $cmd.Parameters["@p$i"].Value = $row[$i]
            }
            $cmd.ExecuteNonQuery() | Out-Null
            $rowsInserted++

            # Batch commit every BatchSize rows -- begin new transaction immediately
            # Note: do NOT Dispose() the old $tx before reassigning -- let GC handle it.
            # Disposing mid-loop invalidates the SQLiteCommand.Transaction reference on some
            # PSSQLite/System.Data.SQLite builds and causes subsequent rows to be lost.
            if ($rowsInserted % $BatchSize -eq 0) {
                $tx.Commit()
                $tx = $conn.BeginTransaction()
                $cmd.Transaction = $tx
                Write-Verbose "Invoke-VBBulkInsert: Committed batch at $rowsInserted rows"

                # Update shared state every batch -- polling loop renders the bar on the main thread
                if ($null -ne $ProgressState) {
                    $ProgressState['InsertRowsInserted'] = [long]$rowsInserted
                }
            }
        }

        # Step 6 -- Final commit for the last partial batch
        $tx.Commit()
        $tx = $null

        # Step 7 -- Write ImportLog audit row
        $serverName = ''
        try {
            if ($SourceFile -match '^\\\\([^\\]+)\\') {
                $serverName = $Matches[1]
            }
        } catch { }

        $logCmd = $conn.CreateCommand()
        $logCmd.CommandText = @"
INSERT OR IGNORE INTO ImportLog
    (FilePath, FileName, FileHash, RecordCount, ErrorCount, DurationSeconds, ImportedAt, ServerName)
VALUES
    (@fp, @fn, @fh, @rc, @ec, @ds, @ia, @sn)
"@
        $null = $logCmd.Parameters.AddWithValue('@fp', $SourceFile)
        $null = $logCmd.Parameters.AddWithValue('@fn', (Split-Path $SourceFile -Leaf))
        $null = $logCmd.Parameters.AddWithValue('@fh', $FileHash)
        $null = $logCmd.Parameters.AddWithValue('@rc', $RecordCount)
        $null = $logCmd.Parameters.AddWithValue('@ec', $ErrorCount)
        $null = $logCmd.Parameters.AddWithValue('@ds', $DurationSeconds)
        $null = $logCmd.Parameters.AddWithValue('@ia', $importedAt)
        $null = $logCmd.Parameters.AddWithValue('@sn', $serverName)
        $logCmd.ExecuteNonQuery() | Out-Null
        $logCmd.Dispose()

        # Signal the polling loop that insert is complete with final row count
        if ($null -ne $ProgressState) {
            $ProgressState['InsertRowsInserted'] = [long]$rowsInserted
            $ProgressState['InsertDone']        = $true
        }

        Write-Verbose "Invoke-VBBulkInsert: Inserted $rowsInserted rows. ImportLog updated."
    }
    catch {
        if ($null -ne $tx) {
            try { $tx.Rollback() } catch { }
        }
        throw
    }
    finally {
        if ($null -ne $cmd)  { try { $cmd.Dispose()  } catch { } }
        if ($null -ne $tx)   { try { $tx.Rollback()  } catch { } }
        if ($null -ne $conn) { try { $conn.Close(); $conn.Dispose() } catch { } }
    }

    return $rowsInserted
}
