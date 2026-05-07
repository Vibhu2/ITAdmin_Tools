# ============================================================
# FUNCTION : Initialize-VBDNSLogDatabase
# VERSION  : 1.0.0
# CHANGED  : 2026-05-07 -- Initial build
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Create and configure the DNSLogDB SQLite database schema
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Creates and configures the DNSLogDB SQLite database with the correct schema,
    indexes, and WAL journal mode.

.DESCRIPTION
    One-time setup step required before any log files can be imported. Creates the
    .db file at the specified path (or connects to an existing one) and ensures the
    following objects exist:

    Tables:
        DNSLog      -- Main data table. One row per parsed PACKET line.
        ImportLog   -- Audit table. One row per imported log file.

    Indexes (on DNSLog):
        idx_LogDateTime   -- Primary time filter column
        idx_IPAddress     -- Client IP lookups and top-talker queries
        idx_Protocol      -- UDP/TCP filter
        idx_ResponseCode  -- NXDOMAIN / error analysis
        idx_QueryType     -- Record type breakdown
        idx_QueryName     -- Domain name lookups
        idx_SourceFile    -- Per-file filtering

    PRAGMA:
        journal_mode = WAL (Write-Ahead Logging, set permanently on the database)

    All CREATE TABLE and CREATE INDEX statements use IF NOT EXISTS — safe to re-run
    against an existing database without losing data.

.PARAMETER DatabasePath
    Full path to the .db file to create or connect to.
    The parent directory must already exist.

.EXAMPLE
    Initialize-VBDNSLogDatabase -DatabasePath 'C:\DNS\dns_analysis.db'

.EXAMPLE
    Initialize-VBDNSLogDatabase -DatabasePath 'C:\DNS\dns_analysis.db' -Verbose

.OUTPUTS
    None. Emits Verbose messages for each object created or verified.

.NOTES
    Version  : 1.0.0
    Author   : Vibhu Bhatnagar
    Modified : 2026-05-07
    Category : Public

    Requires: PSSQLite module (Install-Module PSSQLite)
    Run before: Import-VBDNSLog
#>

function Initialize-VBDNSLogDatabase {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath
    )

    # Step 1 -- Validate the path is a .db file path, not a folder or an existing non-database file

    # Guard: path must not be an existing directory
    if (Test-Path $DatabasePath -PathType Container) {
        throw "Initialize-VBDNSLogDatabase: '-DatabasePath' must be a path to a .db FILE, not a folder.`nYou passed : '$DatabasePath'`nExample    : '$DatabasePath\dns_analysis.db'"
    }

    # Guard: if the file already exists, check it is a SQLite database (magic bytes: 'SQLite format 3')
    if (Test-Path $DatabasePath -PathType Leaf) {
        $ext = [System.IO.Path]::GetExtension($DatabasePath).ToLower()
        if ($ext -notin @('.db', '.sqlite', '.sqlite3', '')) {
            throw ("Initialize-VBDNSLogDatabase: The file '$DatabasePath' does not look like a SQLite database " +
                   "(extension '$ext'). '-DatabasePath' must point to a .db file, not a log file.`n" +
                   "Your log file goes to -InputPath on Import-VBDNSLog, not here.")
        }

        # Read first 16 bytes and check SQLite magic header
        $magic = [System.IO.File]::ReadAllBytes($DatabasePath) | Select-Object -First 16
        $header = [System.Text.Encoding]::ASCII.GetString($magic).TrimEnd([char]0)
        if (-not $header.StartsWith('SQLite format 3')) {
            throw ("Initialize-VBDNSLogDatabase: '$DatabasePath' already exists but is NOT a SQLite database " +
                   "(it appears to be a plain text or other file).`n" +
                   "If this is your DNS log file, pass it to Import-VBDNSLog -InputPath instead.`n" +
                   "Choose a different path for -DatabasePath, e.g.: '$(Split-Path $DatabasePath -Parent)\dns_analysis.db'")
        }
    }

    # Guard: warn if no extension given (common mistake — omitting .db)
    if ([System.IO.Path]::GetExtension($DatabasePath) -eq '') {
        Write-Warning "Initialize-VBDNSLogDatabase: '-DatabasePath' has no file extension. Did you mean '$DatabasePath.db'?"
    }

    # Guard: parent directory must exist
    $parentDir = Split-Path $DatabasePath -Parent
    if (-not [string]::IsNullOrWhiteSpace($parentDir) -and -not (Test-Path $parentDir)) {
        throw "Initialize-VBDNSLogDatabase: Parent directory does not exist: '$parentDir'`nCreate it first: New-Item -Path '$parentDir' -ItemType Directory -Force"
    }

    Write-Verbose "Initialize-VBDNSLogDatabase: Target database: '$DatabasePath'"

    # Step 2 -- Set WAL journal mode (permanent setting on the database file)
    Invoke-SqliteQuery -DataSource $DatabasePath -Query "PRAGMA journal_mode = WAL;" | Out-Null
    Write-Verbose "Initialize-VBDNSLogDatabase: journal_mode = WAL"

    # Step 3 -- Create DNSLog table
    $createDNSLog = @"
CREATE TABLE IF NOT EXISTS DNSLog (
    Id            INTEGER PRIMARY KEY AUTOINCREMENT,
    LogDateTime   TEXT    NOT NULL,
    LogDate       TEXT,
    LogTime       TEXT,
    ThreadId      TEXT,
    PacketId      TEXT,
    Protocol      TEXT,
    Direction     TEXT,
    IPAddress     TEXT,
    IPVersion     TEXT,
    IsPrivate     INTEGER,
    TransactionId TEXT,
    PacketKind    TEXT,
    Opcode        TEXT,
    FlagsHex      TEXT,
    FlagsChar     TEXT,
    ResponseCode  TEXT,
    Status        TEXT,
    Error         TEXT,
    QueryType     TEXT,
    QueryName     TEXT,
    SourceFile    TEXT,
    ImportedAt    TEXT
);
"@
    Invoke-SqliteQuery -DataSource $DatabasePath -Query $createDNSLog
    Write-Verbose "Initialize-VBDNSLogDatabase: Table DNSLog -- OK"

    # Step 4 -- Create ImportLog table
    $createImportLog = @"
CREATE TABLE IF NOT EXISTS ImportLog (
    Id              INTEGER PRIMARY KEY AUTOINCREMENT,
    FilePath        TEXT    UNIQUE NOT NULL,
    FileName        TEXT,
    FileHash        TEXT,
    RecordCount     INTEGER,
    ErrorCount      INTEGER,
    DurationSeconds REAL,
    ImportedAt      TEXT,
    ServerName      TEXT
);
"@
    Invoke-SqliteQuery -DataSource $DatabasePath -Query $createImportLog
    Write-Verbose "Initialize-VBDNSLogDatabase: Table ImportLog -- OK"

    # Step 5 -- Create indexes
    $indexes = @(
        @{ Name = 'idx_LogDateTime';  Column = 'LogDateTime'  },
        @{ Name = 'idx_IPAddress';    Column = 'IPAddress'    },
        @{ Name = 'idx_Protocol';     Column = 'Protocol'     },
        @{ Name = 'idx_ResponseCode'; Column = 'ResponseCode' },
        @{ Name = 'idx_QueryType';    Column = 'QueryType'    },
        @{ Name = 'idx_QueryName';    Column = 'QueryName'    },
        @{ Name = 'idx_SourceFile';   Column = 'SourceFile'   }
    )

    foreach ($idx in $indexes) {
        $sql = "CREATE INDEX IF NOT EXISTS $($idx.Name) ON DNSLog ($($idx.Column));"
        Invoke-SqliteQuery -DataSource $DatabasePath -Query $sql
        Write-Verbose "Initialize-VBDNSLogDatabase: Index $($idx.Name) -- OK"
    }

    Write-Verbose "Initialize-VBDNSLogDatabase: Initialisation complete."
}
