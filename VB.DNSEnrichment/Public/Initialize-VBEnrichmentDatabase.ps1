function Initialize-VBEnrichmentDatabase {
<#
.SYNOPSIS
    Create or migrate the SQLite database used by VB.DNSEnrichment.

.DESCRIPTION
    Idempotent. Safe to call on every module run. Creates the database file (and
    parent folder) if missing, applies any pending migrations from Sql/*.sql in
    numeric order, and records each applied version in the SchemaVersion table.

    Migration files must be named NNN_description.sql (e.g. 001_init.sql,
    002_add_index.sql). The numeric prefix is the version number. Files are
    discovered from $PSScriptRoot\..\Sql relative to this function's location.

.PARAMETER DatabasePath
    Full path to the SQLite database file. The parent folder is created if missing.
    Defaults to "$env:LOCALAPPDATA\VB.DNSEnrichment\enrichment.db".

.OUTPUTS
    [PSCustomObject] -- ComputerName, Status, Path, Created, Version, MigrationsApplied, Error, CollectionTime.

.EXAMPLE
    Initialize-VBEnrichmentDatabase

.EXAMPLE
    Initialize-VBEnrichmentDatabase -DatabasePath 'D:\Data\enrichment.db'

.NOTES
    Version:      1.0.0
    MinPSVersion: 5.1
    Author:       VB
    ChangeLog:
        1.0.0 -- 2026-05-10 -- Initial release
#>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter()]
        [string]$DatabasePath = (Join-Path $env:LOCALAPPDATA 'VB.DNSEnrichment\enrichment.db')
    )

    $computer       = $env:COMPUTERNAME
    $collectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')

    try {
        # --- Step 1 -- Ensure parent folder exists ---
        $parent = Split-Path -Path $DatabasePath -Parent
        if (-not (Test-Path -LiteralPath $parent)) {
            New-Item -Path $parent -ItemType Directory -Force | Out-Null
            Write-Verbose "[Init-DB] Created folder: $parent"
        }

        $created = -not (Test-Path -LiteralPath $DatabasePath)
        if ($created) {
            Write-Verbose "[Init-DB] Database file does not exist; will be created on first query"
        }

        # --- Step 2 -- Locate the Sql migration folder ---
        $sqlFolder = Join-Path $PSScriptRoot '..\Sql'
        $sqlFolder = (Resolve-Path -LiteralPath $sqlFolder -ErrorAction Stop).Path

        $migrationFiles = Get-ChildItem -Path $sqlFolder -Filter '*.sql' -File |
            Where-Object { $_.Name -match '^(\d+)_' } |
            Sort-Object { [int]($_.Name -replace '^(\d+)_.*', '$1') }

        if (-not $migrationFiles) {
            throw "No migration files found in $sqlFolder"
        }

        # --- Step 3 -- Bootstrap SchemaVersion table on first run ---
        # If the database is brand new there's no SchemaVersion table yet, so the
        # SELECT below will fail. Create it pre-emptively as a no-op DDL.
        $bootstrapSql = @'
CREATE TABLE IF NOT EXISTS SchemaVersion (
    Version    INTEGER PRIMARY KEY,
    AppliedAt  TEXT NOT NULL
);
'@
        Invoke-VBSqliteCommand -DatabasePath $DatabasePath -NonQuery -Query $bootstrapSql | Out-Null

        # --- Step 4 -- Read currently-applied versions ---
        $appliedRows = Invoke-VBSqliteCommand -DatabasePath $DatabasePath `
            -Query 'SELECT Version FROM SchemaVersion'
        $applied = @{}
        foreach ($row in $appliedRows) {
            $applied[[int]$row.Version] = $true
        }

        # --- Step 5 -- Apply each pending migration in order ---
        $migrationsApplied = @()
        foreach ($file in $migrationFiles) {
            $version = [int]($file.Name -replace '^(\d+)_.*', '$1')
            if ($applied.ContainsKey($version)) {
                Write-Verbose "[Init-DB] Migration $version already applied -- skipping"
                continue
            }

            Write-Verbose "[Init-DB] Applying migration $($file.Name)"
            $sql = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
            Invoke-VBSqliteCommand -DatabasePath $DatabasePath -NonQuery -Query $sql | Out-Null

            Invoke-VBSqliteCommand -DatabasePath $DatabasePath -NonQuery `
                -Query 'INSERT OR IGNORE INTO SchemaVersion (Version, AppliedAt) VALUES (@v, @t)' `
                -SqlParameters @{ v = $version; t = (Get-Date).ToString('o') } | Out-Null

            $migrationsApplied += $file.Name
        }

        # --- Step 6 -- Read the final highest applied version ---
        $current = Invoke-VBSqliteCommand -DatabasePath $DatabasePath `
            -Query 'SELECT MAX(Version) AS V FROM SchemaVersion'
        $currentVersion = if ($current -and $current.V) { [int]$current.V } else { 0 }

        [PSCustomObject]@{
            ComputerName      = $computer
            Status            = 'Success'
            Path              = $DatabasePath
            Created           = $created
            Version           = $currentVersion
            MigrationsApplied = $migrationsApplied
            Error             = $null
            CollectionTime    = $collectionTime
        }
    }
    catch {
        [PSCustomObject]@{
            ComputerName      = $computer
            Status            = 'Failed'
            Path              = $DatabasePath
            Created           = $false
            Version           = 0
            MigrationsApplied = @()
            Error             = $_.Exception.Message
            CollectionTime    = $collectionTime
        }
    }
}
