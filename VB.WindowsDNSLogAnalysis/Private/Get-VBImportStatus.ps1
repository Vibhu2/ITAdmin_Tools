# --- Get-VBImportStatus ---
# VERSION  : 1.0.1
# CHANGED  : 2026-06-04 -- Replace .NET SHA256 crypto API with Get-FileHash (AV compatibility)
# CHANGED  : 2026-05-07 -- Initial build
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Check whether a log file has already been imported via SHA256 hash
# ENCODING : UTF-8 with BOM

<#
.SYNOPSIS
    Checks the ImportLog table to determine whether a log file has already been imported.

.DESCRIPTION
    Computes the SHA256 hash of the specified file and queries the ImportLog table for a
    matching FileHash. Returns a status object indicating whether the file is already
    in the database.

    SHA256 is used rather than path comparison because:
    - Log files are often renamed or rotated (dns.log -> dns_20260501.log)
    - A content hash detects the same data regardless of filename
    - A new file at the same path with different content gets a different hash and is imported

    The returned object also carries the computed FileHash so the caller (Import-VBDNSLog)
    does not need to hash the file a second time before writing to ImportLog.

.PARAMETER FilePath
    Full path to the log file to check.

.PARAMETER DatabasePath
    Full path to the SQLite .db file containing the ImportLog table.

.OUTPUTS
    [PSCustomObject] with properties:
        FilePath        -- Full path of the checked file
        FileHash        -- SHA256 hash string (always computed and returned)
        AlreadyImported -- $true if the hash exists in ImportLog, $false otherwise
        ImportedAt      -- ISO8601 timestamp of the original import, or $null if not yet imported

.NOTES
    Version  : 1.0.1
    Author   : Vibhu Bhatnagar
    Modified : 2026-06-04
    Category : Private
    Called by: Import-VBDNSLog (before starting the parse pipeline)

    Requires: PSSQLite module (Install-Module PSSQLite)
#>

function Get-VBImportStatus {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath
    )
    # Step 1 -- Compute SHA256 hash of the file
    $logHash = (Get-FileHash -LiteralPath $FilePath -Algorithm SHA256).Hash.ToLower()

    # Step 2 -- Query ImportLog for matching hash
    $importedAt      = $null
    $alreadyImported = $false
    try {
        $result = Invoke-SqliteQuery -DataSource $DatabasePath `
            -Query "SELECT ImportedAt FROM ImportLog WHERE FileHash = @hash LIMIT 1" `
            -SqlParameters @{ hash = $logHash } `
            -ErrorAction Stop
        if ($null -ne $result -and $result.Count -gt 0) {
            $alreadyImported = $true
            $importedAt      = $result[0].ImportedAt
        }
    }
    catch {
        # If ImportLog does not exist yet (uninitialised DB), treat as not imported
        Write-Verbose "Get-VBImportStatus: Could not query ImportLog -- $($_.Exception.Message)"
        $alreadyImported = $false
    }
    # Step 3 -- Return status object
    return [PSCustomObject]@{
        FilePath        = $FilePath
        FileHash        = $logHash
        AlreadyImported = $alreadyImported
        ImportedAt      = $importedAt
    }
}