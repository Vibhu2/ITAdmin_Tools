function Invoke-VBLoadOUITable {
<#
.SYNOPSIS
    Load (and if necessary download/refresh) the IEEE OUI CSV into a hashtable.
    Private helper used only by Get-VBOUIVendor.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$OUIFilePath
    )

    $OUI_URL       = 'https://standards-oui.ieee.org/oui/oui.csv'
    $REFRESH_DAYS  = 30
    $MIN_SIZE_BYTES = 1MB

    # Ensure parent folder exists
    $parent = Split-Path -Path $OUIFilePath -Parent
    if (-not (Test-Path -LiteralPath $parent)) {
        New-Item -Path $parent -ItemType Directory -Force | Out-Null
    }

    $needsDownload = $false

    if (-not (Test-Path -LiteralPath $OUIFilePath)) {
        Write-Verbose "[OUI] oui.csv not found -- downloading from IEEE"
        $needsDownload = $true
    }
    else {
        $item = Get-Item -LiteralPath $OUIFilePath
        if ($item.Length -lt $MIN_SIZE_BYTES) {
            Write-Verbose "[OUI] oui.csv too small ($([int]($item.Length/1KB)) KB) -- re-downloading"
            $needsDownload = $true
        }
        elseif (((Get-Date) - $item.LastWriteTime).TotalDays -gt $REFRESH_DAYS) {
            Write-Verbose "[OUI] oui.csv is $([int]((Get-Date) - $item.LastWriteTime).TotalDays) days old (threshold $REFRESH_DAYS) -- refreshing"
            $needsDownload = $true
        }
    }

    if ($needsDownload) {
        try {
            if ($PSVersionTable.PSVersion.Major -lt 6) {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            }

            Write-Verbose "[OUI] Downloading OUI database from $OUI_URL ..."
            $tmpPath = "$OUIFilePath.tmp"
            Invoke-WebRequest -Uri $OUI_URL -OutFile $tmpPath -UseBasicParsing -ErrorAction Stop

            # Validate the download before replacing
            $tmpItem = Get-Item -LiteralPath $tmpPath
            if ($tmpItem.Length -lt $MIN_SIZE_BYTES) {
                Remove-Item -LiteralPath $tmpPath -Force -ErrorAction SilentlyContinue
                throw "Downloaded file too small ($([int]($tmpItem.Length/1KB)) KB) -- download may have failed"
            }

            # Replace old file
            if (Test-Path -LiteralPath $OUIFilePath) {
                Remove-Item -LiteralPath $OUIFilePath -Force
            }
            Move-Item -LiteralPath $tmpPath -Destination $OUIFilePath -Force
            Write-Verbose "[OUI] Download complete: $([math]::Round($tmpItem.Length/1MB,1)) MB"
        }
        catch {
            Write-Warning "[OUI] Download failed: $($_.Exception.Message). Vendor lookup will be unavailable this session."
            return @{}
        }
    }

    # Import CSV into hashtable keyed by 6-char uppercase OUI (Assignment column)
    Write-Verbose "[OUI] Loading OUI table into memory..."
    $table = @{}
    try {
        $rows = Import-Csv -LiteralPath $OUIFilePath -Encoding UTF8 -ErrorAction Stop
        foreach ($row in $rows) {
            # IEEE CSV columns: Registry, Assignment, "Organization Name", "Organization Address"
            $oui    = $row.Assignment.ToUpperInvariant().Trim()
            $vendor = $row.'Organization Name'.Trim()
            if ($oui.Length -eq 6 -and -not [string]::IsNullOrWhiteSpace($vendor)) {
                $table[$oui] = $vendor
            }
        }
        Write-Verbose "[OUI] Table loaded: $($table.Count) OUI entries"
    }
    catch {
        Write-Warning "[OUI] Failed to parse oui.csv: $($_.Exception.Message)"
    }

    $table
}
