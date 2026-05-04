# ============================================================
# SCRIPT   : Get-VBPrintServerInventory.ps1
# VERSION  : 1.0.0
# CHANGED  : 04-05-2026 -- Initial release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Collects all printer details from the local print server and
#            produces a full inventory report plus a migration-ready CSV
#	     Main purpose of this script is to collect all the require info for printer migration 
# ENCODING : UTF-8 with BOM
# RUN ON   : Print server (local, elevated)
# ============================================================

#Requires -RunAsAdministrator
#Requires -Version 5.1

[CmdletBinding()]
param(
    [string]$OutputFolder = 'C:\Temp'
)

#region --- Setup ---

if (-not (Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}

$ServerName       = $env:COMPUTERNAME
$Timestamp        = Get-Date -Format 'yyyyMMdd_HHmmss'
$InventoryCsv     = Join-Path $OutputFolder "PrinterInventory_${ServerName}_${Timestamp}.csv"
$MigrationCsv     = Join-Path $OutputFolder "PrinterMigration_${ServerName}_${Timestamp}.csv"
$Results          = [System.Collections.Generic.List[object]]::new()

Write-Host "`n  Print Server Inventory -- $ServerName" -ForegroundColor Cyan
Write-Host "  $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')`n" -ForegroundColor Gray

#endregion

#region --- Collect Data ---

Write-Host "  Querying printers and ports..." -ForegroundColor Gray

try {
    $AllPrinters = Get-Printer -ErrorAction Stop
}
catch {
    Write-Error "Failed to query printers: $_"
    exit 1
}

$AllPorts = Get-PrinterPort -ErrorAction SilentlyContinue

foreach ($Printer in $AllPrinters) {

    $Port      = $AllPorts | Where-Object { $_.Name -eq $Printer.PortName }
    $PortType  = 'Other'
    $IPAddress = $null
    $IPStatus  = 'Unresolved'

    # --- Determine port type and IP ---
    if ($Printer.PortName -match '^WSD-') {

        $PortType = 'WSD'

        # Attempt 1: DNS resolution of printer device name
        $IPAddress = try {
            [System.Net.Dns]::GetHostAddresses($Printer.Name) |
                Where-Object { $_.AddressFamily -eq 'InterNetwork' } |
                Select-Object -First 1 -ExpandProperty IPAddressToString
        }
        catch { $null }

        # Attempt 2: Ping and read IPv4 response
        if (-not $IPAddress) {
            $Ping = Test-Connection -ComputerName $Printer.Name -Count 1 -ErrorAction SilentlyContinue
            if ($Ping) {
                $IPAddress = $Ping.IPV4Address.IPAddressToString
            }
        }

        $IPStatus = if ($IPAddress) { 'Resolved' } else { 'Unresolved -- assign static IP' }
    }
    elseif ($Port -and $Port.PrinterHostAddress) {
        $PortType  = 'TCP/IP'
        $IPAddress = $Port.PrinterHostAddress
        $IPStatus  = 'Resolved'
    }
    elseif ($Printer.PortName -match '^USB') {
        $PortType = 'USB'
        $IPStatus = 'N/A -- local port'
    }
    elseif ($Printer.PortName -match '^(LPT|COM)') {
        $PortType = 'Local'
        $IPStatus = 'N/A -- local port'
    }

    # --- Build UNC path ---
    $UNCPath = if ($Printer.Shared -and $Printer.ShareName) {
        "\\$ServerName\$($Printer.ShareName)"
    }
    elseif ($Printer.Shared) {
        "\\$ServerName\$($Printer.Name)"
    }
    else {
        'Not shared'
    }

    $Results.Add([PSCustomObject]@{
        ServerName   = $ServerName
        PrinterName  = $Printer.Name
        ShareName    = if ($Printer.ShareName) { $Printer.ShareName } else { '—' }
        UNCPath      = $UNCPath
        DriverName   = $Printer.DriverName
        PortName     = $Printer.PortName
        PortType     = $PortType
        IPAddress    = if ($IPAddress) { $IPAddress } else { '—' }
        IPStatus     = $IPStatus
        Shared       = $Printer.Shared
        Published    = $Printer.Published
        InfPath      = $Printer.InfPath
    })
}

#endregion

#region --- Display Report ---

Write-Host "`n  ── Printer Inventory ────────────────────────────────────────────`n"

$Results | Format-Table `
    @{ L = 'ShareName';   E = { $_.ShareName };   A = 'Left'   },
    @{ L = 'UNCPath';     E = { $_.UNCPath };      A = 'Left'   },
    @{ L = 'DriverName';  E = { $_.DriverName };   A = 'Left'   },
    @{ L = 'PortType';    E = { $_.PortType };      A = 'Center' },
    @{ L = 'IPAddress';   E = { $_.IPAddress };    A = 'Left'   },
    @{ L = 'IPStatus';    E = { $_.IPStatus };     A = 'Left'   },
    @{ L = 'Shared';      E = { $_.Shared };       A = 'Center' } `
    -AutoSize

# --- Summary ---
$Total      = $Results.Count
$Shared     = @($Results | Where-Object { $_.Shared }).Count
$Resolved   = @($Results | Where-Object { $_.IPStatus -eq 'Resolved' }).Count
$Unresolved = @($Results | Where-Object { $_.IPStatus -like 'Unresolved*' }).Count

Write-Host "  Total printers : $Total"
Write-Host "  Shared         : $Shared"
Write-Host "  IP Resolved    : $Resolved" -ForegroundColor Green
if ($Unresolved -gt 0) {
    Write-Host "  IP Unresolved  : $Unresolved  ← assign static IPs before migration" -ForegroundColor Yellow
}

# --- Flag unresolved printers ---
if ($Unresolved -gt 0) {
    Write-Host "`n  ── Needs Attention (no IP found) ────────────────────────────────`n" -ForegroundColor Yellow
    $Results |
        Where-Object { $_.IPStatus -like 'Unresolved*' } |
        Format-Table PrinterName, ShareName, PortName, PortType -AutoSize
}

#endregion

#region --- Export Files ---

# Full inventory
$Results | Export-Csv -Path $InventoryCsv -NoTypeInformation -Encoding UTF8
Write-Host "`n  Full inventory : $InventoryCsv" -ForegroundColor Cyan

# Migration-ready CSV (shared + IP resolved only)
$MigrationRows = $Results |
    Where-Object { $_.Shared -and $_.IPStatus -eq 'Resolved' } |
    Select-Object `
        @{ N = 'OldPath';    E = { $_.UNCPath } },
        @{ N = 'NewPath';    E = { $_.IPAddress } },
        @{ N = 'DriverName'; E = { $_.DriverName } }

if ($MigrationRows) {
    $MigrationRows | Export-Csv -Path $MigrationCsv -NoTypeInformation -Encoding UTF8
    Write-Host "  Migration CSV  : $MigrationCsv" -ForegroundColor Green
    Write-Host "`n  Migration CSV preview:`n"
    $MigrationRows | Format-Table -AutoSize
}
else {
    Write-Host "`n  No migration-ready rows -- verify printer IPs and sharing." -ForegroundColor Yellow
}

Write-Host ""

#endregion
