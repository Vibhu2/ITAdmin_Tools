# ============================================================
# SCRIPT   : Example-PrinterMigration.ps1
# PURPOSE  : Working example -- install a printer driver and migrate
#            user printer mappings using the rich CSV format.
#            Deploy via RMM; run locally on each target machine.
# VERSION  : 1.0.0
# CHANGED  : 16-06-2026 -- Initial example
# AUTHOR   : Vibhu Bhatnagar
# ENCODING : UTF-8 with BOM
# ============================================================
#
# HOW TO USE
# ----------
# 1. Edit the CONFIGURATION block below (CSV path, log path).
# 2. Populate the CSV -- see Docs\PrinterMappings_Sample.csv for format.
# 3. Deploy this script via RMM to run locally on each target machine.
#    The CSV filters automatically to the local machine name.
# 4. Always do a -WhatIf run first (Step 2 below) before the live run.
#
# REQUIREMENTS
# ------------
#   - PowerShell 5.1
#   - Run as Administrator (required for registry hive access and driver install)
#   - PrintManagement module (built-in on Windows 8 / Server 2012+)
#   - VB.WorkstationReport module loaded (dot-sourced or imported)
#   - CSV accessible from the target machine at $CsvPath
# ============================================================

#Requires -RunAsAdministrator
$ErrorActionPreference = 'Stop'

# ============================================================
# --- CONFIGURATION ---
# ============================================================

$CsvPath  = '\\FileServer\Deploy\PrinterMappings_Sample.csv'
$LogPath  = "\\FileServer\Logs\PrinterMigration_$env:COMPUTERNAME.csv"

# ============================================================
# --- MAIN LOGIC ---
# ============================================================

# Step 1 -- Load the module (adjust path to your deployment location)
$ModulePath = Join-Path $PSScriptRoot '..\VB.WorkstationReport.psm1'
if (Test-Path -Path $ModulePath) {
    Import-Module -Name $ModulePath -Force
}
else {
    throw "Module not found at '$ModulePath'. Adjust the path in the CONFIGURATION block."
}

# Step 2 -- WhatIf run: review what will change before committing
Write-Host "`n[WhatIf] Previewing changes -- no modifications will be made." -ForegroundColor Cyan
Set-VBUserPrinterMigration -MappingCsv $CsvPath -WhatIf

# Step 3 -- Confirm before proceeding (remove this block for unattended RMM deployment)
$confirm = Read-Host "`nProceed with live migration? (Y/N)"
if ($confirm -ne 'Y') {
    Write-Host "Migration cancelled." -ForegroundColor Yellow
    exit 0
}

# Step 4 -- Run the live migration; capture results
Write-Host "`n[Live] Running printer migration on $env:COMPUTERNAME ..." -ForegroundColor Green
$results = Set-VBUserPrinterMigration -MappingCsv $CsvPath -Verbose

# Step 5 -- Display summary to console
Write-Host "`n--- Results ---" -ForegroundColor Cyan
$results | Format-Table -AutoSize

$migrated      = @($results | Where-Object { $_.Action -eq 'Migrated' }).Count
$alreadyDone   = @($results | Where-Object { $_.Action -eq 'AlreadyMigrated' }).Count
$skipped       = @($results | Where-Object { $_.Action -eq 'Skipped' }).Count
$failed        = @($results | Where-Object { $_.Status -eq 'Failed' }).Count

Write-Host "Migrated: $migrated  |  Already done: $alreadyDone  |  Skipped: $skipped  |  Failed: $failed" -ForegroundColor White

# Step 6 -- Export results to log share
$results | Export-Csv -Path $LogPath -NoTypeInformation -Append -Encoding UTF8
Write-Host "Results written to: $LogPath" -ForegroundColor Green

# Step 7 -- Warn if any driver required a reboot
# Note: RebootRequired comes through as a Write-Warning during migration.
# Check the output for Failed rows and review the log if any issues occurred.
if ($failed -gt 0) {
    Write-Warning "$failed mapping(s) failed on $env:COMPUTERNAME -- review $LogPath for details."
}
