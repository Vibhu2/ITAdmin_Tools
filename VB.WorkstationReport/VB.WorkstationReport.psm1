#Requires -Version 5.1
# ============================================================
# VB.WorkstationReport.psm1 -- loader only, zero logic here
# Version : 1.3.1
# Author  : Vibhu Bhatnagar
# ============================================================

# Runtime dependency check -- VB.NextCloud must be installed for upload functions to work
if (-not (Get-Module -ListAvailable -Name 'VB.NextCloud')) {
    Write-Warning "VB.WorkstationReport: VB.NextCloud module not found. Install it with: Install-Module VB.NextCloud. Upload functionality (Invoke-VBWorkstationReport) will fail without it."
}

# Load all private helpers first (not exported)
Get-ChildItem -Path "$PSScriptRoot\Private\*.ps1" -ErrorAction SilentlyContinue | ForEach-Object { . $_.FullName }

# Load and export all public functions
Get-ChildItem -Path "$PSScriptRoot\Public\*.ps1" | ForEach-Object {
    . $_.FullName
    Export-ModuleMember -Function $_.BaseName
}
