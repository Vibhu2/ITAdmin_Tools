#Requires -Version 5.1
# ============================================================
# VB.WorkstationReport.psm1 -- loader only, zero logic here
# Version : 1.3.3 -- Recurse added to Private and Public loaders to support subfolders
# Author  : Vibhu Bhatnagar
# ============================================================

# Runtime dependency check -- VB.NextCloud must be installed for upload functions to work
if (-not (Get-Module -ListAvailable -Name 'VB.NextCloud')) {
    Write-Warning "VB.WorkstationReport: VB.NextCloud module not found. Install it with: Install-Module VB.NextCloud. Upload functionality (Invoke-VBWorkstationReport) will fail without it."
}

# Load all private helpers first (not exported) -- Recurse covers any Private subfolders
Get-ChildItem -Path "$PSScriptRoot\Private\" -Filter '*.ps1' -Recurse -ErrorAction SilentlyContinue |
    ForEach-Object { . $_.FullName }

# Load and export all public functions -- Recurse covers subfolders like Printer_Management
Get-ChildItem -Path "$PSScriptRoot\Public\" -Filter '*.ps1' -Recurse |
    ForEach-Object {
        . $_.FullName
        Export-ModuleMember -Function $_.BaseName
    }
