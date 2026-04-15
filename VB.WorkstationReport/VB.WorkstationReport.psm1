#Requires -Version 5.1
# ============================================================
# VB.WorkstationReport.psm1 -- loader only, zero logic here
# Version : 1.3.0
# Author  : Vibhu Bhatnagar
# ============================================================

# Load all private helpers first (not exported)
Get-ChildItem -Path "$PSScriptRoot\Private\*.ps1" | ForEach-Object { . $_.FullName }

# Load and export all public functions
Get-ChildItem -Path "$PSScriptRoot\Public\*.ps1" | ForEach-Object {
    . $_.FullName
    Export-ModuleMember -Function $_.BaseName
}
