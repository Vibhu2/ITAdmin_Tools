# ============================================================
# MODULE   : VB.WindowsDNSLogAnalysis.psm1
# VERSION  : 1.1.0
# CHANGED  : 2026-05-07 -- Initial build
#            2026-05-07 -- Renamed from DNSLogDB to VB.WindowsDNSLogAnalysis
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Module loader -- dot-sources all Private and Public functions
# ENCODING : UTF-8 with BOM
# ============================================================

#Requires -Version 7.0

# Verify PSSQLite is available before loading -- hard dependency
if (-not (Get-Module -Name PSSQLite -ListAvailable)) {
    throw "VB.WindowsDNSLogAnalysis requires the PSSQLite module. Install it with: Install-Module PSSQLite"
}

# Dot-source Private functions first (helpers required by Public functions)
$Private = Get-ChildItem -Path "$PSScriptRoot\Private\*.ps1" -ErrorAction SilentlyContinue
$Public  = Get-ChildItem -Path "$PSScriptRoot\Public\*.ps1"  -ErrorAction SilentlyContinue

foreach ($file in @($Private + $Public)) {
    try {
        . $file.FullName
    }
    catch {
        Write-Error "VB.WindowsDNSLogAnalysis: Failed to dot-source '$($file.FullName)': $($_.Exception.Message)"
    }
}

# Export Public functions only -- Private functions are internal and never exported
Export-ModuleMember -Function $Public.BaseName
