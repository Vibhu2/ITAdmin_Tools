# ============================================================
# MODULE   : VB.DNSEnrichment.psm1
# VERSION  : 0.1.0
# CHANGED  : 2026-05-10 -- Round 1: skeleton + context + storage
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Module loader -- dot-sources all Private and Public functions
# ENCODING : UTF-8 with BOM
# ============================================================

# Hard dependency: PSSQLite (declared in manifest, but verify here for clear error)
if (-not (Get-Module -Name PSSQLite -ListAvailable)) {
    throw "VB.DNSEnrichment requires the PSSQLite module. Install with: Install-Module PSSQLite -Scope CurrentUser"
}

# Dot-source Private functions first (helpers required by Public functions)
$Private = Get-ChildItem -Path "$PSScriptRoot\Private\*.ps1" -ErrorAction SilentlyContinue
$Public  = Get-ChildItem -Path "$PSScriptRoot\Public\*.ps1"  -ErrorAction SilentlyContinue

foreach ($file in @($Private + $Public)) {
    try {
        . $file.FullName
    }
    catch {
        Write-Error "VB.DNSEnrichment: Failed to dot-source '$($file.FullName)': $($_.Exception.Message)"
    }
}

# Export Public functions only -- Private functions are internal and never exported
Export-ModuleMember -Function $Public.BaseName
