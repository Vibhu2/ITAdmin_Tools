#
# VB.AdminTools.psm1 -- Module loader
# Author  : Vibhu Bhatnagar
#

$functionFiles = Get-ChildItem -Path "$PSScriptRoot\*.ps1" -ErrorAction SilentlyContinue

foreach ($file in $functionFiles) {
    . $file.FullName
}

Export-ModuleMember -Function *
