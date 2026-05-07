#
# VB.AdminTools.psm1 -- Module loader
# Author  : Vibhu Bhatnagar
#

$functionFiles = Get-ChildItem -Path $PSScriptRoot -Filter "*.ps1" -Recurse -ErrorAction SilentlyContinue

foreach ($file in $functionFiles) {
    . $file.FullName
}

Export-ModuleMember -Function *
