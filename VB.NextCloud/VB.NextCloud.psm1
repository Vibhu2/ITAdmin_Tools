# ============================================================
# MODULE  : VB.NextCloud.psm1
# VERSION : 1.2.0
# CHANGED : 15-04-2026 -- Initial PSGallery release
# AUTHOR  : Vibhu Bhatnagar
# PURPOSE : Module loader -- zero logic here
# ENCODING: UTF-8 with BOM
# ------------------------------------------------------------
# This file is the loader only.
# All function logic lives in Public\ and Private\.
# To add a function: drop a .ps1 into the correct subfolder.
# To remove a function: delete the .ps1 file.
# This file never needs to change.
# ============================================================

# Step 1 -- Load all private helpers first (internal, never exported)
Get-ChildItem -Path "$PSScriptRoot\Private" -Recurse -Filter '*.ps1' -ErrorAction SilentlyContinue |
    ForEach-Object {
        . $_.FullName
    }

# Step 2 -- Load all public functions (exported via FunctionsToExport = '*' in manifest)
Get-ChildItem -Path "$PSScriptRoot\Public" -Recurse -Filter '*.ps1' |
    ForEach-Object {
        . $_.FullName
    }
