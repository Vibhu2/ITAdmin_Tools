# ============================================================
# FUNCTION : Write-VBSectionHeader
# VERSION  : 1.0.0
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Display formatted section header in console
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Display a formatted section header in the console.

.DESCRIPTION
    Writes a decorative section header with border and title centered within
    specified width. Useful for separating report sections and improving output
    readability. Write-Host output is intentional for console display.

.PARAMETER Title
    The title text to display. Required parameter.

.PARAMETER BorderColor
    Color for the border and title. Accepts: Cyan, Green, Yellow, Red, White,
    Gray, Magenta. Defaults to Cyan.

.PARAMETER TextColor
    Deprecated. Use BorderColor instead. Maintained for backward compatibility.

.PARAMETER Width
    Width of the header in characters. Defaults to 80.

.PARAMETER BorderChar
    Character used for border decoration. Defaults to '='.

.EXAMPLE
    Write-VBSectionHeader -Title 'Security Assessment'
    Displays a section header with cyan colored borders.

.EXAMPLE
    Write-VBSectionHeader -Title 'Important Notice' -BorderColor Red -Width 100
    Displays a red section header with 100 character width.

.EXAMPLE
    Write-VBSectionHeader -Title 'Server Inventory' -BorderChar '*'
    Displays a section header with asterisk borders.

.OUTPUTS
    None. Outputs to console via Write-Host.

.NOTES
    Version  : 1.0.0
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : Utility
#>

function Write-VBSectionHeader {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Cyan', 'Green', 'Yellow', 'Red', 'White', 'Gray', 'Magenta')]
        [string]$BorderColor = 'Cyan',
        [Parameter(Mandatory = $false)]
        [int]$Width = 80,
        [Parameter(Mandatory = $false)]
        [char]$BorderChar = '='
    )

    # Step 1 -- Calculate padding for centering
    $padding = [Math]::Max(0, $Width - $Title.Length - 2)
    $leftPad = [Math]::Floor($padding / 2)
    $rightPad = $padding - $leftPad

    # Step 2 -- Build border line
    $borderLine = $BorderChar.ToString() * $Width

    # Step 3 -- Build title line with padding
    $leftPadding = $BorderChar.ToString() * $leftPad
    $rightPadding = $BorderChar.ToString() * $rightPad
    $titleLine = "$leftPadding $Title $rightPadding"

    # Step 4 -- Adjust right padding if needed to ensure exact width
    if ($titleLine.Length -ne $Width) {
        $rightPadding = $BorderChar.ToString() * ($rightPad + ($Width - $titleLine.Length))
        $titleLine = "$leftPadding $Title $rightPadding"
    }

    # Step 5 -- Output header
    Write-Host ''
    Write-Host $borderLine -ForegroundColor $BorderColor
    Write-Host $titleLine -ForegroundColor $BorderColor
    Write-Host $borderLine -ForegroundColor $BorderColor
    Write-Host ''
}
