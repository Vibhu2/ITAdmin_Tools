function ConvertTo-VBNormalisedMAC {
<#
.SYNOPSIS
    Normalise a MAC address by stripping separators and uppercasing.

.DESCRIPTION
    Accepts MAC addresses in any of the common formats (00:1A:2B:3C:4D:5E,
    00-1A-2B-3C-4D-5E, 001A.2B3C.4D5E, 001A2B3C4D5E) and returns a 12-character
    uppercase string. Returns $null when the input is null/empty or doesn't
    contain exactly 12 hex characters after stripping separators.

.PARAMETER MACAddress
    The MAC address string to normalise.

.OUTPUTS
    [string] -- 12-char uppercase MAC, or $null on invalid input.

.EXAMPLE
    ConvertTo-VBNormalisedMAC -MACAddress '00:1a:2b:3c:4d:5e'
    # Returns: 001A2B3C4D5E

.NOTES
    Version:      1.0.0
    MinPSVersion: 5.1
    Author:       VB
    ChangeLog:
        1.0.0 -- 2026-05-10 -- Initial release
#>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(ValueFromPipeline)]
        [string]$MACAddress
    )

    process {
        if ([string]::IsNullOrWhiteSpace($MACAddress)) {
            return $null
        }

        # Strip common separators -- colon, dash, dot, space
        $stripped = $MACAddress -replace '[:\-\.\s]', ''
        $stripped = $stripped.ToUpperInvariant()

        # Must be exactly 12 hex characters
        if ($stripped -notmatch '^[0-9A-F]{12}$') {
            return $null
        }

        return $stripped
    }
}
