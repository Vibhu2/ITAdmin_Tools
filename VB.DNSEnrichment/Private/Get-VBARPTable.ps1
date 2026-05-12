function Get-VBARPTable {
<#
.SYNOPSIS
    Parse arp -a output into a hashtable keyed by IP address.
    Private helper used by Get-VBARPEntry.
#>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$IPFilter
    )

    $table = @{}

    $arpArgs = if ($IPFilter) { @($IPFilter) } else { @() }
    $arpLines = & arp.exe -a @arpArgs 2>$null

    if ($LASTEXITCODE -ne 0 -and -not $IPFilter) {
        Write-Warning "[ARP] arp.exe returned exit code $LASTEXITCODE"
        return $table
    }

    # arp -a output format: "  <IP>          <MAC>         <type>"
    # e.g.:  192.168.1.1     00-11-22-33-44-55     dynamic
    foreach ($line in $arpLines) {
        if ($line -match '^\s+([\d\.]+)\s+([\w\-]+)\s+(static|dynamic)\s*$') {
            $ip   = $Matches[1].Trim()
            $mac  = $Matches[2].Trim()
            $type = $Matches[3].Trim()
            # Skip multicast/broadcast entries (01-xx, ff-ff-ff-ff-ff-ff)
            if ($mac -match '^01-' -or $mac -eq 'ff-ff-ff-ff-ff-ff') { continue }
            $table[$ip] = [PSCustomObject]@{
                MAC  = $mac
                Type = $type
            }
        }
    }

    $table
}
