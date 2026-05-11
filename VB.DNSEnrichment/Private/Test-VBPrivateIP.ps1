function Test-VBPrivateIP {
<#
.SYNOPSIS
    Test whether an IP address is in private/RFC1918, CGNAT, link-local, or loopback space.

.DESCRIPTION
    Returns $true for any of:
        RFC1918:    10.0.0.0/8, 172.16.0.0/12, 192.168.0.0/16
        CGNAT:      100.64.0.0/10
        Link-local: 169.254.0.0/16
        Loopback:   127.0.0.0/8
    Otherwise returns $false. Used by the orchestrator to validate input IPs and skip
    public addresses with a warning.

.PARAMETER IPAddress
    The IPv4 address string to test.

.OUTPUTS
    [bool]

.EXAMPLE
    Test-VBPrivateIP -IPAddress '192.168.1.45'

.NOTES
    Version:      1.0.0
    MinPSVersion: 5.1
    Author:       VB
    ChangeLog:
        1.0.0 -- 2026-05-10 -- Initial release
#>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [string]$IPAddress
    )

    process {
        try {
            $parsed = [System.Net.IPAddress]::Parse($IPAddress)
        }
        catch {
            return $false
        }

        if ($parsed.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) {
            return $false
        }

        $bytes = $parsed.GetAddressBytes()
        $a = [int]$bytes[0]
        $b = [int]$bytes[1]

        # RFC1918 -- 10.0.0.0/8
        if ($a -eq 10) { return $true }

        # RFC1918 -- 172.16.0.0/12 (172.16.0.0 - 172.31.255.255)
        if ($a -eq 172 -and $b -ge 16 -and $b -le 31) { return $true }

        # RFC1918 -- 192.168.0.0/16
        if ($a -eq 192 -and $b -eq 168) { return $true }

        # CGNAT -- 100.64.0.0/10 (100.64.0.0 - 100.127.255.255)
        if ($a -eq 100 -and $b -ge 64 -and $b -le 127) { return $true }

        # Link-local -- 169.254.0.0/16
        if ($a -eq 169 -and $b -eq 254) { return $true }

        # Loopback -- 127.0.0.0/8
        if ($a -eq 127) { return $true }

        return $false
    }
}
