# ============================================================
# FUNCTION : Test-VBPrivateIP
# VERSION  : 1.0.0
# CHANGED  : 2026-05-07 -- Initial build
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Classify an IP address as private/reserved or public
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Returns $true if the supplied IP address falls within a private or reserved range.

.DESCRIPTION
    Checks an IP address string against RFC1918 private IPv4 ranges, loopback,
    link-local, and IPv6 private/local ranges. Used by Invoke-VBDNSLogParser via the
    IP classification cache — this function is called only once per unique IP address,
    not once per log line.

    IPv4 ranges checked:
        10.0.0.0/8       RFC1918
        172.16.0.0/12    RFC1918
        192.168.0.0/16   RFC1918
        127.0.0.0/8      Loopback
        169.254.0.0/16   Link-local (APIPA)

    IPv6 ranges checked:
        ::1              Loopback
        fe80::/10        Link-local
        fc00::/7         Unique local (ULA)

.PARAMETER IPAddress
    The IP address string to classify. Accepts both IPv4 and IPv6.

.OUTPUTS
    [bool] $true if private/reserved, $false if public.

.NOTES
    Version  : 1.0.0
    Author   : Vibhu Bhatnagar
    Modified : 2026-05-07
    Category : Private
    Called by: Invoke-VBDNSLogParser (once per unique IP via cache)

    Regex patterns are compiled at module load time (script-scope) and reused.
    Do not move pattern compilation inside the function body.
#>

# Module-level compiled regex constants — compiled once at dot-source time, reused per call
$script:_PrivateIPv4Regex = [regex]::new(
    '^(10\.\d{1,3}\.\d{1,3}\.\d{1,3}|' +
    '172\.(1[6-9]|2\d|3[01])\.\d{1,3}\.\d{1,3}|' +
    '192\.168\.\d{1,3}\.\d{1,3}|' +
    '127\.\d{1,3}\.\d{1,3}\.\d{1,3}|' +
    '169\.254\.\d{1,3}\.\d{1,3})$',
    [System.Text.RegularExpressions.RegexOptions]::Compiled
)

function Test-VBPrivateIP {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$IPAddress
    )

    if ([string]::IsNullOrWhiteSpace($IPAddress)) {
        return $false
    }

    # Step 1 -- IPv4 classification via compiled regex
    if ($IPAddress -match '^\d') {
        return $script:_PrivateIPv4Regex.IsMatch($IPAddress)
    }

    # Step 2 -- IPv6 classification via StartsWith (faster than regex for prefix checks)
    $ipLower = $IPAddress.ToLower()

    # Loopback
    if ($ipLower -eq '::1') { return $true }

    # Link-local fe80::/10 — covers fe80 through febf
    if ($ipLower.StartsWith('fe80:') -or
        $ipLower.StartsWith('fe81:') -or $ipLower.StartsWith('fe82:') -or
        $ipLower.StartsWith('fe83:') -or $ipLower.StartsWith('fe84:') -or
        $ipLower.StartsWith('fe85:') -or $ipLower.StartsWith('fe86:') -or
        $ipLower.StartsWith('fe87:') -or $ipLower.StartsWith('fe88:') -or
        $ipLower.StartsWith('fe89:') -or $ipLower.StartsWith('fe8a:') -or
        $ipLower.StartsWith('fe8b:') -or $ipLower.StartsWith('fe8c:') -or
        $ipLower.StartsWith('fe8d:') -or $ipLower.StartsWith('fe8e:') -or
        $ipLower.StartsWith('fe8f:') -or $ipLower.StartsWith('fe90:') -or
        $ipLower.StartsWith('fe91:') -or $ipLower.StartsWith('fe92:') -or
        $ipLower.StartsWith('fe93:') -or $ipLower.StartsWith('fe94:') -or
        $ipLower.StartsWith('fe95:') -or $ipLower.StartsWith('fe96:') -or
        $ipLower.StartsWith('fe97:') -or $ipLower.StartsWith('fe98:') -or
        $ipLower.StartsWith('fe99:') -or $ipLower.StartsWith('fe9a:') -or
        $ipLower.StartsWith('fe9b:') -or $ipLower.StartsWith('fe9c:') -or
        $ipLower.StartsWith('fe9d:') -or $ipLower.StartsWith('fe9e:') -or
        $ipLower.StartsWith('fe9f:') -or $ipLower.StartsWith('fea0:') -or
        $ipLower.StartsWith('fea1:') -or $ipLower.StartsWith('fea2:') -or
        $ipLower.StartsWith('fea3:') -or $ipLower.StartsWith('fea4:') -or
        $ipLower.StartsWith('fea5:') -or $ipLower.StartsWith('fea6:') -or
        $ipLower.StartsWith('fea7:') -or $ipLower.StartsWith('fea8:') -or
        $ipLower.StartsWith('fea9:') -or $ipLower.StartsWith('feaa:') -or
        $ipLower.StartsWith('feab:') -or $ipLower.StartsWith('feac:') -or
        $ipLower.StartsWith('fead:') -or $ipLower.StartsWith('feae:') -or
        $ipLower.StartsWith('feaf:') -or $ipLower.StartsWith('feb0:') -or
        $ipLower.StartsWith('feb1:') -or $ipLower.StartsWith('feb2:') -or
        $ipLower.StartsWith('feb3:') -or $ipLower.StartsWith('feb4:') -or
        $ipLower.StartsWith('feb5:') -or $ipLower.StartsWith('feb6:') -or
        $ipLower.StartsWith('feb7:') -or $ipLower.StartsWith('feb8:') -or
        $ipLower.StartsWith('feb9:') -or $ipLower.StartsWith('feba:') -or
        $ipLower.StartsWith('febb:') -or $ipLower.StartsWith('febc:') -or
        $ipLower.StartsWith('febd:') -or $ipLower.StartsWith('febe:') -or
        $ipLower.StartsWith('febf:')) {
        return $true
    }

    # Unique local fc00::/7 — covers fc00 through fdff
    if ($ipLower.StartsWith('fc') -or $ipLower.StartsWith('fd')) {
        return $true
    }

    return $false
}
