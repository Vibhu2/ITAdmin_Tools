# ============================================================
# SCRIPT   : Get-VBDNSLogSummary.ps1
# VERSION  : 4.1.0
# CHANGED  : 05-05-2026 -- Added Get-VBDNSTopTalkers and Get-VBDNSTopDomains
#                          returning pipeline-exportable PSCustomObjects with
#                          Rank and TotalInDataset. Added Resolve-VBDNSHostname
#                          private helper (PTR + device-type detection).
#                          Rewrote Format-VBDNSTopTalkers and
#                          Format-VBDNSTopDomains as display-only consumers.
# PREV     : 4.0.2 -- Added Write-Progress with byte-accurate % via
#                          BaseStream.Position, updated every 2000 lines
# PREV     : 4.0.1 -- Fixed regex to match actual Windows DNS Server log
#                          format: added ThreadID + PACKET context + Xid fields
# AUTHOR   : VB (based on RealTime LLC v2.0)
# PURPOSE  : Universal DNS debug log analysis -- extract, filter, export and
#            summarise DNS query records from Windows DNS Server debug logs
# ENCODING : UTF-8 with BOM -- do not re-save without BOM
# ------------------------------------------------------------
# CHANGELOG (last 5 only -- full history in Git)
# v4.1.0 -- 05-05-2026 -- Get-VBDNSTopTalkers, Get-VBDNSTopDomains (pipeline-
#                          exportable), Resolve-VBDNSHostname (PTR + device
#                          type), Format- functions rewritten as display-only
# v4.0.2 -- 05-05-2026 -- Write-Progress byte-accurate % via BaseStream.Position
# v4.0.1 -- 05-05-2026 -- Fixed $logPattern: ThreadID + PACKET context + Xid
#                          fields were missing, causing 0 matches on real logs
# v4.0.0 -- 05-05-2026 -- Full field extraction, multi-file, CSV export,
#                          Format-VBDNSTopTalkers, Format-VBDNSTopDomains,
#                          ConvertTo-VBDNSLabel helper, finalized help blocks
# v3.0.0 -- 05-05-2026 -- VB prefix, PSCustomObject output, output-presentation
#                          separation, streaming reader, $null-left comparisons
# ============================================================

$ErrorActionPreference = 'Stop'

# --- CONFIGURATION ---
# Update these paths per environment before running
$LOG_INPUT  = Join-Path $env:USERPROFILE 'Documents\RealTime Connect\Files\SPI_DNS_DC2_logs.txt'
$LOG_OUTPUT = 'C:\Intel\SPIDNSSummary.csv'

# --- HELPER FUNCTIONS ---

function ConvertTo-VBDNSLabel {
<#
.SYNOPSIS
    Decodes a Windows DNS debug log label string into a readable FQDN.

.DESCRIPTION
    Windows DNS Server debug logs encode domain names in DNS wire-format
    length-prefix notation. Each label segment is prefixed by its character
    count in parentheses, for example (7)example(3)com(0) represents
    example.com. The trailing (0) is the root label terminator.

    This function decodes that format into a standard dot-separated fully
    qualified domain name in lowercase. On decode failure the raw input
    string is returned unchanged and a warning is written, so the pipeline
    is never broken by a malformed entry.

    DNS name compression pointer labels (byte sequences beginning 0xC0) are
    not supported and will be returned as-is with a warning.

.PARAMETER EncodedLabel
    The DNS label string as it appears in the debug log, for example
    (3)www(7)example(3)com(0). Required.

.OUTPUTS
    [string] Decoded FQDN in lowercase, or the original raw string if
    decoding fails.

.EXAMPLE
    ConvertTo-VBDNSLabel -EncodedLabel '(7)example(3)com(0)'
    Returns: example.com

.EXAMPLE
    ConvertTo-VBDNSLabel -EncodedLabel '(3)www(7)example(3)com(0)'
    Returns: www.example.com

.EXAMPLE
    ConvertTo-VBDNSLabel -EncodedLabel '(1)a(4)root(4)arpa(0)'
    Returns: a.root.arpa

.NOTES
    Version  : 4.0.0
    Author   : VB
    Modified : 05-05-2026
    Scope    : Private helper -- called by Get-VBDNSLogSummary only
    Limits   : DNS name compression (pointer labels starting 0xC0) not
               supported. Compressed labels are returned raw with a warning.
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EncodedLabel
    )

    try {
        $labels = [regex]::Matches($EncodedLabel, '\(\d+\)([^(]*)') |
            ForEach-Object { $_.Groups[1].Value.Trim() } |
            Where-Object   { $_ -ne '' }

        if ($null -eq $labels -or @($labels).Count -eq 0) {
            return $EncodedLabel
        }

        return ($labels -join '.').ToLower()
    }
    catch {
        Write-Warning "DNS label decode failed for '$EncodedLabel': $($_.Exception.Message)"
        return $EncodedLabel
    }
}


function Resolve-VBDNSHostname {
<#
.SYNOPSIS
    Resolves an IP address to a hostname and classifies the device type.

.DESCRIPTION
    Performs a reverse DNS (PTR) lookup for the supplied IP address using the
    .NET System.Net.Dns class, which is compatible with Windows PowerShell 5.1.

    The resolved hostname is then matched against a set of well-known naming
    patterns to infer device type, make, and model:

      - Cisco IP Phones  : hostnames containing SEP followed by a MAC address
                           (e.g. SEP001122334455). Make = Cisco, Model = IP Phone.
      - Polycom Phones   : hostnames starting with polycom, spip, vvx, or trio.
                           Make = Polycom, Model inferred from VVX/Trio prefix.
      - Yealink Phones   : hostnames starting with yealink or ylk.
                           Make = Yealink, Model = IP Phone.
      - Grandstream      : hostnames starting with grandstream or gxp or gxv.
                           Make = Grandstream, Model = IP Phone.
      - Avaya Phones     : hostnames starting with avaya or 96xx or j1xx.
                           Make = Avaya, Model = IP Phone.
      - Servers          : hostnames containing srv, server, dc, ad, sql, exch,
                           ftp, mail, vcenter, esx, node, or ending in -svr.
      - Network Gear     : hostnames containing sw, rtr, router, switch, ap,
                           firewall, fw, gw, gateway, or pix.
      - Printers         : hostnames containing prt, print, mfp, copier, or
                           starting with hp, canon, konica, xerox, ricoh.
      - Workstations     : hostnames matching common PC naming conventions
                           (pc-, ws-, dt-, lt-, laptop, desktop).
      - Unknown          : anything not matching the above patterns.

    If the PTR lookup fails (NXDOMAIN, timeout, access denied, etc.) the
    Hostname field is set to '' and DeviceType is set to 'Unknown'. Lookup
    failures are intentionally silent to keep bulk operations fast; use
    -Verbose to see individual failure messages.

.PARAMETER IPAddress
    The IP address to resolve. IPv4 and IPv6 are both accepted. Required.

.OUTPUTS
    [PSCustomObject] with properties:
      IPAddress  [string]  -- the input IP address
      Hostname   [string]  -- resolved hostname, or '' on failure
      DeviceType [string]  -- Server / Phone / NetworkGear / Printer /
                              Workstation / Unknown
      DeviceMake [string]  -- Cisco / Polycom / Yealink / Grandstream /
                              Avaya / '' (empty when not a phone)
      DeviceModel[string]  -- model hint where detectable, else ''

.EXAMPLE
    Resolve-VBDNSHostname -IPAddress '10.0.0.50'

    Returns hostname from PTR lookup and classified device type.

.EXAMPLE
    $talkers | ForEach-Object { Resolve-VBDNSHostname -IPAddress $_.IPAddress }

    Batch resolves a list of IP addresses.

.NOTES
    Version  : 4.1.0
    Author   : VB
    Modified : 05-05-2026
    Scope    : Private helper -- called by Get-VBDNSTopTalkers
    Limits   : PTR resolution depends on correct reverse DNS zone configuration.
               Lookup failures return empty Hostname, DeviceType = Unknown.
               Bulk resolution over large IP sets may be slow if DNS is not
               cached -- use -ResolveHostnames selectively in Get-VBDNSTopTalkers.
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$IPAddress
    )

    $hostname   = ''
    $deviceType = 'Unknown'
    $deviceMake = ''
    $deviceModel = ''

    # --- PTR lookup ---
    try {
        $entry    = [System.Net.Dns]::GetHostEntry($IPAddress)
        $hostname = $entry.HostName
    }
    catch {
        Write-Verbose "PTR lookup failed for $IPAddress : $($_.Exception.Message)"
    }

    # --- Device classification (only if hostname resolved) ---
    if ($hostname -ne '') {
        $h = $hostname.ToLower()

        if ($h -match 'sep[0-9a-f]{12}') {
            # Cisco IP phone -- SEP + 12-hex MAC (e.g. SEP001122334455)
            $deviceType  = 'Phone'
            $deviceMake  = 'Cisco'
            $deviceModel = 'IP Phone'
        }
        elseif ($h -match '^(polycom|spip|vvx|trio)') {
            $deviceType  = 'Phone'
            $deviceMake  = 'Polycom'
            $deviceModel = if ($h -match '^vvx')  { 'VVX Series'  }
                           elseif ($h -match '^trio') { 'Trio Series' }
                           else { 'IP Phone' }
        }
        elseif ($h -match '^(yealink|ylk)') {
            $deviceType  = 'Phone'
            $deviceMake  = 'Yealink'
            $deviceModel = 'IP Phone'
        }
        elseif ($h -match '^(grandstream|gxp|gxv)') {
            $deviceType  = 'Phone'
            $deviceMake  = 'Grandstream'
            $deviceModel = 'IP Phone'
        }
        elseif ($h -match '^(avaya|96xx|j1[0-9][0-9])') {
            $deviceType  = 'Phone'
            $deviceMake  = 'Avaya'
            $deviceModel = 'IP Phone'
        }
        elseif ($h -match '(srv|server|\-svr|dc\d*\.|dc\d*$|\-dc\d|\-ad\d|sql|exch|ftp|mail|vcenter|esx|node)') {
            $deviceType = 'Server'
        }
        elseif ($h -match '(sw\d|\-sw\d|rtr|router|switch|\-ap\d|firewall|\-fw\d|\-gw\d|gateway|pix)') {
            $deviceType = 'NetworkGear'
        }
        elseif ($h -match '(prt|print|mfp|copier)' -or $h -match '^(hp|canon|konica|xerox|ricoh)') {
            $deviceType = 'Printer'
        }
        elseif ($h -match '(^pc\-|^ws\-|^dt\-|^lt\-|laptop|desktop)') {
            $deviceType = 'Workstation'
        }
    }

    [PSCustomObject]@{
        IPAddress   = $IPAddress
        Hostname    = $hostname
        DeviceType  = $deviceType
        DeviceMake  = $deviceMake
        DeviceModel = $deviceModel
    }
}


function Get-VBDNSLogSummary {
<#
.SYNOPSIS
    Extracts and filters DNS query records from Windows DNS Server debug log files.

.DESCRIPTION
    Reads one or more Windows DNS Server debug log files line by line using a
    StreamReader for memory-efficient processing of large files. Each line is
    matched against the standard DNS debug log format to extract: timestamp,
    protocol (UDP/TCP), direction (Rcv/Snd), client IP address, query type
    (A/AAAA/PTR/MX etc.), decoded query name (FQDN), and packet kind (Q/R).

    One PSCustomObject is emitted per unique combination of IPAddress,
    Direction, QueryType and QueryName. Deduplication is global across all
    input files when multiple files are provided. Results stream immediately
    through the pipeline as each unique record is parsed -- the file is never
    fully buffered in memory.

    Results can be filtered by protocol, direction, query type, time range,
    and private IP exclusion. All filters are optional and additive.

    Accepts single file paths, arrays of paths, and pipeline input from
    Get-ChildItem via the FullName property alias.

.PARAMETER InputFilePath
    One or more paths to DNS debug log files. Accepts pipeline input from
    Get-ChildItem via the FullName and Path aliases. Each file path is
    validated individually -- a missing file emits a failure object and
    processing continues with remaining files rather than terminating.

.PARAMETER Protocol
    Filters results by transport protocol. Valid values: All, UDP, TCP.
    Default is All.

.PARAMETER Direction
    Filters results by packet direction. Rcv captures incoming client queries.
    Snd captures outgoing server responses. Default is All.

    Note: your DNS Server debug logging configuration must have the relevant
    direction enabled for those entries to appear in the log file at all.

.PARAMETER QueryType
    Filters results to one or more DNS record types. Common values: A, AAAA,
    PTR, MX, CNAME, SOA, NS, SRV, TXT. When omitted all query types are
    returned. Accepts an array: -QueryType A,AAAA or -QueryType @('A','PTR').

.PARAMETER StartTime
    Returns only log entries timestamped at or after this datetime. Timestamps
    are parsed using InvariantCulture assuming the US date format produced by
    Windows DNS Server (M/d/yyyy h:mm:ss tt). Entries with unparseable
    timestamps pass through this filter unaffected.

.PARAMETER EndTime
    Returns only log entries timestamped at or before this datetime. See
    StartTime for notes on timestamp parsing and timezone assumptions.

.PARAMETER ExcludePrivateIPs
    Excludes RFC 1918 private IPv4 addresses: 10.0.0.0/8, 172.16.0.0/12,
    and 192.168.0.0/16. IPv6 addresses are never excluded by this switch.

.OUTPUTS
    [PSCustomObject] with properties in this order:
      ComputerName   -- name of the machine running this script
      SourceFile     -- full path of the source log file for this record
      LogTimestamp   -- [datetime] parsed from the log line; $null if unparseable
      Status         -- Success or Failed
      Error          -- $null on success; exception or reason message on failure
      IPAddress      -- client IPv4 or IPv6 address extracted from the log line
      IPVersion      -- IPv4 or IPv6
      IsPrivate      -- [bool] $true for RFC 1918 IPv4 ranges; always $false for IPv6
      Protocol       -- UDP or TCP
      Direction      -- Rcv (server received / client sent) or Snd (server sent)
      QueryType      -- DNS record type: A, AAAA, PTR, MX, CNAME, SOA, NS, SRV, TXT
      QueryName      -- decoded FQDN e.g. example.com; raw encoded value if decode fails
      PacketKind     -- Q (query), R (response), or combined value as it appears in log
      CollectionTime -- local datetime this function ran, format dd-MM-yyyy HH:mm:ss

.EXAMPLE
    Get-VBDNSLogSummary -InputFilePath 'C:\RTL\DNSLogs.txt'

    Extracts all unique DNS records from a single log file.

.EXAMPLE
    Get-VBDNSLogSummary -InputFilePath 'C:\RTL\DNSLogs.txt' -Protocol UDP -Direction Rcv -ExcludePrivateIPs

    Returns unique UDP incoming queries from public IP addresses only.

.EXAMPLE
    Get-VBDNSLogSummary -InputFilePath 'C:\RTL\DNSLogs.txt' -QueryType A,AAAA

    Returns only records where the queried type is A or AAAA.

.EXAMPLE
    Get-VBDNSLogSummary -InputFilePath 'C:\RTL\DNSLogs.txt' -StartTime '2025-04-14 08:00:00' -EndTime '2025-04-14 18:00:00'

    Returns records logged between 08:00 and 18:00 on 14 April 2025.

.EXAMPLE
    Get-ChildItem -Path 'C:\RTL\' -Filter 'DNSLogs*.txt' | Get-VBDNSLogSummary

    Processes all matching log files. Deduplication is global across all files.

.EXAMPLE
    $data = Get-VBDNSLogSummary -InputFilePath 'C:\RTL\DNSLogs.txt'
    $data | Export-VBDNSLogSummary -OutputFilePath 'C:\Intel\report.csv'
    $data | Format-VBDNSLogSummary
    $data | Format-VBDNSTopTalkers -Count 20
    $data | Format-VBDNSTopDomains -Count 20

    Collect once into a variable, then pass to multiple consumers without
    re-reading the log file. Note: for very large log files prefer piping
    directly to a single consumer to avoid buffering the full result set.

.NOTES
    Version   : 4.0.0
    Author    : VB
    Modified  : 05-05-2026
    Category  : DNS / Network Analysis
    Requires  : Windows DNS Server debug log (DNS Manager > Debug Logging)
    PS Ver    : 5.1
    Log fmt   : Standard Windows DNS Server debug log format
    Dedup key : IPAddress + Direction + QueryType + QueryName (global across files)
#>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('FullName', 'Path')]
        [string[]]$InputFilePath,

        [Parameter()]
        [ValidateSet('All', 'UDP', 'TCP')]
        [string]$Protocol = 'All',

        [Parameter()]
        [ValidateSet('All', 'Rcv', 'Snd')]
        [string]$Direction = 'All',

        [Parameter()]
        [string[]]$QueryType,

        [Parameter()]
        [datetime]$StartTime = [datetime]::MinValue,

        [Parameter()]
        [datetime]$EndTime = [datetime]::MaxValue,

        [Parameter()]
        [switch]$ExcludePrivateIPs
    )

    begin {
        # Dedup key: IPAddress|Direction|QueryType|QueryName -- global across all files
        $seenKeys = New-Object 'System.Collections.Generic.HashSet[string]'

        # Master pattern -- matches actual Windows DNS Server debug log PACKET lines
        # Actual format: Date Time ThreadID PACKET PacketID Protocol Direction IP Xid Q/R [Flags] QueryType QueryName
        # Example: 10/14/2025 12:25:16 AM 2C94 PACKET 00000000015117E0 UDP Rcv 10.209.1.165 7419 Q [0001 D NOERROR] A (8)automate(11)realtime-it(3)com(0)
        # EVENT lines (server start/stop) do not contain PACKET and are skipped automatically
        $logPattern = '^(\d{1,2}/\d{1,2}/\d{4})\s+(\d{1,2}:\d{2}:\d{2}\s+[AP]M)\s+\w+\s+PACKET\s+\w+\s+(UDP|TCP)\s+(Rcv|Snd)\s+(\S+)\s+\w+\s+(R?\s*[A-Z?])\s+\[([^\]]+)\]\s+(\S+)\s+(.+)$'

        # RFC 1918 private IPv4 ranges
        $privateIPv4Patterns = @('^10\.', '^172\.(1[6-9]|2[0-9]|3[0-1])\.', '^192\.168\.')

        # Timestamp format as produced by Windows DNS Server debug logging
        # Using InvariantCulture to avoid locale-sensitive parse failures
        $tsFormat  = 'M/d/yyyy h:mm:ss tt'
        $tsCulture = [System.Globalization.CultureInfo]::InvariantCulture
    }

    process {
        foreach ($filePath in $InputFilePath) {

            # Per-target validation -- emit failure object and continue rather than terminate
            if (-not (Test-Path -Path $filePath -PathType Leaf)) {
                Write-Warning "File not found: $filePath"
                [PSCustomObject]@{
                    ComputerName   = $env:COMPUTERNAME
                    SourceFile     = $filePath
                    LogTimestamp   = $null
                    Status         = 'Failed'
                    Error          = "File not found: $filePath"
                    IPAddress      = $null
                    IPVersion      = $null
                    IsPrivate      = $null
                    Protocol       = $null
                    Direction      = $null
                    QueryType      = $null
                    QueryName      = $null
                    PacketKind     = $null
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
                continue
            }

            $reader = $null

            try {
                Write-Verbose "Processing: $filePath"
                $fileSize  = (Get-Item -Path $filePath).Length
                $lineCount = 0
                $fileName  = Split-Path -Path $filePath -Leaf
                $reader    = New-Object System.IO.StreamReader -ArgumentList $filePath
                $line      = $reader.ReadLine()

                while ($null -ne $line) {
                    $lineCount++

                    # Update progress every 2000 lines -- calling every line would make Write-Progress the bottleneck
                    if ($lineCount % 2000 -eq 0) {
                        $pct = if ($fileSize -gt 0) { [int](($reader.BaseStream.Position / $fileSize) * 100) } else { 0 }
                        Write-Progress -Activity 'Processing DNS log' `
                                       -Status   $fileName `
                                       -PercentComplete $pct `
                                       -CurrentOperation "$lineCount lines read"
                    }

                    if ($line -match $logPattern) {

                        # Step 1 -- Parse timestamp; $null on failure, never crash
                        $logTimestamp = $null
                        $tsStr        = "$($Matches[1]) $($Matches[2])".Trim()
                        try {
                            $logTimestamp = [datetime]::ParseExact($tsStr, $tsFormat, $tsCulture)
                        }
                        catch {
                            Write-Debug "Timestamp parse failed: $tsStr"
                        }

                        # Step 2 -- Extract all captured fields
                        $parsedProtocol  = $Matches[3]
                        $parsedDirection = $Matches[4]
                        $parsedIP        = $Matches[5].Trim()
                        $parsedKind      = $Matches[6].Trim()
                        $parsedQueryType = $Matches[8].Trim()
                        $rawQueryName    = $Matches[9].Trim()

                        # Step 3 -- Determine IP version and private status
                        $ipVersion = if ($parsedIP -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') { 'IPv4' } else { 'IPv6' }

                        $isPrivate = $false
                        if ($ipVersion -eq 'IPv4') {
                            foreach ($privatePattern in $privateIPv4Patterns) {
                                if ($parsedIP -match $privatePattern) {
                                    $isPrivate = $true
                                    break
                                }
                            }
                        }

                        # Step 4 -- Decode DNS label to FQDN
                        $queryName = ConvertTo-VBDNSLabel -EncodedLabel $rawQueryName

                        # Step 5 -- Apply all filters
                        $include = $true

                        if ($Protocol -ne 'All' -and $parsedProtocol -ne $Protocol) {
                            $include = $false
                        }

                        if ($include -and $Direction -ne 'All' -and $parsedDirection -ne $Direction) {
                            $include = $false
                        }

                        if ($include -and $null -ne $QueryType -and $QueryType.Count -gt 0 -and $QueryType -notcontains $parsedQueryType) {
                            $include = $false
                        }

                        if ($include -and $ExcludePrivateIPs -and $isPrivate) {
                            $include = $false
                        }

                        if ($include -and $null -ne $logTimestamp) {
                            if ($logTimestamp -lt $StartTime -or $logTimestamp -gt $EndTime) {
                                $include = $false
                            }
                        }

                        # Step 6 -- Deduplicate and emit
                        if ($include) {
                            $dedupKey = "$parsedIP|$parsedDirection|$parsedQueryType|$queryName"

                            if ($seenKeys.Add($dedupKey)) {
                                [PSCustomObject]@{
                                    ComputerName   = $env:COMPUTERNAME
                                    SourceFile     = $filePath
                                    LogTimestamp   = $logTimestamp
                                    Status         = 'Success'
                                    Error          = $null
                                    IPAddress      = $parsedIP
                                    IPVersion      = $ipVersion
                                    IsPrivate      = $isPrivate
                                    Protocol       = $parsedProtocol
                                    Direction      = $parsedDirection
                                    QueryType      = $parsedQueryType
                                    QueryName      = $queryName
                                    PacketKind     = $parsedKind
                                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                                }
                            }
                        }
                    }

                    $line = $reader.ReadLine()
                }
            }
            catch {
                [PSCustomObject]@{
                    ComputerName   = $env:COMPUTERNAME
                    SourceFile     = $filePath
                    LogTimestamp   = $null
                    Status         = 'Failed'
                    Error          = $_.Exception.Message
                    IPAddress      = $null
                    IPVersion      = $null
                    IsPrivate      = $null
                    Protocol       = $null
                    Direction      = $null
                    QueryType      = $null
                    QueryName      = $null
                    PacketKind     = $null
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
            finally {
                Write-Progress -Activity 'Processing DNS log' -Completed
                if ($null -ne $reader) { $reader.Close() }
            }
        }
    }
}


function Format-VBDNSLogSummary {
<#
.SYNOPSIS
    Displays a statistical summary from Get-VBDNSLogSummary output to the console.

.DESCRIPTION
    Accepts pipeline input from Get-VBDNSLogSummary and buffers all records
    before computing and displaying a formatted statistics block. Output covers
    unique record and IP counts, protocol split (UDP/TCP), direction split
    (Rcv/Snd), IP version breakdown, private IP count, source file count, and
    any error records encountered during parsing.

    Nothing is returned to the pipeline. All output is via Write-Host.

.PARAMETER InputObject
    One or more PSCustomObjects from Get-VBDNSLogSummary. Accepts pipeline input.

.PARAMETER OutputFile
    Optional. When provided, the output file path is appended to the summary
    display for reference. Has no effect on what is written to disk.

.EXAMPLE
    Get-VBDNSLogSummary -InputFilePath 'C:\RTL\DNSLogs.txt' | Format-VBDNSLogSummary

    Displays a full statistics summary to the console.

.EXAMPLE
    $data | Format-VBDNSLogSummary -OutputFile 'C:\Intel\report.csv'

    Displays statistics and includes the output file path in the summary.

.NOTES
    Version  : 4.0.0
    Author   : VB
    Modified : 05-05-2026
    Category : DNS / Network Analysis
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [PSCustomObject[]]$InputObject,

        [Parameter()]
        [string]$OutputFile
    )

    begin   { $all = [System.Collections.Generic.List[PSCustomObject]]::new() }
    process { foreach ($item in $InputObject) { $all.Add($item) } }

    end {
        $ok   = $all | Where-Object { $_.Status -eq 'Success' }
        $fail = $all | Where-Object { $_.Status -eq 'Failed' }

        $uniqueIPs    = ($ok | Select-Object -ExpandProperty IPAddress -Unique)
        $uniqueIPsCount = if ($null -eq $uniqueIPs) { 0 } else { @($uniqueIPs).Count }

        $sourcesArr   = ($ok | Select-Object -ExpandProperty SourceFile -Unique)
        $sourcesCount = if ($null -eq $sourcesArr) { 0 } else { @($sourcesArr).Count }

        Write-Host ''
        Write-Host '--------- DNS Log Analysis Summary ---------' -ForegroundColor Cyan
        Write-Host "Unique Records       : $(@($ok).Count)"
        Write-Host "Unique Client IPs    : $uniqueIPsCount"
        Write-Host ''
        Write-Host "UDP                  : $(@($ok | Where-Object { $_.Protocol  -eq 'UDP' }).Count)"
        Write-Host "TCP                  : $(@($ok | Where-Object { $_.Protocol  -eq 'TCP' }).Count)"
        Write-Host ''
        Write-Host "Incoming (Rcv)       : $(@($ok | Where-Object { $_.Direction -eq 'Rcv' }).Count)"
        Write-Host "Outgoing (Snd)       : $(@($ok | Where-Object { $_.Direction -eq 'Snd' }).Count)"
        Write-Host ''
        Write-Host "IPv4                 : $(@($ok | Where-Object { $_.IPVersion -eq 'IPv4' }).Count)"
        Write-Host "IPv6                 : $(@($ok | Where-Object { $_.IPVersion -eq 'IPv6' }).Count)"
        Write-Host "Private IPs          : $(@($ok | Where-Object { $_.IsPrivate -eq $true  }).Count)"
        Write-Host ''
        Write-Host "Source Files         : $sourcesCount"

        if ($OutputFile) {
            Write-Host "Output File          : $OutputFile"
        }

        if (@($fail).Count -gt 0) {
            Write-Host ''
            Write-Host "Parse Errors         : $(@($fail).Count)" -ForegroundColor Yellow
            foreach ($f in $fail) {
                Write-Host "  $($f.SourceFile) -- $($f.Error)" -ForegroundColor Yellow
            }
        }

        Write-Host '--------------------------------------------' -ForegroundColor Cyan
        Write-Host ''
    }
}


function Get-VBDNSTopTalkers {
<#
.SYNOPSIS
    Returns the top N client IP addresses by DNS query volume as pipeline objects.

.DESCRIPTION
    Accepts pipeline input from Get-VBDNSLogSummary, groups all successful records
    by client IP address, sorts by count descending, and emits one PSCustomObject
    per ranked entry. The output is pipeline-friendly: it can be piped to
    Format-VBDNSTopTalkers for display, or to Export-Csv for file export, or to
    both via Tee-Object.

    When -ResolveHostnames is specified, each unique IP address is passed to
    Resolve-VBDNSHostname (private helper) which performs a PTR reverse DNS lookup
    and attempts to classify the device as Server, Phone, NetworkGear, Printer,
    Workstation, or Unknown. PTR lookup failure is silent; unresolved IPs get an
    empty Hostname and DeviceType = 'Unknown'. Bulk resolution may be slow if the
    reverse DNS zone is sparse -- use this switch selectively.

.PARAMETER InputObject
    One or more PSCustomObjects from Get-VBDNSLogSummary. Accepts pipeline input.
    Records with Status -ne 'Success' are silently excluded.

.PARAMETER Count
    Maximum number of ranked entries to return. Default is 10. Minimum is 1.
    Actual results may be fewer if the dataset has fewer unique IPs.

.PARAMETER ResolveHostnames
    When set, performs a reverse DNS PTR lookup for each unique IP address and
    classifies the device type. Adds Hostname, DeviceType, DeviceMake, and
    DeviceModel properties to the output. Without this switch those properties
    are always empty strings to keep the output shape consistent.

.OUTPUTS
    [PSCustomObject] per ranked entry with properties:
      Rank           [int]    -- 1-based rank position (1 = busiest)
      Count          [int]    -- number of unique DNS records for this IP
      TotalInDataset [int]    -- total unique records across all IPs (denominator)
      IPAddress      [string] -- client IP address
      Hostname       [string] -- resolved PTR hostname, or '' if not resolved
      DeviceType     [string] -- Server / Phone / NetworkGear / Printer /
                                 Workstation / Unknown
      DeviceMake     [string] -- vendor name for phones (Cisco/Polycom/etc.), else ''
      DeviceModel    [string] -- model hint where detectable, else ''

.EXAMPLE
    Get-VBDNSLogSummary -InputFilePath 'C:\RTL\DNSLogs.txt' | Get-VBDNSTopTalkers

    Returns the top 10 clients as objects. Pipe to Format-VBDNSTopTalkers to display.

.EXAMPLE
    Get-VBDNSLogSummary -InputFilePath 'C:\RTL\DNSLogs.txt' |
        Get-VBDNSTopTalkers -Count 500 -ResolveHostnames |
        Export-Csv 'C:\Intel\TopTalkers.csv' -NoTypeInformation -Encoding UTF8

    Exports the top 500 clients with PTR resolution to CSV.

.EXAMPLE
    $data = Get-VBDNSLogSummary -InputFilePath 'C:\RTL\DNSLogs.txt'
    $talkers = $data | Get-VBDNSTopTalkers -Count 50 -ResolveHostnames
    $talkers | Format-VBDNSTopTalkers
    $talkers | Export-Csv 'C:\Intel\TopTalkers.csv' -NoTypeInformation -Encoding UTF8

    Buffer results, display on screen, and export to CSV in one pass.

.NOTES
    Version  : 4.1.0
    Author   : VB
    Modified : 05-05-2026
    Category : DNS / Network Analysis
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [PSCustomObject[]]$InputObject,

        [Parameter()]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$Count = 10,

        [Parameter()]
        [switch]$ResolveHostnames
    )

    begin   { $all = [System.Collections.Generic.List[PSCustomObject]]::new() }
    process { foreach ($item in $InputObject) { $all.Add($item) } }

    end {
        $ok = @($all | Where-Object { $_.Status -eq 'Success' })

        if ($ok.Count -eq 0) {
            Write-Verbose 'Get-VBDNSTopTalkers: no successful records in input.'
            return
        }

        $grouped = $ok |
            Group-Object -Property IPAddress |
            Sort-Object  -Property Count -Descending

        $totalInDataset = $ok.Count
        $rank = 0

        foreach ($group in ($grouped | Select-Object -First $Count)) {
            $rank++
            $hostname    = ''
            $deviceType  = 'Unknown'
            $deviceMake  = ''
            $deviceModel = ''

            if ($ResolveHostnames) {
                $resolved    = Resolve-VBDNSHostname -IPAddress $group.Name
                $hostname    = $resolved.Hostname
                $deviceType  = $resolved.DeviceType
                $deviceMake  = $resolved.DeviceMake
                $deviceModel = $resolved.DeviceModel
            }

            [PSCustomObject]@{
                Rank           = $rank
                Count          = $group.Count
                TotalInDataset = $totalInDataset
                IPAddress      = $group.Name
                Hostname       = $hostname
                DeviceType     = $deviceType
                DeviceMake     = $deviceMake
                DeviceModel    = $deviceModel
            }
        }
    }
}



function Get-VBDNSTopDomains {
<#
.SYNOPSIS
    Returns the top N most queried domain names as pipeline objects.

.DESCRIPTION
    Accepts pipeline input from Get-VBDNSLogSummary, groups all successful records
    by decoded query name (FQDN), sorts by count descending, and emits one
    PSCustomObject per ranked entry. The output is pipeline-friendly: it can be
    piped to Format-VBDNSTopDomains for display, or to Export-Csv for file export,
    or to both via Tee-Object.

.PARAMETER InputObject
    One or more PSCustomObjects from Get-VBDNSLogSummary. Accepts pipeline input.
    Records with Status -ne 'Success' are silently excluded.

.PARAMETER Count
    Maximum number of ranked entries to return. Default is 10. Minimum is 1.
    Actual results may be fewer if the dataset has fewer unique query names.

.OUTPUTS
    [PSCustomObject] per ranked entry with properties:
      Rank           [int]    -- 1-based rank position (1 = most queried)
      Count          [int]    -- number of unique DNS records for this domain
      TotalInDataset [int]    -- total unique records across all domains (denominator)
      QueryName      [string] -- the decoded FQDN

.EXAMPLE
    Get-VBDNSLogSummary -InputFilePath 'C:\RTL\DNSLogs.txt' | Get-VBDNSTopDomains

    Returns the top 10 most queried domains as objects.

.EXAMPLE
    Get-VBDNSLogSummary -InputFilePath 'C:\RTL\DNSLogs.txt' |
        Get-VBDNSTopDomains -Count 500 |
        Export-Csv 'C:\Intel\TopDomains.csv' -NoTypeInformation -Encoding UTF8

    Exports the top 500 queried domains to CSV.

.EXAMPLE
    $data = Get-VBDNSLogSummary -InputFilePath 'C:\RTL\DNSLogs.txt'
    $domains = $data | Get-VBDNSTopDomains -Count 50
    $domains | Format-VBDNSTopDomains
    $domains | Export-Csv 'C:\Intel\TopDomains.csv' -NoTypeInformation -Encoding UTF8

    Buffer results, display on screen, and export to CSV in one pass.

.EXAMPLE
    Get-VBDNSLogSummary -InputFilePath 'C:\RTL\DNSLogs.txt' -QueryType PTR |
        Get-VBDNSTopDomains -Count 20

    Returns the 20 most reverse-queried IP strings (PTR records only).

.NOTES
    Version  : 4.1.0
    Author   : VB
    Modified : 05-05-2026
    Category : DNS / Network Analysis
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [PSCustomObject[]]$InputObject,

        [Parameter()]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$Count = 10
    )

    begin   { $all = [System.Collections.Generic.List[PSCustomObject]]::new() }
    process { foreach ($item in $InputObject) { $all.Add($item) } }

    end {
        $ok = @($all | Where-Object { $_.Status -eq 'Success' })

        if ($ok.Count -eq 0) {
            Write-Verbose 'Get-VBDNSTopDomains: no successful records in input.'
            return
        }

        $grouped = $ok |
            Group-Object -Property QueryName |
            Sort-Object  -Property Count -Descending

        $totalInDataset = $ok.Count
        $rank = 0

        foreach ($group in ($grouped | Select-Object -First $Count)) {
            $rank++
            [PSCustomObject]@{
                Rank           = $rank
                Count          = $group.Count
                TotalInDataset = $totalInDataset
                QueryName      = $group.Name
            }
        }
    }
}


function Format-VBDNSTopTalkers {
<#
.SYNOPSIS
    Displays pre-aggregated top-talker results from Get-VBDNSTopTalkers.

.DESCRIPTION
    Accepts pipeline input from Get-VBDNSTopTalkers and renders a colour-coded
    ranked table to the console. Nothing is returned to the pipeline -- this
    function is purely presentational.

    The header line shows how many entries are being displayed and the total
    number of unique records in the source dataset (e.g. "Showing 50 of 847").
    Each row shows: rank (#), query count, IP address, and -- when hostname
    resolution was requested -- Hostname, DeviceType, DeviceMake, and DeviceModel.

    To both display and export in one pass, buffer Get-VBDNSTopTalkers into a
    variable first, then pipe to Format-VBDNSTopTalkers and Export-Csv separately.

.PARAMETER InputObject
    One or more PSCustomObjects from Get-VBDNSTopTalkers. Accepts pipeline input.

.EXAMPLE
    Get-VBDNSLogSummary -InputFilePath 'C:\RTL\DNSLogs.txt' |
        Get-VBDNSTopTalkers | Format-VBDNSTopTalkers

    Displays the top 10 clients on screen.

.EXAMPLE
    $talkers = Get-VBDNSLogSummary -InputFilePath 'C:\RTL\DNSLogs.txt' |
                   Get-VBDNSTopTalkers -Count 50 -ResolveHostnames
    $talkers | Format-VBDNSTopTalkers
    $talkers | Export-Csv 'C:\Intel\TopTalkers.csv' -NoTypeInformation -Encoding UTF8

    Display and export from the same buffered result.

.NOTES
    Version  : 4.1.0
    Author   : VB
    Modified : 05-05-2026
    Category : DNS / Network Analysis
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [PSCustomObject[]]$InputObject
    )

    begin   { $rows = [System.Collections.Generic.List[PSCustomObject]]::new() }
    process { foreach ($item in $InputObject) { $rows.Add($item) } }

    end {
        if ($rows.Count -eq 0) {
            Write-Host ''
            Write-Host '  [Top Talkers] No records to display.' -ForegroundColor Yellow
            Write-Host ''
            return
        }

        # Determine if hostname columns are populated
        $hasHostnames = ($rows | Where-Object { $_.Hostname -ne '' }).Count -gt 0

        $showing = $rows.Count
        $total   = $rows[0].TotalInDataset

        Write-Host ''
        Write-Host "--------- Top DNS Clients by Query Count   [Showing $showing of $total unique records] ---------" -ForegroundColor Cyan

        if ($hasHostnames) {
            Write-Host ('  {0,4}  {1,-7}  {2,-40}  {3,-45}  {4,-12}  {5,-12}  {6}' -f
                '#', 'Count', 'IP Address', 'Hostname', 'DeviceType', 'DeviceMake', 'DeviceModel')
            Write-Host ('  {0,4}  {1,-7}  {2,-40}  {3,-45}  {4,-12}  {5,-12}  {6}' -f
                '----', '-------', ('-' * 40), ('-' * 45), '------------', '------------', '------------')
            foreach ($row in $rows) {
                Write-Host ('  {0,4}  {1,-7}  {2,-40}  {3,-45}  {4,-12}  {5,-12}  {6}' -f
                    $row.Rank, $row.Count, $row.IPAddress,
                    $row.Hostname, $row.DeviceType, $row.DeviceMake, $row.DeviceModel)
            }
        }
        else {
            Write-Host ('  {0,4}  {1,-7}  {2}' -f '#', 'Count', 'IP Address')
            Write-Host ('  {0,4}  {1,-7}  {2}' -f '----', '-------', ('-' * 40))
            foreach ($row in $rows) {
                Write-Host ('  {0,4}  {1,-7}  {2}' -f $row.Rank, $row.Count, $row.IPAddress)
            }
        }

        Write-Host ('-' * 80) -ForegroundColor Cyan
        Write-Host ''
    }
}


function Format-VBDNSTopDomains {
<#
.SYNOPSIS
    Displays pre-aggregated top-domain results from Get-VBDNSTopDomains.

.DESCRIPTION
    Accepts pipeline input from Get-VBDNSTopDomains and renders a colour-coded
    ranked table to the console. Nothing is returned to the pipeline -- this
    function is purely presentational.

    The header line shows how many entries are being displayed and the total
    number of unique records in the source dataset (e.g. "Showing 50 of 847").
    Each row shows: rank (#), query count, and the decoded FQDN (QueryName).

    To both display and export in one pass, buffer Get-VBDNSTopDomains into a
    variable first, then pipe to Format-VBDNSTopDomains and Export-Csv separately.

.PARAMETER InputObject
    One or more PSCustomObjects from Get-VBDNSTopDomains. Accepts pipeline input.

.EXAMPLE
    Get-VBDNSLogSummary -InputFilePath 'C:\RTL\DNSLogs.txt' |
        Get-VBDNSTopDomains | Format-VBDNSTopDomains

    Displays the top 10 queried domains on screen.

.EXAMPLE
    $domains = Get-VBDNSLogSummary -InputFilePath 'C:\RTL\DNSLogs.txt' |
                   Get-VBDNSTopDomains -Count 50
    $domains | Format-VBDNSTopDomains
    $domains | Export-Csv 'C:\Intel\TopDomains.csv' -NoTypeInformation -Encoding UTF8

    Display and export from the same buffered result.

.NOTES
    Version  : 4.1.0
    Author   : VB
    Modified : 05-05-2026
    Category : DNS / Network Analysis
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [PSCustomObject[]]$InputObject
    )

    begin   { $rows = [System.Collections.Generic.List[PSCustomObject]]::new() }
    process { foreach ($item in $InputObject) { $rows.Add($item) } }

    end {
        if ($rows.Count -eq 0) {
            Write-Host ''
            Write-Host '  [Top Domains] No records to display.' -ForegroundColor Yellow
            Write-Host ''
            return
        }

        $showing = $rows.Count
        $total   = $rows[0].TotalInDataset

        Write-Host ''
        Write-Host "--------- Top Queried Domain Names   [Showing $showing of $total unique records] ---------" -ForegroundColor Cyan
        Write-Host ('  {0,4}  {1,-7}  {2}' -f '#', 'Count', 'Domain Name')
        Write-Host ('  {0,4}  {1,-7}  {2}' -f '----', '-------', ('-' * 60))

        foreach ($row in $rows) {
            Write-Host ('  {0,4}  {1,-7}  {2}' -f $row.Rank, $row.Count, $row.QueryName)
        }

        Write-Host ('-' * 80) -ForegroundColor Cyan
        Write-Host ''
    }
}


function Export-VBDNSLogSummary {
<#
.SYNOPSIS
    Exports DNS log records from Get-VBDNSLogSummary to a file on disk.

.DESCRIPTION
    Accepts pipeline input from Get-VBDNSLogSummary and writes records to disk
    in one of two formats selected by the -Format parameter.

    CSV (default): All 14 output object properties are written as a proper
    comma-separated file using Export-Csv with UTF-8 encoding and no type
    information header row. Suitable for import into Excel, Splunk, SIEM tools,
    or any downstream script expecting structured data.

    Text (legacy): One IP address per line, sorted ascending. When
    -IncludeProtocol is specified, each line is prefixed with the protocol
    separated by a comma (Protocol,IPAddress). When -IncludeStatistics is
    specified, a statistics block is appended after the IP list. This format
    preserves backward compatibility with scripts that consumed the v3.x output.

    -IncludeStatistics appends a statistics block to Text format output only.
    For CSV output the full data is already present in every row; statistics
    can be computed from the CSV data by the consumer.

    The output file is always created or overwritten. Nothing is returned to
    the pipeline.

.PARAMETER InputObject
    One or more PSCustomObjects from Get-VBDNSLogSummary. Accepts pipeline input.

.PARAMETER OutputFilePath
    Full path to the output file. Created or overwritten. Required.

.PARAMETER Format
    Output format. CSV (default) writes all fields as a comma-separated file.
    Text writes one IP address per line (legacy format compatible with v3.x).

.PARAMETER IncludeProtocol
    Text format only. Prepends the protocol (UDP or TCP) to each line as
    Protocol,IPAddress. Has no effect when Format is CSV since protocol is
    always present as a named column in CSV output.

.PARAMETER IncludeStatistics
    Text format only. Appends a summary statistics block after the IP list.
    Has no effect when Format is CSV.

.OUTPUTS
    No pipeline output. Always produces a file at OutputFilePath.

.EXAMPLE
    Get-VBDNSLogSummary -InputFilePath 'C:\RTL\DNSLogs.txt' |
        Export-VBDNSLogSummary -OutputFilePath 'C:\Intel\report.csv'

    Exports all fields as CSV. Default format.

.EXAMPLE
    Get-VBDNSLogSummary -InputFilePath 'C:\RTL\DNSLogs.txt' |
        Export-VBDNSLogSummary -OutputFilePath 'C:\Intel\IPs.txt' -Format Text -IncludeStatistics

    Exports IP list as plain text with a statistics block appended.

.EXAMPLE
    Get-VBDNSLogSummary -InputFilePath 'C:\RTL\DNSLogs.txt' |
        Export-VBDNSLogSummary -OutputFilePath 'C:\Intel\IPs.txt' -Format Text -IncludeProtocol

    Exports Protocol,IPAddress pairs as plain text (v3.x legacy format).

.EXAMPLE
    Get-ChildItem -Path 'C:\RTL\' -Filter '*.txt' |
        Get-VBDNSLogSummary -ExcludePrivateIPs |
        Export-VBDNSLogSummary -OutputFilePath 'C:\Intel\PublicIPs.csv'

    Processes multiple log files and exports deduplicated public-IP records as CSV.

.NOTES
    Version  : 4.0.0
    Author   : VB
    Modified : 05-05-2026
    Category : DNS / Network Analysis
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [PSCustomObject[]]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$OutputFilePath,

        [Parameter()]
        [ValidateSet('CSV', 'Text')]
        [string]$Format = 'CSV',

        [Parameter()]
        [switch]$IncludeProtocol,

        [Parameter()]
        [switch]$IncludeStatistics
    )

    begin   { $all = [System.Collections.Generic.List[PSCustomObject]]::new() }
    process { foreach ($item in $InputObject) { $all.Add($item) } }

    end {
        try {
            $ok = @($all | Where-Object { $_.Status -eq 'Success' } | Sort-Object -Property IPAddress)

            if ($Format -eq 'CSV') {
                # CSV -- all 14 properties, no type header, UTF-8 encoding
                $ok | Export-Csv -Path $OutputFilePath -NoTypeInformation -Encoding UTF8
            }
            else {
                # Text -- IP list, optional protocol prefix, optional statistics block
                $lines = [System.Collections.Generic.List[string]]::new()

                if ($IncludeProtocol) {
                    $lines.Add('Protocol,IPAddress')
                    foreach ($item in $ok) {
                        $lines.Add("$($item.Protocol),$($item.IPAddress)")
                    }
                }
                else {
                    foreach ($item in $ok) {
                        $lines.Add($item.IPAddress)
                    }
                }

                if ($IncludeStatistics) {
                    $lines.Add('')
                    $lines.Add('--------- DNS Log Analysis Statistics ---------')
                    $lines.Add("Unique Records       : $($ok.Count)")
                    $lines.Add("UDP                  : $(@($ok | Where-Object { $_.Protocol  -eq 'UDP' }).Count)")
                    $lines.Add("TCP                  : $(@($ok | Where-Object { $_.Protocol  -eq 'TCP' }).Count)")
                    $lines.Add("Incoming (Rcv)       : $(@($ok | Where-Object { $_.Direction -eq 'Rcv' }).Count)")
                    $lines.Add("Outgoing (Snd)       : $(@($ok | Where-Object { $_.Direction -eq 'Snd' }).Count)")
                    $lines.Add("IPv4                 : $(@($ok | Where-Object { $_.IPVersion -eq 'IPv4' }).Count)")
                    $lines.Add("IPv6                 : $(@($ok | Where-Object { $_.IPVersion -eq 'IPv6' }).Count)")
                    $lines.Add("Private IPs          : $(@($ok | Where-Object { $_.IsPrivate -eq $true  }).Count)")
                }

                $lines | Set-Content -Path $OutputFilePath -Encoding UTF8
            }

            Write-Verbose "Exported $($ok.Count) records to: $OutputFilePath ($Format)"
        }
        catch {
            Write-Warning "Failed to write output file: $($_.Exception.Message)"
        }
    }
}

<#
# --- MAIN LOGIC ---
# Note: $dnsResults buffers all records in memory so the result can be passed
# to multiple consumers without re-reading the log. For very large log files
# (millions of lines) pipe directly to a single consumer instead.
#
# New in v4.1.0:
#   Get-VBDNSTopTalkers  -- returns PSCustomObjects (pipeline + Export-Csv ready)
#   Get-VBDNSTopDomains  -- returns PSCustomObjects (pipeline + Export-Csv ready)
#   Format-VBDNSTopTalkers -- display only, accepts Get-VBDNSTopTalkers output
#   Format-VBDNSTopDomains -- display only, accepts Get-VBDNSTopDomains output
#
# Workflow: Get- -> buffer -> Format- (screen) + Export-Csv (file)

# Step 1 -- Extract unique DNS records from log file
$dnsResults = Get-VBDNSLogSummary -InputFilePath $LOG_INPUT -Verbose

# Step 2 -- Export raw records to CSV
$dnsResults | Export-VBDNSLogSummary -OutputFilePath $LOG_OUTPUT

# Step 3 -- Display statistics summary to console
$dnsResults | Format-VBDNSLogSummary -OutputFile $LOG_OUTPUT

# Step 4 -- Aggregate top talkers: display on screen + export to CSV
$topTalkers = $dnsResults | Get-VBDNSTopTalkers -Count 50 -ResolveHostnames
$topTalkers | Format-VBDNSTopTalkers
$topTalkers | Export-Csv -Path 'C:\Intel\TopTalkers.csv' -NoTypeInformation -Encoding UTF8

# Step 5 -- Aggregate top domains: display on screen + export to CSV
$topDomains = $dnsResults | Get-VBDNSTopDomains -Count 50
$topDomains | Format-VBDNSTopDomains
$topDomains | Export-Csv -Path 'C:\Intel\TopDomains.csv' -NoTypeInformation -Encoding UTF8
#>
