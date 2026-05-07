# ============================================================
# FUNCTION : Invoke-VBDNSLogParser
# VERSION  : 1.0.0
# CHANGED  : 2026-05-07 -- Initial build
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : High-performance DNS debug log parsing engine
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Parses a Windows DNS debug log file and returns all PACKET records as a buffer of object arrays.

.DESCRIPTION
    The hot-loop parsing engine for the DNSLogDB module. Reads the log file using
    StreamReader (not Get-Content) and applies a multi-stage optimised pipeline:

    Stage 1  -- '*PACKET*' string pre-filter (skips ~75% of all lines before regex)
    Stage 2  -- Compiled regex match (reused across all lines, not recompiled per call)
    Stage 3  -- IP classification via hashtable cache (Test-VBPrivateIP called once per unique IP)
    Stage 4  -- DNS wire-format decode via ConvertFrom-VBDNSName (IndexOf/Substring, no regex)
    Stage 5  -- RCODE to Status/Error mapping (inline string compare, no function call)
    Stage 6  -- Append object[] to pre-sized List (no PSCustomObject, no reflection)

    The returned List[object[]] is consumed directly by Invoke-VBBulkInsert for SQLite
    parameter binding. No PSCustomObject is ever created in the import path.

    object[] column order (fixed contract with Invoke-VBBulkInsert):
        0   LogDateTime   ISO8601 string yyyy-MM-ddTHH:mm:ss
        1   LogDate       Raw date from log (M/d/yyyy)
        2   LogTime       Raw time from log (h:mm:ss tt)
        3   ThreadId      Hex thread ID
        4   PacketId      Internal packet identifier (hex)
        5   Protocol      UDP or TCP
        6   Direction     Rcv or Snd
        7   IPAddress     Remote client IP
        8   IPVersion     IPv4 or IPv6
        9   IsPrivate     1 or 0 (integer)
       10   TransactionId DNS Xid (hex)
       11   PacketKind    Q or R
       12   Opcode        Q, N, U, or ?
       13   FlagsHex      e.g. 8081
       14   FlagsChar     e.g. DR
       15   ResponseCode  NOERROR, NXDOMAIN, SERVFAIL, REFUSED, FORMERR, NOTIMPL
       16   Status        Success or Error
       17   Error         RCODE string if Error, empty string if Success
       18   QueryType     A, AAAA, MX, PTR, SOA, NS, CNAME, TXT, SRV, ANY...
       19   QueryName     Decoded FQDN
       20   SourceFile    Full path of the log file

    WARNING: The column order is a contract. If you add or reorder columns, update
    Invoke-VBBulkInsert parameter binding at the same time or data will be silently
    inserted into the wrong columns.

.PARAMETER FilePath
    Full path to the Windows DNS debug log file to parse.

.PARAMETER ExcludePrivateIPs
    If specified, PACKET lines where the client IP is classified as private are
    skipped entirely and not added to the buffer. Reduces storage on busy internal servers.

.OUTPUTS
    [System.Collections.Generic.List[object[]]] One entry per parsed PACKET line.

.NOTES
    Version  : 1.0.0
    Author   : Vibhu Bhatnagar
    Modified : 2026-05-07
    Category : Private
    Called by: Import-VBDNSLog (inside ForEach-Object -Parallel thread)

    Requires: ConvertFrom-VBDNSName, Test-VBPrivateIP (both Private functions)
    Requires: PowerShell 7.0+ (module requirement — not enforced here)
#>

function Invoke-VBDNSLogParser {
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[object[]]])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [switch]$ExcludePrivateIPs
    )

    # ------------------------------------------------------------------
    # BEGIN -- compile regex once, initialise cache and buffer
    # ------------------------------------------------------------------

    # Compiled regex -- reused across all lines, not recompiled per call
    #
    # Actual Windows DNS debug log PACKET line format:
    #   DATE TIME THREADID PACKET PACKETID PROTOCOL DIRECTION IP XID PACKETKIND [FLAGSHEX FLAGSCHAR RCODE] QUERYTYPE QUERYNAME
    #
    # Example:
    #   10/14/2025 12:25:16 AM 2C94 PACKET  00000000015117E0 UDP Rcv 10.209.1.165    7419   Q [0001   D   NOERROR] A      (8)automate(11)realtime-it(3)com(0)
    #
    # Groups:
    #   1  Protocol     UDP or TCP
    #   2  Direction    Rcv or Snd
    #   3  IPAddress    remote client IP (IPv4 or IPv6)
    #   4  TransactionId  Xid (hex)
    #   5  PacketKind   Q (query) or R (response) -- blank means query in some versions, treat blank as Q
    #   6  FlagsHex     hex flags inside brackets, e.g. 0001 or 8081
    #   7  FlagsChar    flag characters inside brackets, e.g. D or DR (may be empty)
    #   8  ResponseCode RCODE inside brackets, e.g. NOERROR NXDOMAIN
    #   9  QueryType    A AAAA MX PTR etc.
    #  10  RawQueryName wire-format name, e.g. (8)automate(11)realtime-it(3)com(0)
    $packetRegex = [regex]::new(
        '\s+(UDP|TCP)\s+(Rcv|Snd)\s+([\d.]+|[0-9a-fA-F:]+)\s+([0-9a-fA-F]+)\s+(Q|R|)\s*\[([0-9a-fA-F]+)\s+([A-Z]*)\s+(\w+)\]\s+(\w+)\s+(.+)',
        [System.Text.RegularExpressions.RegexOptions]::Compiled
    )

    # Pre-size the buffer for a typical 450MB file (~1.5M PACKET lines)
    $buffer  = [System.Collections.Generic.List[object[]]]::new(500000)

    # IP classification cache -- Test-VBPrivateIP called once per unique IP
    $ipCache = [System.Collections.Generic.Dictionary[string, object]]::new()

    # Culture for ParseExact
    $culture = [System.Globalization.CultureInfo]::InvariantCulture

    # Counters for progress reporting
    $lineCount   = 0
    $errorCount  = 0
    $packetCount = 0

    # File size for percentage progress (StreamReader exposes BaseStream.Position)
    $fileSize  = (Get-Item $FilePath).Length
    $fileName  = Split-Path $FilePath -Leaf

    Write-Verbose "Invoke-VBDNSLogParser: Starting parse of '$FilePath' ($([math]::Round($fileSize/1MB,1)) MB)"

    # ------------------------------------------------------------------
    # MAIN LOOP -- StreamReader line-by-line, flat memory usage
    # ------------------------------------------------------------------

    $reader = $null
    try {
        $reader = [System.IO.StreamReader]::new($FilePath, [System.Text.Encoding]::UTF8, $true, 65536)

        while (($line = $reader.ReadLine()) -ne $null) {
            $lineCount++

            # Progress every 10,000 lines -- percentage based on bytes read vs file size
            if ($lineCount % 10000 -eq 0) {
                $pct = if ($fileSize -gt 0) {
                    [math]::Min(99, [math]::Round($reader.BaseStream.Position / $fileSize * 100))
                } else { -1 }
                Write-Progress -Activity "Parsing DNS Log: $fileName" `
                    -Status "Lines: $($lineCount.ToString('N0')) | PACKET rows: $($packetCount.ToString('N0')) | $pct%" `
                    -PercentComplete $pct
            }

            # Stage 1 -- Pre-filter: skip ~75% of all lines before the regex engine sees them
            # Only PACKET lines carry actual query data
            if ($line -notlike '*PACKET*') { continue }

            try {
                # Stage 2 -- Extract date and time from the start of the line
                # DNS log format: M/d/yyyy h:mm:ss tt <ThreadId> PACKET <PacketId> <rest>
                $parts = $line.TrimStart().Split(' ', 5, [System.StringSplitOptions]::RemoveEmptyEntries)
                if ($parts.Count -lt 5) { continue }

                $logDate  = $parts[0]                                              # M/d/yyyy
                $logTime  = [string]::Concat($parts[1], ' ', $parts[2])           # h:mm:ss tt
                $threadId = $parts[3]                                              # ThreadId (hex)

                # Parse timestamp -- always stored, ParseExact is the fastest path
                $logDateTime = [datetime]::ParseExact(
                    [string]::Concat($logDate, ' ', $logTime),
                    'M/d/yyyy h:mm:ss tt',
                    $culture
                ).ToString('yyyy-MM-ddTHH:mm:ss')

                # Stage 2b -- Match the remaining packet fields with compiled regex
                $m = $packetRegex.Match($line)
                if (-not $m.Success) { continue }

                $protocol      = $m.Groups[1].Value        # UDP or TCP
                $direction     = $m.Groups[2].Value        # Rcv or Snd
                $ipAddress     = $m.Groups[3].Value.Trim() # Client IP
                $transactionId = $m.Groups[4].Value        # DNS Xid
                $packetKind    = if ($m.Groups[5].Value -eq '') { 'Q' } else { $m.Groups[5].Value }  # Q or R (blank = query)
                $opcode        = 'Q'                       # Standard query (not separately encoded in this log version)
                $flagsHex      = $m.Groups[6].Value        # e.g. 0001 or 8081
                $flagsChar     = $m.Groups[7].Value        # e.g. D or DR (may be empty)
                $responseCode  = $m.Groups[8].Value        # NOERROR, NXDOMAIN, etc.
                $queryType     = $m.Groups[9].Value        # A, AAAA, MX, PTR...
                $rawQueryName  = $m.Groups[10].Value.Trim()

                # Extract PacketId from between PACKET keyword and protocol
                $packetId    = ''
                $packetMatch = [regex]::Match($line, 'PACKET\s+([0-9a-fA-F]+)\s+')
                if ($packetMatch.Success) { $packetId = $packetMatch.Groups[1].Value }

                # Stage 3 -- IP classification via cache (Test-VBPrivateIP called once per unique IP)
                if (-not $ipCache.ContainsKey($ipAddress)) {
                    $ipCache[$ipAddress] = @{
                        Version   = if ($ipAddress -match '^\d') { 'IPv4' } else { 'IPv6' }
                        IsPrivate = if (Test-VBPrivateIP -IPAddress $ipAddress) { 1 } else { 0 }
                    }
                }
                $cached    = $ipCache[$ipAddress]
                $ipVersion = $cached.Version
                $isPrivate = $cached.IsPrivate

                # Apply ExcludePrivateIPs filter at parse time (storage decision, not query decision)
                if ($ExcludePrivateIPs -and $isPrivate -eq 1) { continue }

                # Stage 4 -- DNS wire-format name decode (IndexOf/Substring, no regex)
                $queryName = ConvertFrom-VBDNSName -RawName $rawQueryName

                # Stage 5 -- RCODE to Status/Error mapping (inline, no function call)
                if ($responseCode -eq 'NOERROR') {
                    $status = 'Success'
                    $error  = ''
                } else {
                    $status = 'Error'
                    $error  = $responseCode
                }

                # Stage 6 -- Append object[] to buffer (cheapest possible allocation)
                # Column order is a fixed contract with Invoke-VBBulkInsert
                $buffer.Add([object[]]@(
                    $logDateTime,    # 0  LogDateTime
                    $logDate,        # 1  LogDate
                    $logTime,        # 2  LogTime
                    $threadId,       # 3  ThreadId
                    $packetId,       # 4  PacketId
                    $protocol,       # 5  Protocol
                    $direction,      # 6  Direction
                    $ipAddress,      # 7  IPAddress
                    $ipVersion,      # 8  IPVersion
                    $isPrivate,      # 9  IsPrivate
                    $transactionId,  # 10 TransactionId
                    $packetKind,     # 11 PacketKind
                    $opcode,         # 12 Opcode
                    $flagsHex,       # 13 FlagsHex
                    $flagsChar,      # 14 FlagsChar
                    $responseCode,   # 15 ResponseCode
                    $status,         # 16 Status
                    $error,          # 17 Error
                    $queryType,      # 18 QueryType
                    $queryName,      # 19 QueryName
                    $FilePath        # 20 SourceFile
                ))

                $packetCount++
            }
            catch {
                # Fail gracefully -- never crash the pipeline on a single bad line
                $errorCount++
                Write-Verbose "Invoke-VBDNSLogParser: Parse error on line $lineCount`: $($_.Exception.Message)"
            }
        }
    }
    finally {
        if ($null -ne $reader) { $reader.Close() }
        Write-Progress -Activity "Parsing DNS Log" -Completed
    }

    Write-Verbose "Invoke-VBDNSLogParser: Finished. Lines=$lineCount | PACKET=$packetCount | Errors=$errorCount"

    # Return a wrapper object -- never use Add-Member on the List[object[]] buffer.
    # Add-Member pipes the List through the PS pipeline which unrolls it into individual
    # object[] elements and destroys the List type, causing the insert loop to iterate
    # over scalar string values instead of object[] rows.
    return [PSCustomObject]@{
        Buffer     = $buffer
        ErrorCount = $errorCount
    }
}
