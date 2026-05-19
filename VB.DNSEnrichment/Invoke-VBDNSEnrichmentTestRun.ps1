# ============================================================
# SCRIPT   : Invoke-VBDNSEnrichmentTestRun.ps1
# VERSION  : 1.0.0
# AUTHOR   : VB
# PURPOSE  : Top-level orchestration wrapper for the VB.DNSEnrichment
#            test suite.  Accepts CSV input, provisions or mounts a
#            SQLite database interactively, auto-detects environment
#            capabilities (DHCP role, AD, SNMP), then delegates every
#            test section to Test-VBDNSEnrichmentModule.ps1.
#            Output is tee'd to a timestamped log file next to the DB.
#            A per-IP summary CSV is written on completion.
# REQUIRES : VB.DNSEnrichment v0.4.0, PSSQLite
# ENCODING : UTF-8 with BOM
# ============================================================

#Requires -Version 5.1

[CmdletBinding()]
param(
    # ---- Input -------------------------------------------------
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$CsvPath,

    # ---- DB provisioning (prompted if omitted) -----------------
    [Parameter()]
    [string]$DatabaseFolder,

    [Parameter()]
    [string]$DatabaseName,

    # ---- Optional env overrides --------------------------------
    # If the DHCP Server Role is detected locally it is auto-populated.
    # Supply this only when the role lives on a remote server.
    [Parameter()]
    [string]$DHCPServer,

    # SNMP community strings — defaults used if not supplied.
    [Parameter()]
    [string[]]$SNMPCommunityStrings = @('public'),

    # Switch IPs for Layer 10 — left empty if not supplied.
    [Parameter()]
    [string[]]$SwitchTargets = @(),

    # ---- Test behaviour ----------------------------------------
    [Parameter()]
    [switch]$SkipActiveProbes,

    [Parameter()]
    [switch]$ForceRefresh
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ScriptStart = Get-Date

# ============================================================
# REGION: Helpers
# ============================================================

function Write-Banner {
    param([string]$Text, [string]$Colour = 'White')
    Write-Host ""
    Write-Host "  $Text" -ForegroundColor $Colour -BackgroundColor DarkBlue
    Write-Host ""
}

function Write-Step {
    param([string]$Text)
    Write-Host "  >> $Text" -ForegroundColor Cyan
}

function Write-Info {
    param([string]$Text)
    Write-Host "     $Text" -ForegroundColor Gray
}

function Write-Warn {
    param([string]$Text)
    Write-Host "  [!] $Text" -ForegroundColor Yellow
}

function Write-Err {
    param([string]$Text)
    Write-Host "  [X] $Text" -ForegroundColor Red
}

function Prompt-WithDefault {
    param(
        [string]$Prompt,
        [string]$Default = ''
    )
    $display = if ($Default) { "$Prompt [$Default]: " } else { "$Prompt: " }
    $answer  = Read-Host -Prompt $display
    if ([string]::IsNullOrWhiteSpace($answer)) { $Default } else { $answer.Trim() }
}

# ============================================================
# REGION: CSV ingestion
# ============================================================

Write-Banner "VB.DNSEnrichment -- Test Orchestrator  |  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" 'White'

Write-Step "Reading CSV: $CsvPath"
try {
    $csvRows = Import-Csv -Path $CsvPath -Encoding UTF8
} catch {
    Write-Err "Failed to read CSV: $($_.Exception.Message)"
    exit 1
}

if ($csvRows.Count -eq 0) {
    Write-Err "CSV is empty -- nothing to process."
    exit 1
}

# Auto-detect IP column
$ipColumnCandidates = @('IPAddress','IP_Address','IP Address','IP','ip_address','ip address','ipaddress')
$ipColumn = $null
$firstRow  = $csvRows | Select-Object -First 1
foreach ($candidate in $ipColumnCandidates) {
    if ($firstRow.PSObject.Properties.Name -contains $candidate) {
        $ipColumn = $candidate
        break
    }
}

if (-not $ipColumn) {
    Write-Warn "Could not auto-detect IP column. Available columns:"
    $firstRow.PSObject.Properties.Name | ForEach-Object { Write-Info "  - $_" }
    $ipColumn = Prompt-WithDefault -Prompt "Enter the exact column name containing IP addresses"
    if (-not ($firstRow.PSObject.Properties.Name -contains $ipColumn)) {
        Write-Err "Column '$ipColumn' not found in CSV."
        exit 1
    }
}

Write-Info "Using IP column: '$ipColumn'"

$rawIPs = $csvRows | ForEach-Object { $_.$ipColumn } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() }

if ($rawIPs.Count -eq 0) {
    Write-Err "No IP addresses found in column '$ipColumn'."
    exit 1
}

# Deduplicate
$IPAddresses = $rawIPs | Sort-Object -Unique
Write-Info "Loaded $($IPAddresses.Count) unique IP(s) from $($csvRows.Count) CSV row(s)."

# ============================================================
# REGION: Database provisioning
# ============================================================

Write-Step "Database setup"

if (-not $DatabaseFolder) {
    $DatabaseFolder = Prompt-WithDefault -Prompt "Enter folder path for the SQLite database" -Default (Join-Path $env:USERPROFILE 'Desktop\VBEnrichmentDB')
}

if (-not (Test-Path -LiteralPath $DatabaseFolder)) {
    Write-Info "Folder does not exist -- creating: $DatabaseFolder"
    try {
        New-Item -ItemType Directory -Path $DatabaseFolder -Force | Out-Null
    } catch {
        Write-Err "Cannot create folder: $($_.Exception.Message)"
        exit 1
    }
}

if (-not $DatabaseName) {
    $DatabaseName = Prompt-WithDefault -Prompt "Enter the database filename (without extension)" -Default 'VBEnrichment'
}

# Normalise: strip .db extension if the user typed it, then re-add
$DatabaseName = [System.IO.Path]::GetFileNameWithoutExtension($DatabaseName)
$dbPath = Join-Path $DatabaseFolder "$DatabaseName.db"

if (Test-Path -LiteralPath $dbPath) {
    Write-Info "Existing database found -- will mount: $dbPath"
} else {
    Write-Info "No database at that path -- will create: $dbPath"
}

# ============================================================
# REGION: Environment auto-detection
# ============================================================

Write-Step "Detecting local environment capabilities"

# DHCP Server Role
if (-not $DHCPServer) {
    $dhcpRolePresent = $false
    try {
        $dhcpService = Get-Service -Name 'DHCPServer' -ErrorAction SilentlyContinue
        if ($dhcpService -and $dhcpService.Status -eq 'Running') {
            $dhcpRolePresent = $true
        }
    } catch {
        # Service query failed -- role absent
    }

    if ($dhcpRolePresent) {
        $DHCPServer = $env:COMPUTERNAME
        Write-Info "DHCP Server role detected on this machine -- using: $DHCPServer"
    } else {
        $dhcpInput = Prompt-WithDefault -Prompt "DHCP Server FQDN or IP (leave blank to skip)" -Default ''
        if (-not [string]::IsNullOrWhiteSpace($dhcpInput)) {
            $DHCPServer = $dhcpInput
        }
    }
}

# AD availability
$adAvailable = $false
try {
    $adModule = Get-Module -Name ActiveDirectory -ListAvailable -ErrorAction SilentlyContinue
    if ($adModule) {
        $adAvailable = $true
        Write-Info "ActiveDirectory module found -- AD layer will be active."
    } else {
        Write-Info "ActiveDirectory module not found -- AD layer will be skipped."
    }
} catch {
    Write-Info "Could not check for ActiveDirectory module -- AD layer may be skipped."
}

# PSSQLite check
$psSqliteAvailable = $null -ne (Get-Module -Name PSSQLite -ListAvailable -ErrorAction SilentlyContinue)
if (-not $psSqliteAvailable) {
    Write-Err "PSSQLite module is not installed. Install it with:  Install-Module PSSQLite -Scope CurrentUser"
    exit 1
}
Write-Info "PSSQLite available."

# SNMP availability hint (best-effort)
$snmpHint = $false
try {
    $snmpSvc = Get-Service -Name 'SNMP' -ErrorAction SilentlyContinue
    if ($snmpSvc) { $snmpHint = $true }
} catch { }
Write-Info "SNMP service present on this host: $snmpHint (module detects full availability at runtime)"

# Summary
Write-Host ""
Write-Host ("  {0,-28} {1}" -f "CSV file:",$CsvPath)          -ForegroundColor Gray
Write-Host ("  {0,-28} {1}" -f "IPs to test:",$IPAddresses.Count) -ForegroundColor Gray
Write-Host ("  {0,-28} {1}" -f "Database path:",$dbPath)       -ForegroundColor Gray
Write-Host ("  {0,-28} {1}" -f "DHCP Server:",$(if($DHCPServer){$DHCPServer}else{'(not configured)'})) -ForegroundColor Gray
Write-Host ("  {0,-28} {1}" -f "SNMP community strings:",($SNMPCommunityStrings -join ', ')) -ForegroundColor Gray
Write-Host ("  {0,-28} {1}" -f "Switch targets:",$(if($SwitchTargets.Count -gt 0){$SwitchTargets -join ', '}else{'(none)'})) -ForegroundColor Gray
Write-Host ("  {0,-28} {1}" -f "SkipActiveProbes:",$SkipActiveProbes.IsPresent) -ForegroundColor Gray
Write-Host ("  {0,-28} {1}" -f "ForceRefresh:",$ForceRefresh.IsPresent) -ForegroundColor Gray
Write-Host ""

$confirm = Prompt-WithDefault -Prompt "Proceed with these settings? (Y/N)" -Default 'Y'
if ($confirm -notmatch '^[Yy]') {
    Write-Warn "Aborted by user."
    exit 0
}

# ============================================================
# REGION: Log file setup (tee to file)
# ============================================================

$logFile = Join-Path $DatabaseFolder ("TestRun_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
Write-Step "Logging output to: $logFile"

# Start transcript to capture all console output
try {
    Start-Transcript -Path $logFile -Append -NoClobber | Out-Null
} catch {
    Write-Warn "Could not start transcript: $($_.Exception.Message)"
}

# ============================================================
# REGION: Build context -- inject DB path
# ============================================================

Write-Step "Loading VB.DNSEnrichment module"

$modulePsd1 = Join-Path $PSScriptRoot 'VB.DNSEnrichment.psd1'
if (-not (Test-Path -LiteralPath $modulePsd1)) {
    Write-Err "Module manifest not found at: $modulePsd1"
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
    exit 1
}

try {
    Import-Module $modulePsd1 -Force -ErrorAction Stop
    Write-Info "Module loaded: v$((Get-Module VB.DNSEnrichment).Version)"
} catch {
    Write-Err "Module import failed: $($_.Exception.Message)"
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
    exit 1
}

# Build context params
$ctxParams = @{
    DatabasePath         = $dbPath
    SNMPCommunityStrings = $SNMPCommunityStrings
    SwitchTargets        = $SwitchTargets
    Quiet                = $true
}
if ($DHCPServer) { $ctxParams['DHCPServer'] = $DHCPServer }

Write-Step "Building enrichment context"
try {
    $ctx = Get-VBEnrichmentContext @ctxParams
    Write-Info "Context ready  PSEdition=$($ctx.PSEdition)  CanUseParallel=$($ctx.CanUseParallel)"
    Write-Info "  AD=$($ctx.ADAvailable)  DHCP=$($ctx.DHCPAvailable)  SNMP=$($ctx.SNMPAvailable)  mDNS=$($ctx.mDNSAvailable)"
} catch {
    Write-Err "Get-VBEnrichmentContext failed: $($_.Exception.Message)"
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
    exit 1
}

# Initialize (or verify) the database
Write-Step "Initialising database"
try {
    Initialize-VBEnrichmentDatabase -DatabasePath $dbPath
    Write-Info "Database ready: $dbPath"
} catch {
    Write-Err "Database initialisation failed: $($_.Exception.Message)"
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
    exit 1
}

# ============================================================
# REGION: Delegate to Test-VBDNSEnrichmentModule.ps1
# ============================================================

Write-Step "Launching test suite (Test-VBDNSEnrichmentModule.ps1)"
Write-Host ""

$testScript = Join-Path $PSScriptRoot 'Test-VBDNSEnrichmentModule.ps1'
if (-not (Test-Path -LiteralPath $testScript)) {
    Write-Err "Test script not found: $testScript"
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
    exit 1
}

$testParams = @{
    IPAddress    = $IPAddresses
    OutputPath   = $DatabaseFolder
    ForceRefresh = $ForceRefresh
}
if ($DHCPServer)              { $testParams['DHCPServer']            = $DHCPServer }
if ($SNMPCommunityStrings)    { $testParams['SNMPCommunityStrings']  = $SNMPCommunityStrings }
if ($SwitchTargets.Count -gt 0) { $testParams['SwitchTargets']      = $SwitchTargets }
if ($SkipActiveProbes)        { $testParams['SkipActiveProbes']      = $true }

$testExitCode = 0
try {
    & $testScript @testParams
    $testExitCode = $LASTEXITCODE
} catch {
    Write-Err "Test script threw an exception: $($_.Exception.Message)"
    $testExitCode = 1
}

# ============================================================
# REGION: Database update verification + summary CSV export
# ============================================================

Write-Host ""
Write-Step "Verifying database state and exporting summary"

$summaryPath = $null
try {
    $dbRows = Get-VBEnrichmentResult -IPAddress $IPAddresses -Context $ctx

    # IPs that enrichment did not produce a DB row for (e.g. public IPs rejected)
    $storedIPs   = @($dbRows | ForEach-Object { $_.IPAddress })
    $missingIPs  = $IPAddresses | Where-Object { $storedIPs -notcontains $_ }

    Write-Info "IPs requested : $($IPAddresses.Count)"
    Write-Info "IPs in DB     : $($storedIPs.Count)"

    if ($missingIPs.Count -gt 0) {
        Write-Warn "$($missingIPs.Count) IP(s) have no DB row (likely rejected as non-private or errored):"
        $missingIPs | ForEach-Object { Write-Info "  - $_" }
    }

    # For IPs that were already in the DB but may have stale data, run a targeted upsert
    # to ensure the DB reflects the freshest results from this test run.
    if ($storedIPs.Count -gt 0 -and -not $ForceRefresh) {
        Write-Info "Running targeted refresh upsert for $($storedIPs.Count) DB row(s)..."
        try {
            $refreshResults = Invoke-VBIPEnrichment -IPAddress $storedIPs -Context $ctx -ForceRefresh
            $upsertCount = @($refreshResults).Count
            Write-Info "Upsert complete -- $upsertCount row(s) updated in DB."
        } catch {
            Write-Warn "Upsert pass failed: $($_.Exception.Message)"
        }
    }

    # Re-query after upsert to get the final state for export
    $summaryRows = Get-VBEnrichmentResult -IPAddress $IPAddresses -Context $ctx

    if ($summaryRows -and @($summaryRows).Count -gt 0) {
        $summaryPath = Join-Path $DatabaseFolder ("EnrichmentSummary_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
        $summaryRows | Export-VBEnrichmentResult -Format CSV -Path $summaryPath
        Write-Info "Summary CSV written ($(@($summaryRows).Count) rows): $summaryPath"
    } else {
        Write-Warn "No enrichment rows in DB for the tested IPs -- summary CSV skipped."
    }
} catch {
    Write-Warn "DB verification / export step failed: $($_.Exception.Message)"
}

# ============================================================
# REGION: Final banner
# ============================================================

$elapsed = [int](New-TimeSpan -Start $ScriptStart -End (Get-Date)).TotalSeconds
Write-Host ""
Write-Host ("  " + ("=" * 60)) -ForegroundColor DarkGray
$statusColour = if ($testExitCode -ne 0) { 'Red' } else { 'Green' }
$statusText   = if ($testExitCode -ne 0) { 'COMPLETED WITH FAILURES' } else { 'ALL TESTS PASSED' }
Write-Host ("  $statusText  |  {0}s elapsed" -f $elapsed) -ForegroundColor $statusColour
Write-Host ("  DB    : $dbPath") -ForegroundColor Gray
Write-Host ("  Log   : $logFile") -ForegroundColor Gray
if ($summaryPath -and (Test-Path -LiteralPath $summaryPath -ErrorAction SilentlyContinue)) {
    Write-Host ("  CSV   : $summaryPath") -ForegroundColor Gray
}
Write-Host ("  " + ("=" * 60)) -ForegroundColor DarkGray
Write-Host ""

Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
exit $testExitCode
