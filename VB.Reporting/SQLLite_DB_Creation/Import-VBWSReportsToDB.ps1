# ============================================================
# SCRIPT  : Import-VBWSReportsToDB.ps1
# VERSION : 2.0.7
# CHANGED : 05-06-2026 -- SessionId TEXT not INTEGER; no PK cols in Int lists; bool empty guard
#            paths, boolean conversion, transaction wrapping
# AUTHOR  : Vibhu
# PURPOSE : Import YCN client workstation report CSVs into SQLite
# ENCODING: UTF-8 with BOM -- do not re-save without BOM
# ------------------------------------------------------------
# REQUIRES: PSSQLite module -- Install-Module PSSQLite -Scope CurrentUser
# ------------------------------------------------------------
# FILENAME PATTERN: <ClientName>_<ReportName>_<DocType>.csv
#   e.g. DSI_UPM_WS_Report.csv, DSI_NIC_WS_Status.csv
#
# USAGE:
#   # Basic -- creates DSI_Reports.db next to the script
#   .\Import-VBWSReportsToDB.ps1 -CsvFolder "C:\Reports\DSI"
#
#   # Different client -- change $CLIENT_NAME below, then:
#   .\Import-VBWSReportsToDB.ps1 -CsvFolder "C:\Reports\AGG"
#
#   # Custom DB path
#   .\Import-VBWSReportsToDB.ps1 -CsvFolder "C:\Reports\DSI" -DatabasePath "C:\DB\DSI.db"
#
#   # Overwrite existing DB
#   .\Import-VBWSReportsToDB.ps1 -CsvFolder "C:\Reports\DSI" -Force
#
#   # See step-by-step progress
#   .\Import-VBWSReportsToDB.ps1 -CsvFolder "C:\Reports\DSI" -Verbose
# ------------------------------------------------------------
# CHANGELOG
# v2.0.0 -- 05-06-2026 -- Full rewrite: proper typed schema, two ingest paths
# v1.2.2 -- 05-06-2026 -- Fixed INSERT to use positional ? params
# v1.2.1 -- 05-06-2026 -- Renamed to Import-VBWSReportsToDB
# ============================================================

#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CsvFolder,

    [Parameter()]
    [string]$DatabasePath = '',

    [Parameter()]
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# --- CONFIGURATION ---

# Client name prefix in all CSV filenames (e.g. DSI, AGG, BEB, SPI).
# Change this one value when running for a different client.
$CLIENT_NAME = 'DSI'

# Resolve default DB path now that CLIENT_NAME is known
if ($DatabasePath -eq '') {
    $DatabasePath = Join-Path $PSScriptRoot ($CLIENT_NAME + '_Reports.db')
}

# Schema DDL path -- expected next to this script
$SCHEMA_PATH = Join-Path $PSScriptRoot 'VB_WSReports_Schema.sql'

# YES/NO and True/False values that map to INTEGER 1
$TRUTHY_VALUES = @('YES', 'TRUE', 'YES', '1', 'ENABLED')

# --- TABLE DEFINITIONS ---
# Each entry:
#   DocType  -- 'Report' or 'Status' (completes the filename)
#   Table    -- SQLite table name
#   PK       -- column name(s) that form the primary key
#   Ingest   -- 'Replace' (INSERT OR REPLACE) or 'Transaction' (DELETE+INSERT)
#   Bool     -- columns to convert to 1/0
#   Int      -- columns to store as INTEGER (non-boolean numerics)
#   Real     -- columns to store as REAL
#   Columns  -- ordered list matching the CSV headers exactly

$TABLE_DEFS = [ordered]@{

    'AzJoinStatus_WS' = @{
        DocType = 'Report'
        Table   = 'AzJoinStatus_WS'
        Ingest  = 'Replace'
        Bool    = @('AzureAdJoined','EnterpriseJoined','DomainJoined','NgcSet',
                    'WorkplaceJoined','AzureAdPrt','EnterprisePrt','IsDeviceJoined',
                    'IsUserAzureAD','PolicyEnabled','PostLogonEnabled','DeviceEligible',
                    'SessionIsNotRemote','AutoDetectSettings')
        Int     = @()
        Real    = @()
        Columns = @('ComputerName','Status','CollectionTime','AzureAdJoined',
                    'EnterpriseJoined','DomainJoined','DomainName','DeviceName',
                    'NgcSet','WorkplaceJoined','WamDefaultSet','AzureAdPrt',
                    'AzureAdPrtAuthority','EnterprisePrt','EnterprisePrtAuthority',
                    'DiagnosticsReference','UserContext','ClientTime',
                    'ADConnectivityTest','ADConfigurationTest','DRSDiscoveryTest',
                    'DRSConnectivityTest','TokenacquisitionTest','FallbacktoSync-Join',
                    'PreviousRegistration','ErrorPhase','ClientErrorCode',
                    'AutoDetectSettings','Auto-ConfigurationURL','ProxyServerList',
                    'ProxyBypassList','AutoDetectPACStatus','ExecutingAccountName',
                    'AccessType','IsDeviceJoined','IsUserAzureAD','PolicyEnabled',
                    'PostLogonEnabled','DeviceEligible','SessionIsNotRemote',
                    'CertEnrollment','PreReqResult','Formoreinformation')
        # Map CSV column names that differ from DB column names
        ColMap  = @{ 'FallbacktoSync-Join' = 'FallbacktoSync_Join'
                     'Auto-ConfigurationURL' = 'Auto_ConfigurationURL'
                     'Formoreinformation' = 'DiagnosticURL' }
        PKColumns = @('ComputerName')
    }

    'CSC_WS' = @{
        DocType = 'Report'
        Table   = 'CSC_WS'
        Ingest  = 'Replace'
        Bool    = @('CacheActive','CacheEnabled')
        Int     = @('FileCount')
        Real    = @()
        Columns = @('ComputerName','NetCachePolicy','CSCServiceStatus','CSCStatus',
                    'CacheActive','CacheEnabled','CacheLocation','FileCount',
                    'GroupPolicyName','FileList','CollectionTime','Status')
        ColMap  = @{}
        PKColumns = @('ComputerName')
    }

    'DiskInventory_WS' = @{
        DocType = 'Report'
        Table   = 'DiskInventory_WS'
        Ingest  = 'Transaction'
        Bool    = @('IsSystemDisk','IsVolumeBased')
        Int     = @('AllocationUnitSize','PartitionCount','DiskNumber','SpindleSpeed')
        Real    = @('TotalSizeGB','UsedGB','FreeGB','FreePercent')
        Columns = @('ComputerName','DriveLetter','Label','FileSystem','TotalSizeGB',
                    'UsedGB','FreeGB','FreePercent','AllocationUnitSize',
                    'OperationalStatus','IsSystemDisk','StorageType','DiskType',
                    'MediaType','BusType','SpindleSpeed','Usage','FriendlyName',
                    'UniqueId','HealthStatus','SerialNumber','FirmwareVersion',
                    'InterfaceType','WMICaption','WmiStatus','PartitionCount',
                    'DiskNumber','IsVolumeBased','CollectionStatus','CollectionTime')
        ColMap  = @{}
        PKColumns = @('ComputerName','DriveLetter')
    }

    'GPO_WS' = @{
        DocType = 'Report'
        Table   = 'GPO_WS'
        Ingest  = 'Transaction'
        Bool    = @('Enabled','AccessDenied')
        Int     = @('Version')
        Real    = @()
        Columns = @('ComputerName','GPOName','Enabled','AccessDenied','Version',
                    'GUID','Status')
        ColMap  = @{}
        PKColumns = @('ComputerName','GPOName')
    }

    'LoggedOnUsers_WS' = @{
        DocType = 'Report'
        Table   = 'LoggedOnUsers_WS'
        Ingest  = 'Transaction'
        Bool    = @()
        Int     = @('LogonType','LogonId')
        Real    = @()
        Columns = @('ComputerName','Status','Name','Domain','SID','LogonType',
                    'LogonId','Session','SessionId','State','IdleTime',
                    'LogonTime','Error','CollectionTime')
        ColMap  = @{ 'Name' = 'Name' }
        PKColumns = @('ComputerName','SID','SessionId')
    }

    'NIC_WS' = @{
        DocType = 'Status'
        Table   = 'NIC_WS'
        Ingest  = 'Transaction'
        Bool    = @('DHCPEnabled')
        Int     = @('SubnetMask')
        Real    = @()
        Columns = @('ComputerName','InterfaceName','InterfaceDescription','IPAddress',
                    'SubnetMask','DefaultGateway','DNSServers','IPType',
                    'DHCPEnabled','Status','CollectionTime')
        ColMap  = @{}
        PKColumns = @('ComputerName','InterfaceName')
    }

    'ODFB_WS' = @{
        DocType = 'Report'
        Table   = 'ODFB_WS'
        Ingest  = 'Transaction'
        Bool    = @()
        Int     = @()
        Real    = @()
        Columns = @('ComputerName','UserName','SID','OneDriveType','UserEmail',
                    'UserFolder','KFMStatus','KFMFolders','CollectionTime','Status')
        ColMap  = @{}
        PKColumns = @('ComputerName','UserName')
    }

    'SysGPO_AppliedGPOs' = @{
        DocType = 'Report'
        Table   = 'SysGPO_AppliedGPOs'
        Ingest  = 'Transaction'
        Bool    = @()
        Int     = @()
        Real    = @()
        Columns = @('ComputerName','GPOName')
        ColMap  = @{}
        PKColumns = @('ComputerName','GPOName')
    }

    'SysGPO_SecurityGroups' = @{
        DocType = 'Report'
        Table   = 'SysGPO_SecurityGroups'
        Ingest  = 'Transaction'
        Bool    = @()
        Int     = @()
        Real    = @()
        Columns = @('ComputerName','SecurityGroup')
        ColMap  = @{}
        PKColumns = @('ComputerName','SecurityGroup')
    }

    'SysGPO_SystemInfo' = @{
        DocType = 'Report'
        Table   = 'SysGPO_SystemInfo'
        Ingest  = 'Replace'
        Bool    = @('SlowLink')
        Int     = @()
        Real    = @()
        Columns = @('ComputerName','Status','OsConfiguration','OsVersion','SiteName',
                    'SlowLink','LastApplied','AppliedFrom','SlowLinkThreshold',
                    'DomainName','DomainType','AppliedGPOs','SecurityGroups',
                    'Error','CollectionTime')
        ColMap  = @{}
        PKColumns = @('ComputerName')
    }

    'SystemADType_WS' = @{
        DocType = 'Report'
        Table   = 'SystemADType_WS'
        Ingest  = 'Replace'
        Bool    = @('AzureAdJoined','DomainJoined')
        Int     = @()
        Real    = @()
        Columns = @('ComputerName','AzureAdJoined','DomainJoined','JoinType')
        ColMap  = @{}
        PKColumns = @('ComputerName')
    }

    'UFR_WS' = @{
        DocType = 'Report'
        Table   = 'UFR_WS'
        Ingest  = 'Transaction'
        Bool    = @()
        Int     = @()
        Real    = @()
        Columns = @('ComputerName','UserSID','UserName','FolderType','ValueName',
                    'ValueData','Status')
        ColMap  = @{}
        PKColumns = @('ComputerName','UserName','ValueName')
    }

    'UP_WS' = @{
        DocType = 'Report'
        Table   = 'UP_WS'
        Ingest  = 'Transaction'
        Bool    = @('Loaded')
        Int     = @()
        Real    = @()
        Columns = @('ComputerName','SID','Domain','Username','ProfilePath',
                    'LastUseTime','Loaded','CollectionTime','Status')
        ColMap  = @{}
        PKColumns = @('ComputerName','SID')
    }

    'UPM_WS' = @{
        DocType = 'Report'
        Table   = 'UPM_WS'
        Ingest  = 'Transaction'
        Bool    = @()
        Int     = @('PrinterCount')
        Real    = @()
        Columns = @('ComputerName','Username','NetworkPrinters','DefaultPrinter',
                    'PrinterDevices','PrinterCount','LastProfileUpdate',
                    'CollectionTime','Status')
        ColMap  = @{}
        PKColumns = @('ComputerName','Username')
    }

    'UserGPO_AppliedGPOs' = @{
        DocType = 'Report'
        Table   = 'UserGPO_AppliedGPOs'
        Ingest  = 'Transaction'
        Bool    = @()
        Int     = @()
        Real    = @()
        Columns = @('ComputerName','UserName','GPOName')
        ColMap  = @{}
        PKColumns = @('ComputerName','UserName','GPOName')
    }

    'UserGPO_SecurityGroups' = @{
        DocType = 'Report'
        Table   = 'UserGPO_SecurityGroups'
        Ingest  = 'Transaction'
        Bool    = @()
        Int     = @()
        Real    = @()
        Columns = @('ComputerName','UserName','SecurityGroup')
        ColMap  = @{}
        PKColumns = @('ComputerName','UserName','SecurityGroup')
    }

    'UserGPO_SystemInfo' = @{
        DocType = 'Report'
        Table   = 'UserGPO_SystemInfo'
        Ingest  = 'Transaction'
        Bool    = @('SlowLink')
        Int     = @()
        Real    = @()
        Columns = @('ComputerName','Status','UserName','LastApplied','AppliedFrom',
                    'SlowLink','DomainName','DomainType','CollectionTime')
        ColMap  = @{}
        PKColumns = @('ComputerName','UserName')
    }

    'USF_WS' = @{
        DocType = 'Report'
        Table   = 'USF_WS'
        Ingest  = 'Transaction'
        Bool    = @()
        Int     = @()
        Real    = @()
        Columns = @('UserSID','UserName','FolderType','ValueName','ValueData',
                    'CollectionTime','Status','ComputerName')
        ColMap  = @{}
        PKColumns = @('ComputerName','UserName','ValueName')
    }
}

# --- HELPER FUNCTIONS ---

function Convert-VBFieldValue {
    # Converts a raw CSV string to the appropriate typed value for SQLite.
    param(
        [string]$Value,
        [string]$ColumnName,
        [string[]]$BoolColumns,
        [string[]]$IntColumns,
        [string[]]$RealColumns
    )
    if ($BoolColumns -contains $ColumnName) {
        if ([string]::IsNullOrWhiteSpace($Value)) { return 0 }
        if ($TRUTHY_VALUES -contains $Value.ToUpper()) { return 1 } else { return 0 }
    }
    if ($IntColumns -contains $ColumnName) {
        $parsed = 0
        if ([int]::TryParse($Value, [ref]$parsed)) { return $parsed }
        return $null   # NULL is correct for unparseable integers -- keeps column typed
    }
    if ($RealColumns -contains $ColumnName) {
        $parsed = 0.0
        if ([double]::TryParse($Value, [ref]$parsed)) { return $parsed }
        return $null   # NULL is correct for unparseable reals
    }
    # Default: TEXT -- empty string becomes 'N/A'
    if ([string]::IsNullOrWhiteSpace($Value)) { return 'N/A' }
    return $Value
}

function Invoke-VBReplaceIngest {
    # Single-row-per-computer tables: INSERT OR REPLACE
    # SqlParameters requires IDictionary -- keys become @key in SQL.
    # Using p0..pN index-suffixed keys to guarantee uniqueness across all column names.
    param(
        [string]$Table,
        [string[]]$Columns,
        [hashtable]$ColMap,
        [string[]]$BoolColumns,
        [string[]]$IntColumns,
        [string[]]$RealColumns,
        [object[]]$Rows,
        [string]$Database
    )

    $dbCols   = $Columns | ForEach-Object { if ($ColMap.ContainsKey($_)) { $ColMap[$_] } else { $_ } }
    $paramKeys = 0..($Columns.Count - 1) | ForEach-Object { "p$_" }
    $quotedCols = $dbCols | ForEach-Object { """$_""" }
    $paramRefs  = $paramKeys | ForEach-Object { "@$_" }
    $sql = "INSERT OR REPLACE INTO [$Table] ($($quotedCols -join ', ')) VALUES ($($paramRefs -join ', '));"

    foreach ($row in $Rows) {
        $params = [ordered]@{}
        for ($i = 0; $i -lt $Columns.Count; $i++) {
            $params[$paramKeys[$i]] = Convert-VBFieldValue -Value $row.($Columns[$i]) `
                -ColumnName $Columns[$i] -BoolColumns $BoolColumns `
                -IntColumns $IntColumns -RealColumns $RealColumns
        }
        Invoke-SqliteQuery -DataSource $Database -Query $sql -SqlParameters $params
    }
    return $Rows.Count
}

function Invoke-VBTransactionIngest {
    # Multi-row-per-computer tables: DELETE + INSERT per computer.
    # Uses New-SQLiteConnection + .NET BeginTransaction for atomicity.
    # SqlParameters requires IDictionary -- index-suffixed keys p0..pN guarantee uniqueness.
    param(
        [string]$Table,
        [string[]]$Columns,
        [hashtable]$ColMap,
        [string[]]$BoolColumns,
        [string[]]$IntColumns,
        [string[]]$RealColumns,
        [object[]]$Rows,
        [string]$Database,
        [string[]]$PKColumns = @()
    )

    $dbCols    = $Columns | ForEach-Object { if ($ColMap.ContainsKey($_)) { $ColMap[$_] } else { $_ } }
    $paramKeys = 0..($Columns.Count - 1) | ForEach-Object { "p$_" }
    $quotedCols = $dbCols | ForEach-Object { """$_""" }
    $paramRefs  = $paramKeys | ForEach-Object { "@$_" }
    $insertSql  = "INSERT OR REPLACE INTO [$Table] ($($quotedCols -join ', ')) VALUES ($($paramRefs -join ', '));"
    $deleteSql  = "DELETE FROM [$Table] WHERE ComputerName = @cn;"

    $byComputer = $Rows | Group-Object -Property ComputerName

    $conn = New-SQLiteConnection -DataSource $Database
    try {
        foreach ($group in $byComputer) {
            $txn = $conn.BeginTransaction()
            try {
                Invoke-SqliteQuery -SQLiteConnection $conn -Query $deleteSql `
                    -SqlParameters @{ cn = $group.Name }
                foreach ($row in $group.Group) {
                    $params = [ordered]@{}
                    for ($i = 0; $i -lt $Columns.Count; $i++) {
                        $params[$paramKeys[$i]] = Convert-VBFieldValue -Value $row.($Columns[$i]) `
                            -ColumnName $Columns[$i] -BoolColumns $BoolColumns `
                            -IntColumns $IntColumns -RealColumns $RealColumns
                    }
                    Invoke-SqliteQuery -SQLiteConnection $conn -Query $insertSql -SqlParameters $params
                }
                $txn.Commit()
            }
            catch {
                $txn.Rollback()
                throw
            }
            finally {
                $txn.Dispose()
            }
        }
    }
    finally {
        $conn.Close()
        $conn.Dispose()
    }
    return $Rows.Count
}

# --- MAIN LOGIC ---

# Step 1 -- Validate PSSQLite is available
Write-Verbose 'Step 1 -- Checking PSSQLite module'
if (-not (Get-Module -ListAvailable -Name PSSQLite)) {
    throw 'PSSQLite module not found. Run: Install-Module PSSQLite -Scope CurrentUser'
}
Import-Module PSSQLite -ErrorAction Stop

# Step 2 -- Validate inputs
Write-Verbose 'Step 2 -- Validating inputs'
if (-not (Test-Path -Path $CsvFolder -PathType Container)) {
    throw "CSV folder not found: $CsvFolder"
}
if (-not (Test-Path -Path $SCHEMA_PATH)) {
    throw "Schema file not found: $SCHEMA_PATH -- ensure VB_WSReports_Schema.sql is next to this script"
}

# Step 3 -- Handle existing database
Write-Verbose 'Step 3 -- Preparing database'
if (Test-Path -Path $DatabasePath) {
    if ($Force) {
        Remove-Item -Path $DatabasePath -Force
        Write-Warning "Existing database removed: $DatabasePath"
    } else {
        throw "Database already exists: $DatabasePath -- use -Force to overwrite"
    }
}

# Step 4 -- Apply schema (all CREATE TABLE + CREATE INDEX statements)
# Pass the full SQL in one call -- PSSQLite handles multi-statement DDL correctly.
# Splitting on semicolons breaks multi-line CREATE TABLE blocks.
Write-Verbose 'Step 4 -- Applying schema'
$schemaSql = Get-Content -Path $SCHEMA_PATH -Raw -Encoding UTF8
Invoke-SqliteQuery -DataSource $DatabasePath -Query $schemaSql
Write-Verbose 'Schema applied'

# Step 5 -- Import CSV data
Write-Verbose 'Step 5 -- Importing CSV data'
$Results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($reportName in $TABLE_DEFS.Keys) {
    $def      = $TABLE_DEFS[$reportName]
    $csvFile  = $CLIENT_NAME + '_' + $reportName + '_' + $def.DocType + '.csv'
    $csvPath  = Join-Path $CsvFolder $csvFile
    $rowsImported = 0
    $importStatus = 'TableOnly'

    if (Test-Path -Path $csvPath) {
        try {
            $rows = Import-Csv -Path $csvPath -Encoding UTF8

            if ($rows.Count -gt 0) {
                $ingestParams = @{
                    Table       = $def.Table
                    Columns     = $def.Columns
                    ColMap      = $def.ColMap
                    BoolColumns = $def.Bool
                    IntColumns  = $def.Int
                    RealColumns = $def.Real
                    Rows        = $rows
                    Database    = $DatabasePath
                }

                $ingestParams['PKColumns'] = $def.PKColumns
                if ($def.Ingest -eq 'Replace') {
                    $rowsImported = Invoke-VBReplaceIngest @ingestParams
                } else {
                    $rowsImported = Invoke-VBTransactionIngest @ingestParams
                }
                $importStatus = 'Imported'
            } else {
                $importStatus = 'EmptyCSV'
            }
        }
        catch {
            $importStatus = 'Error: ' + $_.Exception.Message
        }
    } else {
        Write-Warning "CSV not found, table created empty: $csvFile"
    }

    $Results.Add([PSCustomObject]@{
        Table        = $def.Table
        Ingest       = $def.Ingest
        CsvFile      = $csvFile
        RowsImported = $rowsImported
        Status       = $importStatus
    })
}

# Step 6 -- Summary
Write-Host "`nDatabase : $DatabasePath" -ForegroundColor Cyan
Write-Host "Client   : $CLIENT_NAME`n" -ForegroundColor Cyan
$Results | Format-Table -AutoSize
