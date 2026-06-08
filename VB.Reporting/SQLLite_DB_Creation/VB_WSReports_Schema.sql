-- ============================================================
-- FILE    : VB_WSReports_Schema.sql
-- VERSION : 1.0.0
-- DATE    : 05-06-2026
-- AUTHOR  : Vibhu
-- PURPOSE : SQLite schema for YCN client workstation reports
-- SOURCE  : VB.ServerInventory Schema v0.2 (Opus-reviewed)
-- ------------------------------------------------------------
-- USAGE   : Applied automatically by Import-VBWSReportsToDB.ps1
--           Can also be run manually: sqlite3 DSI_Reports.db < VB_WSReports_Schema.sql
-- ------------------------------------------------------------
-- GLOBAL CONVENTIONS
--   Identity columns (ComputerName, SID, UserName) : TEXT NOT NULL COLLATE NOCASE
--   CollectionTime                                 : TEXT (ISO 8601, UTC)
--   Status                                         : TEXT
--   YES/NO, True/False columns                     : INTEGER (1/0)
--   PASS/FAIL/SKIPPED test columns                 : TEXT (carry error codes)
-- ============================================================

PRAGMA journal_mode = WAL;
PRAGMA foreign_keys = ON;

-- ============================================================
-- SINGLE-ROW-PER-COMPUTER TABLES
-- Ingest: INSERT OR REPLACE
-- ============================================================

-- 1. Azure AD / domain join status
CREATE TABLE IF NOT EXISTS AzJoinStatus_WS (
    ComputerName            TEXT NOT NULL COLLATE NOCASE,
    CollectionTime          TEXT,
    Status                  TEXT,
    AzureAdJoined           INTEGER,        -- YES/NO -> 1/0
    EnterpriseJoined        INTEGER,
    DomainJoined            INTEGER,
    DomainName              TEXT,
    DeviceName              TEXT,
    NgcSet                  INTEGER,
    WorkplaceJoined         INTEGER,
    WamDefaultSet           TEXT,           -- carries ERROR (0x...) values
    AzureAdPrt              INTEGER,
    AzureAdPrtAuthority     TEXT,
    EnterprisePrt           INTEGER,
    EnterprisePrtAuthority  TEXT,
    DiagnosticsReference    TEXT,
    UserContext             TEXT,
    ClientTime              TEXT,
    ADConnectivityTest      TEXT,           -- PASS/FAIL/SKIPPED + error codes
    ADConfigurationTest     TEXT,
    DRSDiscoveryTest        TEXT,
    DRSConnectivityTest     TEXT,
    TokenacquisitionTest    TEXT,
    FallbacktoSync_Join     TEXT,           -- ENABLED/DISABLED
    PreviousRegistration    TEXT,
    ErrorPhase              TEXT,
    ClientErrorCode         TEXT,
    AutoDetectSettings      INTEGER,        -- YES/NO -> 1/0
    Auto_ConfigurationURL   TEXT,
    ProxyServerList         TEXT,
    ProxyBypassList         TEXT,
    AutoDetectPACStatus     TEXT,
    ExecutingAccountName    TEXT,
    AccessType              TEXT,
    IsDeviceJoined          INTEGER,
    IsUserAzureAD           INTEGER,
    PolicyEnabled           INTEGER,
    PostLogonEnabled        INTEGER,
    DeviceEligible          INTEGER,
    SessionIsNotRemote      INTEGER,
    CertEnrollment          TEXT,
    PreReqResult            TEXT,
    DiagnosticURL           TEXT,
    PRIMARY KEY (ComputerName)
);

-- 2. Client Side Caching (offline files)
CREATE TABLE IF NOT EXISTS CSC_WS (
    ComputerName        TEXT NOT NULL COLLATE NOCASE,
    CollectionTime      TEXT,
    Status              TEXT,
    NetCachePolicy      TEXT,
    CSCServiceStatus    TEXT,
    CSCStatus           TEXT,
    CacheActive         INTEGER,            -- True/False -> 1/0
    CacheEnabled        INTEGER,
    CacheLocation       TEXT,
    FileCount           INTEGER,
    GroupPolicyName     TEXT,
    FileList            TEXT,
    PRIMARY KEY (ComputerName)
);

-- 10. System-context GPO system info
CREATE TABLE IF NOT EXISTS SysGPO_SystemInfo (
    ComputerName        TEXT NOT NULL COLLATE NOCASE,
    CollectionTime      TEXT,
    Status              TEXT,
    OsConfiguration     TEXT,
    OsVersion           TEXT,
    SiteName            TEXT,
    SlowLink            INTEGER,            -- Yes/No -> 1/0
    LastApplied         TEXT,
    AppliedFrom         TEXT,
    SlowLinkThreshold   TEXT,
    DomainName          TEXT,
    DomainType          TEXT,
    AppliedGPOs         TEXT,
    SecurityGroups      TEXT,
    Error               TEXT,
    PRIMARY KEY (ComputerName)
);

-- 11. AD join type summary
CREATE TABLE IF NOT EXISTS SystemADType_WS (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    CollectionTime  TEXT,
    Status          TEXT,
    AzureAdJoined   INTEGER,               -- YES/NO -> 1/0
    DomainJoined    INTEGER,
    JoinType        TEXT,
    PRIMARY KEY (ComputerName)
);

-- ============================================================
-- PER-SID TABLES (multi-row on RDS/Parallels session hosts)
-- Ingest: DELETE WHERE ComputerName = @cn, then INSERT, in transaction
-- ============================================================

-- 7. OneDrive for Business
CREATE TABLE IF NOT EXISTS ODFB_WS (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    UserName        TEXT NOT NULL COLLATE NOCASE,
    SID             TEXT COLLATE NOCASE,   -- identity column, data only
    CollectionTime  TEXT,
    Status          TEXT,
    OneDriveType    TEXT,
    UserEmail       TEXT,
    UserFolder      TEXT,
    KFMStatus       TEXT,
    KFMFolders      TEXT,
    PRIMARY KEY (ComputerName, UserName)
);

-- 13. User profiles
CREATE TABLE IF NOT EXISTS UP_WS (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    SID             TEXT NOT NULL COLLATE NOCASE,
    CollectionTime  TEXT,
    Status          TEXT,
    Domain          TEXT,
    Username        TEXT COLLATE NOCASE,
    ProfilePath     TEXT,
    LastUseTime     TEXT,
    Loaded          INTEGER,               -- True/False -> 1/0
    PRIMARY KEY (ComputerName, SID)
);

-- 14. User printer mappings
CREATE TABLE IF NOT EXISTS UPM_WS (
    ComputerName        TEXT NOT NULL COLLATE NOCASE,
    Username            TEXT NOT NULL COLLATE NOCASE,
    CollectionTime      TEXT,
    Status              TEXT,
    NetworkPrinters     TEXT,
    DefaultPrinter      TEXT,
    PrinterDevices      TEXT,
    PrinterCount        INTEGER,
    LastProfileUpdate   TEXT,
    PRIMARY KEY (ComputerName, Username)
);

-- 5. Logged-on users
CREATE TABLE IF NOT EXISTS LoggedOnUsers_WS (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    SID             TEXT NOT NULL COLLATE NOCASE,
    SessionId       TEXT NOT NULL,     -- identifier, not arithmetic
    CollectionTime  TEXT,
    Status          TEXT,
    Name            TEXT COLLATE NOCASE,
    Domain          TEXT,
    LogonType       INTEGER,
    LogonId         INTEGER,               -- LUID, informational only
    Session         TEXT,
    State           TEXT,
    IdleTime        TEXT,
    LogonTime       TEXT,
    Error           TEXT,
    PRIMARY KEY (ComputerName, SID, SessionId)
);

-- ============================================================
-- PER-RESOURCE TABLES (multi-row per computer)
-- Ingest: DELETE WHERE ComputerName = @cn, then INSERT, in transaction
-- ============================================================

-- 3. Disk / volume inventory
CREATE TABLE IF NOT EXISTS DiskInventory_WS (
    ComputerName        TEXT NOT NULL COLLATE NOCASE,
    DriveLetter         TEXT NOT NULL,
    CollectionTime      TEXT,
    CollectionStatus    TEXT,
    OperationalStatus   TEXT,
    HealthStatus        TEXT,
    WmiStatus           TEXT,
    IsSystemDisk        INTEGER,           -- True/False -> 1/0
    IsVolumeBased       INTEGER,
    TotalSizeGB         REAL,
    UsedGB              REAL,
    FreeGB              REAL,
    FreePercent         REAL,
    AllocationUnitSize  INTEGER,
    PartitionCount      INTEGER,
    DiskNumber          INTEGER,
    SpindleSpeed        INTEGER,
    Label               TEXT,
    FileSystem          TEXT,
    StorageType         TEXT,
    DiskType            TEXT,
    MediaType           TEXT,
    BusType             TEXT,
    Usage               TEXT,
    FriendlyName        TEXT,
    UniqueId            TEXT,
    SerialNumber        TEXT,
    FirmwareVersion     TEXT,
    InterfaceType       TEXT,
    WMICaption          TEXT,
    PRIMARY KEY (ComputerName, DriveLetter)
);

-- 4. Computer-scoped GPOs
CREATE TABLE IF NOT EXISTS GPO_WS (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    GPOName         TEXT NOT NULL,
    CollectionTime  TEXT,
    Status          TEXT,
    Enabled         INTEGER,               -- True/False -> 1/0
    AccessDenied    INTEGER,
    Version         INTEGER,
    GUID            TEXT,
    PRIMARY KEY (ComputerName, GPOName)
);

-- 6. Network adapters
CREATE TABLE IF NOT EXISTS NIC_WS (
    ComputerName            TEXT NOT NULL COLLATE NOCASE,
    InterfaceName           TEXT NOT NULL,
    CollectionTime          TEXT,
    Status                  TEXT,
    InterfaceDescription    TEXT,
    IPAddress               TEXT,
    SubnetMask              INTEGER,       -- CIDR prefix length
    DefaultGateway          TEXT,
    DNSServers              TEXT,
    IPType                  TEXT,
    DHCPEnabled             INTEGER,       -- True/False -> 1/0
    PRIMARY KEY (ComputerName, InterfaceName)
);

-- 8. System-context applied GPOs
CREATE TABLE IF NOT EXISTS SysGPO_AppliedGPOs (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    GPOName         TEXT NOT NULL,
    PRIMARY KEY (ComputerName, GPOName)
);

-- 9. System-context security groups
CREATE TABLE IF NOT EXISTS SysGPO_SecurityGroups (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    SecurityGroup   TEXT NOT NULL,
    PRIMARY KEY (ComputerName, SecurityGroup)
);

-- ============================================================
-- PER-USER TABLES (multi-row per computer)
-- Ingest: DELETE WHERE ComputerName = @cn, then INSERT, in transaction
-- ============================================================

-- 12. User folder redirection (registry)
CREATE TABLE IF NOT EXISTS UFR_WS (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    UserName        TEXT NOT NULL COLLATE NOCASE,
    UserSID         TEXT COLLATE NOCASE,   -- identity data column
    ValueName       TEXT NOT NULL,
    CollectionTime  TEXT,
    Status          TEXT,
    FolderType      TEXT,
    ValueData       TEXT,
    PRIMARY KEY (ComputerName, UserName, ValueName)
);

-- 15. User-context applied GPOs
CREATE TABLE IF NOT EXISTS UserGPO_AppliedGPOs (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    UserName        TEXT NOT NULL COLLATE NOCASE,
    GPOName         TEXT NOT NULL,
    PRIMARY KEY (ComputerName, UserName, GPOName)
);

-- 16. User-context security groups
CREATE TABLE IF NOT EXISTS UserGPO_SecurityGroups (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    UserName        TEXT NOT NULL COLLATE NOCASE,
    SecurityGroup   TEXT NOT NULL,
    PRIMARY KEY (ComputerName, UserName, SecurityGroup)
);

-- 17. User-context GPO system info
CREATE TABLE IF NOT EXISTS UserGPO_SystemInfo (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    UserName        TEXT NOT NULL COLLATE NOCASE,
    CollectionTime  TEXT,
    Status          TEXT,
    LastApplied     TEXT,
    AppliedFrom     TEXT,
    SlowLink        INTEGER,               -- Yes/No -> 1/0
    DomainName      TEXT,
    DomainType      TEXT,
    PRIMARY KEY (ComputerName, UserName)
);

-- 18. User shell folders (registry)
CREATE TABLE IF NOT EXISTS USF_WS (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    UserName        TEXT NOT NULL COLLATE NOCASE,
    UserSID         TEXT COLLATE NOCASE,   -- identity data column
    ValueName       TEXT NOT NULL,
    CollectionTime  TEXT,
    Status          TEXT,
    FolderType      TEXT,
    ValueData       TEXT,
    PRIMARY KEY (ComputerName, UserName, ValueName)
);

-- ============================================================
-- INDEXES
-- No standalone ComputerName indexes -- PK leftmost-prefix covers those
-- ============================================================

-- AzJoinStatus_WS
CREATE INDEX IF NOT EXISTS idx_azjoin_aadjoined     ON AzJoinStatus_WS (AzureAdJoined);
CREATE INDEX IF NOT EXISTS idx_azjoin_domainjoined  ON AzJoinStatus_WS (DomainJoined);
CREATE INDEX IF NOT EXISTS idx_azjoin_prereq        ON AzJoinStatus_WS (PreReqResult);

-- CSC_WS
CREATE INDEX IF NOT EXISTS idx_csc_cacheenabled     ON CSC_WS (CacheEnabled);

-- DiskInventory_WS
CREATE INDEX IF NOT EXISTS idx_disk_freepercent     ON DiskInventory_WS (FreePercent);

-- NIC_WS
CREATE INDEX IF NOT EXISTS idx_nic_iptype           ON NIC_WS (IPType);

-- ODFB_WS
CREATE INDEX IF NOT EXISTS idx_odfb_type            ON ODFB_WS (OneDriveType);
CREATE INDEX IF NOT EXISTS idx_odfb_kfm             ON ODFB_WS (KFMStatus);

-- SysGPO_SystemInfo
CREATE INDEX IF NOT EXISTS idx_sysgpo_domain        ON SysGPO_SystemInfo (DomainName);
CREATE INDEX IF NOT EXISTS idx_sysgpo_os            ON SysGPO_SystemInfo (OsVersion);

-- SystemADType_WS
CREATE INDEX IF NOT EXISTS idx_adtype_jointype      ON SystemADType_WS (JoinType);

-- Per-user lookups (UserName is never leftmost PK column in these tables)
CREATE INDEX IF NOT EXISTS idx_ufr_user             ON UFR_WS (UserName);
CREATE INDEX IF NOT EXISTS idx_usf_user             ON USF_WS (UserName);
CREATE INDEX IF NOT EXISTS idx_usergpo_applied_user ON UserGPO_AppliedGPOs (UserName);
CREATE INDEX IF NOT EXISTS idx_usergpo_secgrp_user  ON UserGPO_SecurityGroups (UserName);
CREATE INDEX IF NOT EXISTS idx_usergpo_sysinfo_user ON UserGPO_SystemInfo (UserName);

-- Optional -- uncomment when running staleness or failed-collection reports
-- CREATE INDEX IF NOT EXISTS idx_azjoin_collected  ON AzJoinStatus_WS (CollectionTime);
-- CREATE INDEX IF NOT EXISTS idx_azjoin_status     ON AzJoinStatus_WS (Status);
