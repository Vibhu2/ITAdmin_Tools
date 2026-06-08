---
title: VB.ServerInventory — SQLite Schema
version: 0.2
status: review-revised
date: 2026-06-05
supersedes: v0.1 (Opus review draft)
---

# VB.ServerInventory — SQLite Schema (v0.2)

> [!note] What changed from v0.1
> This revision applies the outcomes of the design review: a **delete-then-insert ingest pattern** for all multi-row tables, **composite primary keys** on the RDS-affected tables (`ODFB_WS`, `UPM_WS`, `LoggedOnUsers_WS`, `UserGPO_SystemInfo`), removal of **redundant `ComputerName` indexes**, and **`COLLATE NOCASE`** on all identity columns. The full reasoning is in the [Design Rationale](#design-rationale) section at the end.

> [!warning] Column completeness
> The `CREATE TABLE` statements below contain every column whose name and type were settled in the design draft, plus the three global columns. Tables where the draft referred to columns generically (e.g. "all test result columns") carry a marked placeholder line — `-- + source columns ...` — where the remaining columns from the collection scripts must be added. **All unspecified source columns are `TEXT` unless a type decision says otherwise.** Fill these in before running the DDL.

---

## Global conventions

These apply to every table and are stated once here rather than repeated in the rationale per-table:

| Convention | Decision |
|---|---|
| Identity columns (`ComputerName`, `SID`, `UserName`) | `TEXT NOT NULL COLLATE NOCASE` |
| `CollectionTime` | `TEXT` — ISO 8601, zero-padded, **UTC** |
| `Status` | `TEXT` (`Success`, `Failed`, `Not Joined`, …) |
| YES/NO and True/False columns | `INTEGER` (1/0) |
| PASS/FAIL/SKIPPED test columns | `TEXT` (they carry error codes, e.g. `FAIL [0x80070005]`) |
| Single-row-per-computer ingest | `INSERT OR REPLACE` |
| Multi-row-per-computer ingest | `DELETE … WHERE ComputerName = @cn` then `INSERT`, in one transaction |

---

## Schema — CREATE TABLE

### Single-row-per-computer tables

```sql
-- 1. Azure AD / domain join status -----------------------------------------
CREATE TABLE IF NOT EXISTS AzJoinStatus_WS (
    ComputerName        TEXT NOT NULL COLLATE NOCASE,
    CollectionTime      TEXT,
    Status              TEXT,
    AzureAdJoined       INTEGER,   -- 0/1
    EnterpriseJoined    INTEGER,
    DomainJoined        INTEGER,
    NgcSet              INTEGER,
    WorkplaceJoined     INTEGER,
    AzureAdPrt          INTEGER,
    EnterprisePrt       INTEGER,
    IsDeviceJoined      INTEGER,
    IsUserAzureAD       INTEGER,
    PolicyEnabled       INTEGER,
    PostLogonEnabled    INTEGER,
    DeviceEligible      INTEGER,
    SessionIsNotRemote  INTEGER,
    AutoDetectSettings  INTEGER,
    WamDefaultSet       TEXT,      -- carries ERROR (0x...) values
    PreReqResult        TEXT,      -- test result, carries error codes
    -- + source columns: remaining test-result columns (all TEXT, carry error codes)
    PRIMARY KEY (ComputerName)
);

-- 2. Client Side Caching (offline files) -----------------------------------
CREATE TABLE IF NOT EXISTS CSC_WS (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    CollectionTime  TEXT,
    Status          TEXT,
    FileCount       INTEGER,
    CacheActive     INTEGER,   -- 0/1
    CacheEnabled    INTEGER,   -- 0/1
    -- + source columns (TEXT unless noted)
    PRIMARY KEY (ComputerName)
);

-- 10. System-context GPO system info ----------------------------------------
CREATE TABLE IF NOT EXISTS SysGPO_SystemInfo (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    CollectionTime  TEXT,
    Status          TEXT,
    DomainName      TEXT,
    OsVersion       TEXT,
    SlowLink        INTEGER,   -- Yes/No -> 1/0
    -- + source columns (TEXT unless noted)
    PRIMARY KEY (ComputerName)
);

-- 11. AD join type (summary) ------------------------------------------------
CREATE TABLE IF NOT EXISTS SystemADType_WS (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    CollectionTime  TEXT,
    Status          TEXT,
    JoinType        TEXT,
    AzureAdJoined   INTEGER,   -- 0/1
    DomainJoined    INTEGER,   -- 0/1
    -- + source columns (TEXT unless noted)
    PRIMARY KEY (ComputerName)
);
```

### Per-SID tables (multi-row on session hosts)

```sql
-- 7. OneDrive for Business --------------------------------------------------
-- PK changed from ComputerName -> ComputerName + SID (RDS hosts have many users)
CREATE TABLE IF NOT EXISTS ODFB_WS (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    SID             TEXT NOT NULL COLLATE NOCASE,
    CollectionTime  TEXT,
    Status          TEXT,
    OneDriveType    TEXT,
    KFMStatus       TEXT,
    -- + source columns (TEXT unless noted)
    PRIMARY KEY (ComputerName, SID)
);

-- 13. User profiles ---------------------------------------------------------
CREATE TABLE IF NOT EXISTS UP_WS (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    SID             TEXT NOT NULL COLLATE NOCASE,
    CollectionTime  TEXT,
    Status          TEXT,
    Loaded          INTEGER,   -- 0/1
    -- + source columns (TEXT unless noted)
    PRIMARY KEY (ComputerName, SID)
);

-- 14. User printer mappings -------------------------------------------------
-- PK changed from ComputerName -> ComputerName + SID (registry-sourced per HKU\<SID>)
CREATE TABLE IF NOT EXISTS UPM_WS (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    SID             TEXT NOT NULL COLLATE NOCASE,
    CollectionTime  TEXT,
    Status          TEXT,
    PrinterCount    INTEGER,
    -- + source columns (TEXT unless noted)
    PRIMARY KEY (ComputerName, SID)
);

-- 5. Logged-on users --------------------------------------------------------
-- PK changed from ComputerName + SID -> ComputerName + SID + SessionId
-- Assumes ingest is filtered to interactive/RDP sessions to match quser.
CREATE TABLE IF NOT EXISTS LoggedOnUsers_WS (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    SID             TEXT NOT NULL COLLATE NOCASE,
    SessionId       INTEGER NOT NULL,
    UserName        TEXT COLLATE NOCASE,
    CollectionTime  TEXT,
    Status          TEXT,
    LogonType       INTEGER,
    LogonId         INTEGER,   -- LUID, regenerates per boot; informational only
    -- + source columns (TEXT unless noted)
    PRIMARY KEY (ComputerName, SID, SessionId)
);
```

### Per-resource tables (multi-row)

```sql
-- 3. Disk / volume inventory ------------------------------------------------
CREATE TABLE IF NOT EXISTS DiskInventory_WS (
    ComputerName        TEXT NOT NULL COLLATE NOCASE,
    DriveLetter         TEXT NOT NULL,
    CollectionTime      TEXT,
    Status              TEXT,
    TotalSizeGB         REAL,
    UsedGB              REAL,
    FreeGB              REAL,
    FreePercent         REAL,
    AllocationUnitSize  INTEGER,
    PartitionCount      INTEGER,
    DiskNumber          INTEGER,
    IsSystemDisk        INTEGER,   -- 0/1
    IsVolumeBased       INTEGER,   -- 0/1
    -- + source columns (TEXT unless noted)
    PRIMARY KEY (ComputerName, DriveLetter)
);

-- 4. Computer-scoped GPOs ---------------------------------------------------
CREATE TABLE IF NOT EXISTS GPO_WS (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    GPOName         TEXT NOT NULL,
    CollectionTime  TEXT,
    Status          TEXT,
    Enabled         INTEGER,   -- 0/1
    AccessDenied    INTEGER,   -- 0/1
    Version         INTEGER,
    -- + source columns (TEXT unless noted)
    PRIMARY KEY (ComputerName, GPOName)
);

-- 6. Network adapters -------------------------------------------------------
CREATE TABLE IF NOT EXISTS NIC_WS (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    InterfaceName   TEXT NOT NULL,
    CollectionTime  TEXT,
    Status          TEXT,
    IPType          TEXT,
    SubnetMask      INTEGER,   -- CIDR prefix length
    DHCPEnabled     INTEGER,   -- 0/1
    -- + source columns (TEXT unless noted)
    PRIMARY KEY (ComputerName, InterfaceName)
);

-- 8. System-context applied GPOs --------------------------------------------
CREATE TABLE IF NOT EXISTS SysGPO_AppliedGPOs (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    GPOName         TEXT NOT NULL,
    CollectionTime  TEXT,
    Status          TEXT,
    -- + source columns (all TEXT)
    PRIMARY KEY (ComputerName, GPOName)
);

-- 9. System-context security groups -----------------------------------------
CREATE TABLE IF NOT EXISTS SysGPO_SecurityGroups (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    SecurityGroup   TEXT NOT NULL,
    CollectionTime  TEXT,
    Status          TEXT,
    -- + source columns (all TEXT)
    PRIMARY KEY (ComputerName, SecurityGroup)
);
```

### Per-user tables (multi-row)

```sql
-- 12. User folder redirection -----------------------------------------------
CREATE TABLE IF NOT EXISTS UFR_WS (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    UserName        TEXT NOT NULL COLLATE NOCASE,
    ValueName       TEXT NOT NULL,
    CollectionTime  TEXT,
    Status          TEXT,
    -- + source columns (all TEXT)
    PRIMARY KEY (ComputerName, UserName, ValueName)
);

-- 15. User-context applied GPOs ---------------------------------------------
CREATE TABLE IF NOT EXISTS UserGPO_AppliedGPOs (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    UserName        TEXT NOT NULL COLLATE NOCASE,
    GPOName         TEXT NOT NULL,
    CollectionTime  TEXT,
    Status          TEXT,
    -- + source columns (all TEXT)
    PRIMARY KEY (ComputerName, UserName, GPOName)
);

-- 16. User-context security groups ------------------------------------------
CREATE TABLE IF NOT EXISTS UserGPO_SecurityGroups (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    UserName        TEXT NOT NULL COLLATE NOCASE,
    SecurityGroup   TEXT NOT NULL,
    CollectionTime  TEXT,
    Status          TEXT,
    -- + source columns (all TEXT)
    PRIMARY KEY (ComputerName, UserName, SecurityGroup)
);

-- 17. User-context GPO system info ------------------------------------------
-- PK changed from ComputerName -> ComputerName + UserName (match siblings 15/16)
CREATE TABLE IF NOT EXISTS UserGPO_SystemInfo (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    UserName        TEXT NOT NULL COLLATE NOCASE,
    CollectionTime  TEXT,
    Status          TEXT,
    SlowLink        INTEGER,   -- Yes/No -> 1/0
    -- + source columns (TEXT unless noted)
    PRIMARY KEY (ComputerName, UserName)
);

-- 18. User shell folders ----------------------------------------------------
CREATE TABLE IF NOT EXISTS USF_WS (
    ComputerName    TEXT NOT NULL COLLATE NOCASE,
    UserName        TEXT NOT NULL COLLATE NOCASE,
    ValueName       TEXT NOT NULL,
    CollectionTime  TEXT,
    Status          TEXT,
    -- + source columns (all TEXT)
    PRIMARY KEY (ComputerName, UserName, ValueName)
);
```

---

## Schema — CREATE INDEX

> [!info] No `ComputerName` indexes
> Every table is keyed on `ComputerName` (alone or as the leftmost column of a composite PK). SQLite auto-creates an index for the PK and serves leftmost-prefix lookups from it for free, so a standalone `ComputerName` index would be pure write overhead. They are intentionally absent.

```sql
-- AzJoinStatus_WS: join-state and prereq filtering
CREATE INDEX IF NOT EXISTS idx_azjoin_aadjoined    ON AzJoinStatus_WS (AzureAdJoined);
CREATE INDEX IF NOT EXISTS idx_azjoin_domainjoined ON AzJoinStatus_WS (DomainJoined);
CREATE INDEX IF NOT EXISTS idx_azjoin_prereq       ON AzJoinStatus_WS (PreReqResult);

-- CSC_WS
CREATE INDEX IF NOT EXISTS idx_csc_cacheenabled    ON CSC_WS (CacheEnabled);

-- DiskInventory_WS: low-disk-space reporting
CREATE INDEX IF NOT EXISTS idx_disk_freepercent    ON DiskInventory_WS (FreePercent);

-- NIC_WS
CREATE INDEX IF NOT EXISTS idx_nic_iptype          ON NIC_WS (IPType);

-- ODFB_WS
CREATE INDEX IF NOT EXISTS idx_odfb_type           ON ODFB_WS (OneDriveType);
CREATE INDEX IF NOT EXISTS idx_odfb_kfm            ON ODFB_WS (KFMStatus);

-- SysGPO_SystemInfo
CREATE INDEX IF NOT EXISTS idx_sysgpo_domain       ON SysGPO_SystemInfo (DomainName);
CREATE INDEX IF NOT EXISTS idx_sysgpo_os           ON SysGPO_SystemInfo (OsVersion);

-- SystemADType_WS
CREATE INDEX IF NOT EXISTS idx_adtype_jointype     ON SystemADType_WS (JoinType);

-- Per-user lookups (UserName is never the leftmost PK column here)
CREATE INDEX IF NOT EXISTS idx_ufr_user            ON UFR_WS (UserName);
CREATE INDEX IF NOT EXISTS idx_usergpo_applied_user ON UserGPO_AppliedGPOs (UserName);
CREATE INDEX IF NOT EXISTS idx_usergpo_secgrp_user  ON UserGPO_SecurityGroups (UserName);
CREATE INDEX IF NOT EXISTS idx_usergpo_sysinfo_user ON UserGPO_SystemInfo (UserName);
CREATE INDEX IF NOT EXISTS idx_usf_user            ON USF_WS (UserName);

-- Optional fleet-monitoring indexes — add only if you run these reports.
-- Staleness ("machines not reporting since X"):
-- CREATE INDEX IF NOT EXISTS idx_azjoin_collected ON AzJoinStatus_WS (CollectionTime);
-- Failed-collection filtering:
-- CREATE INDEX IF NOT EXISTS idx_azjoin_status    ON AzJoinStatus_WS (Status);
```

---

## Ingest pattern

### Single-row-per-computer tables

`AzJoinStatus_WS`, `CSC_WS`, `SysGPO_SystemInfo`, `SystemADType_WS`.

```sql
INSERT OR REPLACE INTO CSC_WS (ComputerName, CollectionTime, Status, FileCount, CacheActive, CacheEnabled, /* ... */)
VALUES (@cn, @t, @status, @filecount, @cacheactive, @cacheenabled /* , ... */);
```

### All other (multi-row) tables

Delete the computer's existing rows, then insert the current set — wrapped in one transaction so a reader never sees a half-loaded machine.

```sql
BEGIN;
DELETE FROM DiskInventory_WS WHERE ComputerName = @cn;
-- INSERT one row per current drive for @cn
INSERT INTO DiskInventory_WS (ComputerName, DriveLetter, /* ... */) VALUES (@cn, @drive, /* ... */);
-- ... repeat per drive ...
COMMIT;
```

> [!tip] Out-of-order arrival
> If collections from different sites can land out of execution order and you need true "newest-by-timestamp wins" on the single-row tables, replace `INSERT OR REPLACE` with:
> ```sql
> INSERT INTO AzJoinStatus_WS (...) VALUES (...)
> ON CONFLICT(ComputerName) DO UPDATE SET ...
>   WHERE excluded.CollectionTime > AzJoinStatus_WS.CollectionTime;
> ```
> If ingest is always in order, keep `INSERT OR REPLACE` — it's simpler and correct.

---

## Design Rationale

This section explains *why* each decision was made and *why it matters*, so the design can be defended and reused.

### Global decisions

**Identity columns are `TEXT NOT NULL COLLATE NOCASE`.**
`ComputerName`, `SID`, and `UserName` are the keys everything joins and de-duplicates on. SQLite's default text comparison is case-sensitive, so `ABC` and `abc`, or `DOMAIN\User` and `DOMAIN\user`, are treated as different values. Different sources emit different casing (DNS vs NetBIOS, registry vs WMI), so without normalization the same machine or user can produce two rows that never collide on the primary key — meaning `INSERT OR REPLACE` never overwrites the stale one and the duplicate lives forever. `COLLATE NOCASE` makes the key comparison case-insensitive, so the collision happens and fresh data wins. *Why it matters:* it is the difference between "fresh data wins" actually holding and silently accumulating duplicate ghost rows. (The alternative — uppercasing on write — works too, but baking the collation into the schema is self-documenting and can't be forgotten by a future ingest path.)

**`CollectionTime` stays `TEXT`, ISO 8601, zero-padded, UTC.**
SQLite has no native datetime type, and ISO-8601 text sorts and compares correctly as a string — `'2026-06-05T09:00:00Z' > '2026-06-04T23:00:00Z'` is true lexically. That only holds if every value is the *same* zero-padded format; a single non-padded or differently-formatted value breaks the ordering invisibly. UTC removes per-site timezone and DST ambiguity across the multi-site fleet. *Why it matters:* staleness reports, "newest wins" logic, and any time-range query all rely on string comparison being trustworthy.

**`Status` stays `TEXT`.**
It carries discrete labels (`Success`, `Failed`, `Not Joined`) that are read by humans and filtered by exact match. There is nothing numeric to gain. *Why it matters:* keeps the values self-explaining in reports without a lookup table.

**YES/NO and True/False columns are `INTEGER` (1/0).**
Storing booleans as integers lets you write `WHERE AzureAdJoined = 1` and aggregate with `SUM()`/`COUNT()` directly, rather than string-matching `'Yes'`. *Why it matters:* the whole point of the database is fleet-wide reporting, and boolean filters/counts are the most common query shape.

**PASS/FAIL/SKIPPED test columns stay `TEXT`.**
These are not booleans — they carry diagnostic payloads like `FAIL [0x80070005]`. Coercing them to 1/0 would discard the error code, which is the single most useful piece of data for troubleshooting. *Why it matters:* the error code is why you'd query the column at all.

### Primary-key decisions

**Single-row tables key on `ComputerName`.**
`AzJoinStatus_WS`, `CSC_WS`, `SysGPO_SystemInfo`, `SystemADType_WS` describe one fact per machine, so the machine name is the natural unique key and re-collection cleanly overwrites the prior row. *Why it matters:* simplest correct key; `INSERT OR REPLACE` alone keeps them current.

**`ODFB_WS` and `UPM_WS` changed to `ComputerName + SID`.**
The v0.1 assumption of one row per computer is false on Parallels RAS / RDS session hosts, where many users have OneDrive and per-user printer mappings *simultaneously*. With `PK = ComputerName`, every user but the last collected is silently dropped by `REPLACE`. `UPM_WS` is read from `HKU\<SID>` anyway, so it is per-SID by nature — a single-row PK fights the data's actual shape. SID is preferred over UserName because it survives renames and disambiguates identical display names across domains. *Why it matters:* on exactly the hosts where this data is most operationally important (shared session servers), the old key caused silent data loss.

**`LoggedOnUsers_WS` changed to `ComputerName + SID + SessionId`.**
A single SID can hold more than one session (console plus RDP), so `ComputerName + SID` collapses them. `SessionId` (the terminal-services session number) is chosen over `LogonId` because it maps directly to `quser` output and is human-readable, whereas `LogonId` is a LUID that regenerates on every reboot. This PK assumes ingest is filtered to interactive/RDP sessions; if raw `Win32_LogonSession` (including network/service logons) is loaded instead, the table fills with transient noise and the key must include `LogonId`. *Why it matters:* it keeps one row per real interactive session and lines the table up with the `quser` cross-check workflow.

**`UserGPO_SystemInfo` changed to `ComputerName + UserName`.**
Its siblings `UserGPO_AppliedGPOs` and `UserGPO_SecurityGroups` are already keyed per user, but v0.1 keyed system info per computer — so on a multi-user machine everyone's GPO system info collapsed to one row. Matching the sibling granularity fixes the inconsistency. *Why it matters:* per-user GPO reporting was incomplete on shared machines, and the mismatch would confuse anyone joining the three user-GPO tables.

**Composite keys on the remaining multi-row tables** (`DiskInventory_WS`, `GPO_WS`, `NIC_WS`, `SysGPO_*`, `UFR_WS`, `UP_WS`, `UserGPO_Applied/SecurityGroups`, `USF_WS`) use the stable natural identifier for each row (drive letter, GPO name, interface name, security group, value name, SID). *Why it matters:* a meaningful natural key lets re-collection map a current row onto its prior version instead of duplicating it.

### Ingest decisions

**`INSERT OR REPLACE` for single-row tables.**
Mechanically it is a DELETE-then-INSERT, which means it resets unsupplied columns, changes the rowid, and fires delete/insert triggers. None of that bites here: the full row is supplied every collection, there are no foreign keys, and rowids are not referenced externally. So the simplest option is also the correct one. *Why it matters:* avoids unnecessary `ON CONFLICT` boilerplate on tables that don't need it.

**Delete-then-insert (in a transaction) for multi-row tables.**
This is the most important correctness fix in the revision. `INSERT OR REPLACE` only touches rows that *collide* on the primary key; it never removes rows that have *disappeared* since the last collection. So when a drive is removed, a GPO is unlinked, or a user logs off, the old row lingers forever and "fresh data wins" quietly fails. Deleting the computer's rows first and re-inserting the current set guarantees the table reflects the latest collection exactly. Wrapping it in a transaction means a concurrent reader never sees a machine mid-reload (rows deleted, not yet re-inserted). *Why it matters:* without this, multi-row tables steadily fill with stale entries that make every historical-looking query wrong, and the errors are invisible until someone notices a decommissioned drive still "exists."

**`ON CONFLICT DO UPDATE … WHERE excluded.CollectionTime > …` only if needed.**
This enforces newest-by-timestamp rather than last-by-execution-order, which matters only if collections can arrive out of order (e.g. queued per site). It is offered as an option rather than baked in, because adding it where ingest is already ordered is needless complexity. *Why it matters:* protects against an older collection overwriting a newer one — but only a real risk in some topologies.

### Index decisions

**No standalone `ComputerName` indexes.**
SQLite auto-indexes the primary key, and a composite PK answers leftmost-prefix lookups (`WHERE ComputerName = …`) directly from that index. A separate `ComputerName` index would duplicate it and add write cost with zero read benefit. Because every collection writes every table, trimming useless indexes is a measurable win, not just tidiness. *Why it matters:* this is a write-heavy workload; each redundant index taxes every insert.

**`UserName` indexes kept.**
`UserName` is never the leftmost column of any primary key, so per-user queries ("show me everything for user X across the fleet") cannot use the PK index and genuinely benefit from a dedicated one. *Why it matters:* per-user reporting is a real access pattern that would otherwise full-scan.

**Reporting-column indexes** (`FreePercent`, `CacheEnabled`, `IPType`, `AzureAdJoined`, `DomainJoined`, `PreReqResult`, `OneDriveType`, `KFMStatus`, `JoinType`, `DomainName`, `OsVersion`) back the specific fleet questions these tables exist to answer — low disk space, join state, OS spread, OneDrive rollout. *Why it matters:* they target known query shapes; indexing is justified by an actual report, not added speculatively.

**`CollectionTime` / `Status` indexes left commented out.**
Useful for staleness and failed-collection reports, but only worth their write cost if those reports are actually run. *Why it matters:* keeps the default schema lean; opt in when the query exists.
