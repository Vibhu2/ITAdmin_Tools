---
title: VB.WSReports — SQLite Database Documentation
version: 1.0.0
date: 05-06-2026
author: Vibhu
status: final
scope: YCN multi-tenant workstation reporting database
files:
  - VB_WSReports_Schema.sql
  - Import-VBWSReportsToDB.ps1
---

# VB.WSReports — SQLite Database Documentation

## Purpose

This system imports Windows workstation inventory and configuration reports
collected across YCN's managed tenants into a per-client SQLite database.
The database is used for fleet-wide reporting and querying — identifying
machines not joined to Azure AD, drives running low on space, users without
OneDrive KFM enabled, GPO application status, and similar operational queries.

One database is created per client. The same schema and import script serve
all clients — only the client name prefix changes.

---

## Files

| File | Purpose |
|---|---|
| `VB_WSReports_Schema.sql` | DDL — all 18 CREATE TABLE and 20 CREATE INDEX statements |
| `Import-VBWSReportsToDB.ps1` | PowerShell import script — reads CSVs, applies schema, ingests data |

Both files must be in the same folder. The script reads the schema file at
runtime and applies it before any data is loaded.

---

## Prerequisites

PSSQLite PowerShell module — install once per machine:

```powershell
Install-Module PSSQLite -Scope CurrentUser
```

---

## Usage

```powershell
# Basic run -- creates DSI_Reports.db next to the script
.\Import-VBWSReportsToDB.ps1 -CsvFolder "C:\Reports\DSI"

# Re-run / overwrite existing database
.\Import-VBWSReportsToDB.ps1 -CsvFolder "C:\Reports\DSI" -Force

# Custom database output path
.\Import-VBWSReportsToDB.ps1 -CsvFolder "C:\Reports\DSI" -DatabasePath "C:\DB\DSI.db" -Force

# Verbose step-by-step progress
.\Import-VBWSReportsToDB.ps1 -CsvFolder "C:\Reports\DSI" -Verbose -Force
```

### Switching clients

Open `Import-VBWSReportsToDB.ps1` and change line 57:

```powershell
$CLIENT_NAME = 'AGG'   # was 'DSI'
```

The script builds CSV filenames as `<CLIENT_NAME>_<ReportName>_<DocType>.csv`
and names the output database `<CLIENT_NAME>_Reports.db`. Changing one
variable handles the entire client switch.

---

## CSV Filename Pattern

```
<ClientName> _ <ReportName> _ <DocType> .csv
    DSI            UPM_WS        Report
    DSI            NIC_WS        Status
```

- `ClientName` — matches `$CLIENT_NAME` in the script (e.g. `DSI`, `AGG`, `BEB`)
- `ReportName` — stable identifier, used as the SQLite table name
- `DocType` — `Report` for all reports, `Status` for `NIC_WS` only

---

## Source Reports and Tables

### Single-row-per-computer tables

One row per machine. Ingest uses `INSERT OR REPLACE` — fresh data silently
replaces the prior row for that computer.

| # | CSV ReportName | Table | Description |
|---|---|---|---|
| 1 | `AzJoinStatus_WS` | `AzJoinStatus_WS` | Azure AD / domain join status and diagnostics |
| 2 | `CSC_WS` | `CSC_WS` | Client Side Caching (offline files) configuration |
| 10 | `SysGPO_SystemInfo` | `SysGPO_SystemInfo` | System-context GPO application summary |
| 11 | `SystemADType_WS` | `SystemADType_WS` | AD join type summary (on-prem, AAD, hybrid) |

### Per-SID / per-user tables (multi-row on session hosts)

These tables have one row per user per computer. On Parallels RAS or RDS
session hosts, many users are active simultaneously, so the PK must include
the user identifier. Ingest uses DELETE + INSERT in a transaction.

| # | CSV ReportName | Table | PK | Description |
|---|---|---|---|---|
| 5 | `LoggedOnUsers_WS` | `LoggedOnUsers_WS` | ComputerName + SID + SessionId | Active logon sessions |
| 7 | `ODFB_WS` | `ODFB_WS` | ComputerName + UserName | OneDrive for Business status per user |
| 13 | `UP_WS` | `UP_WS` | ComputerName + SID | Local user profiles |
| 14 | `UPM_WS` | `UPM_WS` | ComputerName + Username | User printer mappings |

### Per-resource tables (multi-row per computer)

One row per resource instance (drive, GPO, NIC, security group). Ingest uses
DELETE + INSERT in a transaction.

| # | CSV ReportName | Table | PK | Description |
|---|---|---|---|---|
| 3 | `DiskInventory_WS` | `DiskInventory_WS` | ComputerName + DriveLetter | Disk / volume inventory |
| 4 | `GPO_WS` | `GPO_WS` | ComputerName + GPOName | Computer-scoped GPO list |
| 6 | `NIC_WS` | `NIC_WS` | ComputerName + InterfaceName | Network adapter configuration |
| 8 | `SysGPO_AppliedGPOs` | `SysGPO_AppliedGPOs` | ComputerName + GPOName | System-context applied GPOs |
| 9 | `SysGPO_SecurityGroups` | `SysGPO_SecurityGroups` | ComputerName + SecurityGroup | System-context security groups |

### Per-user tables (multi-row per user per computer)

One row per registry value per user. Ingest uses DELETE + INSERT in a
transaction.

| # | CSV ReportName | Table | PK | Description |
|---|---|---|---|---|
| 12 | `UFR_WS` | `UFR_WS` | ComputerName + UserName + ValueName | User folder redirection (registry) |
| 15 | `UserGPO_AppliedGPOs` | `UserGPO_AppliedGPOs` | ComputerName + UserName + GPOName | User-context applied GPOs |
| 16 | `UserGPO_SecurityGroups` | `UserGPO_SecurityGroups` | ComputerName + UserName + SecurityGroup | User-context security groups |
| 17 | `UserGPO_SystemInfo` | `UserGPO_SystemInfo` | ComputerName + UserName | User-context GPO application summary |
| 18 | `USF_WS` | `USF_WS` | ComputerName + UserName + ValueName | User shell folders (registry) |

---

## Schema Design Decisions

### Global conventions

Every decision below applies to all 18 tables.

**Identity columns are `TEXT NOT NULL COLLATE NOCASE`.**
`ComputerName`, `SID`, and `UserName` are the join and deduplication keys
across every table. Different collection sources emit different casing — DNS
returns `DESKTOP-01`, NetBIOS returns `desktop-01`, WMI may return either.
Without `COLLATE NOCASE`, these are treated as different values and `INSERT OR
REPLACE` never finds the existing row to replace, silently creating duplicate
ghost rows. `COLLATE NOCASE` makes the PK comparison case-insensitive so the
collision happens and fresh data wins.

**`CollectionTime` is stored as `TEXT`.**
SQLite has no native datetime type. ISO 8601 text (`2026-06-05T08:21:35Z`)
sorts and compares correctly as a string, provided every value is consistently
zero-padded and in UTC. This is the collection script's responsibility.

**YES/NO and True/False columns are stored as `INTEGER` (1/0).**
Boolean storage as integers enables `WHERE AzureAdJoined = 1` and direct
aggregation with `SUM()` and `COUNT()`. Storing `'YES'` or `'True'` requires
string matching and prevents numeric aggregation entirely.

**PASS/FAIL/SKIPPED test columns are stored as `TEXT`.**
These columns carry diagnostic payloads such as `FAIL [0x80070005]`. Coercing
them to 1/0 would discard the error code, which is the most operationally
useful part of the value.

**PK columns are never typed `INTEGER` or `REAL`.**
Identifiers used as primary key components must be `TEXT`. An integer PK
column that receives an empty CSV value produces `NULL`, which violates `NOT
NULL` and causes the insert to fail. `SessionId` was initially typed
`INTEGER NOT NULL` and caused exactly this failure — it was corrected to
`TEXT NOT NULL` because it is a session identifier, not a number used for
arithmetic.

**Empty TEXT values are stored as `N/A`.**
The source CSVs contain empty fields in any column, including PK columns such
as `ValueName`. `N/A` is stored instead of `NULL` so that every TEXT column
has a consistent non-null sentinel, queries can filter on `WHERE col = 'N/A'`
to identify missing data, and PK constraints are never violated by empty
strings. Integer and real columns store `NULL` on empty or unparseable values
— `N/A` cannot be stored in a typed numeric column without breaking
aggregation queries.

### Primary key decisions

**Single-row tables key on `ComputerName` alone.**
`AzJoinStatus_WS`, `CSC_WS`, `SysGPO_SystemInfo`, and `SystemADType_WS`
describe exactly one fact per machine. `INSERT OR REPLACE` on `ComputerName`
is sufficient to keep them current.

**`ODFB_WS` and `UPM_WS` key on `ComputerName + UserName`.**
The initial assumption of one row per computer is incorrect on Parallels RAS
and RDS session hosts, where many users have simultaneous OneDrive sessions
and per-user printer mappings. A single-column PK on `ComputerName` would
silently drop every user except the last one collected. `UserName` is used
rather than `SID` because the `UPM_WS` CSV does not contain a SID column.
SID is stored as a data column in `ODFB_WS` for future use — a dedicated
`UserName → SID` mapping report is planned.

**`LoggedOnUsers_WS` keys on `ComputerName + SID + SessionId`.**
A single SID can hold more than one simultaneous session (console plus RDP).
`SessionId` is the terminal-services session number, which maps directly to
`quser` output. It is typed `TEXT` not `INTEGER` because the CSV can contain
empty values that would violate a `NOT NULL INTEGER` constraint.

**`UserGPO_SystemInfo` keys on `ComputerName + UserName`.**
Its siblings `UserGPO_AppliedGPOs` and `UserGPO_SecurityGroups` are both
keyed per user. Keying system info per computer only would collapse all users'
GPO system info into one row on multi-user machines, producing incomplete and
inconsistent results when joining the three user-GPO tables.

**Composite keys on all remaining multi-row tables** use the natural stable
identifier for each row — drive letter, GPO name, interface name, security
group name, registry value name. This allows re-collection to map a new row
onto its prior version rather than duplicating it.

### Ingest decisions

**`INSERT OR REPLACE` for single-row tables.**
The full row is supplied on every collection and there are no foreign keys or
external rowid references. `INSERT OR REPLACE` is a DELETE + INSERT
internally, which resets the row cleanly. It is the simplest correct option
for these tables.

**DELETE + INSERT in a transaction for all multi-row tables.**
`INSERT OR REPLACE` only touches rows that collide on the primary key. It
never removes rows whose key no longer exists in the current collection. When
a drive is removed, a GPO is unlinked, a user logs off, or a NIC is
decommissioned, the old row would persist indefinitely and any query against
that table would return stale data. Deleting all rows for a given computer
before re-inserting the current set guarantees the table reflects the latest
collection exactly. Wrapping the delete and all inserts in a single .NET
`BeginTransaction()` / `Commit()` block ensures a concurrent reader never
sees the table in a half-loaded state.

Raw `BEGIN`/`COMMIT` SQL strings were tried first and failed because
`Invoke-SqliteQuery` opens and closes its own connection on each call — the
transaction issued on one call does not exist on the next. The fix was to
open a persistent `New-SQLiteConnection` and call `.BeginTransaction()` on
that connection object, so all operations within the loop share the same
connection and the same transaction.

### Index decisions

**No standalone `ComputerName` indexes.**
Every table has `ComputerName` as the leftmost PK column (alone or in a
composite). SQLite auto-creates a B-tree index on the PK. A leftmost-prefix
lookup (`WHERE ComputerName = 'X'`) uses this index directly at no extra cost.
A separate `ComputerName` index would duplicate it and add write overhead on
every insert. Because this is a write-heavy workload (every collection rewrites
every table), trimming redundant indexes is a measurable win.

**`UserName` indexes are kept.**
`UserName` is never the leftmost PK column in any table. Queries that look up
everything for a specific user across the fleet cannot use the PK index and
would full-scan without a dedicated index. Per-user queries are a real
operational access pattern.

**Reporting-column indexes** back specific known query shapes:

| Index | Query it supports |
|---|---|
| `idx_azjoin_aadjoined` | Machines not joined to Azure AD |
| `idx_azjoin_domainjoined` | Machines not domain-joined |
| `idx_azjoin_prereq` | Machines that will not provision (`WillNotProvision`) |
| `idx_csc_cacheenabled` | Machines with offline files enabled |
| `idx_disk_freepercent` | Drives below a free-space threshold |
| `idx_nic_iptype` | Static vs dynamic IP assignment fleet-wide |
| `idx_odfb_type` / `idx_odfb_kfm` | OneDrive type and KFM rollout status |
| `idx_sysgpo_domain` / `idx_sysgpo_os` | OS version spread and domain distribution |
| `idx_adtype_jointype` | Join type breakdown across the fleet |

---

## Script Architecture

### Parameters

| Parameter | Type | Required | Default | Description |
|---|---|---|---|---|
| `CsvFolder` | string | yes | — | Folder containing the client's CSV files |
| `DatabasePath` | string | no | `<script folder>\<CLIENT_NAME>_Reports.db` | Output SQLite database path |
| `Force` | switch | no | off | Delete and recreate an existing database |

### Configuration block

`$CLIENT_NAME` — the only value that changes per client. Controls both the
CSV filename prefix and the default database filename.

`$SCHEMA_PATH` — resolved to `VB_WSReports_Schema.sql` in the same folder as
the script. The schema file must be co-located.

`$TRUTHY_VALUES` — the set of source strings that map to integer `1` for
boolean columns: `YES`, `TRUE`, `1`, `ENABLED`.

### Table definitions (`$TABLE_DEFS`)

An ordered hashtable — one entry per report. Each entry is a hashtable with
these keys:

| Key | Type | Purpose |
|---|---|---|
| `DocType` | string | `Report` or `Status` — completes the CSV filename |
| `Table` | string | SQLite table name |
| `Ingest` | string | `Replace` or `Transaction` — selects the ingest function |
| `Bool` | string[] | CSV column names to convert to 1/0 |
| `Int` | string[] | CSV column names to store as INTEGER (non-boolean numerics only — never PK columns) |
| `Real` | string[] | CSV column names to store as REAL |
| `Columns` | string[] | Ordered list of CSV column headers, exactly as they appear in the file |
| `ColMap` | hashtable | CSV column name → DB column name mappings for columns that differ (e.g. hyphens replaced with underscores) |
| `PKColumns` | string[] | Primary key column names — used for documentation and future validation |

### Helper functions

**`Convert-VBFieldValue`**
Converts a raw CSV string to the correct typed value before it is passed to
SQLite. Decision order:

1. Bool column → `1` if value is in `$TRUTHY_VALUES`, `0` otherwise. Empty
   string → `0`.
2. Int column → parse as integer, return `$null` on failure.
3. Real column → parse as double, return `$null` on failure.
4. Everything else (TEXT) → return value as-is. Empty or whitespace → `N/A`.

**`Invoke-VBReplaceIngest`**
Used for single-row-per-computer tables. Builds one `INSERT OR REPLACE`
statement with named parameters `@p0..@pN` and executes it for each row.
Named index-suffixed parameters are used rather than positional `?` because
PSSQLite's `SqlParameters` accepts only `IDictionary` types — arrays are not
supported.

**`Invoke-VBTransactionIngest`**
Used for all multi-row tables. Opens a persistent `SQLiteConnection`, then for
each group of rows sharing a `ComputerName`:
1. Begins a .NET transaction on the connection.
2. Executes `DELETE FROM [Table] WHERE ComputerName = @cn`.
3. Inserts all current rows for that computer using `INSERT OR REPLACE`.
4. Commits the transaction. On error, rolls back and re-throws.

The connection is closed and disposed in a `finally` block regardless of
outcome.

### Execution steps

| Step | What happens |
|---|---|
| 1 | PSSQLite module presence check — fail fast with install instructions if missing |
| 2 | Validate CSV folder and schema file paths exist |
| 3 | If database exists: delete if `-Force`, throw if not |
| 4 | Apply schema — full SQL file passed to `Invoke-SqliteQuery` in one call |
| 5 | For each table def: build filename, import CSV, select ingest path, execute |
| 6 | Print summary table showing table, ingest type, CSV filename, row count, status |

---

## Common Queries

```sql
-- Machines not joined to Azure AD
SELECT ComputerName, DomainName, JoinType
FROM SystemADType_WS
WHERE AzureAdJoined = 0;

-- Drives with less than 15% free space
SELECT ComputerName, DriveLetter, TotalSizeGB, FreeGB, FreePercent
FROM DiskInventory_WS
WHERE FreePercent < 15
ORDER BY FreePercent;

-- Machines where Azure AD join prereq will not provision
SELECT ComputerName, PreReqResult, ADConfigurationTest, DRSDiscoveryTest
FROM AzJoinStatus_WS
WHERE PreReqResult = 'WillNotProvision';

-- OneDrive KFM not configured
SELECT ComputerName, UserName, UserEmail, OneDriveType, KFMStatus
FROM ODFB_WS
WHERE KFMStatus != 'KFM Configured';

-- All GPOs applied to a specific computer
SELECT g.GPOName
FROM SysGPO_AppliedGPOs g
WHERE g.ComputerName = 'DESKTOP-K97GLEF';

-- Join machine info with disk and NIC for a full computer summary
SELECT
    a.ComputerName,
    a.JoinType,
    s.OsVersion,
    n.IPAddress,
    n.IPType,
    d.DriveLetter,
    d.FreePercent
FROM SystemADType_WS a
LEFT JOIN SysGPO_SystemInfo s ON s.ComputerName = a.ComputerName
LEFT JOIN NIC_WS n            ON n.ComputerName = a.ComputerName
LEFT JOIN DiskInventory_WS d  ON d.ComputerName = a.ComputerName
                              AND d.IsSystemDisk = 1
ORDER BY a.ComputerName;
```

---

## Adding a New Report

1. Add the CSV to the client's report folder following the naming pattern:
   `<ClientName>_<ReportName>_Report.csv`

2. Add a `CREATE TABLE IF NOT EXISTS` block to `VB_WSReports_Schema.sql`
   following the global conventions. Add any useful indexes below the
   existing index block.

3. Add an entry to `$TABLE_DEFS` in `Import-VBWSReportsToDB.ps1`:
   - Set `Ingest` to `Replace` if the report has one row per computer,
     `Transaction` if it can have multiple rows per computer.
   - List every CSV column in `Columns` in header order.
   - Classify boolean columns in `Bool`, numeric counts in `Int`, decimal
     values in `Real`.
   - Add `ColMap` entries for any CSV column names that contain characters
     invalid in SQL identifiers (hyphens, spaces, commas).
   - Set `PKColumns` to match the PRIMARY KEY declared in the schema.

4. Run the script with `-Force` to recreate the database with the new table.

---

## Known Limitations

- **No historical data.** Each import replaces the previous data for each
  computer. If you need point-in-time history, a `CollectionDate` column and
  an append-only ingest pattern would be required.

- **`UPM_WS` has no SID column.** Printer mappings are keyed on
  `ComputerName + Username`. A future `Username → SID` mapping report will
  allow joining printer data to SID-keyed tables. When that report is added,
  the `UPM_WS` PK can be migrated to `ComputerName + SID`.

- **`CollectionTime` format consistency.** The database stores `CollectionTime`
  as TEXT. String-based date comparison only works correctly when every value
  is in the same ISO 8601 zero-padded UTC format. Mixed formats from different
  collection script versions will produce incorrect sort and range query results.

- **One database per client.** Cross-client fleet queries (e.g. all machines
  across all tenants with Azure AD join failures) require either attaching
  multiple databases in SQLite or externalising to a shared database.
