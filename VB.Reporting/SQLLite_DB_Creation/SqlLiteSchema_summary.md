Excellent — this is thorough and well-reasoned. Opus addressed every open question. Let me read the truncated section before confirming I have everything:I have everything I need. The schema is complete and well-reasoned. Here's a quick summary of the key changes from my draft that Opus corrected:

| My Draft | Opus Correction | Reason |
| --- | --- | --- |
| `ODFB_WS` PK = `ComputerName` | `ComputerName + SID` | RDS/Parallels hosts have many simultaneous users |
| `UPM_WS` PK = `ComputerName` | `ComputerName + SID` | Registry-sourced per `HKU\<SID>`, same reason |
| `LoggedOnUsers_WS` PK = `ComputerName + SID` | `ComputerName + SID + SessionId` | Same user can have console + RDP sessions simultaneously |
| `UserGPO_SystemInfo` PK = `ComputerName` | `ComputerName + UserName` | Must match siblings `UserGPO_AppliedGPOs` and `UserGPO_SecurityGroups` |
| Multi-row tables: `INSERT OR REPLACE` | `DELETE WHERE ComputerName = @cn` + `INSERT` in transaction | `REPLACE` only handles collisions — never removes disappeared rows (detached drives, unlinked GPOs, logged-off users) |
| `ComputerName` secondary indexes | None | PK auto-index already covers leftmost-prefix lookups — redundant write cost |
| Identity columns `TEXT NOT NULL` | `TEXT NOT NULL COLLATE NOCASE` | Prevents duplicate ghost rows from case differences across DNS/NetBIOS/WMI sources |

No open questions remain. Ready to rewrite the import script with:

1. Proper DDL from the schema
2. Two ingest paths — `INSERT OR REPLACE` for single-row tables, `DELETE` + `INSERT` in transaction for multi-row tables
3. YES/NO and True/False values converted to `1`/`0` on import
4. `UPM_WS` and `LoggedOnUsers_WS` using `SID` as part of the PK — which means the script needs to read `SID` from the CSV correctly

One thing to confirm before I start: the `UPM_WS` CSV currently doesn't have a `SID` column in the headers — it has `Username`. Since Opus says the PK should be `ComputerName + SID` because the data comes from `HKU\<SID>`, does your collection script actually output a `SID` column in `UPM_WS`, or is that something that needs to be added to the collection script first?
