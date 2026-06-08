**Prerequisites — one time only:**

```powershell
Install-Module PSSQLite -Scope CurrentUser
```

---

**Run it:**

```powershell
.\Import-VBDsiReportsToDB.ps1 -CsvFolder 'C:\Path\To\CSVs'
```

That creates `DSI_Reports.db` in the same folder as the script.

---

**Common variations:**

```powershell
# Put the DB somewhere specific
.\Import-VBDsiReportsToDB.ps1 -CsvFolder 'C:\Reports\DSI' -DatabasePath 'C:\DB\DSI.db'

# Re-run and overwrite an existing DB
.\Import-VBDsiReportsToDB.ps1 -CsvFolder 'C:\Reports\DSI' -Force

# See step-by-step progress
.\Import-VBDsiReportsToDB.ps1 -CsvFolder 'C:\Reports\DSI' -Verbose
```

---

**For a different client (e.g. AGG):**

1. Open the script, change line 59: 

   ```powershell
   $CLIENT_NAME = 'AGG'
   
   ```
2. Run pointing at that client's CSV folder: 

   ```powershell
   .\Import-VBDsiReportsToDB.ps1 -CsvFolder 'C:\Reports\AGG'
   
   ```

   Creates `AGG_Reports.db` automatically.

---

**Expected output on success:**

```shell
Database: C:\Scripts\DSI_Reports.db

Table                    CsvFile                              RowsImported Status
-----                    -------                              ------------ ------
AzJoinStatus_WS          DSI_AzJoinStatus_WS_Report.csv                2 Imported
CSC_WS                   DSI_CSC_WS_Report.csv                         2 Imported
DiskInventory_WS         DSI_DiskInventory_WS_Report.csv               2 Imported
...
NIC_WS                   DSI_NIC_WS_Status.csv                         1 Imported
...
```

Any row showing `TableOnly` means that CSV wasn't found in the folder — table still exists in the DB, just empty. Any `Error:` row means the import failed for that report and the message tells you why.
