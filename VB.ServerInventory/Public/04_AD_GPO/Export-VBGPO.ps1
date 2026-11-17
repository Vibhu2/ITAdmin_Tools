# ============================================================
# SCRIPT  : Export-GPO_AIO.ps1
# VERSION : 1.0
# CHANGED : 17-11-2026 — Initial release
# AUTHOR  : Admin
# PURPOSE : Backup + HTML reports +XML reports+ GUID-to-name mapping for all GPOs
# ============================================================

# --- CONFIGURATION ---
$ROOT_PATH   = 'C:\Realtime\GPO_AIO'
$BACKUP_PATH = Join-Path $ROOT_PATH 'Backup'
$REPORT_PATH_HTML = Join-Path $ROOT_PATH 'Reports_HTML'
$REPORT_PATH_XML = Join-Path $ROOT_PATH 'Reports_xml'
$CSV_PATH    = Join-Path $ROOT_PATH 'GPO_Mapping.csv'

# --- MAIN LOGIC ---
# Step 1 - Create folders
foreach ($path in $BACKUP_PATH, $REPORT_PATH_HTML, $REPORT_PATH_XML) {
    New-Item -Path $path -ItemType Directory -Force | Out-Null
}

# Step 2 - Backup all GPOs (folders are named by BackupId)
$Backups = Backup-GPO -All -Path $BACKUP_PATH

# Step 3 - HTML report per GPO + build mapping
$Mapping = foreach ($gpo in Get-GPO -All) {
    $safeName   = $gpo.DisplayName -replace '[\\/:*?"<>|]', '_'
    $reportFile = Join-Path $REPORT_PATH_HTML "$safeName.html"
    $backup     = $Backups | Where-Object { $_.GpoId -eq $gpo.Id }

    Get-GPOReport -Guid $gpo.Id -ReportType Html -Path $reportFile | Out-Null

    [PSCustomObject]@{
        ComputerName     = $env:COMPUTERNAME
        DisplayName      = $gpo.DisplayName
        GpoId            = $gpo.Id
        BackupId         = $backup.Id
        BackupFolder     = "{$($backup.Id)}"
        ReportFile       = "$safeName.html"
        GpoStatus        = $gpo.GpoStatus
        ModificationTime = $gpo.ModificationTime.ToString('dd-MM-yyyy HH:mm:ss')
    }
}

# Step 4 - xml report per GPO + build mapping
$Mapping = foreach ($gpo in Get-GPO -All) {
    $safeName   = $gpo.DisplayName -replace '[\\/:*?"<>|]', '_'
    $reportFile = Join-Path $REPORT_PATH_XML "$safeName.xml"
    $backup     = $Backups | Where-Object { $_.GpoId -eq $gpo.Id }

    Get-GPOReport -Guid $gpo.Id -ReportType xml -Path $reportFile | Out-Null

    [PSCustomObject]@{
        ComputerName     = $env:COMPUTERNAME
        DisplayName      = $gpo.DisplayName
        GpoId            = $gpo.Id
        BackupId         = $backup.Id
        BackupFolder     = "{$($backup.Id)}"
        ReportFile       = "$safeName.xml"
        GpoStatus        = $gpo.GpoStatus
        ModificationTime = $gpo.ModificationTime.ToString('dd-MM-yyyy HH:mm:ss')
    }
}
# Step 5 - Mapping CSV at root
$Mapping | Export-Csv -Path $CSV_PATH -NoTypeInformation