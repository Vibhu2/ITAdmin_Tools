# =========================
# CONFIGURATION (EDIT HERE ONLY)
# =========================

# Common Variables Used Across Multiple Reports
$BasePath        = "C:\Users\Vibhu.Bhatnagar\Nextcloud\Realtime-IT\Reports\DPC_Reports"
$ExportReport    = "C:\Realtime"

# File patterns
$CSC_Pattern     = "*_CNC.csv"
$UFR_Pattern     = "*_UFR.csv"
$UPM_Pattern     = "*_UPM.csv"
$NIC_Pattern     = "*_NIC.csv"
$GPO_Pattern     = "*_GPO.csv"
$ODFB_Pattern    = "*_ODFB.csv"
$UP_Pattern      = "*_UP.csv"

# Output files
$CSC_Output      = "DPC_CSC_WS_Report.csv"
$UFR_Output      = "DPC_UFR_WS_Report.csv"
$UPM_Output      = "DPC_UPM_WS_Report.csv"
$NIC_Output      = "DPC_NIC_WS_Status.csv"
$GPO_Output      = "DPC_GPO_WS_Report.csv"
$ODFB_Output     = "DPC_ODFB_WS_Report.csv"
$UP_Output       = "DPC_UP_WS_Report.csv"

# =========================
# SCRIPT START
# =========================

Set-Location $BasePath

# -------------------------
# CSC Report
# -------------------------
$CscFiles  = Get-ChildItem -Path .\ -Filter $CSC_Pattern
$CscData   = $CscFiles | ForEach-Object { Import-Csv $_.FullName }
$CscData   | Export-Csv -Path (Join-Path $ExportReport $CSC_Output) -NoTypeInformation

# -------------------------
# Folder Redirection Report
# -------------------------
$UfrFiles  = Get-ChildItem -Path .\ -Filter $UFR_Pattern
$UfrData   = $UfrFiles | ForEach-Object { Import-Csv $_.FullName }
$UfrData   | Export-Csv -Path (Join-Path $ExportReport $UFR_Output) -NoTypeInformation

# -------------------------
# Network Printer Report
# -------------------------
$UpmFiles  = Get-ChildItem -Path .\ -Filter $UPM_Pattern
$UpmData   = $UpmFiles | ForEach-Object { Import-Csv $_.FullName }
$UpmData   | Export-Csv -Path (Join-Path $ExportReport $UPM_Output) -NoTypeInformation

# -------------------------
# Network Details Report
# -------------------------
$NicFiles  = Get-ChildItem -Path .\ -Filter $NIC_Pattern
$NicData   = $NicFiles | ForEach-Object { Import-Csv $_.FullName }
$NicData   | Export-Csv -Path (Join-Path $ExportReport $NIC_Output) -NoTypeInformation

# -------------------------
# GPO Report
# -------------------------
$GpoFiles  = Get-ChildItem -Path .\ -Filter $GPO_Pattern
$GpoData   = $GpoFiles | ForEach-Object { Import-Csv $_.FullName }
$GpoData   | Export-Csv -Path (Join-Path $ExportReport $GPO_Output) -NoTypeInformation

# -------------------------
# OneDrive Status Report
# -------------------------
$OdfbFiles = Get-ChildItem -Path .\ -Filter $ODFB_Pattern
$OdfbData  = $OdfbFiles | ForEach-Object { Import-Csv $_.FullName }
$OdfbData  | Export-Csv -Path (Join-Path $ExportReport $ODFB_Output) -NoTypeInformation

# -------------------------
# User Profile Report
# -------------------------
$UpFiles   = Get-ChildItem -Path .\ -Filter $UP_Pattern
$UpData    = $UpFiles | ForEach-Object { Import-Csv $_.FullName }
$UpData    | Export-Csv -Path (Join-Path $ExportReport $UP_Output) -NoTypeInformation

# =========================
# SUMMARY
# =========================

Clear-Host

Write-Host '=================================================================='
Write-Host '  DSI Workstation Report -- Export Summary'
Write-Host "  $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')"
Write-Host '=================================================================='
Write-Host ''
Write-Host ('  {0,-6}  {1,-26}  {2,7}  {3,10}' -f 'Report', 'Name', 'Files', 'Records')
Write-Host ('  {0,-6}  {1,-26}  {2,7}  {3,10}' -f '------', '--------------------------', '-------', '----------')
Write-Host ('  {0,-6}  {1,-26}  {2,7}  {3,10}' -f 'CSC',  'CSC Sync',                   $CscFiles.Count,  $CscData.Count)
Write-Host ('  {0,-6}  {1,-26}  {2,7}  {3,10}' -f 'UFR',  'User Folder Redirection',     $UfrFiles.Count,  $UfrData.Count)
Write-Host ('  {0,-6}  {1,-26}  {2,7}  {3,10}' -f 'UPM',  'User Printer Management',     $UpmFiles.Count,  $UpmData.Count)
Write-Host ('  {0,-6}  {1,-26}  {2,7}  {3,10}' -f 'NIC',  'Network Config Report',       $NicFiles.Count,  $NicData.Count)
Write-Host ('  {0,-6}  {1,-26}  {2,7}  {3,10}' -f 'GPO',  'Group Policy Usage',          $GpoFiles.Count,  $GpoData.Count)
Write-Host ('  {0,-6}  {1,-26}  {2,7}  {3,10}' -f 'ODFB', 'OneDrive Folder Report',      $OdfbFiles.Count, $OdfbData.Count)
Write-Host ('  {0,-6}  {1,-26}  {2,7}  {3,10}' -f 'UP',   'User Profile Details',        $UpFiles.Count,   $UpData.Count)
Write-Host ''
Write-Host '  All reports exported successfully.'
Write-Host '=================================================================='