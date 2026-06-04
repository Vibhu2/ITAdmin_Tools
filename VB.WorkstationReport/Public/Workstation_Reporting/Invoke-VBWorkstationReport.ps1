# ============================================================
# FUNCTION : Invoke-VBWorkstationReport
# MODULE   : VB.WorkstationReport
# VERSION  : 1.5.1
# CHANGED  : 16-04-2026 -- All parameters made mandatory, defaults removed.
#                          OutputPath validated upfront -- warns and stops if missing.
#                          OutputPath validated again before upload.
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Orchestrates workstation data collection and uploads reports to Nextcloud
# ENCODING : UTF-8 with BOM
# ============================================================

function Invoke-VBWorkstationReport {
    <#
    .SYNOPSIS
    Generates a full workstation report and uploads it to Nextcloud.

    .DESCRIPTION
    Invoke-VBWorkstationReport orchestrates seven data collection functions:
      - Get-VBNetworkInterface           -> <DomainName>_<COMPUTERNAME>_NIC.csv
      - Get-VBOneDriveFolderBackupStatus -> <DomainName>_<COMPUTERNAME>_ODFB.csv
      - Get-VBSyncCenterStatus           -> <DomainName>_<COMPUTERNAME>_CNC.csv
      - Get-VBUserFolderRedirections     -> <DomainName>_<COMPUTERNAME>_UFR.csv
      - Get-VBUserPrinterMappings        -> <DomainName>_<COMPUTERNAME>_UPM.csv
      - Get-VBUserProfile                -> <DomainName>_<COMPUTERNAME>_UP.csv
      - Get-VBUserShellFolders           -> <DomainName>_<COMPUTERNAME>_USF.csv
      - Get-vbuptimenabled              -> <DomainName>_<COMPUTERNAME>_UPT.csv
      - Get-vbsystnemgpresult           -> <DomainName>_<COMPUTERNAME>_sysGpResult-systeminfo.csv
      - Get-vbsystnemgpresult           -> <DomainName>_<COMPUTERNAME>_sysGpResult-appliedgpos.csv
      - Get-vbsystnemgpresult           -> <DomainName>_<COMPUTERNAME>_sysGpResult-securitygroups.csv
      - Get-VBLoggedOnUser           -> <DomainName_><COMPUTERNAME>_LoggedOnUsers.csv
      - Get-VBLoggedOnUserGpResult -> <DomainName>_<COMPUTERNAME>_UserGpResult-systeminfo.csv'
      - Get-VBLoggedOnUserGpResult -> <DomainName>_<COMPUTERNAME>_UserGpResult-appliedgpos.csv'
      - Get-VBLoggedOnUserGpResult -> <DomainName>_<COMPUTERNAME>_UserGpResult-securitygroups.csv'
      - Get-VBDiskInventory                  -> <DomainName>_<COMPUTERNAME>_DiskInventory.csv

    All parameters are mandatory -- no defaults are set. OutputPath must exist before
    the function runs; if it does not exist the function warns and stops immediately.
    Before upload, OutputPath is validated again and upload is skipped with a warning
    if no CSVs are found.

    .PARAMETER Credential
    PSCredential for Nextcloud authentication. Mandatory.

    .PARAMETER NextcloudBaseUrl
    Base URL of the Nextcloud instance. Mandatory. Example: 'https://vault.dediserve.com'

    .PARAMETER NextcloudDestination
    Remote folder path on Nextcloud where reports will be uploaded. Mandatory.
    Example: 'Realtime-IT/Reports'

    .PARAMETER OutputPath
    Local folder path where CSV reports are written before upload. Mandatory.
    The folder must already exist -- this function will not create it.
    Example: 'C:\Realtime\Reports'

    .PARAMETER SkipUpload
    When specified, reports are generated and saved locally but not uploaded to Nextcloud.
    Useful for testing or when Nextcloud is unavailable.

    .EXAMPLE
    $cred = New-Object PSCredential('username', (ConvertTo-SecureString 'AppPassword' -AsPlainText -Force))
    Invoke-VBWorkstationReport -Credential $cred `
        -NextcloudBaseUrl 'https://vault.dediserve.com' `
        -NextcloudDestination 'Realtime-IT/Reports' `
        -OutputPath 'C:\Realtime\Reports'

    .EXAMPLE
    Invoke-VBWorkstationReport -Credential $cred `
        -NextcloudBaseUrl 'https://vault.dediserve.com' `
        -NextcloudDestination 'Realtime-IT/Reports' `
        -OutputPath 'C:\Realtime\Reports' `
        -SkipUpload

    Generates all seven reports locally without uploading to Nextcloud.

    .OUTPUTS
    PSCustomObject
    Returns a summary object with:
    - ComputerName    : Machine the report was collected from
    - OutputPath      : Local folder where CSVs were written
    - ReportsGenerated: Number of CSV files successfully created
    - UploadResults   : Array of upload result objects (empty when -SkipUpload is used)
    - Errors          : Semicolon-separated error messages (null on full success)
    - Duration        : Total execution time
    - Status          : 'Success', 'PartialFailure', or 'Failed'
    - CollectionTime  : Timestamp of the run

    .NOTES
    Version : 1.5.1
    Author  : Vibhu Bhatnagar
    Category: Windows Workstation Administration

    Requirements:
    - PowerShell 5.1 or higher
    - Administrative privileges (for registry hive access)
    - OutputPath directory must exist before calling this function
    - Network access to Nextcloud (unless -SkipUpload is specified)
    - VB.WorkstationReport and VB.NextCloud modules must be installed

    Security note:
    Never store credentials in plain text inside scripts. Use Get-Credential interactively,
    or retrieve credentials from a secrets manager / encrypted credential store.
    #>

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [PSCredential]$Credential,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$NextcloudBaseUrl,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$NextcloudDestination,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$OutputPath,

        [switch]$SkipUpload
    )

    $startTime = Get-Date
    $collectionTime = $startTime.ToString('dd-MM-yyyy HH:mm:ss')
    $computerName = $env:COMPUTERNAME
    $DomainName = $env:USERDOMAIN
    $csvFiles = [System.Collections.Generic.List[string]]::new()
    $errors = [System.Collections.Generic.List[string]]::new()

    # -- Pre-flight: OutputPath must exist ------------------------------------------
    if (-not (Test-Path -Path $OutputPath -PathType Container)) {
        Write-Warning "Invoke-VBWorkstationReport: OutputPath '$OutputPath' does not exist. Create the folder first and re-run."
        return [PSCustomObject]@{
            ComputerName     = $computerName
            OutputPath       = $OutputPath
            ReportsGenerated = 0
            UploadResults    = @()
            Errors           = "OutputPath '$OutputPath' does not exist."
            Duration         = '00:00.000'
            Status           = 'Failed'
            CollectionTime   = $collectionTime

        }
    }

    try {
        # Clear previous reports
        Write-Verbose "Clearing existing CSVs from: $OutputPath"
        Remove-Item -Path (Join-Path $OutputPath '*.csv') -Force -ErrorAction SilentlyContinue

        # -- Report 1: Network Interfaces ---------------------------------------------
        $csvPath = Join-Path $OutputPath "${DomainName}_${computerName}_NIC.csv"
        Write-Verbose "Collecting network interfaces..."
        try {
            Get-VBNetworkInterface |
            Export-Csv -Path $csvPath -NoTypeInformation -Force
            $csvFiles.Add($csvPath)
            Write-Verbose "Saved: $csvPath"
        } catch {
            $errors.Add("NIC: $($_.Exception.Message)")
            Write-Warning "Network interface collection failed: $($_.Exception.Message)"
        }

        # -- Report 2: OneDrive Folder Backup Status ----------------------------------
        $csvPath = Join-Path $OutputPath "${DomainName}_${computerName}_ODFB.csv"
        Write-Verbose "Collecting OneDrive folder backup status..."
        try {
            Get-VBOneDriveFolderBackupStatus |
            Export-Csv -Path $csvPath -NoTypeInformation -Force
            $csvFiles.Add($csvPath)
            Write-Verbose "Saved: $csvPath"
        } catch {
            $errors.Add("ODFB: $($_.Exception.Message)")
            Write-Warning "OneDrive folder backup collection failed: $($_.Exception.Message)"
        }

        # -- Report 3: Sync Center Status ---------------------------------------------
        $csvPath = Join-Path $OutputPath "${DomainName}_${computerName}_CSC.csv"
        Write-Verbose "Collecting Sync Center status..."
        try {
            Get-VBSyncCenterStatus |
            Export-Csv -Path $csvPath -NoTypeInformation -Force
            $csvFiles.Add($csvPath)
            Write-Verbose "Saved: $csvPath"
        } catch {
            $errors.Add("CSC: $($_.Exception.Message)")
            Write-Warning "Sync Center collection failed: $($_.Exception.Message)"
        }

        # -- Report 4: User Folder Redirections ---------------------------------------
        $csvPath = Join-Path $OutputPath "${DomainName}_${computerName}_UFR.csv"
        Write-Verbose "Collecting folder redirections..."
        try {
            Get-VBUserFolderRedirections -TableOutput |
            Export-Csv -Path $csvPath -NoTypeInformation -Force
            $csvFiles.Add($csvPath)
            Write-Verbose "Saved: $csvPath"
        } catch {
            $errors.Add("UFR: $($_.Exception.Message)")
            Write-Warning "Folder redirection collection failed: $($_.Exception.Message)"
        }

        # -- Report 5: User Printer Mappings ------------------------------------------
        $csvPath = Join-Path $OutputPath "${DomainName}_${computerName}_UPM.csv"
        Write-Verbose "Collecting printer mappings..."
        try {
            Get-VBUserPrinterMappings -TableOutput |
            Export-Csv -Path $csvPath -NoTypeInformation -Force
            $csvFiles.Add($csvPath)
            Write-Verbose "Saved: $csvPath"
        } catch {
            $errors.Add("UPM: $($_.Exception.Message)")
            Write-Warning "Printer mapping collection failed: $($_.Exception.Message)"
        }

        # -- Report 6: User Profiles --------------------------------------------------
        $csvPath = Join-Path $OutputPath "${DomainName}_${computerName}_UP.csv"
        Write-Verbose "Collecting user profiles..."
        try {
            Get-VBUserProfile |
            Export-Csv -Path $csvPath -NoTypeInformation -Force
            $csvFiles.Add($csvPath)
            Write-Verbose "Saved: $csvPath"
        } catch {
            $errors.Add("UP: $($_.Exception.Message)")
            Write-Warning "User profile collection failed: $($_.Exception.Message)"
        }

        # -- Report 7: User Shell Folders ---------------------------------------------
        $csvPath = Join-Path $OutputPath "${DomainName}_${computerName}_USF.csv"
        Write-Verbose "Collecting user shell folders..."
        try {
            Get-VBUserShellFolders |
            Export-Csv -Path $csvPath -NoTypeInformation -Force
            $csvFiles.Add($csvPath)
            Write-Verbose "Saved: $csvPath"
        } catch {
            $errors.Add("USF: $($_.Exception.Message)")
            Write-Warning "User shell folder collection failed: $($_.Exception.Message)"
        }
        # -- Report 8 AzJoinStatus -------------------------------------------------------------

        $csvPath = Join-Path $OutputPath "${DomainName}_${computerName}_AzJoinStatus.csv"
        Write-Verbose "Collecting user Device Azure Joined Status..."
        try {
            Get-VBAzureJoinStatus | Export-Csv -Path (Join-Path $OutputPath "${DomainName}_${computerName}_AzJoinStatus.csv") -NoTypeInformation -Force
        } catch {
            $errors.Add("DsRegStatus: $($_.Exception.Message)")
            Write-Warning "User Device Azure Joined Status collection failed: $($_.Exception.Message)"
        }
        # -- Report 9 System GPO Result ---------------------------------------------------------
        Write-Verbose "Collecting system GPO result data..."
        try {
            $gpResult = Get-VBSystemGpResult

            $csvPath = Join-Path $OutputPath "${DomainName}_${computerName}_sysGpResult-systeminfo.csv"
            $gpResult | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
            $csvFiles.Add($csvPath)
            Write-Verbose "Saved: $csvPath"

            $csvPath = Join-Path $OutputPath "${DomainName}_${computerName}_sysGpResult-appliedgpos.csv"
            $gpResult.AppliedGPOs | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
            $csvFiles.Add($csvPath)
            Write-Verbose "Saved: $csvPath"

            $csvPath = Join-Path $OutputPath "${DomainName}_${computerName}_sysGpResult-securitygroups.csv"
            $gpResult.SecurityGroups | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
            $csvFiles.Add($csvPath)
            Write-Verbose "Saved: $csvPath"

        } catch {
            $errors.Add("SystemGPOResult: $($_.Exception.Message)")
            Write-Warning "System GPO result collection failed: $($_.Exception.Message)"
        }

        # -- Report 10 Logged On Users -----------------------------------------------------------
        Write-Verbose "Collecting logged on user sessions..."
        try {
            $csvPath = Join-Path $OutputPath "${DomainName}_${computerName}_LoggedOnUsers.csv"
            Get-VBLoggedOnUser | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
            $csvFiles.Add($csvPath)
            Write-Verbose "Saved: $csvPath"
        } catch {
            $errors.Add("LoggedOnUsers: $($_.Exception.Message)")
            Write-Warning "Logged on user session collection failed: $($_.Exception.Message)"
        }

        # -- Report 11 Logged On User GP Results -------------------------------------------------
        Write-Verbose "Collecting logged on user GPO results..."
        try {
            $csvPath = Join-Path $OutputPath "${DomainName}_${ComputerName}_UserGpResult-SystemInfo.csv"
            Get-VBLoggedOnUserGpResult | Select-Object -Property ComputerName, Status, UserName, LastApplied, AppliedFrom, SlowLink, DomainName, DomainType, CollectionTime | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
            $csvFiles.Add($csvPath)
            Write-Verbose "Saved: $csvPath"

            $csvPath = Join-Path $OutputPath "${DomainName}_${ComputerName}_UserGpResult-AppliedGPOs.csv"
            Get-VBLoggedOnUserGpResult | Select-Object -ExpandProperty AppliedGPOs | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
            $csvFiles.Add($csvPath)
            Write-Verbose "Saved: $csvPath"

            $csvPath = Join-Path $OutputPath "${DomainName}_${ComputerName}_UserGpResult-SecurityGroups.csv"
            Get-VBLoggedOnUserGpResult | Select-Object -ExpandProperty SecurityGroups | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
            $csvFiles.Add($csvPath)
            Write-Verbose "Saved: $csvPath"

        } catch {
            $errors.Add("UserGpResult-systeminfo: $($_.Exception.Message)")
            $errors.Add("UserGpResult-AppliedGPOs: $($_.Exception.Message)")
            $errors.Add("UserGpResult-SecurityGroups: $($_.Exception.Message)")
            Write-Warning "Logged on user GPO result collection failed: $($_.Exception.Message)"
        }
        # -- Report 12 Disk Inventory -------------------------------------------------
        $csvPath = Join-Path $OutputPath "${DomainName}_${computerName}_DiskInventory.csv"
        Write-Verbose "Collecting disk inventory..."
        try {
            Get-VBDiskInventory |
            Export-Csv -Path $csvPath -NoTypeInformation -Force
            $csvFiles.Add($csvPath)
            Write-Verbose "Saved: $csvPath"
        } catch {
            $errors.Add("DiskInventory: $($_.Exception.Message)")
            Write-Warning "Disk inventory collection failed: $($_.Exception.Message)"
        }
        #--Report 13 Logon Server / Join Type ---------------------------------------------------------
        $csvPath = Join-Path $OutputPath "${DomainName}_${computerName}_SystemADType.csv"
        Write-Verbose "Collecting system AD join type and logon server information  "
        try {
            Get-VBAzureJoinStatus | Get-VBJoinType |
            Export-Csv -Path $csvPath -NoTypeInformation -Force
            $csvFiles.Add($csvPath)
            Write-Verbose "Saved: $csvPath"
        } catch {
            $errors.Add("SystemADType: $($_.Exception.Message)")
            Write-Warning "System AD join type and logon server collection failed: $($_.Exception.Message)"
        }
        # -- Upload -------------------------------------------------------------------
        $uploadResults = @()
        if (-not $SkipUpload) {
            # Validate OutputPath still exists and has CSVs before attempting upload
            if (-not (Test-Path -Path $OutputPath -PathType Container)) {
                $msg = "Upload skipped: OutputPath '$OutputPath' no longer exists."
                $errors.Add($msg)
                Write-Warning $msg
            } elseif ($csvFiles.Count -eq 0) {
                $msg = "Upload skipped: no CSV files were generated in '$OutputPath'."
                $errors.Add($msg)
                Write-Warning $msg
            } else {
                Write-Verbose "Uploading $($csvFiles.Count) file(s) to $NextcloudBaseUrl/$NextcloudDestination"
                $uploadResults = $csvFiles |
                Set-VBNextcloudFile -BaseUrl $NextcloudBaseUrl `
                    -Credential $Credential -DestinationPath $NextcloudDestination -Overwrite

                $failedUploads = @($uploadResults | Where-Object { $_.Status -eq 'Failed' })
                if ($failedUploads.Count -gt 0) {
                    $failedUploads | ForEach-Object {
                        $errors.Add("Upload failed for $($_.SourceFile): $($_.Error)")
                        Write-Warning "Upload failed: $($_.SourceFile) -- $($_.Error)"
                    }
                }
            }
        } else {
            Write-Verbose 'Upload skipped (-SkipUpload specified).'
        }

        $duration = (Get-Date) - $startTime
        $statusCode = if ($errors.Count -eq 0) { 'Success' }
        elseif ($csvFiles.Count -gt 0) { 'PartialFailure' }
        else { 'Failed' }

        [PSCustomObject]@{
            ComputerName     = $computerName
            OutputPath       = $OutputPath
            ReportsGenerated = $csvFiles.Count
            UploadResults    = $uploadResults
            Errors           = if ($errors.Count) { $errors -join '; ' } else { $null }
            Duration         = '{0:mm\:ss\.fff}' -f $duration
            Status           = $statusCode
            CollectionTime   = $collectionTime
        }
    } catch {
        Write-Error -Message "Invoke-VBWorkstationReport failed: $($_.Exception.Message)"
        [PSCustomObject]@{
            ComputerName     = $computerName
            OutputPath       = $OutputPath
            ReportsGenerated = $csvFiles.Count
            UploadResults    = @()
            Errors           = $_.Exception.Message
            Duration         = '{0:mm\:ss\.fff}' -f ((Get-Date) - $startTime)
            Status           = 'Failed'
            CollectionTime   = $collectionTime
        }
    }
}
