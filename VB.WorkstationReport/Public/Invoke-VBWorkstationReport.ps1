# ============================================================
# FUNCTION : Invoke-VBWorkstationReport
# MODULE   : VB.WorkstationReport
# VERSION  : 1.1.1
# CHANGED  : 14-04-2026 -- Standards compliance fixes
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Orchestrates workstation data collection and uploads reports to Nextcloud
# ENCODING : UTF-8 with BOM
# ============================================================

function Invoke-VBWorkstationReport {
    <#
    .SYNOPSIS
    Generates a full workstation report and uploads it to Nextcloud.

    .DESCRIPTION
    Invoke-VBWorkstationReport orchestrates the four data collection functions:
      - Get-VBUserPrinterMappings  -> <COMPUTERNAME>_UPM.csv
      - Get-VBSyncCenterStatus     -> <COMPUTERNAME>_CNC.csv
      - Get-VBUserFolderRedirections -> <COMPUTERNAME>_UFR.csv
      - Get-VBNetworkInterface     -> <COMPUTERNAME>_NI.csv

    Each report is saved as a CSV to the local output path, then all CSVs are uploaded
    to the configured Nextcloud destination folder using Set-VBNextcloudFile.

    Existing CSVs in the output path are cleared before collection begins. Credentials
    are always passed as a PSCredential -- plain-text passwords are not accepted.

    .PARAMETER Credential
    PSCredential for Nextcloud authentication. Use Get-Credential or build a PSCredential
    from a SecureString. Required.

    .PARAMETER NextcloudBaseUrl
    Base URL of the Nextcloud instance. Defaults to 'https://vault.dediserve.com'.

    .PARAMETER NextcloudDestination
    Remote folder path where reports will be uploaded. Defaults to 'Vibhu/Reports'.

    .PARAMETER OutputPath
    Local folder path where CSV reports are written before upload.
    Defaults to 'C:\Realtime'. The folder is created if it does not exist.

    .PARAMETER SkipUpload
    When specified, reports are generated and saved locally but not uploaded to Nextcloud.
    Useful for testing or when Nextcloud is unavailable.

    .EXAMPLE
    $cred = Get-Credential
    Invoke-VBWorkstationReport -Credential $cred

    Runs the full report against the local machine with default paths and uploads to Nextcloud.

    .EXAMPLE
    $secPw  = ConvertTo-SecureString 'MyAppPassword' -AsPlainText -Force
    $cred   = New-Object PSCredential('myuser', $secPw)
    Invoke-VBWorkstationReport -Credential $cred `
        -NextcloudBaseUrl 'https://cloud.example.com' `
        -NextcloudDestination 'IT/Workstations' `
        -OutputPath 'D:\Reports'

    Generates reports to D:\Reports and uploads to a custom Nextcloud path.

    .EXAMPLE
    Invoke-VBWorkstationReport -Credential $cred -SkipUpload -OutputPath 'C:\Temp\Reports'

    Generates reports locally without uploading -- useful for offline testing.

    .OUTPUTS
    PSCustomObject
    Returns a summary object with:
    - ComputerName    : Machine the report was collected from
    - OutputPath      : Local folder where CSVs were written
    - ReportsGenerated: Number of CSV files successfully created
    - UploadResults   : Array of upload result objects (empty when -SkipUpload is used)
    - Duration        : Total execution time
    - Status          : 'Success', 'PartialFailure', or 'Failed'
    - CollectionTime  : Timestamp of the run

    .NOTES
    Version : 1.1.1
    Author  : Vibhu Bhatnagar
    Category: Windows Workstation Administration

    Requirements:
    - PowerShell 5.1 or higher
    - Administrative privileges (for registry hive access)
    - Network access to Nextcloud (unless -SkipUpload is specified)
    - WorkstationReport module functions must be loaded

    Security note:
    Never store credentials in plain text inside scripts. Use Get-Credential interactively,
    or retrieve credentials from a secrets manager / encrypted credential store.
    #>

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [PSCredential]$Credential,

        [ValidateNotNullOrEmpty()]
        [string]$NextcloudBaseUrl = 'https://vault.dediserve.com',

        [ValidateNotNullOrEmpty()]
        [string]$NextcloudDestination = 'Vibhu/Reports',

        [ValidateNotNullOrEmpty()]
        [string]$OutputPath = 'C:\Realtime',

        [switch]$SkipUpload
    )

    $startTime      = Get-Date
    $collectionTime = $startTime.ToString('dd-MM-yyyy HH:mm:ss')
    $computerName   = $env:COMPUTERNAME
    $csvFiles       = [System.Collections.Generic.List[string]]::new()
    $errors         = [System.Collections.Generic.List[string]]::new()

    try {
        # Ensure output directory exists
        if (-not (Test-Path $OutputPath)) {
            Write-Verbose "Creating output directory: $OutputPath"
            $null = New-Item -Path $OutputPath -ItemType Directory -Force
        }

        # Clear previous reports
        Write-Verbose "Clearing existing CSVs from: $OutputPath"
        Remove-Item -Path (Join-Path $OutputPath '*.csv') -Force -ErrorAction SilentlyContinue

        # -- Report 1: User Printer Mappings -----------------------------------------
        $csvPath = Join-Path $OutputPath "${computerName}_UPM.csv"
        Write-Verbose "Collecting printer mappings..."
        try {
            Get-VBUserPrinterMappings -TableOutput |
                Export-Csv -Path $csvPath -NoTypeInformation -Force
            $csvFiles.Add($csvPath)
            Write-Verbose "Saved: $csvPath"
        }
        catch {
            $errors.Add("UPM: $($_.Exception.Message)")
            Write-Warning "Printer mapping collection failed: $($_.Exception.Message)"
        }

        # -- Report 2: Sync Center Status ---------------------------------------------
        $csvPath = Join-Path $OutputPath "${computerName}_CNC.csv"
        Write-Verbose "Collecting Sync Center status..."
        try {
            Get-VBSyncCenterStatus |
                Export-Csv -Path $csvPath -NoTypeInformation -Force
            $csvFiles.Add($csvPath)
            Write-Verbose "Saved: $csvPath"
        }
        catch {
            $errors.Add("CNC: $($_.Exception.Message)")
            Write-Warning "Sync Center collection failed: $($_.Exception.Message)"
        }

        # -- Report 3: User Folder Redirections ---------------------------------------
        $csvPath = Join-Path $OutputPath "${computerName}_UFR.csv"
        Write-Verbose "Collecting folder redirections..."
        try {
            Get-VBUserFolderRedirections -TableOutput |
                Export-Csv -Path $csvPath -NoTypeInformation -Force
            $csvFiles.Add($csvPath)
            Write-Verbose "Saved: $csvPath"
        }
        catch {
            $errors.Add("UFR: $($_.Exception.Message)")
            Write-Warning "Folder redirection collection failed: $($_.Exception.Message)"
        }

        # -- Report 4: Network Interfaces ---------------------------------------------
        $csvPath = Join-Path $OutputPath "${computerName}_NI.csv"
        Write-Verbose "Collecting network interfaces..."
        try {
            Get-VBNetworkInterface |
                Export-Csv -Path $csvPath -NoTypeInformation -Force
            $csvFiles.Add($csvPath)
            Write-Verbose "Saved: $csvPath"
        }
        catch {
            $errors.Add("NI: $($_.Exception.Message)")
            Write-Warning "Network interface collection failed: $($_.Exception.Message)"
        }

        # -- Upload --------------------------------------------------------------------
        $uploadResults = @()
        if (-not $SkipUpload -and $csvFiles.Count -gt 0) {
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
        elseif ($SkipUpload) {
            Write-Verbose 'Upload skipped (-SkipUpload specified).'
        }

        $duration   = (Get-Date) - $startTime
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
    }
    catch {
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
