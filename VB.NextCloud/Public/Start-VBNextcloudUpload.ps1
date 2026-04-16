# ============================================================
# FUNCTION : Start-VBNextcloudUpload
# MODULE   : VB.NextCloud
# VERSION  : 1.2.1
# CHANGED  : 15-04-2026 -- Fix -WhatIf not propagating to Set-VBNextcloudFile inner calls
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Batch uploads files to Nextcloud via WebDAV (wrapper around Set-VBNextcloudFile)
# ENCODING : UTF-8 with BOM
# ============================================================

function Start-VBNextcloudUpload {
    <#
    .SYNOPSIS
    Uploads one or more files to a Nextcloud server via WebDAV.

    .DESCRIPTION
    Start-VBNextcloudUpload is a batch wrapper around Set-VBNextcloudFile. It accepts an array
    of file paths, ensures the destination folder exists (using New-VBNextcloudFolder), and
    uploads each file in sequence. Files that do not exist locally are skipped with a warning
    and a Failed result object rather than terminating the pipeline.

    Credentials are passed as a PSCredential object. Plain-text username/password parameters
    are intentionally not supported -- use Get-Credential or ConvertTo-SecureString to build
    a PSCredential before calling this function.

    .PARAMETER Files
    One or more local file paths to upload. Accepts pipeline input.

    .PARAMETER BaseUrl
    Base URL of the Nextcloud instance, e.g. 'https://cloud.example.com'.

    .PARAMETER Credential
    PSCredential object containing the Nextcloud username and password.

    .PARAMETER DestinationPath
    Remote folder path where files will be uploaded, e.g. 'Vibhu/Reports'.
    The folder is created automatically if it does not exist.

    .PARAMETER WebDAVPath
    WebDAV endpoint path appended to BaseUrl. Defaults to 'remote.php/webdav'.

    .PARAMETER Overwrite
    When specified, replaces existing remote files. Without this switch, existing
    files cause a Failed result for that file without aborting the batch.

    .EXAMPLE
    $cred = Get-Credential
    Start-VBNextcloudUpload -Files 'C:\Reports\report1.csv','C:\Reports\report2.csv' `
        -BaseUrl 'https://cloud.example.com' -Credential $cred `
        -DestinationPath 'Vibhu/Reports'

    Uploads two files to the Vibhu/Reports folder.

    .EXAMPLE
    Get-ChildItem 'C:\Realtime\*.csv' | Select-Object -ExpandProperty FullName |
        Start-VBNextcloudUpload -BaseUrl 'https://cloud.example.com' `
            -Credential $cred -DestinationPath 'Vibhu/Reports' -Overwrite

    Pipes all CSV files from C:\Realtime and uploads them, overwriting existing remote files.

    .EXAMPLE
    $results = Start-VBNextcloudUpload -Files $filePaths -BaseUrl $url `
        -Credential $cred -DestinationPath 'Backups'
    $results | Where-Object { $_.Status -eq 'Failed' } | Select-Object SourceFile, Error

    Uploads a batch and filters the results to show only failed uploads.

    .OUTPUTS
    PSCustomObject
    One result object per file, each containing:
    - SourceFile     : Local file path
    - TargetPath     : Remote destination path
    - FileSize       : File size in bytes
    - Overwritten    : Whether an existing file was replaced
    - Status         : 'Success' or 'Failed'
    - Error          : Error message (only present on failure)
    - CollectionTime : Timestamp of the operation

    .NOTES
    Version : 1.2.1
    Author  : Vibhu Bhatnagar

    Requirements:
    - PowerShell 5.1 or higher
    - Network access to the Nextcloud instance
    - Valid Nextcloud credentials with write access to the destination folder
    - New-VBNextcloudFolder and Set-VBNextcloudFile must be available (included in this module)
    #>

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Files,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$BaseUrl,

        [Parameter(Mandatory)]
        [PSCredential]$Credential,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$DestinationPath,

        [ValidateNotNullOrEmpty()]
        [string]$WebDAVPath = 'remote.php/webdav',

        [switch]$Overwrite
    )

    begin {
        # Ensure the destination folder exists before the first upload
        Write-Verbose "Ensuring destination folder exists: $DestinationPath"
        $null = New-VBNextcloudFolder -BaseUrl $BaseUrl -Credential $Credential `
            -FolderPath $DestinationPath -WebDAVPath $WebDAVPath
    }

    process {
        foreach ($file in $Files) {
            if (Test-Path -Path $file -PathType Leaf) {
                Write-Verbose "Uploading: $file"
                Set-VBNextcloudFile -FilePath $file -BaseUrl $BaseUrl -Credential $Credential `
                    -DestinationPath $DestinationPath -WebDAVPath $WebDAVPath -Overwrite:$Overwrite `
                    -WhatIf:$WhatIfPreference
            }
            else {
                Write-Warning "File not found, skipping: $file"
                [PSCustomObject]@{
                    ComputerName   = $env:COMPUTERNAME
                    SourceFile     = $file
                    TargetPath     = $null
                    FileSize       = $null
                    Overwritten    = $false
                    Error          = 'File not found'
                    Status         = 'Failed'
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
