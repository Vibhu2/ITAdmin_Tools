# ============================================================
# FUNCTION : Set-VBNextcloudFile
# MODULE   : VB.NextCloud
# VERSION  : 1.1.1
# CHANGED  : 14-04-2026 -- Standards compliance fixes
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Uploads a single file to Nextcloud via WebDAV PUT
# ENCODING : UTF-8 with BOM
# ============================================================

function Set-VBNextcloudFile {
    <#
    .SYNOPSIS
    Uploads a single file to a Nextcloud server via WebDAV.

    .DESCRIPTION
    Set-VBNextcloudFile uploads a local file to a specified path on a Nextcloud server using
    the WebDAV protocol. It checks for an existing remote file before uploading and honours
    the -Overwrite switch to control replacement behaviour. TLS 1.2 is enforced for all
    connections. The function returns a structured result object whether the upload succeeds
    or fails, making it safe to use in pipelines and automated workflows.

    .PARAMETER FilePath
    Full local path to the file to upload. Accepts pipeline input. The file must exist.

    .PARAMETER BaseUrl
    Base URL of the Nextcloud instance, e.g. 'https://cloud.example.com'.

    .PARAMETER Credential
    PSCredential object containing the Nextcloud username and password.

    .PARAMETER DestinationPath
    Remote folder path on the Nextcloud server where the file will be placed,
    e.g. 'Vibhu/Reports'. The folder must already exist -- use New-VBNextcloudFolder first.

    .PARAMETER WebDAVPath
    WebDAV endpoint path appended to BaseUrl. Defaults to 'remote.php/webdav'.

    .PARAMETER Overwrite
    When specified, replaces an existing remote file. Without this switch, the function
    throws if the file already exists.

    .EXAMPLE
    Set-VBNextcloudFile -FilePath 'C:\Reports\report.csv' `
        -BaseUrl 'https://cloud.example.com' `
        -Credential (Get-Credential) `
        -DestinationPath 'Vibhu/Reports'

    Uploads report.csv to the Vibhu/Reports folder on the Nextcloud server.

    .EXAMPLE
    Get-ChildItem 'C:\Reports\*.csv' | Select-Object -ExpandProperty FullName |
        Set-VBNextcloudFile -BaseUrl 'https://cloud.example.com' `
            -Credential $cred -DestinationPath 'Vibhu/Reports' -Overwrite

    Uploads all CSV files from C:\Reports via the pipeline, replacing any existing
    remote files.

    .EXAMPLE
    $result = Set-VBNextcloudFile -FilePath 'C:\data.csv' -BaseUrl 'https://cloud.example.com' `
        -Credential $cred -DestinationPath 'Backups' -WhatIf

    Uses -WhatIf to preview the upload without actually transferring any data.

    .OUTPUTS
    PSCustomObject
    Returns an object with:
    - SourceFile     : Local file path
    - TargetPath     : Remote destination path
    - FileSize       : File size in bytes
    - Overwritten    : Whether an existing file was replaced
    - Status         : 'Success' or 'Failed'
    - Error          : Error message (only present on failure)
    - CollectionTime : Timestamp of the operation

    .NOTES
    Version : 1.1.1
    Author  : Vibhu Bhatnagar

    Requirements:
    - PowerShell 5.1 or higher
    - Network access to the Nextcloud instance
    - Valid Nextcloud credentials with write access to the destination folder
    #>

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [string]$FilePath,

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

    process {
        try {
            if (-not (Test-Path -Path $FilePath -PathType Leaf)) {
                throw "File not found: $FilePath"
            }

            $fileInfo   = Get-Item -Path $FilePath
            $webDavUrl  = '{0}/{1}' -f $BaseUrl.TrimEnd('/'), $WebDAVPath.Trim('/')
            $targetPath = '{0}/{1}' -f $DestinationPath.Trim('/').Replace('\', '/'), $fileInfo.Name
            $uploadUrl  = '{0}/{1}' -f $webDavUrl, $targetPath

            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

            Write-Verbose "Checking remote path: $uploadUrl"

            # Check if the file already exists on the remote
            $fileExists = $false
            try {
                $null = Invoke-WebRequest -Uri $uploadUrl -Method Head -Credential $Credential `
                    -ErrorAction Stop -UseBasicParsing
                $fileExists = $true
                Write-Debug "Remote file exists: $uploadUrl"
            }
            catch {
                $fileExists = $false
            }

            if ($fileExists -and -not $Overwrite) {
                throw "File already exists at '$targetPath'. Use -Overwrite to replace."
            }

            if ($PSCmdlet.ShouldProcess($uploadUrl, 'Upload file')) {
                Write-Verbose "Uploading '$($fileInfo.Name)' ($($fileInfo.Length) bytes) to '$uploadUrl'"
                $fileContent = [System.IO.File]::ReadAllBytes($FilePath)
                $null = Invoke-RestMethod -Method Put -Uri $uploadUrl -Body $fileContent -Credential $Credential
                Write-Verbose "Upload complete: $targetPath"
            }

            [PSCustomObject]@{
                ComputerName   = $env:COMPUTERNAME
                SourceFile     = $FilePath
                TargetPath     = $targetPath
                FileSize       = $fileInfo.Length
                Overwritten    = $fileExists
                Status         = 'Success'
                CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
            }
        }
        catch {
            Write-Error -Message $_.Exception.Message -ErrorAction Continue
            [PSCustomObject]@{
                ComputerName   = $env:COMPUTERNAME
                SourceFile     = $FilePath
                TargetPath     = $null
                FileSize       = $null
                Overwritten    = $false
                Error          = $_.Exception.Message
                Status         = 'Failed'
                CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
            }
        }
    }
}
