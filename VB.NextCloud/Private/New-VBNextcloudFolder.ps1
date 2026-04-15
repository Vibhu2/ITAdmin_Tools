# ============================================================
# FUNCTION : New-VBNextcloudFolder
# MODULE   : VB.NextCloud
# VERSION  : 1.1.2
# CHANGED  : 15-04-2026 -- Updated caller reference to Start-VBNextcloudUpload
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Creates a folder on Nextcloud via WebDAV MKCOL (private helper)
# ENCODING : UTF-8 with BOM
# ============================================================

function New-VBNextcloudFolder {
    <#
    .SYNOPSIS
    Creates a folder on a Nextcloud server via WebDAV MKCOL.

    .DESCRIPTION
    New-VBNextcloudFolder sends a WebDAV MKCOL request to create a folder at the specified
    path on a Nextcloud server. If the folder already exists the function returns an
    'AlreadyExists' status without error, making it safe to call idempotently before
    uploading files. This is a private helper function used by Start-VBNextcloudUpload.

    .PARAMETER BaseUrl
    Base URL of the Nextcloud instance, e.g. 'https://cloud.example.com'.

    .PARAMETER Credential
    PSCredential object containing the Nextcloud username and password.

    .PARAMETER FolderPath
    Remote folder path to create, e.g. 'Vibhu/Reports'. Intermediate folders must
    already exist -- Nextcloud does not auto-create parent directories.

    .PARAMETER WebDAVPath
    WebDAV endpoint path appended to BaseUrl. Defaults to 'remote.php/webdav'.

    .EXAMPLE
    New-VBNextcloudFolder -BaseUrl 'https://cloud.example.com' `
        -Credential $cred -FolderPath 'Vibhu/Reports'

    Creates the Reports folder under Vibhu on the Nextcloud server.

    .EXAMPLE
    New-VBNextcloudFolder -BaseUrl 'https://cloud.example.com' `
        -Credential $cred -FolderPath 'Vibhu/Reports' -WhatIf

    Previews folder creation without making any changes.

    .OUTPUTS
    PSCustomObject
    Returns an object with:
    - FolderPath     : Remote path that was targeted
    - FolderUrl      : Full WebDAV URL used for the MKCOL request
    - Created        : Boolean -- $true if the folder was created, $false if it already existed
    - Status         : 'Success', 'AlreadyExists', or 'Failed'
    - Error          : Error message (only present on failure)
    - CollectionTime : Timestamp of the operation

    .NOTES
    Version : 1.1.1
    Author  : Vibhu Bhatnagar

    Private function -- not exported by the module. Called internally by Start-VBNextcloudUpload.
    #>

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$BaseUrl,

        [Parameter(Mandatory)]
        [PSCredential]$Credential,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$FolderPath,

        [ValidateNotNullOrEmpty()]
        [string]$WebDAVPath = 'remote.php/webdav'
    )

    process {
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

            $webDavUrl  = '{0}/{1}' -f $BaseUrl.TrimEnd('/'), $WebDAVPath.Trim('/')
            $targetPath = $FolderPath.Trim('/').Replace('\', '/')
            $folderUrl  = '{0}/{1}' -f $webDavUrl, $targetPath

            Write-Verbose "Checking for existing folder: $folderUrl"

            # Check if the folder already exists
            $folderExists = $false
            try {
                $null = Invoke-WebRequest -Uri $folderUrl -Method Head -Credential $Credential `
                    -ErrorAction Stop -UseBasicParsing
                $folderExists = $true
                Write-Debug "Folder already exists: $folderUrl"
            }
            catch {
                $folderExists = $false
            }

            if ($folderExists) {
                Write-Verbose "Folder exists, skipping creation: $targetPath"
                return [PSCustomObject]@{
                    ComputerName   = $env:COMPUTERNAME
                    FolderPath     = $targetPath
                    FolderUrl      = $folderUrl
                    Created        = $false
                    Status         = 'AlreadyExists'
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }

            if ($PSCmdlet.ShouldProcess($folderUrl, 'Create Nextcloud folder')) {
                Write-Verbose "Creating folder: $targetPath"
                $null = Invoke-WebRequest -Uri $folderUrl -Method 'MKCOL' -Credential $Credential `
                    -UseBasicParsing -ErrorAction Stop
                Write-Verbose "Folder created: $targetPath"
            }

            [PSCustomObject]@{
                ComputerName   = $env:COMPUTERNAME
                FolderPath     = $targetPath
                FolderUrl      = $folderUrl
                Created        = $true
                Status         = 'Success'
                CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
            }
        }
        catch {
            Write-Error -Message $_.Exception.Message -ErrorAction Continue
            [PSCustomObject]@{
                ComputerName   = $env:COMPUTERNAME
                FolderPath     = $FolderPath
                FolderUrl      = $null
                Created        = $false
                Error          = $_.Exception.Message
                Status         = 'Failed'
                CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
            }
        }
    }
}
