# ============================================================
# FUNCTION : Get-VBNextcloudFiles
# MODULE   : WorkstationReport
# VERSION  : 1.3.0
# CHANGED  : 14-04-2026 -- Standards compliance fixes
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Lists files and folders in a Nextcloud directory via WebDAV PROPFIND
# ENCODING : UTF-8 with BOM
# ============================================================

function Get-VBNextcloudFiles {
    <#
    .SYNOPSIS
    Lists files and folders in a Nextcloud directory via WebDAV PROPFIND.

    .DESCRIPTION
    Get-VBNextcloudFiles queries a Nextcloud server using the WebDAV PROPFIND method to
    retrieve metadata for all items in a specified folder. Returns structured objects for
    each item including name, path, type, size, and last modified date. TLS 1.2 is enforced.
    The folder itself is excluded from the results so only its direct children are returned.

    .PARAMETER BaseUrl
    Base URL of the Nextcloud instance, e.g. 'https://cloud.example.com'.

    .PARAMETER Credential
    PSCredential object containing the Nextcloud username and password.

    .PARAMETER FolderPath
    Remote folder path to list, e.g. 'Vibhu/Reports'. Defaults to the WebDAV root.

    .PARAMETER WebDAVPath
    WebDAV endpoint path appended to BaseUrl. Defaults to 'remote.php/webdav'.

    .EXAMPLE
    Get-VBNextcloudFiles -BaseUrl 'https://cloud.example.com' -Credential (Get-Credential)

    Lists all items in the WebDAV root folder.

    .EXAMPLE
    Get-VBNextcloudFiles -BaseUrl 'https://cloud.example.com' `
        -Credential $cred -FolderPath 'Vibhu/Reports'

    Lists all items in the Vibhu/Reports folder.

    .EXAMPLE
    Get-VBNextcloudFiles -BaseUrl 'https://cloud.example.com' -Credential $cred `
        -FolderPath 'Vibhu/Reports' | Where-Object { -not $_.IsFolder } |
        Select-Object Name, Size, LastModified

    Returns only files (not subfolders) from the specified path.

    .OUTPUTS
    PSCustomObject
    Returns an object per item with:
    - Name         : Display name of the item
    - Path         : Full WebDAV href path
    - IsFolder     : Boolean -- $true if the item is a folder
    - Size         : File size in bytes ($null for folders)
    - LastModified : Last modified date string from the server
    - Status       : 'Success' or 'Failed'
    - Error        : Error message (only present on failure)

    .NOTES
    Version : 1.3.0
    Author  : Vibhu Bhatnagar

    Requirements:
    - PowerShell 5.1 or higher
    - Network access to the Nextcloud instance
    - Valid Nextcloud credentials with read access to the target folder
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$BaseUrl,

        [Parameter(Mandatory)]
        [PSCredential]$Credential,

        [string]$FolderPath = '',

        [ValidateNotNullOrEmpty()]
        [string]$WebDAVPath = 'remote.php/webdav'
    )

    process {
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

            $webDavUrl  = '{0}/{1}' -f $BaseUrl.TrimEnd('/'), $WebDAVPath.Trim('/')
            $targetPath = $FolderPath.Trim('/').Replace('\', '/')
            $listUrl    = if ($targetPath) { '{0}/{1}' -f $webDavUrl, $targetPath } else { $webDavUrl }

            Write-Verbose "Listing Nextcloud folder: $listUrl"

            $propfindBody = @'
<?xml version="1.0"?>
<d:propfind xmlns:d="DAV:">
    <d:prop>
        <d:displayname/>
        <d:getcontentlength/>
        <d:getlastmodified/>
        <d:resourcetype/>
    </d:prop>
</d:propfind>
'@
            $headers = @{
                'Content-Type' = 'application/xml'
                'Depth'        = '1'
            }

            $response = Invoke-WebRequest -Uri $listUrl -Method 'PROPFIND' -Body $propfindBody `
                -Headers $headers -Credential $Credential -UseBasicParsing -ErrorAction Stop

            $xml            = [xml]$response.Content
            $collectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')

            foreach ($item in $xml.multistatus.response) {
                # Normalise the href for comparison -- strip trailing slash
                $itemHref = $item.href.TrimEnd('/')

                # Build the expected self-href for the requested folder to skip it
                $selfHref = if ($targetPath) {
                    ('/{0}/{1}' -f $WebDAVPath.Trim('/'), $targetPath).TrimEnd('/')
                }
                else {
                    ('/{0}' -f $WebDAVPath.Trim('/')).TrimEnd('/')
                }

                if ($itemHref -eq $selfHref) { continue }

                $isFolder = $null -ne $item.propstat.prop.resourcetype.collection

                Write-Debug "Found item: $($item.propstat.prop.displayname) (IsFolder=$isFolder)"

                [PSCustomObject]@{
                    ComputerName   = $env:COMPUTERNAME
                    Name           = $item.propstat.prop.displayname
                    Path           = $item.href
                    IsFolder       = $isFolder
                    Size           = if ($isFolder) { $null } else { [long]$item.propstat.prop.getcontentlength }
                    LastModified   = $item.propstat.prop.getlastmodified
                    CollectionTime = $collectionTime
                    Status         = 'Success'
                }
            }
        }
        catch {
            Write-Error -Message $_.Exception.Message -ErrorAction Continue
            [PSCustomObject]@{
                ComputerName   = $env:COMPUTERNAME
                Name           = $null
                Path           = $FolderPath
                IsFolder       = $null
                Size           = $null
                LastModified   = $null
                Error          = $_.Exception.Message
                Status         = 'Failed'
                CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
            }
        }
    }
}
