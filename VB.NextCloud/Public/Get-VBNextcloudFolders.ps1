# ============================================================
# FUNCTION : Get-VBNextcloudFolders
# MODULE   : VB.NextCloud
# VERSION  : 1.0.0
# CHANGED  : 15-04-2026 -- Initial release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Lists subfolders in a Nextcloud directory via WebDAV PROPFIND
# ENCODING : UTF-8 with BOM
# ============================================================

function Get-VBNextcloudFolders {
    <#
    .SYNOPSIS
    Lists subfolders in a Nextcloud directory via WebDAV PROPFIND.

    .DESCRIPTION
    Get-VBNextcloudFolders queries a Nextcloud server using the WebDAV PROPFIND method to
    retrieve metadata for all subfolders in a specified folder. Only folders are returned --
    files are excluded. The folder itself is excluded from results so only direct children
    are returned. TLS 1.2 is enforced.

    .PARAMETER BaseUrl
    Base URL of the Nextcloud instance, e.g. 'https://cloud.example.com'.

    .PARAMETER Credential
    PSCredential object containing the Nextcloud username and password.

    .PARAMETER FolderPath
    Remote folder path to list, e.g. 'Vibhu/Reports'. Defaults to the WebDAV root.

    .PARAMETER WebDAVPath
    WebDAV endpoint path appended to BaseUrl. Defaults to 'remote.php/webdav'.

    .EXAMPLE
    Get-VBNextcloudFolders -BaseUrl 'https://cloud.example.com' -Credential (Get-Credential)

    Lists all subfolders in the WebDAV root.

    .EXAMPLE
    Get-VBNextcloudFolders -BaseUrl 'https://cloud.example.com' `
        -Credential $cred -FolderPath 'Vibhu'

    Lists all subfolders inside the Vibhu folder.

    .EXAMPLE
    Get-VBNextcloudFolders -BaseUrl 'https://cloud.example.com' -Credential $cred |
        Select-Object Name, Path, LastModified

    Lists root subfolders showing name, path and last modified date.

    .OUTPUTS
    PSCustomObject
    Returns an object per folder with:
    - Name         : Display name of the folder
    - Path         : Full WebDAV href path
    - LastModified : Last modified date string from the server
    - Status       : 'Success' or 'Failed'
    - Error        : Error message (only present on failure)

    .NOTES
    Version : 1.0.0
    Module  : VB.NextCloud
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

            Write-Verbose "Listing Nextcloud folders: $listUrl"

            $propfindBody = @'
<?xml version="1.0"?>
<d:propfind xmlns:d="DAV:">
    <d:prop>
        <d:displayname/>
        <d:getlastmodified/>
        <d:resourcetype/>
    </d:prop>
</d:propfind>
'@

            # Invoke-WebRequest does not support PROPFIND on PS 5.1 -- use HttpWebRequest directly
            $httpRequest             = [System.Net.HttpWebRequest]::Create($listUrl)
            $httpRequest.Method      = 'PROPFIND'
            $httpRequest.ContentType = 'application/xml'
            $httpRequest.Headers.Add('Depth', '1')
            $httpRequest.Credentials = $Credential.GetNetworkCredential()

            $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($propfindBody)
            $httpRequest.ContentLength = $bodyBytes.Length
            $requestStream = $httpRequest.GetRequestStream()
            $requestStream.Write($bodyBytes, 0, $bodyBytes.Length)
            $requestStream.Close()

            $webResponse    = $httpRequest.GetResponse()
            $responseStream = $webResponse.GetResponseStream()
            $reader         = [System.IO.StreamReader]::new($responseStream)
            $responseBody   = $reader.ReadToEnd()
            $reader.Close()
            $webResponse.Close()

            $xml            = [xml]$responseBody
            $collectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')

            foreach ($item in $xml.multistatus.response) {
                # Skip the requested folder itself
                $itemHref = $item.href.TrimEnd('/')
                $selfHref = if ($targetPath) {
                    ('/{0}/{1}' -f $WebDAVPath.Trim('/'), $targetPath).TrimEnd('/')
                }
                else {
                    ('/{0}' -f $WebDAVPath.Trim('/')).TrimEnd('/')
                }

                if ($itemHref -eq $selfHref) { continue }

                # Only return folders
                $isFolder = $null -ne $item.propstat.prop.resourcetype.collection
                if (-not $isFolder) { continue }

                Write-Debug "Found folder: $($item.propstat.prop.displayname)"

                [PSCustomObject]@{
                    ComputerName   = $env:COMPUTERNAME
                    Name           = $item.propstat.prop.displayname
                    Path           = $item.href
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
                LastModified   = $null
                Error          = $_.Exception.Message
                Status         = 'Failed'
                CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
            }
        }
    }
}
