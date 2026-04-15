# ============================================================
# FUNCTION : Get-VBShareInformation
# VERSION  : 1.0.2
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Enumerate file shares on local and remote systems
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Enumerate file shares on local and remote systems.

.DESCRIPTION
    Retrieves information about file system shares on Windows systems, excluding system and administrative shares.
    Supports local and remote queries with alternate credentials.

.PARAMETER ComputerName
    Target computer(s). Defaults to local machine. Accepts pipeline input.
    Supports aliases: Name, Server, Host.

.PARAMETER Credential
    Alternate credentials for remote execution.

.EXAMPLE
    Get-VBShareInformation

.EXAMPLE
    Get-VBShareInformation -ComputerName SERVER01

.EXAMPLE
    'SRV01','SRV02' | Get-VBShareInformation -Credential (Get-Credential)

.OUTPUTS
    [PSCustomObject]: ComputerName, ShareName, Path, Description, ShareState, FolderEnumerationMode, ConcurrentUserLimit, Status, CollectionTime

.NOTES
    Version  : 1.0.2
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : Shares
#>

function Get-VBShareInformation {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'Server', 'Host')]
        [string[]]$ComputerName = $env:COMPUTERNAME,

        [PSCredential]$Credential
    )

    process {
        foreach ($computer in $ComputerName) {
            try {
                # Step 1 -- Query shares from local or remote computer
                if ($computer -eq $env:COMPUTERNAME) {
                    $shares = Get-SmbShare | Where-Object {
                        $_.ShareType -eq 'FileSystemDirectory' -and
                        $_.Name -notmatch '^(ADMIN\$|C\$|D\$|E\$|F\$|G\$|H\$|I\$|IPC\$|NETLOGON|SYSVOL|print\$)'
                    }
                } else {
                    $shares = Invoke-Command -ComputerName $computer -Credential $Credential -ScriptBlock {
                        Get-SmbShare | Where-Object {
                            $_.ShareType -eq 'FileSystemDirectory' -and
                            $_.Name -notmatch '^(ADMIN\$|C\$|D\$|E\$|F\$|G\$|H\$|I\$|IPC\$|NETLOGON|SYSVOL|print\$)'
                        }
                    }
                }

                # Step 2 -- Emit PSCustomObject for each share
                if ($shares) {
                    foreach ($share in $shares) {
                        [PSCustomObject]@{
                            ComputerName             = $computer
                            ShareName                = $share.Name
                            Path                     = $share.Path
                            Description              = $share.Description
                            ShareState               = $share.ShareState
                            FolderEnumerationMode    = $share.FolderEnumerationMode
                            ConcurrentUserLimit      = $share.ConcurrentUserLimit
                            Status                   = 'Success'
                            CollectionTime           = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                        }
                    }
                } else {
                    [PSCustomObject]@{
                        ComputerName   = $computer
                        ShareName      = 'None'
                        Status         = 'No shares found'
                        CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    }
                }
            }
            catch {
                [PSCustomObject]@{
                    ComputerName   = $computer
                    Error          = $_.Exception.Message
                    Status         = 'Failed'
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
