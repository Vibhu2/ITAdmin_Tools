# ============================================================
# FUNCTION : Get-VBShareInformation
# VERSION  : 1.2.0
# CHANGED  : 16-04-2026 -- Separated shares by type (System/File/Printer)
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Enumerate and classify file shares on local and remote systems
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Enumerate and classify shares on local and remote systems.

.DESCRIPTION
    Retrieves all shares and classifies them as System, File, or Printer shares.
    Includes SMB configuration, caching policy, encryption, and availability details.

.PARAMETER ComputerName
    Target computer(s). Defaults to local machine. Accepts pipeline input.
    Aliases: Name, Server, Host

.PARAMETER Credential
    Alternate credentials for remote execution.

.EXAMPLE
    Get-VBShareInformation

    Retrieves all shares from the local computer.

.EXAMPLE
    Get-VBShareInformation | Where-Object { $_.ShareType -eq 'File' }

    Get file shares only.

.OUTPUTS
    [PSCustomObject]: ComputerName, ShareName, ShareType, Path, Description, AvailabilityType,
    ShareState, ConcurrentUserLimit, EncryptData, CachingPolicy, FolderEnumerationMode,
    Status, CollectionTime

.NOTES
    Version  : 1.2.0
    Author   : Vibhu Bhatnagar
    Modified : 16-04-2026
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
                $scriptBlock = {
                    # Step 1 -- Collect all shares
                    $collectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    $allShares = Get-SmbShare

                    $shareInfo = @()

                    # Step 2 -- Classify each share
                    foreach ($share in $allShares) {
                        # Step 3 -- Determine share type
                        $shareType = if ($share.Name -match '^(ADMIN\$|C\$|D\$|E\$|F\$|G\$|H\$|I\$|IPC\$|NETLOGON|SYSVOL|print\$)$') {
                            'System'
                        } elseif ($share.ShareType -eq 'PrintQueue') {
                            'Printer'
                        } else {
                            'File'
                        }

                        # Step 4 -- Extract description safely
                        $description = if ([string]::IsNullOrWhiteSpace($share.Description)) { '(No description)' } else { $share.Description }

                        # Step 5 -- Build share info object
                        $shareInfo += [PSCustomObject]@{
                            ComputerName         = $env:COMPUTERNAME
                            ShareName            = $share.Name
                            ShareType            = $shareType
                            Path                 = $share.Path
                            Description          = $description
                            AvailabilityType     = $share.AvailabilityType
                            ShareState           = $share.ShareState
                            ConcurrentUserLimit  = $share.ConcurrentUserLimit
                            EncryptData          = $share.EncryptData
                            CachingPolicy        = $share.CachingPolicy
                            FolderEnumerationMode = $share.FolderEnumerationMode
                            Status               = 'Success'
                            CollectionTime       = $collectionTime
                        }
                    }

                    return $shareInfo
                }

                # Step 6 -- Local vs remote execution
                if ($computer -eq $env:COMPUTERNAME) {
                    $result = & $scriptBlock
                } else {
                    $params = @{
                        ComputerName = $computer
                        ScriptBlock  = $scriptBlock
                        ErrorAction  = 'Stop'
                    }
                    if ($Credential) { $params.Credential = $Credential }
                    $result = Invoke-Command @params
                }

                $result
            }
            catch {
                # Step 7 -- Error handling
                [PSCustomObject]@{
                    ComputerName   = $computer
                    ShareName      = $null
                    ShareType      = $null
                    Error          = $_.Exception.Message
                    Status         = 'Failed'
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}