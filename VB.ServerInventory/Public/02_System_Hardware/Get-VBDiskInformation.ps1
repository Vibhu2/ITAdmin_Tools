# ============================================================
# FUNCTION : Get-VBDiskInformation
# VERSION  : 1.0.0
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu
# PURPOSE  : Retrieves logical disk information from local or remote computers
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Retrieves logical disk information from local or remote computers.

.DESCRIPTION
    The Get-VBDiskInformation function collects detailed disk information including size,
    used space, free space, and percentage free for all logical drives. It supports both
    local and remote computer queries with pipeline input and credential authentication.

.PARAMETER ComputerName
    Target computer(s). Defaults to local machine. Accepts pipeline input.
    Aliases: Name, Server, Host

.PARAMETER Credential
    Alternate credentials for remote execution.

.EXAMPLE
    Get-VBDiskInformation

    Retrieves disk information from the local computer.

.EXAMPLE
    Get-VBDiskInformation -ComputerName "SERVER01"

    Retrieves disk information from a remote server.

.EXAMPLE
    'SRV01', 'SRV02' | Get-VBDiskInformation -Credential (Get-Credential)

    Uses pipeline input to query multiple servers with alternate credentials.

.OUTPUTS
    [PSCustomObject]: ComputerName, DeviceID, SizeGB, UsedSpaceGB, FreeSpaceGB, PercentFree,
    Status, CollectionTime

.NOTES
    Version  : 1.0.0
    Author   : Vibhu
    Modified : 10-04-2026
    Category : System Hardware
#>

function Get-VBDiskInformation {
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
                # Step 1 -- Local vs remote determination
                if ($computer -eq $env:COMPUTERNAME) {
                    # Local disk collection
                    $Disks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3"
                } else {
                    # Step 2 -- Remote disk collection
                    $CimParams = @{
                        ComputerName = $computer
                        ClassName    = 'Win32_LogicalDisk'
                        Filter       = 'DriveType=3'
                        ErrorAction  = 'Stop'
                    }
                    if ($Credential) { $CimParams.Credential = $Credential }
                    $Disks = Get-CimInstance @CimParams
                }

                # Step 3 -- Process each disk and create output objects
                foreach ($disk in $Disks) {
                    [PSCustomObject]@{
                        ComputerName   = $computer
                        DeviceID       = $disk.DeviceID
                        SizeGB         = [math]::Round($disk.Size / 1GB, 2)
                        UsedSpaceGB    = [math]::Round(($disk.Size - $disk.FreeSpace) / 1GB, 2)
                        FreeSpaceGB    = [math]::Round($disk.FreeSpace / 1GB, 2)
                        PercentFree    = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 2)
                        Status         = 'Success'
                        CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    }
                }
            }
            catch {
                # Step 4 -- Error handling
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
