# ============================================================
# FUNCTION : Get-VBBitLockerRecoveryKey
# VERSION  : 1.0.0
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Retrieve BitLocker recovery keys from target computer
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Retrieve BitLocker recovery keys from target computer(s).

.DESCRIPTION
    Queries BitLocker volumes on target computer to retrieve protection status
    and recovery keys. Supports single volume or all drives. Returns status
    and recovery key information per volume.

.PARAMETER ComputerName
    Target computer(s). Defaults to local machine. Accepts pipeline input.
    Aliases: Name, Server, Host

.PARAMETER Credential
    Alternate credentials for remote execution.

.PARAMETER MountPoint
    Specific drive mount point (e.g., C:, D:). If not specified with -AllDrives,
    queries all available volumes.

.PARAMETER AllDrives
    Query all available drives. Used when MountPoint is not specified.

.EXAMPLE
    Get-VBBitLockerRecoveryKey
    Queries all BitLocker volumes on local computer.

.EXAMPLE
    Get-VBBitLockerRecoveryKey -ComputerName SERVER01 -MountPoint 'C:'
    Queries BitLocker status for C: drive on SERVER01.

.EXAMPLE
    'SERVER01', 'SERVER02' | Get-VBBitLockerRecoveryKey
    Queries all BitLocker volumes on multiple computers via pipeline.

.OUTPUTS
    [PSCustomObject]: ComputerName, Drive, Status, RecoveryKey, CollectionTime

.NOTES
    Version  : 1.0.0
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : Security
#>

function Get-VBBitLockerRecoveryKey {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'Server', 'Host')]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [PSCredential]$Credential,
        [string]$MountPoint,
        [switch]$AllDrives
    )
    process {
        foreach ($computer in $ComputerName) {
            try {
                # Step 1 -- Define remote script block
                $scriptBlock = {
                    param($mountPoint, $allDrives)

                    # Step 1a -- Determine volumes to query
                    if ($allDrives -or [string]::IsNullOrEmpty($mountPoint)) {
                        $volumes = Get-Volume | Where-Object { $_.DriveLetter } | ForEach-Object { "$($_.DriveLetter):" }
                    } else {
                        $volumes = @($mountPoint)
                    }

                    # Step 1b -- Process each volume
                    foreach ($volume in $volumes) {
                        $bitLockerVol = Get-BitLockerVolume -MountPoint $volume -ErrorAction SilentlyContinue

                        if ($null -eq $bitLockerVol) {
                            [PSCustomObject]@{
                                Drive = $volume
                                Status = 'BitLocker not available'
                                RecoveryKey = 'N/A'
                            }
                        }
                        elseif ($bitLockerVol.ProtectionStatus -eq 'On') {
                            $recoveryKey = $bitLockerVol.KeyProtector | Where-Object { $_.RecoveryPassword } | Select-Object -ExpandProperty RecoveryPassword
                            [PSCustomObject]@{
                                Drive = $volume
                                Status = 'BitLocker Enabled'
                                RecoveryKey = if ($recoveryKey) { $recoveryKey } else { 'N/A' }
                            }
                        }
                        elseif ($bitLockerVol.ProtectionStatus -eq 'Off') {
                            [PSCustomObject]@{
                                Drive = $volume
                                Status = 'BitLocker Not Enabled'
                                RecoveryKey = 'N/A'
                            }
                        }
                        else {
                            [PSCustomObject]@{
                                Drive = $volume
                                Status = 'Error checking status'
                                RecoveryKey = 'N/A'
                            }
                        }
                    }
                }

                # Step 2 -- Execute locally or remotely
                if ($computer -eq $env:COMPUTERNAME) {
                    $results = & $scriptBlock -mountPoint $MountPoint -allDrives $AllDrives
                } else {
                    $splat = @{
                        ComputerName = $computer
                        ScriptBlock = $scriptBlock
                        ArgumentList = @($MountPoint, $AllDrives)
                    }
                    if ($Credential) {
                        $splat['Credential'] = $Credential
                    }
                    $results = Invoke-Command @splat
                }

                # Step 3 -- Output results with metadata
                foreach ($item in $results) {
                    [PSCustomObject]@{
                        ComputerName = $computer
                        Drive = $item.Drive
                        Status = $item.Status
                        RecoveryKey = $item.RecoveryKey
                        CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    }
                }
            }
            catch {
                [PSCustomObject]@{
                    ComputerName = $computer
                    Drive = 'N/A'
                    Status = 'Failed'
                    Error = $_.Exception.Message
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
