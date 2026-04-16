# ============================================================
# FUNCTION : Get-VBDiskInventory
# VERSION  : 1.1.0
# CHANGED  : 16-04-2026 -- Added CollectionTime to all output objects
# AUTHOR   : Vibhu
# PURPOSE  : Collects full disk inventory -- volume, physical, and raw disk data
# ------------------------------------------------------------
# CHANGELOG (last 3-5 only -- full history in Git)
# v1.1.0 -- 16-04-2026 -- Added CollectionTime; merged from Get-VBDiskInformation
# v1.0.0 -- [original date] -- Initial release
# ============================================================

function Get-VBDiskInventory {
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
                    # Step 1 -- Collect all disk data sources
                    $allPhysicalDisks     = Get-PhysicalDisk
                    $allVolumes           = Get-Volume | Where-Object { $_.DriveLetter -and $_.DriveType -eq 'Fixed' }
                    $allPartitions        = Get-Partition
                    $allWmiDisks          = Get-CimInstance -ClassName Win32_DiskDrive
                    $collectionTime       = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')

                    $diskInfo             = @()
                    $processedDiskNumbers = @()

                    # Step 2 -- Process volume-based disks
                    foreach ($volume in $allVolumes) {
                        $partition = $allPartitions | Where-Object { $_.AccessPaths -contains "$($volume.DriveLetter):\" }
                        $physDisk  = $allPhysicalDisks | Where-Object { $_.DeviceId -eq $partition.DiskNumber }
                        $wmiDisk   = $allWmiDisks | Where-Object { $_.Index -eq $partition.DiskNumber }

                        if ($physDisk) {
                            $processedDiskNumbers += $physDisk.DeviceId
                            $storageType = if ($physDisk.BusType -eq 'iSCSI') { 'iSCSI' } else { 'Local' }
                            $diskType    = if ($physDisk.BusType -eq 'RAID') { 'RAID' } else { 'Direct Disk' }

                            $diskInfo += [PSCustomObject]@{
                                ComputerName    = $env:COMPUTERNAME
                                DriveLetter     = "$($volume.DriveLetter):"
                                Label           = $volume.FileSystemLabel
                                FileSystem      = $volume.FileSystem
                                TotalSizeGB     = [math]::Round($volume.Size / 1GB, 2)
                                UsedGB          = [math]::Round(($volume.Size - $volume.SizeRemaining) / 1GB, 2)
                                FreeGB          = [math]::Round($volume.SizeRemaining / 1GB, 2)
                                FreePercent     = [math]::Round(($volume.SizeRemaining / $volume.Size) * 100, 1)
                                StorageType     = $storageType
                                DiskType        = $diskType
                                MediaType       = $physDisk.MediaType
                                BusType         = $physDisk.BusType
                                FriendlyName    = $physDisk.FriendlyName
                                HealthStatus    = $physDisk.HealthStatus
                                SerialNumber    = $physDisk.SerialNumber
                                FirmwareVersion = $physDisk.FirmwareVersion
                                InterfaceType   = $wmiDisk.InterfaceType
                                WMICaption      = $wmiDisk.Caption
                                DiskNumber      = $physDisk.DeviceId
                                IsVolumeBased   = $true
                                Status          = 'Success'
                                CollectionTime  = $collectionTime
                            }
                        }
                    }

                    # Step 3 -- Process raw (unpartitioned) disks
                    foreach ($physDisk in $allPhysicalDisks) {
                        if ($physDisk.DeviceId -notin $processedDiskNumbers) {
                            $wmiDisk     = $allWmiDisks | Where-Object { $_.Model -like "*$($physDisk.FriendlyName)*" }
                            $storageType = if ($physDisk.BusType -eq 'iSCSI') { 'iSCSI' } else { 'Local' }
                            $diskType    = if ($physDisk.BusType -eq 'RAID') { 'RAID' } else { 'Direct Disk' }
                            $totalSizeGB = [math]::Round($physDisk.Size / 1GB, 2)

                            $diskInfo += [PSCustomObject]@{
                                ComputerName    = $env:COMPUTERNAME
                                DriveLetter     = 'N/A (Raw Disk)'
                                Label           = 'N/A'
                                FileSystem      = 'N/A'
                                TotalSizeGB     = $totalSizeGB
                                UsedGB          = 0
                                FreeGB          = $totalSizeGB
                                FreePercent     = 100
                                StorageType     = $storageType
                                DiskType        = $diskType
                                MediaType       = $physDisk.MediaType
                                BusType         = $physDisk.BusType
                                FriendlyName    = $physDisk.FriendlyName
                                HealthStatus    = $physDisk.HealthStatus
                                SerialNumber    = $physDisk.SerialNumber
                                FirmwareVersion = $physDisk.FirmwareVersion
                                InterfaceType   = $wmiDisk.InterfaceType
                                WMICaption      = $wmiDisk.Caption
                                DiskNumber      = $physDisk.DeviceId
                                IsVolumeBased   = $false
                                Status          = 'Success'
                                CollectionTime  = $collectionTime
                            }
                        }
                    }

                    return $diskInfo
                }

                # Step 4 -- Local vs remote execution
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
                # Step 5 -- Error handling
                [PSCustomObject]@{
                    ComputerName   = $computer
                    DriveLetter    = $null
                    Error          = $_.Exception.Message
                    Status         = 'Failed'
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}