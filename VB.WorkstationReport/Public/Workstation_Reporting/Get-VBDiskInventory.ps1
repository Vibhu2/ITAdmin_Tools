# ============================================================
# FUNCTION : Get-VBDiskInventory
# VERSION  : 1.2.2
# CHANGED  : 01-05-2026 -- UniqueId now resolved via Win32_PhysicalMedia on NVMe
#                          to replace EUI placeholder (eui.FFFFFFFFFFFFFFFF)
#                          with real manufacturer identifier
# AUTHOR   : Vibhu
# PURPOSE  : Collects full disk inventory -- volume, physical, and raw disk data
# ENCODING : UTF-8 with BOM
# ------------------------------------------------------------
# CHANGELOG (last 3-5 only -- full history in Git)
# v1.2.2 -- 01-05-2026 -- Fixed NVMe UniqueId via Win32_PhysicalMedia fallback
# v1.2.1 -- 01-05-2026 -- Fixed NVMe SerialNumber via AdapterSerialNumber fallback
# v1.2.0 -- 01-05-2026 -- Added 8 new properties; renamed Status -> CollectionStatus
# v1.1.0 -- 16-04-2026 -- Added CollectionTime; merged from Get-VBDiskInformation
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
                    $allPhysicalMedia     = Get-CimInstance -ClassName Win32_PhysicalMedia
                    $systemDrive          = $env:SystemDrive.TrimEnd('\')
                    $collectionTime       = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')

                    $diskInfo             = @()
                    $processedDiskNumbers = @()

                    # Step 2 -- Process volume-based disks
                    foreach ($volume in $allVolumes) {
                        $partition = $allPartitions | Where-Object { $_.AccessPaths -contains "$($volume.DriveLetter):\" }
                        $physDisk  = $allPhysicalDisks | Where-Object { $_.DeviceId -eq $partition.DiskNumber }
                        $wmiDisk   = $allWmiDisks | Where-Object { $_.Index -eq $partition.DiskNumber }
                        $mediaDisk = $allPhysicalMedia | Where-Object { $_.Tag -eq $wmiDisk.Name }

                        if ($physDisk) {
                            $processedDiskNumbers += $physDisk.DeviceId
                            $storageType  = if ($physDisk.BusType -eq 'iSCSI') { 'iSCSI' } else { 'Local' }
                            $diskType     = if ($physDisk.BusType -eq 'RAID') { 'RAID' } else { 'Direct Disk' }
                            $partCount    = ($allPartitions | Where-Object { $_.DiskNumber -eq $physDisk.DeviceId }).Count
                            $isSystemDisk = "$($volume.DriveLetter):" -eq $systemDrive
                            # NVMe: AdapterSerialNumber holds real manufacturer serial; SerialNumber returns EUI/NGUID
                            $serialNumber = if ($physDisk.BusType -eq 'NVMe' -and $physDisk.AdapterSerialNumber) { $physDisk.AdapterSerialNumber.Trim() } else { $physDisk.SerialNumber }
                            # NVMe: Win32_PhysicalMedia.SerialNumber holds real unique ID; Get-PhysicalDisk.UniqueId returns EUI placeholder
                            $uniqueId     = if ($physDisk.BusType -eq 'NVMe' -and $mediaDisk.SerialNumber) { $mediaDisk.SerialNumber.Trim() } else { $physDisk.UniqueId }

                            $diskInfo += [PSCustomObject]@{
                                ComputerName        = $env:COMPUTERNAME
                                DriveLetter         = "$($volume.DriveLetter):"
                                Label               = $volume.FileSystemLabel
                                FileSystem          = $volume.FileSystem
                                TotalSizeGB         = [math]::Round($volume.Size / 1GB, 2)
                                UsedGB              = [math]::Round(($volume.Size - $volume.SizeRemaining) / 1GB, 2)
                                FreeGB              = [math]::Round($volume.SizeRemaining / 1GB, 2)
                                FreePercent         = [math]::Round(($volume.SizeRemaining / $volume.Size) * 100, 1)
                                AllocationUnitSize  = $volume.AllocationUnitSize
                                OperationalStatus   = $volume.OperationalStatus
                                IsSystemDisk        = $isSystemDisk
                                StorageType         = $storageType
                                DiskType            = $diskType
                                MediaType           = $physDisk.MediaType
                                BusType             = $physDisk.BusType
                                SpindleSpeed        = $physDisk.SpindleSpeed
                                Usage               = $physDisk.Usage
                                FriendlyName        = $physDisk.FriendlyName
                                UniqueId            = $uniqueId
                                HealthStatus        = $physDisk.HealthStatus
                                SerialNumber        = $serialNumber
                                FirmwareVersion     = $physDisk.FirmwareVersion
                                InterfaceType       = $wmiDisk.InterfaceType
                                WMICaption          = $wmiDisk.Caption
                                WmiStatus           = $wmiDisk.Status
                                PartitionCount      = $partCount
                                DiskNumber          = $physDisk.DeviceId
                                IsVolumeBased       = $true
                                CollectionStatus    = 'Success'
                                CollectionTime      = $collectionTime
                            }
                        }
                    }

                    # Step 3 -- Process raw (unpartitioned) disks
                    foreach ($physDisk in $allPhysicalDisks) {
                        if ($physDisk.DeviceId -notin $processedDiskNumbers) {
                            $wmiDisk      = $allWmiDisks | Where-Object { $_.Model -like "*$($physDisk.FriendlyName)*" }
                            $mediaDisk    = $allPhysicalMedia | Where-Object { $_.Tag -eq $wmiDisk.Name }
                            $storageType  = if ($physDisk.BusType -eq 'iSCSI') { 'iSCSI' } else { 'Local' }
                            $diskType     = if ($physDisk.BusType -eq 'RAID') { 'RAID' } else { 'Direct Disk' }
                            $totalSizeGB  = [math]::Round($physDisk.Size / 1GB, 2)
                            $partCount    = ($allPartitions | Where-Object { $_.DiskNumber -eq $physDisk.DeviceId }).Count
                            # NVMe: AdapterSerialNumber holds real manufacturer serial; SerialNumber returns EUI/NGUID
                            $serialNumber = if ($physDisk.BusType -eq 'NVMe' -and $physDisk.AdapterSerialNumber) { $physDisk.AdapterSerialNumber.Trim() } else { $physDisk.SerialNumber }
                            # NVMe: Win32_PhysicalMedia.SerialNumber holds real unique ID; Get-PhysicalDisk.UniqueId returns EUI placeholder
                            $uniqueId     = if ($physDisk.BusType -eq 'NVMe' -and $mediaDisk.SerialNumber) { $mediaDisk.SerialNumber.Trim() } else { $physDisk.UniqueId }

                            $diskInfo += [PSCustomObject]@{
                                ComputerName        = $env:COMPUTERNAME
                                DriveLetter         = 'N/A (Raw Disk)'
                                Label               = 'N/A'
                                FileSystem          = 'N/A'
                                TotalSizeGB         = $totalSizeGB
                                UsedGB              = 0
                                FreeGB              = $totalSizeGB
                                FreePercent         = 100
                                AllocationUnitSize  = 'N/A'
                                OperationalStatus   = 'N/A'
                                IsSystemDisk        = $false
                                StorageType         = $storageType
                                DiskType            = $diskType
                                MediaType           = $physDisk.MediaType
                                BusType             = $physDisk.BusType
                                SpindleSpeed        = $physDisk.SpindleSpeed
                                Usage               = $physDisk.Usage
                                FriendlyName        = $physDisk.FriendlyName
                                UniqueId            = $uniqueId
                                HealthStatus        = $physDisk.HealthStatus
                                SerialNumber        = $serialNumber
                                FirmwareVersion     = $physDisk.FirmwareVersion
                                InterfaceType       = $wmiDisk.InterfaceType
                                WMICaption          = $wmiDisk.Caption
                                WmiStatus           = $wmiDisk.Status
                                PartitionCount      = $partCount
                                DiskNumber          = $physDisk.DeviceId
                                IsVolumeBased       = $false
                                CollectionStatus    = 'Success'
                                CollectionTime      = $collectionTime
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
                    ComputerName     = $computer
                    DriveLetter      = $null
                    Error            = $_.Exception.Message
                    CollectionStatus = 'Failed'
                    CollectionTime   = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}