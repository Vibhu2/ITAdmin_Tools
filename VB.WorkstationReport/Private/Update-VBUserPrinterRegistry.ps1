# ============================================================
# FUNCTION : Update-VBUserPrinterRegistry
# MODULE   : VB.WorkstationReport
# VERSION  : 1.0.0
# CHANGED  : 23-04-2026 -- Initial release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Applies printer path migrations to a single mounted user hive
# ENCODING : UTF-8 with BOM
# NOTE     : Private -- caller must mount hive via Mount-VBUserHive before calling
# ============================================================

function Update-VBUserPrinterRegistry {
    <#
    .SYNOPSIS
        Applies printer path migrations to a single mounted user registry hive.

    .DESCRIPTION
        Update-VBUserPrinterRegistry is an internal helper called by Set-VBUserPrinterMigration.
        It accepts a SID whose hive is already mounted under HKEY_USERS and applies a set of
        printer mapping rules, handling all four migration scenarios:

          UNC  -> UNC   Rename the Connections subkey and update Devices/PrinterPorts entries.
          UNC  -> IP    Remove the Connections subkey; add Devices/PrinterPorts pointing to the
                        new TCP/IP port name (IP_x.x.x.x). Machine-level port must already exist.
          IP   -> UNC   Add a Connections subkey for the new UNC path; update Devices/PrinterPorts.
          IP   -> IP    Update Devices/PrinterPorts to point to the new TCP/IP port name.

        Registry keys written per user:
          HKU\{SID}\Printers\Connections\,,server,share       -- UNC connection marker
          HKU\{SID}\Software\...\Devices                      -- printer device list
          HKU\{SID}\Software\...\PrinterPorts                 -- printer port list (if present)
          HKU\{SID}\Software\...\Windows  (Device value)      -- default printer

        Idempotent: if the new printer already exists and the old printer is already gone
        the function returns Action = 'AlreadyMigrated' without making any changes.

        NOTE: This is a private function. Do not call it directly. Use Set-VBUserPrinterMigration.
        The caller is responsible for mounting and dismounting the hive via Mount-VBUserHive
        and Dismount-VBUserHive.

    .PARAMETER SID
        SID of the user whose hive is currently mounted under HKEY_USERS.

    .PARAMETER Username
        Display name used in output objects. Defaults to 'Unknown'.

    .PARAMETER ComputerName
        Computer name used in output objects. Defaults to local machine.

    .PARAMETER Mappings
        Array of PSCustomObjects, each with OldPath, NewPath, and DriverName properties.
        Produced by Set-VBUserPrinterMigration from CSV or hashtable input.

    .EXAMPLE
        # Called internally by Set-VBUserPrinterMigration after Mount-VBUserHive
        $mappings = @(
            [PSCustomObject]@{ OldPath = '\\OldServer\HP01'; NewPath = '10.30.1.50'; DriverName = 'HP LaserJet' }
        )
        Update-VBUserPrinterRegistry -SID $mountResult.SID -Username 'jdoe' -Mappings $mappings

    .OUTPUTS
        PSCustomObject
        Returns one object per mapping rule processed:
          - ComputerName : Target computer
          - Username     : User profile name
          - SID          : User SID
          - OldPath      : Old printer path
          - NewPath      : New printer path
          - Action       : 'Migrated', 'Skipped', 'AlreadyMigrated', or 'Failed'
          - Details      : Semicolon-separated list of registry actions taken
          - Status       : 'Success' or 'Failed'
          - Error        : Error message (only present on failure)
          - Timestamp    : Time of action (dd-MM-yyyy HH:mm:ss)

    .NOTES
        Version  : 1.0.0
        Author   : Vibhu Bhatnagar
        Category : Printer Management (Private)

        UNC path to Connections key name conversion:
          \\server\printer  ->  ,,server,printer
          Leading \\ becomes ,, then remaining \ become ,

        Devices key value format:
          UNC printer  : winspool,Ne00:             (port reassigned by Windows at logon)
          IP printer   : winspool,IP_x.x.x.x        (TCP/IP port name)

        PrinterPorts key value format:
          UNC printer  : winspool,Ne00:,15,45        (adds read/transmit timeouts)
          IP printer   : winspool,IP_x.x.x.x,15,45

        A user logoff/logon may be required for printer changes to fully apply
        in active sessions.
    #>

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$SID,

        [string]$Username     = 'Unknown',
        [string]$ComputerName = $env:COMPUTERNAME,

        [Parameter(Mandatory)]
        [object[]]$Mappings   # PSCustomObjects: OldPath, NewPath, DriverName
    )

    # --- Registry paths ---
    $regBase     = "Registry::HKEY_USERS\$SID"
    $connectPath = "$regBase\Printers\Connections"
    $devicesPath = "$regBase\Software\Microsoft\Windows NT\CurrentVersion\Devices"
    $portsPath   = "$regBase\Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts"
    $windowsPath = "$regBase\Software\Microsoft\Windows NT\CurrentVersion\Windows"

    foreach ($map in $Mappings) {

        # Step 1 -- Normalize paths
        $oldPath = ($map.OldPath -replace '/', '\').TrimEnd('\').Trim()
        $newPath = ($map.NewPath -replace '/', '\').TrimEnd('\').Trim()

        $isOldUNC = $oldPath -match '^\\\\'
        $isNewUNC = $newPath -match '^\\\\'

        # Step 2 -- Pre-compute key and port names
        $oldKeyName  = $null
        $newKeyName  = $null
        $newPortName = $null

        if ($isOldUNC) {
            # \\server\printer -> ,,server,printer
            $oldKeyName = ',,' + ($oldPath.TrimStart('\') -replace '\\', ',')
        }
        if ($isNewUNC) {
            $newKeyName = ',,' + ($newPath.TrimStart('\') -replace '\\', ',')
        }
        else {
            $newPortName = "IP_$newPath"
        }

        # Step 3 -- Find old printer in user registry
        $oldFound      = $false
        $oldDeviceName = $null

        if ($isOldUNC) {
            $oldFound      = Test-Path -Path "$connectPath\$oldKeyName"
            $oldDeviceName = $oldPath   # in Devices key, UNC printer name IS the UNC path
        }
        else {
            # IP -- find by matching port value in Devices key (value contains 'IP_x.x.x.x')
            $portToFind  = "IP_$($oldPath.Trim(','))"
            $deviceProps = Get-ItemProperty -Path $devicesPath -ErrorAction SilentlyContinue
            if ($deviceProps) {
                $members = $deviceProps |
                    Get-Member -MemberType NoteProperty |
                    Where-Object { $_.Name -notlike 'PS*' }
                foreach ($member in $members) {
                    if ($deviceProps.($member.Name) -like "*$portToFind*") {
                        $oldDeviceName = $member.Name
                        $oldFound      = $true
                        break
                    }
                }
            }
        }

        if (-not $oldFound) {
            [PSCustomObject]@{
                ComputerName = $ComputerName
                Username     = $Username
                SID          = $SID
                OldPath      = $map.OldPath
                NewPath      = $map.NewPath
                Action       = 'Skipped'
                Details      = 'Old printer not found in user profile'
                Status       = 'Success'
                Timestamp    = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
            }
            continue
        }

        # Step 4 -- Compute new printer display name (needed for Devices key and default printer)
        $newDisplayName = $null
        if ($isNewUNC) {
            $newDisplayName = $newPath
        }
        elseif ($isOldUNC) {
            # UNC -> IP: use share name as display name (e.g. \\server\HP01 -> HP01)
            $newDisplayName = $oldPath.TrimEnd('\').Split('\')[-1]
        }
        else {
            # IP -> IP: keep existing display name
            $newDisplayName = $oldDeviceName
        }

        # Step 5 -- Idempotency: check if new printer already mapped
        $newAlreadyExists = $false
        if ($isNewUNC) {
            $newAlreadyExists = Test-Path -Path "$connectPath\$newKeyName"
        }
        else {
            $deviceProps = Get-ItemProperty -Path $devicesPath -ErrorAction SilentlyContinue
            if ($deviceProps) {
                $members = $deviceProps |
                    Get-Member -MemberType NoteProperty |
                    Where-Object { $_.Name -notlike 'PS*' }
                foreach ($member in $members) {
                    if ($deviceProps.($member.Name) -like "*$newPortName*") {
                        $newAlreadyExists = $true
                        break
                    }
                }
            }
        }

        # If new already mapped AND old already gone -- fully migrated, nothing to do
        $oldStillPresent = $false
        if ($isOldUNC) {
            $oldStillPresent = Test-Path -Path "$connectPath\$oldKeyName"
        }
        else {
            $oldStillPresent = $null -ne $oldDeviceName
        }

        if ($newAlreadyExists -and (-not $oldStillPresent)) {
            [PSCustomObject]@{
                ComputerName = $ComputerName
                Username     = $Username
                SID          = $SID
                OldPath      = $map.OldPath
                NewPath      = $map.NewPath
                Action       = 'AlreadyMigrated'
                Details      = 'New printer already present, old printer already removed'
                Status       = 'Success'
                Timestamp    = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
            }
            continue
        }

        if (-not $PSCmdlet.ShouldProcess("$Username on $ComputerName", "Migrate $($map.OldPath) -> $($map.NewPath)")) {
            continue
        }

        try {
            $actionsTaken = [System.Collections.Generic.List[string]]::new()

            # Step 6 -- Add new printer entries (if not already present)
            if (-not $newAlreadyExists) {

                # Read old Devices value to carry forward when migrating UNC -> UNC
                $oldDeviceValue = $null
                if ($oldDeviceName -and (Test-Path -Path $devicesPath)) {
                    $oldDeviceValue = (Get-ItemProperty -Path $devicesPath -Name $oldDeviceName -ErrorAction SilentlyContinue).$oldDeviceName
                }

                # Build new Devices value
                $newDeviceValue = $null
                if ($isNewUNC) {
                    # Carry forward port value for UNC -> UNC; default for other types
                    if ($isOldUNC -and $oldDeviceValue) {
                        $newDeviceValue = $oldDeviceValue
                    }
                    else {
                        $newDeviceValue = 'winspool,Ne00:'
                    }
                }
                else {
                    $newDeviceValue = "winspool,$newPortName"
                }

                # Add UNC Connections subkey
                if ($isNewUNC) {
                    if (-not (Test-Path -Path $connectPath)) {
                        New-Item -Path $connectPath -Force | Out-Null
                    }
                    New-Item -Path $connectPath -Name $newKeyName -Force | Out-Null
                    $actionsTaken.Add("Added Connections key: $newKeyName")
                }

                # Add Devices entry
                if (Test-Path -Path $devicesPath) {
                    Set-ItemProperty -Path $devicesPath -Name $newDisplayName -Value $newDeviceValue -Type String -Force
                    $actionsTaken.Add("Added Devices entry: $newDisplayName = $newDeviceValue")
                }

                # Add PrinterPorts entry (optional key -- only write if it exists)
                if (Test-Path -Path $portsPath) {
                    # PrinterPorts value appends timeout values: winspool,Ne00:,15,45
                    $newPortsValue = $newDeviceValue
                    if ($newPortsValue -notlike '*,15,45') {
                        $newPortsValue = "$newPortsValue,15,45"
                    }
                    Set-ItemProperty -Path $portsPath -Name $newDisplayName -Value $newPortsValue -Type String -Force
                    $actionsTaken.Add("Added PrinterPorts entry: $newDisplayName")
                }
            }

            # Step 7 -- Remove old printer entries
            if ($isOldUNC) {
                $oldConnKeyPath = "$connectPath\$oldKeyName"
                if (Test-Path -Path $oldConnKeyPath) {
                    Remove-Item -Path $oldConnKeyPath -Recurse -Force
                    $actionsTaken.Add("Removed Connections key: $oldKeyName")
                }
                if (Test-Path -Path $devicesPath) {
                    Remove-ItemProperty -Path $devicesPath -Name $oldPath -ErrorAction SilentlyContinue
                    $actionsTaken.Add("Removed Devices entry: $oldPath")
                }
                if (Test-Path -Path $portsPath) {
                    Remove-ItemProperty -Path $portsPath -Name $oldPath -ErrorAction SilentlyContinue
                }
            }
            else {
                if ($oldDeviceName) {
                    if (Test-Path -Path $devicesPath) {
                        Remove-ItemProperty -Path $devicesPath -Name $oldDeviceName -ErrorAction SilentlyContinue
                        $actionsTaken.Add("Removed Devices entry: $oldDeviceName")
                    }
                    if (Test-Path -Path $portsPath) {
                        Remove-ItemProperty -Path $portsPath -Name $oldDeviceName -ErrorAction SilentlyContinue
                    }
                }
            }

            # Step 8 -- Update default printer if it was the one being migrated
            $defaultReg = Get-ItemProperty -Path $windowsPath -Name 'Device' -ErrorAction SilentlyContinue
            if ($defaultReg) {
                $defaultName = ($defaultReg.Device -split ',')[0]

                $isDefault = $false
                if ($isOldUNC) {
                    $isDefault = $defaultName -eq $oldPath
                }
                else {
                    $isDefault = $defaultName -eq $oldDeviceName
                }

                if ($isDefault) {
                    $newDefaultValue = $null
                    if ($isNewUNC) {
                        $newDefaultValue = "$newPath,winspool,Ne00:"
                    }
                    else {
                        $newDefaultValue = "$newDisplayName,winspool,$newPortName"
                    }
                    Set-ItemProperty -Path $windowsPath -Name 'Device' -Value $newDefaultValue -Type String -Force
                    $actionsTaken.Add("Updated default printer: $newDisplayName")
                }
            }

            [PSCustomObject]@{
                ComputerName = $ComputerName
                Username     = $Username
                SID          = $SID
                OldPath      = $map.OldPath
                NewPath      = $map.NewPath
                Action       = 'Migrated'
                Details      = ($actionsTaken -join '; ')
                Status       = 'Success'
                Timestamp    = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
            }
        }
        catch {
            [PSCustomObject]@{
                ComputerName = $ComputerName
                Username     = $Username
                SID          = $SID
                OldPath      = $map.OldPath
                NewPath      = $map.NewPath
                Action       = 'Failed'
                Details      = $_.Exception.Message
                Error        = $_.Exception.Message
                Status       = 'Failed'
                Timestamp    = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
            }
        }
    }
}
