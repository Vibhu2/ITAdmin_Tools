# ============================================================
# FUNCTION : Add-VBUserPrinter
# MODULE   : VB.WorkstationReport
# VERSION  : 1.0.0
# CHANGED  : 23-04-2026 -- Initial release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Adds a new UNC or IP printer to all or targeted user profiles
# ENCODING : UTF-8 with BOM
# ------------------------------------------------------------
# CHANGELOG (last 3-5 only -- full history in Git)
# v1.0.0 -- 23-04-2026 -- Initial release
# ============================================================

function Add-VBUserPrinter {
    <#
    .SYNOPSIS
        Adds a new UNC or IP printer to all or targeted user profiles on a machine.

    .DESCRIPTION
        Add-VBUserPrinter injects a printer mapping into user registry hives without
        requiring an existing printer to replace. It is the complement to
        Set-VBUserPrinterMigration, which only operates on printers already mapped
        in a user profile.

        For each target user profile the function:
          1. Mounts the NTUSER.DAT hive if not already loaded (via Mount-VBUserHive)
          2. Checks whether the printer is already present (idempotent -- skips if found)
          3. Writes the appropriate registry entries under HKU\{SID}
          4. Optionally sets the printer as the user's default
          5. Safely dismounts the hive (via Dismount-VBUserHive)

        For IP destinations, a machine-level TCP/IP port and printer are created before
        the per-user loop runs. This step requires the script to execute locally on the
        target machine (e.g. deployed via RMM). A terminating error is thrown if an IP
        printer is requested against a remote ComputerName target.

        Registry keys written per user:
          HKU\{SID}\Printers\Connections\,,server,share   -- UNC connection marker
          HKU\{SID}\Software\...\Devices                  -- printer device list
          HKU\{SID}\Software\...\PrinterPorts             -- printer port list (if present)
          HKU\{SID}\Software\...\Windows  (Device value)  -- default printer (if -SetAsDefault)

        A user logoff/logon may be required for changes to take full effect in active sessions.

    .PARAMETER PrinterPath
        UNC path (\\server\printer) or IP address (e.g. 10.30.1.50) of the printer to add.

    .PARAMETER PrinterName
        Display name for the printer as it appears in the user's Devices list.
        Required for IP printers. For UNC printers defaults to the full UNC path.
        Must match the machine-level printer name when adding IP printers.

    .PARAMETER DriverName
        Printer driver name. Required when PrinterPath is an IP address.
        The driver must already be installed on the machine.
        Example: 'HP LaserJet 400 M401', 'Canon Generic Plus PCL6'
        Run: Get-PrinterDriver | Select-Object Name  to list installed drivers.

    .PARAMETER TargetUser
        Username or SID of a specific user to target. When omitted all non-system
        user profiles on the machine are processed.

    .PARAMETER SetAsDefault
        When specified, sets the added printer as the user's default printer.

    .PARAMETER ComputerName
        Target computer(s). Accepts pipeline input. Defaults to local machine.
        Remote targets are supported for user registry changes only.
        IP printer port/printer creation requires local execution.

    .PARAMETER Credential
        Credentials for remote execution. Not required for local or domain-joined targets.

    .EXAMPLE
        # -------------------------------------------------------------------
        # Adding an IP printer to all users on the local machine
        # -------------------------------------------------------------------
        # Step 1 -- Confirm the driver name on the target machine
        Get-PrinterDriver | Select-Object Name
        # Example output: HP LaserJet 400 M401

        # Step 2 -- Add the printer to all user profiles
        Add-VBUserPrinter -PrinterPath '10.30.1.55' `
                          -PrinterName 'HP_Accounts' `
                          -DriverName  'HP LaserJet 400 M401'

    .EXAMPLE
        # Adding a UNC printer to all user profiles (no DriverName needed)
        Add-VBUserPrinter -PrinterPath '\\PrintServer02\Canon_Reception'

    .EXAMPLE
        # Add an IP printer to a single user and set it as their default
        Add-VBUserPrinter -PrinterPath '10.30.1.55' `
                          -PrinterName 'HP_Accounts' `
                          -DriverName  'HP LaserJet 400 M401' `
                          -TargetUser  'jdoe' `
                          -SetAsDefault

    .EXAMPLE
        # Dry run -- see what would be added without making any changes
        Add-VBUserPrinter -PrinterPath '10.30.1.55' `
                          -PrinterName 'HP_Accounts' `
                          -DriverName  'HP LaserJet 400 M401' `
                          -WhatIf

    .EXAMPLE
        # Capture results and export to CSV for RMM reporting
        $results = Add-VBUserPrinter -PrinterPath '\\PrintServer02\Canon_Reception'
        $results | Export-Csv -Path "\\FileServer\Logs\AddPrinter_$env:COMPUTERNAME.csv" `
            -NoTypeInformation -Encoding UTF8

        # Review what was skipped (already had the printer)
        $results | Where-Object { $_.Action -eq 'AlreadyExists' }

    .EXAMPLE
        # Add a UNC printer to a specific user by SID
        Add-VBUserPrinter -PrinterPath '\\PrintServer02\Finance_HP' `
                          -TargetUser  'S-1-5-21-3456789012-1234567890-123456789-1001'

    .OUTPUTS
        PSCustomObject
        Returns one object per user profile processed:
          - ComputerName : Target computer
          - Username     : User profile name
          - SID          : User SID
          - PrinterPath  : UNC path or IP address supplied
          - PrinterName  : Display name written to the user's Devices key
          - Action       : 'Added', 'AlreadyExists', or 'Failed'
          - SetAsDefault : True if printer was set as default for this user
          - Status       : 'Success' or 'Failed'
          - Error        : Error message (only present on failure)
          - Timestamp    : Time of action (dd-MM-yyyy HH:mm:ss)

    .NOTES
        Version  : 1.0.0
        Author   : Vibhu Bhatnagar
        Category : Printer Management

        Requirements:
        - PowerShell 5.1
        - Administrative privileges
        - PrintManagement module (built-in on Windows 8 / Server 2012+)
        - Printer driver already installed for IP destinations
        - Script must run locally on target machine for IP printer port additions

        Related functions:
        - Set-VBUserPrinterMigration  -- replaces existing printer mappings
        - Get-VBUserPrinterMappings   -- audits current printer mappings per user
    #>

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$PrinterPath,

        [string]$PrinterName,

        [string]$DriverName,

        [string]$TargetUser,

        [switch]$SetAsDefault,

        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Name', 'Server')]
        [string[]]$ComputerName = $env:COMPUTERNAME,

        [PSCredential]$Credential
    )

    begin {
        $ErrorActionPreference = 'Stop'

        # --- Normalize and validate PrinterPath ---
        $PrinterPath = ($PrinterPath -replace '/', '\').TrimEnd('\').Trim()
        $isUNC       = $PrinterPath -match '^\\\\'
        $isIP        = $PrinterPath -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}'

        if (-not $isUNC -and -not $isIP) {
            throw "PrinterPath must be a UNC path (\\server\printer) or an IP address. Received: '$PrinterPath'"
        }

        if ($isIP -and -not $DriverName) {
            throw "DriverName is required when PrinterPath is an IP address. Run: Get-PrinterDriver | Select-Object Name"
        }

        if ($isIP -and -not $PrinterName) {
            throw "PrinterName is required when PrinterPath is an IP address. This becomes the display name in the user's printer list."
        }

        # --- Resolve display name and port name ---
        $resolvedPrinterName = if ($PrinterName) { $PrinterName } else { $PrinterPath }
        $portName            = if ($isIP) { "IP_$PrinterPath" } else { $null }

        # --- Pre-compute registry values ---
        $connectionKeyName = $null
        $deviceValue       = $null

        if ($isUNC) {
            # \\server\printer -> ,,server,printer
            $connectionKeyName = ',,' + ($PrinterPath.TrimStart('\') -replace '\\', ',')
            $deviceValue       = 'winspool,Ne00:'
        }
        else {
            $deviceValue = "winspool,$portName"
        }
    }

    process {
        foreach ($computer in $ComputerName) {
            try {
                # --- Step 1: Machine-level port and printer for IP destinations (local only) ---
                if ($isIP) {
                    if ($computer -ne $env:COMPUTERNAME) {
                        throw "Machine-level printer port additions are not supported for remote targets. Run this script directly on '$computer' via RMM."
                    }

                    if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
                        if ($PSCmdlet.ShouldProcess($portName, 'Add TCP/IP printer port')) {
                            Add-PrinterPort -Name $portName -PrinterHostAddress $PrinterPath -ErrorAction Stop
                            Write-Verbose "Added printer port: $portName"
                        }
                    }

                    if (-not (Get-Printer -Name $resolvedPrinterName -ErrorAction SilentlyContinue)) {
                        if ($PSCmdlet.ShouldProcess($resolvedPrinterName, 'Add printer')) {
                            Add-Printer -Name $resolvedPrinterName -PortName $portName -DriverName $DriverName -ErrorAction Stop
                            Write-Verbose "Added printer: $resolvedPrinterName on port $portName"
                        }
                    }
                }

                # --- Step 2: Resolve target user profiles ---
                $profileParams = @{ ErrorAction = 'Stop' }

                if ($computer -ne $env:COMPUTERNAME) {
                    $profileParams['ComputerName'] = $computer
                    if ($Credential) { $profileParams['Credential'] = $Credential }
                }

                $profiles = Get-VBUserProfile @profileParams

                if ($TargetUser) {
                    $profiles = @($profiles | Where-Object {
                        $_.Username -eq $TargetUser -or $_.SID -eq $TargetUser
                    })

                    if ($profiles.Count -eq 0) {
                        Write-Warning "No profile found for TargetUser '$TargetUser' on $computer"
                        continue
                    }
                }

                # --- Step 3: Per-user registry injection ---
                foreach ($profile in $profiles) {

                    $mountParams = @{ SID = $profile.SID; ErrorAction = 'Stop' }

                    if ($computer -ne $env:COMPUTERNAME) {
                        $mountParams['ComputerName'] = $computer
                        if ($Credential) { $mountParams['Credential'] = $Credential }
                    }

                    $mountResult = Mount-VBUserHive @mountParams

                    if ($mountResult.Status -ne 'Success') {
                        [PSCustomObject]@{
                            ComputerName = $computer
                            Username     = $profile.Username
                            SID          = $profile.SID
                            PrinterPath  = $PrinterPath
                            PrinterName  = $resolvedPrinterName
                            Action       = 'Failed'
                            SetAsDefault = $false
                            Status       = 'Failed'
                            Error        = "Hive mount failed: $($mountResult.Error)"
                            Timestamp    = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                        }
                        continue
                    }

                    try {
                        $regBase     = "Registry::HKEY_USERS\$($mountResult.SID)"
                        $connectPath = "$regBase\Printers\Connections"
                        $devicesPath = "$regBase\Software\Microsoft\Windows NT\CurrentVersion\Devices"
                        $portsPath   = "$regBase\Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts"
                        $windowsPath = "$regBase\Software\Microsoft\Windows NT\CurrentVersion\Windows"

                        # --- Idempotency check ---
                        $alreadyExists = $false

                        if ($isUNC) {
                            $alreadyExists = Test-Path -Path "$connectPath\$connectionKeyName"
                        }
                        else {
                            $deviceProps = Get-ItemProperty -Path $devicesPath -ErrorAction SilentlyContinue
                            if ($deviceProps) {
                                $members = $deviceProps |
                                    Get-Member -MemberType NoteProperty |
                                    Where-Object { $_.Name -notlike 'PS*' }
                                foreach ($member in $members) {
                                    if ($deviceProps.($member.Name) -like "*$portName*") {
                                        $alreadyExists = $true
                                        break
                                    }
                                }
                            }
                        }

                        if ($alreadyExists) {
                            [PSCustomObject]@{
                                ComputerName = $computer
                                Username     = $profile.Username
                                SID          = $profile.SID
                                PrinterPath  = $PrinterPath
                                PrinterName  = $resolvedPrinterName
                                Action       = 'AlreadyExists'
                                SetAsDefault = $false
                                Status       = 'Success'
                                Timestamp    = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                            }
                            continue
                        }

                        if (-not $PSCmdlet.ShouldProcess("$($profile.Username) on $computer", "Add printer $resolvedPrinterName")) {
                            continue
                        }

                        # --- Add UNC Connections subkey ---
                        if ($isUNC) {
                            if (-not (Test-Path -Path $connectPath)) {
                                New-Item -Path $connectPath -Force | Out-Null
                            }
                            New-Item -Path $connectPath -Name $connectionKeyName -Force | Out-Null
                        }

                        # --- Add Devices entry ---
                        if (Test-Path -Path $devicesPath) {
                            Set-ItemProperty -Path $devicesPath -Name $resolvedPrinterName -Value $deviceValue -Type String -Force
                        }

                        # --- Add PrinterPorts entry (only if key exists) ---
                        if (Test-Path -Path $portsPath) {
                            $portsValue = if ($deviceValue -notlike '*,15,45') { "$deviceValue,15,45" } else { $deviceValue }
                            Set-ItemProperty -Path $portsPath -Name $resolvedPrinterName -Value $portsValue -Type String -Force
                        }

                        # --- Set as default if requested ---
                        $wasSetAsDefault = $false
                        if ($SetAsDefault) {
                            $defaultValue = "$resolvedPrinterName,$deviceValue"
                            Set-ItemProperty -Path $windowsPath -Name 'Device' -Value $defaultValue -Type String -Force
                            $wasSetAsDefault = $true
                        }

                        [PSCustomObject]@{
                            ComputerName = $computer
                            Username     = $profile.Username
                            SID          = $profile.SID
                            PrinterPath  = $PrinterPath
                            PrinterName  = $resolvedPrinterName
                            Action       = 'Added'
                            SetAsDefault = $wasSetAsDefault
                            Status       = 'Success'
                            Timestamp    = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                        }
                    }
                    catch {
                        [PSCustomObject]@{
                            ComputerName = $computer
                            Username     = $profile.Username
                            SID          = $profile.SID
                            PrinterPath  = $PrinterPath
                            PrinterName  = $resolvedPrinterName
                            Action       = 'Failed'
                            SetAsDefault = $false
                            Status       = 'Failed'
                            Error        = $_.Exception.Message
                            Timestamp    = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                        }
                    }
                    finally {
                        $mountResult | Dismount-VBUserHive | Out-Null
                    }
                }
            }
            catch {
                [PSCustomObject]@{
                    ComputerName = $computer
                    Username     = 'N/A'
                    SID          = 'N/A'
                    PrinterPath  = $PrinterPath
                    PrinterName  = $resolvedPrinterName
                    Action       = 'Failed'
                    SetAsDefault = $false
                    Status       = 'Failed'
                    Error        = $_.Exception.Message
                    Timestamp    = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
