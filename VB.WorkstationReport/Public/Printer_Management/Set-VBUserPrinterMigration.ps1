# ============================================================
# FUNCTION : Set-VBUserPrinterMigration
# MODULE   : VB.WorkstationReport
# VERSION  : 2.0.1
# CHANGED  : 16-06-2026 -- Capture Install-VBPrinterDriver result; Write-Warning when reboot required
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Migrates user printer mappings between UNC paths and IP addresses
# ENCODING : UTF-8 with BOM
# ------------------------------------------------------------
# CHANGELOG (last 3-5 only -- full history in Git)
# v2.0.1 -- 16-06-2026 -- Capture driver install result; emit Write-Warning when RebootRequired is set
# v2.0.0 -- 16-06-2026 -- Rich CSV: ComputerName/Username/DriverPath/DefaultPrinter columns; driver install via Install-VBPrinterDriver helper
# v1.1.3 -- 04-05-2026 -- Existing printer port updated via Set-Printer when wrong; fixes UNC path still showing in printer properties
# v1.1.2 -- 04-05-2026 -- Username filter matches short name, DOMAIN\user, and profile folder name (fixes Entra ID / dot-format accounts)
# v1.1.1 -- 04-05-2026 -- Add-Printer "already exists" treated as skip; machine-level errors no longer abort per-user loop
# v1.1.0 -- 04-05-2026 -- -TargetUser replaced with -Username[] and -SID[] arrays; warn per missing user
# ============================================================
function Set-VBUserPrinterMigration {
    <#
    .SYNOPSIS
        Migrates user printer mappings from UNC paths to IP addresses or vice versa
        across all or targeted user profiles on a machine.

    .DESCRIPTION
        Set-VBUserPrinterMigration reads printer mapping rules from a CSV file or
        hashtable and applies them to each matching user profile on the target computer.

        For each user it mounts the registry hive (if not already loaded), applies the
        mapping changes via Update-VBUserPrinterRegistry, then safely dismounts the hive.
        Machine-level TCP/IP printer ports and printers are added once per machine before
        the per-user loop runs.

        Supports four migration scenarios:
          UNC  -> UNC   Server migration (\\old\printer -> \\new\printer)
          UNC  -> IP    Replace shared printer with direct IP printer
          IP   -> UNC   Replace direct IP printer with shared printer
          IP   -> IP    Port change on direct IP printer

        REMOTE EXECUTION NOTE:
        Machine-level printer port and printer additions (required for IP destinations)
        are NOT supported when ComputerName targets a remote machine. The script must
        run locally on each target workstation (e.g. deployed via RMM). A terminating
        error is raised for that machine (reported as a Failed result object) if a
        remote target requires machine-level port additions.

    .PARAMETER ComputerName
        Target computer(s). Accepts pipeline input. Defaults to local machine.
        Remote targets are supported for user registry changes only.
        'localhost', '127.0.0.1', '.' and the local FQDN are all recognised as local.

    .PARAMETER Credential
        Credentials for remote execution. Not required for local or domain-joined targets.

    .PARAMETER PrinterMappings
        Hashtable of OldPath = NewPath pairs.
        Example: @{ '\\OldServer\HP01' = '10.30.1.50'; '10.30.1.60' = '\\NewServer\Canon02' }
        For IP destinations, supply -DriverName. Use -MappingCsv for per-printer driver control.

    .PARAMETER MappingCsv
        Path to a CSV file. Two formats are supported:

        Rich format (recommended) -- includes ComputerName and Username columns.
        The script filters rows to $env:COMPUTERNAME automatically and targets only
        the users listed. Run locally on each machine (e.g. via RMM).
        Columns: ComputerName, Username, OldPath, NewPath, DriverName, DriverPath, DefaultPrinter
          - DriverName     : Required when NewPath is an IP address.
          - DriverPath     : Optional. Path to driver source (.inf, .cab, or folder) used
                             to auto-install the driver if not already present.
          - DefaultPrinter : Optional. Display name of the printer to set as the user's
                             default after migration. Leave blank to preserve current default.
                             For UNC destinations the display name is the full UNC path.
                             For IP destinations migrated from UNC it is the share name
                             (e.g. \Server\HP_Floor2 -> HP_Floor2).

        Legacy format -- no ComputerName column. Behaviour unchanged from v1.x.
        Columns: OldPath, NewPath, DriverName [, DriverPath]
        Use -Username / -SID to target specific users. CSV takes priority over
        -PrinterMappings if both are supplied.

    .PARAMETER DriverName
        Driver name to use when adding IP printers via -PrinterMappings hashtable.
        Applies to all IP destinations in the hashtable. For per-printer driver control
        use -MappingCsv with a DriverName column instead.

    .PARAMETER PrinterNames
        Optional hashtable mapping OldPath to a friendly display name for IP destinations.
        Use with -PrinterMappings when you want explicit printer names without a CSV file.
        Example: @{ '\\\\PrintServer\\HP01' = 'hpprinter410' }
        For per-printer name control via CSV, add a PrinterName column to -MappingCsv instead.

    .PARAMETER Username
        One or more usernames to target. Only profiles matching these usernames are
        processed. A warning is emitted for any username not found on the machine.
        When omitted (and -SID is also omitted) all non-system profiles are processed.

    .PARAMETER SID
        One or more SIDs to target. Only profiles matching these SIDs are processed.
        A warning is emitted for any SID not found on the machine.
        Can be combined with -Username (OR logic -- either match is included).

    .PARAMETER BackupMappings
        When specified, saves a snapshot of each user's current printer mappings to
        -BackupPath before applying changes. Requires -BackupPath.
        If the backup fails for a user, migration is SKIPPED for that user and a
        Failed result object is emitted -- changes are never applied without a backup.

    .PARAMETER BackupPath
        Full path to the CSV file where backup snapshots are written.
        Required when -BackupMappings is specified. The parent directory must exist.

    .EXAMPLE
        # -------------------------------------------------------------------
        # STEP 1 -- Create the mapping CSV (save as C:\Temp\PrinterMappings.csv)
        # -------------------------------------------------------------------
        # Required columns : OldPath, NewPath
        # Optional column  : DriverName  (required when NewPath is an IP address)
        # DriverName can be left blank for UNC -> UNC rows
        #
        # Sample CSV covering all four migration types:
        #
        #   OldPath,NewPath,DriverName
        #   \\PrintServer01\HP_Floor2,10.30.1.50,HP LaserJet 400 M401
        #   \\PrintServer01\Canon_HR,10.30.1.51,Canon Generic Plus PCL6
        #   10.30.1.60,\\PrintServer02\Ricoh_Reception,
        #   \\PrintServer01\Zebra_Labels,\\PrintServer02\Zebra_Labels,
        #
        # Row breakdown:
        #   Row 1 -- UNC -> IP  (DriverName required)
        #   Row 2 -- UNC -> IP  (DriverName required)
        #   Row 3 -- IP  -> UNC (DriverName blank -- not needed)
        #   Row 4 -- UNC -> UNC (DriverName blank -- not needed)
        #
        # To create it from PowerShell:
        $csv = @"
                OldPath,NewPath,DriverName
                \\PrintServer01\HP_Floor2,10.30.1.50,HP LaserJet 400 M401
                \\PrintServer01\Canon_HR,10.30.1.51,Canon Generic Plus PCL6
                10.30.1.60,\\PrintServer02\Ricoh_Reception,
                \\PrintServer01\Zebra_Labels,\\PrintServer02\Zebra_Labels,
                "@
        $csv | Out-File -FilePath 'C:\Temp\PrinterMappings.csv' -Encoding UTF8

        # -------------------------------------------------------------------
        # STEP 2 -- Run the migration (deploy via RMM, runs locally on machine)
        # -------------------------------------------------------------------
        Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv'

    .EXAMPLE
        Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv' -WhatIf

        Dry run -- shows what would be changed without making any modifications.
        No registry keys, ports, or printers are created or removed.
        Always run -WhatIf first when deploying to a new environment.

    .EXAMPLE
        Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv' `
            -BackupMappings -BackupPath 'C:\Realtime\Reports\PrinterBackup.csv'

        Saves a before-snapshot of each user's current printer mappings to CSV first,
        then applies the migration. Append-safe -- multiple machines can write to the
        same backup CSV when deployed via RMM.

    .EXAMPLE
        Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv' -Username 'jdoe','asmith'

        Migrates printers for jdoe and asmith only. All other profiles on the machine
        are skipped. A warning is emitted for any username not found on this machine.

    .EXAMPLE
        Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv' -SID 'S-1-5-21-123456789-1001'

        Targets a specific user by SID. Useful when deploying to a fleet via RMM --
        the script self-selects on machines where that SID exists.

    .EXAMPLE
        # UNC -> UNC server migration via hashtable (no DriverName needed)
        $mappings = @{
            '\\OldPrintServer\HP01'    = '\\NewPrintServer\HP01'
            '\\OldPrintServer\Canon02' = '\\NewPrintServer\Canon02'
        }
        Set-VBUserPrinterMigration -PrinterMappings $mappings

    .EXAMPLE
        # UNC -> IP migration via hashtable (all printers share the same driver)
        $mappings = @{
            '\\PrintServer\HP01' = '10.30.1.50'
            '\\PrintServer\HP02' = '10.30.1.51'
        }
        Set-VBUserPrinterMigration -PrinterMappings $mappings -DriverName 'HP LaserJet 400 M401'

    .EXAMPLE
        # IP -> UNC migration via hashtable
        $mappings = @{
            '10.30.1.50' = '\\NewPrintServer\HP01'
            '10.30.1.51' = '\\NewPrintServer\HP02'
        }
        Set-VBUserPrinterMigration -PrinterMappings $mappings

    .EXAMPLE
        # Capture full migration results and export to CSV for reporting
        $results = Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv'
        $results | Export-Csv -Path 'C:\Realtime\Reports\MigrationResults.csv' -NoTypeInformation -Encoding UTF8

        # Review failures only
        $results | Where-Object { $_.Status -eq 'Failed' } | Format-Table

    .EXAMPLE
        # RMM deployment pattern -- run locally, log results to network share
        $results = Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv' `
            -BackupMappings -BackupPath "\\FileServer\Logs\PrinterBackup_$env:COMPUTERNAME.csv"
        $results | Export-Csv -Path "\\FileServer\Logs\PrinterMigration_$env:COMPUTERNAME.csv" `
            -NoTypeInformation -Encoding UTF8

    .OUTPUTS
        PSCustomObject
        Returns one object per user per printer mapping action. All objects share the
        same property set in the same order (safe for Export-Csv):
          - ComputerName : Target computer
          - Username     : User profile name
          - SID          : User SID
          - OldPath      : Old printer path from mapping rule
          - NewPath      : New printer path from mapping rule
          - Action       : 'Migrated', 'Skipped', 'AlreadyMigrated', or 'Failed'
          - Details      : Registry actions taken or reason for skip/failure
          - Status       : 'Success' or 'Failed'
          - Error        : Error message (empty on success)
          - Timestamp    : Time of action (dd-MM-yyyy HH:mm:ss)

    .NOTES
        Version  : 1.1.5
        Author   : Vibhu Bhatnagar
        Category : Printer Management

        Requirements:
        - PowerShell 5.1
        - Administrative privileges
        - PrintManagement module (built-in on Windows 8 / Server 2012+)
        - Required printer drivers already installed for IP destinations
          (verified before Add-Printer -- a clear error is raised if missing)
        - Script must run locally on target machine for IP printer port additions
    #>

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Name', 'Server')]
        [string[]]$ComputerName = $env:COMPUTERNAME,

        [PSCredential]$Credential,

        [hashtable]$PrinterMappings,

        [string]$MappingCsv,

        [string]$DriverName,

        [hashtable]$PrinterNames,

        [string[]]$Username,

        [string[]]$SID,

        [switch]$BackupMappings,

        [string]$BackupPath
    )

    begin {
        $ErrorActionPreference = 'Stop'

        # --- Private helper: strict IPv4 validation ---
        # Anchored regex rejects trailing junk (e.g. '10.1.1.1.bad');
        # IPAddress.TryParse rejects out-of-range octets (e.g. '999.1.1.1').
        function Test-VBIPv4Address {
            param([string]$Value)

            if ($Value -notmatch '^\d{1,3}(\.\d{1,3}){3}$') { return $false }

            $parsed = $null
            return [System.Net.IPAddress]::TryParse($Value, [ref]$parsed)
        }

        # --- Canonical output property set (single shape for all result objects) ---
        $outputProperties = @(
            'ComputerName', 'Username', 'SID', 'OldPath', 'NewPath',
            'Action', 'Details', 'Status', 'Error', 'Timestamp'
        )

        # --- Validate backup parameters ---
        if ($BackupMappings) {
            if (-not $BackupPath) {
                throw '-BackupPath is required when -BackupMappings is specified.'
            }

            $backupDir = Split-Path -Path $BackupPath -Parent
            if ($backupDir -and -not (Test-Path -Path $backupDir)) {
                throw "Backup directory does not exist: $backupDir"
            }
        }

        # --- Step 1: Load and validate mappings ---
        $normalizedMappings = [System.Collections.Generic.List[object]]::new()
        $perUserMappings    = $null  # populated when rich CSV (ComputerName column) detected

        if ($MappingCsv) {
            if (-not (Test-Path -Path $MappingCsv)) {
                throw "Mapping CSV not found: $MappingCsv"
            }

            $csvRows = Import-Csv -Path $MappingCsv -ErrorAction Stop

            # Detect rich CSV format (has ComputerName column)
            $isRichCsv = $csvRows.Count -gt 0 -and
                         ($csvRows[0].PSObject.Properties.Name -contains 'ComputerName')

            if ($isRichCsv) {
                # Filter to this machine only
                $csvRows = @($csvRows | Where-Object { $_.ComputerName.Trim() -eq $env:COMPUTERNAME })

                if ($csvRows.Count -eq 0) {
                    Write-Verbose "No rows matching ComputerName '$env:COMPUTERNAME' in CSV -- nothing to do."
                    return
                }

                # Build per-user mapping structure and normalizedMappings for machine-level setup
                $perUserMappings = @{}
                foreach ($row in $csvRows) {
                    if (-not $row.OldPath -or -not $row.NewPath) {
                        throw "CSV row is missing OldPath or NewPath: $($row | Out-String)"
                    }

                    $isNewIP = $row.NewPath.Trim() -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}'
                    if ($isNewIP -and -not $row.DriverName) {
                        throw "DriverName is required for IP destinations. Missing for: $($row.OldPath.Trim()) -> $($row.NewPath.Trim()) (Username: $($row.Username.Trim()))"
                    }

                    $username = $row.Username.Trim()
                    if (-not $perUserMappings.ContainsKey($username)) {
                        $perUserMappings[$username] = [System.Collections.Generic.List[object]]::new()
                    }

                    $mappingObj = [PSCustomObject]@{
                        OldPath        = $row.OldPath.Trim()
                        NewPath        = $row.NewPath.Trim()
                        DriverName     = if ($row.DriverName) { $row.DriverName.Trim() } else { '' }
                        DriverPath     = if ($row.PSObject.Properties['DriverPath'] -and $row.DriverPath) { $row.DriverPath.Trim() } else { '' }
                        DefaultPrinter = if ($row.PSObject.Properties['DefaultPrinter'] -and $row.DefaultPrinter) { $row.DefaultPrinter.Trim() } else { '' }
                    }
                    $perUserMappings[$username].Add($mappingObj)

                    # Also add to normalizedMappings for machine-level port/printer setup
                    $normalizedMappings.Add([PSCustomObject]@{
                        OldPath    = $mappingObj.OldPath
                        NewPath    = $mappingObj.NewPath
                        DriverName = $mappingObj.DriverName
                        DriverPath = $mappingObj.DriverPath
                    })
                }
            }
            else {
                # Legacy CSV format: OldPath, NewPath, DriverName [, DriverPath]
                foreach ($row in $csvRows) {
                    if (-not $row.OldPath -or -not $row.NewPath) {
                        throw "CSV row is missing OldPath or NewPath: $($row | Out-String)"
                    }

                    $isNewIP = $row.NewPath.Trim() -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}'

                    if ($isNewIP -and -not $row.DriverName) {
                        throw "DriverName is required for IP destinations. Missing for: $($row.OldPath.Trim()) -> $($row.NewPath.Trim())"
                    }

                    $normalizedMappings.Add([PSCustomObject]@{
                        OldPath    = $row.OldPath.Trim()
                        NewPath    = $row.NewPath.Trim()
                        DriverName = if ($row.DriverName) { $row.DriverName.Trim() } else { '' }
                        DriverPath = if ($row.PSObject.Properties['DriverPath'] -and $row.DriverPath) { $row.DriverPath.Trim() } else { '' }
                    })
                }
            }
        }
        elseif ($PrinterMappings) {
            foreach ($key in $PrinterMappings.Keys) {
                # Guard against null/empty keys or values -- .Trim() on $null throws
                # an opaque NullReferenceException otherwise
                if ([string]::IsNullOrWhiteSpace([string]$key)) {
                    throw 'PrinterMappings contains a null or empty key. Every entry must be OldPath = NewPath.'
                }
                if ([string]::IsNullOrWhiteSpace([string]$PrinterMappings[$key])) {
                    throw "PrinterMappings value for '$key' is null or empty. Every OldPath key must have a NewPath value."
                }

                $oldPath = ([string]$key).Trim()
                $newPath = ([string]$PrinterMappings[$key]).Trim()
                $isNewIP = Test-VBIPv4Address -Value $newPath

                if ($isNewIP -and -not $DriverName) {
                    throw "Use -DriverName when supplying IP destinations via -PrinterMappings, or use -MappingCsv for per-printer driver control. Missing driver for: $oldPath -> $newPath"
                }

                $normalizedMappings.Add([PSCustomObject]@{
                        OldPath     = $oldPath
                        NewPath     = $newPath
                        DriverName  = if ($isNewIP) { $DriverName } else { '' }
                        PrinterName = if ($PrinterNames -and $PrinterNames.ContainsKey($oldPath)) { $PrinterNames[$oldPath] } else { '' }
                    })
            }
        }
        else {
            throw 'Either -PrinterMappings or -MappingCsv must be supplied.'
        }

        if ($normalizedMappings.Count -eq 0) {
            throw 'No valid printer mappings found in the supplied input.'
        }

        # Mappings with IP destinations require machine-level port/printer setup.
        # Invariant across computers -- computed once here.
        $ipMappings = @($normalizedMappings | Where-Object { Test-VBIPv4Address -Value $_.NewPath })
    }

    process {
        foreach ($computer in $ComputerName) {
            # Robust local-machine detection: name equality alone misses
            # 'localhost', '127.0.0.1', '.' and FQDN forms of the local name.
            $isLocal = $computer -in @('.', 'localhost', '127.0.0.1', '::1') -or
                       ($computer -split '\.')[0] -eq $env:COMPUTERNAME

            try {
                # --- Step 2: Machine-level TCP/IP port and printer setup (local only) ---
                if ($ipMappings.Count -gt 0) {
                    if (-not $isLocal) {
                        throw "Machine-level printer port additions are not supported for remote targets. Run this script directly on '$computer' via RMM."
                    }

                # Install required drivers for IP destinations before creating ports/printers
                $driversToInstall = $ipMappings |
                    Where-Object { $_.DriverName } |
                    Sort-Object DriverName -Unique
                foreach ($driverInfo in $driversToInstall) {
                    $driverResult = Install-VBPrinterDriver -DriverName $driverInfo.DriverName `
                                                             -DriverPath $driverInfo.DriverPath
                    if ($driverResult.RebootRequired) {
                        Write-Warning "Driver '$($driverInfo.DriverName)' requires a reboot. Printer mappings may not function until the machine is restarted."
                    }
                }

                    foreach ($ipMap in $ipMappings) {
                        $portName = "IP_$($ipMap.NewPath)"

                        # Use explicit PrinterName when supplied; otherwise derive from UNC share segment
                        $printerDisplayName = $null
                        if ($ipMap.PrinterName) {
                            $printerDisplayName = $ipMap.PrinterName
                        }
                        elseif ($ipMap.OldPath -match '^\\\\') {
                            $printerDisplayName = $ipMap.OldPath.TrimEnd('\').Split('\')[-1]
                        }

                        # Add TCP/IP port if it does not exist
                        if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
                            if ($PSCmdlet.ShouldProcess($portName, 'Add TCP/IP printer port')) {
                                Add-PrinterPort -Name $portName -PrinterHostAddress $ipMap.NewPath -ErrorAction Stop
                                Write-Verbose "Added printer port: $portName"
                            }
                        }

                        # Add or update machine-level printer if display name is known
                        if ($printerDisplayName) {
                            $existingPrinter = Get-Printer -Name $printerDisplayName -ErrorAction SilentlyContinue

                            if ($existingPrinter) {
                                # Printer exists -- check if it is already on the correct port
                                if ($existingPrinter.PortName -ne $portName) {
                                    if ($PSCmdlet.ShouldProcess($printerDisplayName, "Update port: $($existingPrinter.PortName) -> $portName")) {
                                        Set-Printer -Name $printerDisplayName -PortName $portName -ErrorAction Stop
                                        Write-Verbose "Updated printer '$printerDisplayName' port: $($existingPrinter.PortName) -> $portName"
                                    }
                                }
                                else {
                                    Write-Verbose "Printer '$printerDisplayName' already on correct port ($portName) -- skipping."
                                }
                            }
                            else {
                                # Pre-check the driver -- Add-Printer fails with an opaque
                                # error if the driver is not in the driver store
                                if (-not (Get-PrinterDriver -Name $ipMap.DriverName -ErrorAction SilentlyContinue)) {
                                    throw "Printer driver '$($ipMap.DriverName)' is not installed on $computer. Install the driver before running the migration. Run 'Get-PrinterDriver | Select-Object Name' to list installed drivers."
                                }

                                if ($PSCmdlet.ShouldProcess($printerDisplayName, 'Add printer')) {
                                    Add-Printer -Name $printerDisplayName -PortName $portName -DriverName $ipMap.DriverName -ErrorAction Stop
                                    Write-Verbose "Added printer: $printerDisplayName on port $portName"
                                }
                            }
                        }
                    }
                }

                # --- Step 3: Resolve target user profiles ---
                $profileParams = @{ ErrorAction = 'Stop' }

                if (-not $isLocal) {
                    $profileParams['ComputerName'] = $computer
                    if ($Credential) { $profileParams['Credential'] = $Credential }
                }

                $profiles = Get-VBUserProfile @profileParams

                if ($perUserMappings) {
                    # Rich CSV: filter profiles to usernames defined in the CSV for this machine
                    $profiles = @($profiles | Where-Object {
                        $p = $_
                        $perUserMappings.Keys | Where-Object {
                            $_ -eq $p.Username -or
                            $_ -eq "$($p.Domain)\$($p.Username)" -or
                            $_ -eq (Split-Path $p.ProfilePath -Leaf)
                        }
                    })
                    foreach ($csvUser in $perUserMappings.Keys) {
                        $found = $profiles | Where-Object {
                            $csvUser -eq $_.Username -or
                            $csvUser -eq "$($_.Domain)\$($_.Username)" -or
                            $csvUser -eq (Split-Path $_.ProfilePath -Leaf)
                        }
                        if (-not $found) {
                            Write-Warning "Username '$csvUser' from CSV not found on $computer -- skipped."
                        }
                    }
                    if ($profiles.Count -eq 0) {
                        Write-Warning "No matching profiles found on $computer -- skipping machine."
                        continue
                    }
                }
                elseif ($Username -or $SID) {
                    $profiles = @($profiles | Where-Object {
                        $p = $_
                        # Match any of: short name, DOMAIN\user, or profile folder name
                        # Covers on-prem AD, Entra ID (dot-format), and UPN-style accounts
                        $usernameMatch = $Username -and ($Username | Where-Object {
                            $_ -eq $p.Username -or
                            $_ -eq "$($p.Domain)\$($p.Username)" -or
                            $_ -eq (Split-Path $p.ProfilePath -Leaf)
                        })
                        $usernameMatch -or ($SID -and $p.SID -in $SID)
                    })

                    # Warn for each requested user not found on this machine
                    foreach ($name in $Username) {
                        $found = $profiles | Where-Object {
                            $name -eq $_.Username -or
                            $name -eq "$($_.Domain)\$($_.Username)" -or
                            $name -eq (Split-Path $_.ProfilePath -Leaf)
                        }
                        if (-not $found) {
                            Write-Warning "Username '$name' not found on $computer -- skipped."
                        }
                    }
                    foreach ($sid in $SID) {
                        if ($sid -notin ($profiles | Select-Object -ExpandProperty SID)) {
                            Write-Warning "SID '$sid' not found on $computer -- skipped."
                        }
                    }

                    if ($profiles.Count -eq 0) {
                        Write-Warning "No matching profiles found on $computer -- skipping machine."
                        continue
                    }
                }

                # --- Step 4: Per-user migration ---
                # NOTE: $userProfile (not $profile) -- $profile would shadow the
                # automatic $PROFILE variable
                foreach ($userProfile in $profiles) {

                    $mountParams = @{
                        SID         = $userProfile.SID
                        ErrorAction = 'Stop'
                    }

                    if (-not $isLocal) {
                        $mountParams['ComputerName'] = $computer
                        if ($Credential) { $mountParams['Credential'] = $Credential }
                    }

                    $mountResult = Mount-VBUserHive @mountParams

                    if ($mountResult.Status -ne 'Success') {
                        [PSCustomObject]@{
                            ComputerName = $computer
                            Username     = $userProfile.Username
                            SID          = $userProfile.SID
                            OldPath      = 'N/A'
                            NewPath      = 'N/A'
                            Action       = 'Failed'
                            Details      = "Hive mount failed: $($mountResult.Error)"
                            Status       = 'Failed'
                            Error        = $mountResult.Error
                            Timestamp    = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                        }
                        continue
                    }

                    try {
                        # Optional: backup current printer state before changes.
                        # A failed backup SKIPS migration for this user -- changes
                        # are never applied without the requested safety net.
                        if ($BackupMappings) {
                            try {
                                $backupParams = @{ TableOutput = $true; ErrorAction = 'Stop' }
                                if (-not $isLocal) {
                                    $backupParams['ComputerName'] = $computer
                                    if ($Credential) { $backupParams['Credential'] = $Credential }
                                }

                                $backupData = Get-VBUserPrinterMappings @backupParams |
                                Where-Object { $_.Username -eq $userProfile.Username }

                                if ($backupData) {
                                    $backupData | Export-Csv -Path $BackupPath -NoTypeInformation -Append -Encoding UTF8 -ErrorAction Stop
                                }
                            }
                            catch {
                                Write-Warning "Backup failed for '$($userProfile.Username)' on $computer -- migration skipped for this user. $($_.Exception.Message)"
                                [PSCustomObject]@{
                                    ComputerName = $computer
                                    Username     = $userProfile.Username
                                    SID          = $userProfile.SID
                                    OldPath      = 'N/A'
                                    NewPath      = 'N/A'
                                    Action       = 'Failed'
                                    Details      = "Backup failed -- migration skipped for this user: $($_.Exception.Message)"
                                    Status       = 'Failed'
                                    Error        = $_.Exception.Message
                                    Timestamp    = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                                }
                                continue
                            }
                        }

                        # Resolve per-user mappings (rich CSV) or all mappings (legacy)
                        $userMappings = if ($perUserMappings) {
                            $key = $perUserMappings.Keys | Where-Object {
                                $_ -eq $profile.Username -or
                                $_ -eq "$($profile.Domain)\$($profile.Username)" -or
                                $_ -eq (Split-Path $profile.ProfilePath -Leaf)
                            } | Select-Object -First 1
                            if ($key) { @($perUserMappings[$key]) } else { @() }
                        }
                        else {
                            $normalizedMappings
                        }

                        # Apply registry changes for this user
                        if ($userMappings.Count -gt 0) {
                            Update-VBUserPrinterRegistry `
                                -SID          $mountResult.SID `
                                -Username     $profile.Username `
                                -ComputerName $computer `
                                -Mappings     $userMappings
                        }

                        # Explicitly set default printer if specified in CSV (overrides auto-update)
                        $defaultRow = @($userMappings) | Where-Object { $_.DefaultPrinter } | Select-Object -First 1
                        if ($defaultRow) {
                            $regBase     = "Registry::HKEY_USERS\$($mountResult.SID)"
                            $devicesPath = "$regBase\Software\Microsoft\Windows NT\CurrentVersion\Devices"
                            $windowsPath = "$regBase\Software\Microsoft\Windows NT\CurrentVersion\Windows"

                            $deviceProps = Get-ItemProperty -Path $devicesPath -ErrorAction SilentlyContinue
                            if ($deviceProps -and
                                ($deviceProps.PSObject.Properties.Name -contains $defaultRow.DefaultPrinter)) {
                                $deviceValue = $deviceProps.($defaultRow.DefaultPrinter)
                                if ($PSCmdlet.ShouldProcess(
                                        "$($profile.Username) on $computer",
                                        "Set default printer: $($defaultRow.DefaultPrinter)")) {
                                    Set-ItemProperty -Path $windowsPath -Name 'Device' `
                                        -Value "$($defaultRow.DefaultPrinter),$deviceValue" `
                                        -Type String -Force
                                    Write-Verbose "Set default printer '$($defaultRow.DefaultPrinter)' for $($profile.Username)"
                                }
                            }
                            else {
                                Write-Warning "DefaultPrinter '$($defaultRow.DefaultPrinter)' not found in Devices for $($profile.Username) -- skipped."
                            }
                        }
                    }
                    finally {
                        # Always dismount -- even if Update-VBUserPrinterRegistry throws
                        # or the backup catch issues 'continue'
                        $mountResult | Dismount-VBUserHive | Out-Null
                    }
                }
            }
            catch {
                [PSCustomObject]@{
                    ComputerName = $computer
                    Username     = 'N/A'
                    SID          = 'N/A'
                    OldPath      = 'N/A'
                    NewPath      = 'N/A'
                    Action       = 'Failed'
                    Details      = $_.Exception.Message
                    Status       = 'Failed'
                    Error        = $_.Exception.Message
                    Timestamp    = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
