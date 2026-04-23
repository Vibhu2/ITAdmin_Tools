# ============================================================
# FUNCTION : Set-VBUserPrinterMigration
# MODULE   : VB.WorkstationReport
# VERSION  : 1.0.2
# CHANGED  : 23-04-2026 -- Added CSV format sample and step-by-step creation example
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Migrates user printer mappings between UNC paths and IP addresses
# ENCODING : UTF-8 with BOM
# ------------------------------------------------------------
# CHANGELOG (last 3-5 only -- full history in Git)
# v1.0.2 -- 23-04-2026 -- Added CSV format sample and step-by-step creation example
# v1.0.1 -- 23-04-2026 -- Finalized: expanded help with all migration types and RMM examples
# v1.0.0 -- 23-04-2026 -- Initial release
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
        error is thrown if a remote target requires machine-level port additions.

    .PARAMETER ComputerName
        Target computer(s). Accepts pipeline input. Defaults to local machine.
        Remote targets are supported for user registry changes only.

    .PARAMETER Credential
        Credentials for remote execution. Not required for local or domain-joined targets.

    .PARAMETER PrinterMappings
        Hashtable of OldPath = NewPath pairs.
        Example: @{ '\\OldServer\HP01' = '10.30.1.50'; '10.30.1.60' = '\\NewServer\Canon02' }
        For IP destinations, supply -DriverName. Use -MappingCsv for per-printer driver control.

    .PARAMETER MappingCsv
        Path to a CSV file with columns: OldPath, NewPath, DriverName
        DriverName is required when NewPath is an IP address. CSV takes priority over
        -PrinterMappings if both are supplied.

    .PARAMETER DriverName
        Driver name to use when adding IP printers via -PrinterMappings hashtable.
        Applies to all IP destinations in the hashtable. For per-printer driver control
        use -MappingCsv with a DriverName column instead.

    .PARAMETER TargetUser
        Username or SID of a specific user to migrate. When omitted all non-system
        user profiles on the machine are processed.

    .PARAMETER BackupMappings
        When specified, saves a snapshot of each user's current printer mappings to
        -BackupPath before applying changes. Requires -BackupPath.

    .PARAMETER BackupPath
        Full path to the CSV file where backup snapshots are written.
        Required when -BackupMappings is specified.

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
        Set-VBUserPrinterMigration -MappingCsv 'C:\Temp\PrinterMappings.csv' -TargetUser 'jdoe'

        Migrates printers for a single user only. All other profiles on the machine
        are skipped. Accepts username or SID for -TargetUser.

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
        Returns one object per user per printer mapping action:
          - ComputerName : Target computer
          - Username     : User profile name
          - SID          : User SID
          - OldPath      : Old printer path from mapping rule
          - NewPath      : New printer path from mapping rule
          - Action       : 'Migrated', 'Skipped', 'AlreadyMigrated', or 'Failed'
          - Details      : Registry actions taken or reason for skip/failure
          - Status       : 'Success' or 'Failed'
          - Error        : Error message (only present on failure)
          - Timestamp    : Time of action (dd-MM-yyyy HH:mm:ss)

    .NOTES
        Version  : 1.0.2
        Author   : Vibhu Bhatnagar
        Category : Printer Management

        Requirements:
        - PowerShell 5.1
        - Administrative privileges
        - PrintManagement module (built-in on Windows 8 / Server 2012+)
        - Required printer drivers already installed for IP destinations
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

        [string]$TargetUser,

        [switch]$BackupMappings,

        [string]$BackupPath
    )

    begin {
        $ErrorActionPreference = 'Stop'

        # --- Validate backup parameters ---
        if ($BackupMappings -and -not $BackupPath) {
            throw '-BackupPath is required when -BackupMappings is specified.'
        }

        # --- Step 1: Load and validate mappings ---
        $normalizedMappings = [System.Collections.Generic.List[object]]::new()

        if ($MappingCsv) {
            # CSV takes priority when both supplied
            if (-not (Test-Path -Path $MappingCsv)) {
                throw "Mapping CSV not found: $MappingCsv"
            }

            $csvRows = Import-Csv -Path $MappingCsv -ErrorAction Stop

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
                })
            }
        }
        elseif ($PrinterMappings) {
            foreach ($key in $PrinterMappings.Keys) {
                $oldPath  = $key.Trim()
                $newPath  = $PrinterMappings[$key].Trim()
                $isNewIP  = $newPath -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}'

                if ($isNewIP -and -not $DriverName) {
                    throw "Use -DriverName when supplying IP destinations via -PrinterMappings, or use -MappingCsv for per-printer driver control. Missing driver for: $oldPath -> $newPath"
                }

                $normalizedMappings.Add([PSCustomObject]@{
                    OldPath    = $oldPath
                    NewPath    = $newPath
                    DriverName = if ($isNewIP) { $DriverName } else { '' }
                })
            }
        }
        else {
            throw 'Either -PrinterMappings or -MappingCsv must be supplied.'
        }

        if ($normalizedMappings.Count -eq 0) {
            throw 'No valid printer mappings found in the supplied input.'
        }
    }

    process {
        foreach ($computer in $ComputerName) {
            try {
                # --- Step 2: Machine-level TCP/IP port and printer setup (local only) ---
                $ipMappings = $normalizedMappings |
                    Where-Object { $_.NewPath -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}' }

                if ($ipMappings) {
                    if ($computer -ne $env:COMPUTERNAME) {
                        throw "Machine-level printer port additions are not supported for remote targets. Run this script directly on '$computer' via RMM."
                    }

                    foreach ($ipMap in $ipMappings) {
                        $portName = "IP_$($ipMap.NewPath)"

                        # Only derive display name when old path is UNC (IP->IP handled per-user)
                        $printerDisplayName = $null
                        if ($ipMap.OldPath -match '^\\\\') {
                            $printerDisplayName = $ipMap.OldPath.TrimEnd('\').Split('\')[-1]
                        }

                        # Add TCP/IP port if it does not exist
                        if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
                            if ($PSCmdlet.ShouldProcess($portName, 'Add TCP/IP printer port')) {
                                Add-PrinterPort -Name $portName -PrinterHostAddress $ipMap.NewPath -ErrorAction Stop
                                Write-Verbose "Added printer port: $portName"
                            }
                        }

                        # Add machine-level printer if display name is known and printer does not exist
                        if ($printerDisplayName) {
                            if (-not (Get-Printer -Name $printerDisplayName -ErrorAction SilentlyContinue)) {
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

                # --- Step 4: Per-user migration ---
                foreach ($profile in $profiles) {

                    $mountParams = @{
                        SID         = $profile.SID
                        ErrorAction = 'Stop'
                    }

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
                            OldPath      = 'N/A'
                            NewPath      = 'N/A'
                            Action       = 'Failed'
                            Details      = "Hive mount failed: $($mountResult.Error)"
                            Error        = $mountResult.Error
                            Status       = 'Failed'
                            Timestamp    = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                        }
                        continue
                    }

                    try {
                        # Optional: backup current printer state before changes
                        if ($BackupMappings) {
                            $backupParams = @{ TableOutput = $true; ErrorAction = 'SilentlyContinue' }
                            if ($computer -ne $env:COMPUTERNAME) {
                                $backupParams['ComputerName'] = $computer
                                if ($Credential) { $backupParams['Credential'] = $Credential }
                            }

                            $backupData = Get-VBUserPrinterMappings @backupParams |
                                Where-Object { $_.Username -eq $profile.Username }

                            if ($backupData) {
                                $backupData | Export-Csv -Path $BackupPath -NoTypeInformation -Append -Encoding UTF8
                            }
                        }

                        # Apply registry changes for this user
                        Update-VBUserPrinterRegistry `
                            -SID          $mountResult.SID `
                            -Username     $profile.Username `
                            -ComputerName $computer `
                            -Mappings     $normalizedMappings
                    }
                    finally {
                        # Always dismount -- even if Update-VBUserPrinterRegistry throws
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
                    Error        = $_.Exception.Message
                    Status       = 'Failed'
                    Timestamp    = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
