# ============================================================
# FUNCTION : Get-VBUserPrinterMappings
# MODULE   : WorkstationReport
# VERSION  : 1.1.1
# CHANGED  : 14-04-2026 -- Standards compliance fixes
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Retrieves printer mappings for all user profiles on local or remote computers
# ENCODING : UTF-8 with BOM
# ============================================================

function Get-VBUserPrinterMappings {
    <#
    .SYNOPSIS
    Retrieves printer mappings and default printer settings for all user profiles on
    specified computers.

    .DESCRIPTION
    Get-VBUserPrinterMappings audits printer configurations for all user profiles on local
    or remote computers. It examines user registry hives to identify network printer
    mappings, default printer settings, and available printer devices, while excluding
    virtual and system printers (PDF, XPS, OneNote, Fax).

    Registry hives for inactive profiles are loaded and safely unloaded after inspection.
    A single scriptblock is used for both local and remote execution to eliminate code
    duplication -- local runs use direct invocation (&), remote runs use Invoke-Command.

    .PARAMETER ComputerName
    Computer names to audit. Accepts pipeline input. Defaults to the local computer.

    .PARAMETER Credential
    Credentials for remote computer access. Not required for local execution.

    .PARAMETER TableOutput
    When specified, suppresses console output and returns only structured objects.
    Without this switch, verbose progress is written to the console using Write-Host.

    .EXAMPLE
    Get-VBUserPrinterMappings

    Audits printer mappings for all user profiles on the local computer.

    .EXAMPLE
    Get-VBUserPrinterMappings -ComputerName 'WS001','WS002' -Credential $cred -TableOutput

    Audits printer mappings on two remote workstations and returns structured objects only.

    .EXAMPLE
    'WS001','WS002' | Get-VBUserPrinterMappings |
        Where-Object { $_.PrinterCount -gt 0 } | Format-Table

    Pipeline input -- audits workstations and shows only users with mapped printers.

    .EXAMPLE
    Get-VBUserPrinterMappings -TableOutput |
        Export-Csv -Path 'C:\Reports\PrinterAudit.csv' -NoTypeInformation

    Exports printer audit results to CSV without console output.

    .OUTPUTS
    PSCustomObject
    Returns objects with:
    - ComputerName      : Target computer
    - Username          : User profile name
    - NetworkPrinters   : Semicolon-separated list of network printer connections
    - DefaultPrinter    : Name of the default printer, or 'No default printer'
    - PrinterDevices    : Semicolon-separated list of available printer devices
    - PrinterCount      : Total count of network printers and devices combined
    - LastProfileUpdate : Last write time of the user profile directory
    - CollectionTime    : Timestamp of data collection
    - Status            : 'Success' or 'Failed'
    - Error             : Error message (only present on failure)

    .NOTES
    Version : 1.1.1
    Author  : Vibhu Bhatnagar
    Category: Windows Workstation Administration

    Requirements:
    - PowerShell 5.1 or higher
    - Administrative privileges on target computers
    - PowerShell Remoting enabled for remote targets
    - Registry access permissions for HKEY_USERS and HKEY_LOCAL_MACHINE
    #>

    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Name', 'Server', 'Host')]
        [string[]]$ComputerName = $env:COMPUTERNAME,

        [PSCredential]$Credential,

        [switch]$TableOutput
    )

    process {
        # Core logic defined once -- executed locally or via Invoke-Command
        $scriptBlock = {
            param([bool]$ShowConsole)

            $computerName    = $env:COMPUTERNAME
            $collectionTime  = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
            $results         = [System.Collections.Generic.List[object]]::new()

            $virtualPrinters = @(
                'Microsoft Print to PDF',
                'Microsoft XPS Document Writer',
                'OneNote (Desktop)',
                'OneNote for Windows 10'
            )
            $systemPrinters  = @('Fax')

            if ($ShowConsole) {
                Write-Host "`nStarting printer mapping audit on: $computerName" -ForegroundColor Green
                Write-Host '=================================================' -ForegroundColor Green
            }

            # Collect user profile paths
            $userProfiles = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' |
                Where-Object { $_.PSChildName.Length -gt 8 } |
                ForEach-Object { $_.GetValue('ProfileImagePath') }

            foreach ($profilePath in $userProfiles) {
                $username   = Split-Path $profilePath -Leaf
                $hiveLoaded = $false

                if ($ShowConsole) {
                    Write-Host "`nChecking: $username" -ForegroundColor Cyan
                    Write-Host ('-' * 40) -ForegroundColor Gray
                }

                # Resolve SID for this profile
                $userSID = (
                    Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' |
                    Where-Object { $_.GetValue('ProfileImagePath') -eq $profilePath }
                ).PSChildName

                if (-not $userSID) {
                    if ($ShowConsole) { Write-Host "  Could not resolve SID for: $username" -ForegroundColor Red }
                    $results.Add([PSCustomObject]@{
                        ComputerName      = $computerName
                        Username          = $username
                        NetworkPrinters   = 'No SID Found'
                        DefaultPrinter    = 'No SID Found'
                        PrinterDevices    = 'No SID Found'
                        PrinterCount      = 0
                        LastProfileUpdate = (Get-Item $profilePath -ErrorAction SilentlyContinue).LastWriteTime
                        CollectionTime    = $collectionTime
                        Error             = 'Could not resolve SID for user profile'
                        Status            = 'Failed'
                    })
                    continue
                }

                try {
                    # Load hive if not already mounted
                    if (-not (Test-Path "Registry::HKEY_USERS\$userSID")) {
                        $null = reg load "HKU\$userSID" "$profilePath\NTUSER.DAT"
                        $hiveLoaded = $true
                    }

                    # --- Network printers ---
                    $networkPrinters = @(
                        Get-ChildItem "Registry::HKEY_USERS\$userSID\Printers\Connections" -ErrorAction SilentlyContinue |
                        ForEach-Object { $_.PSChildName.Replace(',', '\') } |
                        Where-Object {
                            $pn = $_
                            -not ($virtualPrinters | Where-Object { $pn -like "*$_*" })
                        }
                    )

                    if ($ShowConsole) {
                        Write-Host '  Mapped Network Printers:' -ForegroundColor Yellow
                        if ($networkPrinters) { $networkPrinters | ForEach-Object { Write-Host "    - $_" -ForegroundColor White } }
                        else { Write-Host '    None' -ForegroundColor Gray }
                    }

                    # --- Default printer ---
                    $defaultPrinterReg  = Get-ItemProperty "Registry::HKEY_USERS\$userSID\Software\Microsoft\Windows NT\CurrentVersion\Windows" `
                        -Name 'Device' -ErrorAction SilentlyContinue
                    $defaultPrinterName = if ($defaultPrinterReg) {
                        $pn = ($defaultPrinterReg.Device -split ',')[0]
                        $isVirtual = $virtualPrinters | Where-Object { $pn -like "*$_*" }
                        $isSystem  = $systemPrinters  | Where-Object { $pn -like "*$_*" }
                        if ($isVirtual -or $isSystem) { 'No default printer' } else { $pn }
                    }
                    else { 'No default printer' }

                    if ($ShowConsole) {
                        Write-Host '  Default Printer:' -ForegroundColor Yellow
                        Write-Host "    - $defaultPrinterName" -ForegroundColor White
                    }

                    # --- Printer devices ---
                    $allDevices    = Get-ItemProperty "Registry::HKEY_USERS\$userSID\Software\Microsoft\Windows NT\CurrentVersion\Devices" `
                        -ErrorAction SilentlyContinue |
                        Select-Object -Property * -ExcludeProperty PS* |
                        Get-Member -MemberType NoteProperty |
                        Select-Object -ExpandProperty Name

                    $printerDevices = @(
                        $allDevices | Where-Object {
                            $d = $_
                            -not ($virtualPrinters | Where-Object { $d -like "*$_*" }) -and
                            -not ($systemPrinters  | Where-Object { $d -like "*$_*" })
                        }
                    )

                    if ($ShowConsole) {
                        Write-Host '  Printer Devices:' -ForegroundColor Yellow
                        if ($printerDevices) { $printerDevices | ForEach-Object { Write-Host "    - $_" -ForegroundColor White } }
                        else { Write-Host '    None' -ForegroundColor Gray }
                    }

                    $results.Add([PSCustomObject]@{
                        ComputerName      = $computerName
                        Username          = $username
                        NetworkPrinters   = if ($networkPrinters) { $networkPrinters -join '; ' } else { 'None' }
                        DefaultPrinter    = $defaultPrinterName
                        PrinterDevices    = if ($printerDevices) { $printerDevices -join '; ' } else { 'None' }
                        PrinterCount      = $networkPrinters.Count + $printerDevices.Count
                        LastProfileUpdate = (Get-Item $profilePath -ErrorAction SilentlyContinue).LastWriteTime
                        CollectionTime    = $collectionTime
                        Status            = 'Success'
                    })
                }
                catch {
                    if ($ShowConsole) { Write-Host "  Error processing $username : $_" -ForegroundColor Red }
                    $results.Add([PSCustomObject]@{
                        ComputerName      = $computerName
                        Username          = $username
                        NetworkPrinters   = 'Error'
                        DefaultPrinter    = 'Error'
                        PrinterDevices    = 'Error'
                        PrinterCount      = 0
                        LastProfileUpdate = (Get-Item $profilePath -ErrorAction SilentlyContinue).LastWriteTime
                        CollectionTime    = $collectionTime
                        Error             = $_.Exception.Message
                        Status            = 'Failed'
                    })
                }
                finally {
                    if ($hiveLoaded) {
                        [System.GC]::Collect()
                        Start-Sleep -Seconds 1
                        $null = reg unload "HKU\$userSID"
                    }
                }
            }

            return $results
        }

        foreach ($computer in $ComputerName) {
            try {
                Write-Verbose "Targeting computer: $computer"

                $showConsole = -not $TableOutput.IsPresent

                if ($computer -eq $env:COMPUTERNAME) {
                    & $scriptBlock $showConsole
                }
                else {
                    $invokeParams = @{
                        ComputerName = $computer
                        ScriptBlock  = $scriptBlock
                        ArgumentList = $showConsole
                        ErrorAction  = 'Stop'
                    }
                    if ($Credential) { $invokeParams['Credential'] = $Credential }
                    Invoke-Command @invokeParams
                }
            }
            catch {
                Write-Error -Message "Failed to connect to '$computer': $($_.Exception.Message)" -ErrorAction Continue
                [PSCustomObject]@{
                    ComputerName    = $computer
                    Username        = 'Connection Error'
                    NetworkPrinters = 'Connection Failed'
                    DefaultPrinter  = 'Connection Failed'
                    PrinterDevices  = 'Connection Failed'
                    PrinterCount    = 0
                    CollectionTime  = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    Error           = $_.Exception.Message
                    Status          = 'Failed'
                }
            }
        }
    }
}
