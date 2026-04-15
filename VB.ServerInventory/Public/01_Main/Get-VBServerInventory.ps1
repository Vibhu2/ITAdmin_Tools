# ============================================================
# FUNCTION : Get-VBServerInventory
# VERSION  : 1.0.0
# CHANGED  : 10-04-2026 -- Initial VB-compliant release (orchestrator rewrite)
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Orchestrates full server inventory by calling all VB inventory functions
# ENCODING : UTF-8 with BOM
# ------------------------------------------------------------
# CHANGELOG (last 3-5 only -- full history in Git)
# v1.0.0 -- 10-04-2026 -- Initial release, replaces monolithic Get-ServerInventory
# ============================================================
#
# PREREQUISITE: All Get-VB* functions must be loaded before calling this function.
# Dot-source each file from Updated_Code or load via a module manifest.
#
# CATEGORY FLAGS -- control which sections run via switch params:
#   -IncludeAD        : Runs all AD / GPO / AD User Health sections
#   -IncludeSecurity  : Runs BitLocker, Firewall, AzureAD sections
#   -IncludePrinting  : Runs printer and share sections
#   -IncludeApps      : Runs installed apps, Store apps, features, updates
#   -SkipStore        : Skips Windows Store apps (slow on some servers)
#   -ExportCSV        : Exports each section to a separate CSV under OutputPath

<#
.SYNOPSIS
    Orchestrates a full server inventory by calling all VB inventory functions.

.DESCRIPTION
    Get-VBServerInventory is the top-level orchestrator for the Realtime server
    inventory toolkit. It calls each individual Get-VB* function, collects the
    results, and optionally exports each section to a dated CSV file.

    All individual Get-VB* functions must be loaded (dot-sourced or module-imported)
    before calling this function. The orchestrator itself contains no collection
    logic -- it delegates entirely to the individual functions.

    Use the category switches to control which sections run. By default only the
    core sections (System, Disk, Network, Services) run. AD, Security, Printing,
    and Apps sections are opt-in to keep runtime reasonable.

.PARAMETER ComputerName
    Target computer(s) to inventory. Defaults to local machine. Accepts pipeline input.

.PARAMETER Credential
    Alternate credentials passed through to all individual inventory functions.

.PARAMETER ExportCSV
    When specified, exports each section to a separate CSV file under OutputPath.

.PARAMETER OutputPath
    Folder path for CSV exports. Defaults to a dated folder on the desktop.
    Created automatically if it does not exist.

.PARAMETER IncludeAD
    Runs AD, GPO, and AD User Health sections.

.PARAMETER IncludeSecurity
    Runs BitLocker, Firewall, and Azure AD join status sections.

.PARAMETER IncludePrinting
    Runs printer and file share sections.

.PARAMETER IncludeApps
    Runs installed applications, Windows features, and updates sections.

.PARAMETER SkipStore
    Skips the Windows Store apps section when -IncludeApps is specified.

.EXAMPLE
    Get-VBServerInventory

    Runs core inventory (System, Disk, Network) on the local machine.

.EXAMPLE
    Get-VBServerInventory -ComputerName SERVER01 -IncludeAD -IncludeSecurity

    Runs core + AD + Security sections against SERVER01.

.EXAMPLE
    Get-VBServerInventory -ComputerName SERVER01 -IncludeAD -IncludeApps -ExportCSV

    Runs core + AD + Apps sections and exports all results to CSV on the desktop.

.EXAMPLE
    'SERVER01','SERVER02' | Get-VBServerInventory -IncludeAD -Credential (Get-Credential)

    Pipelines two servers with alternate credentials, running core + AD sections.

.OUTPUTS
    [PSCustomObject]: Section, ComputerName, Data (PSCustomObject array per section)

.NOTES
    Version  : 1.0.0
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : Main / Orchestrator
#>
function Get-VBServerInventory {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'Server', 'Host')]
        [string[]]$ComputerName = $env:COMPUTERNAME,

        [PSCredential]$Credential,

        [switch]$ExportCSV,

        [string]$OutputPath = (Join-Path $env:USERPROFILE "Desktop\ServerInventory-$(Get-Date -Format 'yyyyMMdd-HHmmss')"),

        [switch]$IncludeAD,
        [switch]$IncludeSecurity,
        [switch]$IncludePrinting,
        [switch]$IncludeApps,
        [switch]$SkipStore
    )

    begin {
        # Step 1 -- Create export folder if needed
        if ($ExportCSV -and -not (Test-Path -Path $OutputPath)) {
            New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
            Write-Verbose "Created output folder: $OutputPath"
        }

        # Step 2 -- Build shared param hashtable for pass-through to all functions
        $SharedParams = @{}
        if ($Credential) { $SharedParams['Credential'] = $Credential }
    }

    process {
        foreach ($computer in $ComputerName) {
            Write-Verbose "Starting inventory for: $computer"
            $AllSections = [System.Collections.Generic.List[PSCustomObject]]::new()

            # Step 3 -- Core sections (always run)
            $Sections = [ordered]@{
                'SystemInfo'        = { Get-VBSystemInfo        -ComputerName $computer @SharedParams }
                'DiskInformation'   = { Get-VBDiskInformation   -ComputerName $computer @SharedParams }
                'NetworkInfo'       = { Get-VBNetworkInformation -ComputerName $computer @SharedParams }
                'DHCPInformation'   = { Get-VBDHCPInformation    -ComputerName $computer @SharedParams }
                'RDSUsers'          = { Get-VBRDSUserInformation -ComputerName $computer @SharedParams }
            }

            # Step 4 -- AD sections (opt-in)
            if ($IncludeAD) {
                $Sections['ActiveDirectory']      = { Get-VBActiveDirectoryInfo          -ComputerName $computer @SharedParams }
                $Sections['GPOInformation']       = { Get-VBGPOInformation               -ComputerName $computer @SharedParams }
                $Sections['GPOComprehensive']     = { Get-VBGPOComprehensiveReport       -ComputerName $computer @SharedParams }
                $Sections['UnusedGPOs']           = { Get-VBUnusedGPOs                  -ComputerName $computer @SharedParams }
                $Sections['GpoConnections']       = { Get-VBGpoConnections              -ComputerName $computer @SharedParams }
                $Sections['InactiveUsers']        = { Get-VBInactiveUsers               -ComputerName $computer @SharedParams }
                $Sections['InactiveComputers']    = { Get-VBInactiveComputers           -ComputerName $computer @SharedParams }
                $Sections['NoLoginAccounts']      = { Get-VBAdAccountWithNoLogin        -ComputerName $computer @SharedParams }
                $Sections['NoPasswordRequired']   = { Get-VBNoPasswordRequiredUsers     -ComputerName $computer @SharedParams }
                $Sections['PasswordNeverExpires'] = { Get-VBPasswordNeverExpiresUsers   -ComputerName $computer @SharedParams }
                $Sections['OldAdminPasswords']    = { Get-VBOldAdminPasswords           -ComputerName $computer @SharedParams }
                $Sections['EmptyADGroups']        = { Get-VBEmptyADGroups              -ComputerName $computer @SharedParams }
                $Sections['ADGroupMemberCount']   = { Get-VBADGroupsWithMemberCount    -ComputerName $computer @SharedParams }
                $Sections['DNSServerInfo']        = { Get-VBDNSServerInfo              -ComputerName $computer @SharedParams }
                $Sections['DHCPDetailedInfo']     = { Get-VBDhcpInfo                   -ComputerName $computer @SharedParams }
            }

            # Step 5 -- Security sections (opt-in)
            if ($IncludeSecurity) {
                $Sections['AzureADJoinStatus']    = { Get-VBAzureADJoinStatus          -ComputerName $computer @SharedParams }
                $Sections['BitLockerRecovery']    = { Get-VBBitLockerRecoveryKey       -ComputerName $computer @SharedParams }
                $Sections['FirewallRules']        = { Get-VBFirewallPortRules          -ComputerName $computer @SharedParams }
                $Sections['ScheduledTasks']       = { Get-VBNonMicrosoftScheduledTasks -ComputerName $computer @SharedParams }
            }

            # Step 6 -- Printing and Shares (opt-in)
            if ($IncludePrinting) {
                $Sections['ShareInformation']     = { Get-VBShareInformation           -ComputerName $computer @SharedParams }
                $Sections['PrinterInformation']   = { Get-VBPrinterInformation        -ComputerName $computer @SharedParams }
                $Sections['PrintingInfo']         = { Get-VBPrintPrintingInfo          -ComputerName $computer @SharedParams }
            }

            # Step 7 -- Apps and Features (opt-in)
            if ($IncludeApps) {
                $Sections['InstalledApplications'] = { Get-VBInstalledApplications    -ComputerName $computer @SharedParams }
                $Sections['WindowsFeatures']       = { Get-VBWindowsFeaturesInfo      -ComputerName $computer @SharedParams }
                $Sections['WindowsUpdates']        = { Get-VBWindowsUpdateInfo        -ComputerName $computer @SharedParams }
                if (-not $SkipStore) {
                    $Sections['WindowsStoreApps']  = { Get-VBWindowsStoreApps         -ComputerName $computer @SharedParams }
                }
            }

            # Step 8 -- Execute each section, collect results, optionally export CSV
            foreach ($sectionName in $Sections.Keys) {
                Write-Verbose "  Running section: $sectionName"
                try {
                    $SectionData = & $Sections[$sectionName]

                    if ($ExportCSV -and $SectionData) {
                        $CsvFile = Join-Path $OutputPath "$computer-$sectionName.csv"
                        $SectionData | Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8
                        Write-Verbose "    Exported: $CsvFile"
                    }

                    $AllSections.Add([PSCustomObject]@{
                        ComputerName = $computer
                        Section      = $sectionName
                        RecordCount  = @($SectionData).Count
                        Data         = $SectionData
                        Status       = 'Success'
                        CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    })
                }
                catch {
                    $AllSections.Add([PSCustomObject]@{
                        ComputerName = $computer
                        Section      = $sectionName
                        RecordCount  = 0
                        Data         = $null
                        Error        = $_.Exception.Message
                        Status       = 'Failed'
                        CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    })
                    Write-Warning "Section '$sectionName' failed for $computer -- $($_.Exception.Message)"
                }
            }

            # Step 9 -- Return all section results for this computer
            $AllSections
        }
    }
}
