# ============================================================
# FUNCTION : Get-VBSyncCenterStatus
# MODULE   : WorkstationReport
# VERSION  : 1.1.1
# CHANGED  : 14-04-2026 -- Standards compliance fixes
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Retrieves Windows Sync Center (CSC) configuration and status
# ENCODING : UTF-8 with BOM
# ============================================================

function Get-VBSyncCenterStatus {
    <#
    .SYNOPSIS
    Retrieves the Windows Sync Center (Offline Files / CSC) configuration and status.

    .DESCRIPTION
    Get-VBSyncCenterStatus examines the Windows Client Side Caching (CSC) subsystem by
    querying registry settings, WMI/CIM classes, and the cache directory. It reports on
    Group Policy configuration, the CSC driver and service startup types, cache active/
    enabled state, cache location, file count, and optionally the full list of cached files.

    A single scriptblock is used for both local and remote execution to eliminate code
    duplication -- local runs use direct invocation (&), remote runs use Invoke-Command.

    .PARAMETER ComputerName
    Computer names to query. Accepts pipeline input. Defaults to the local computer.

    .PARAMETER Credential
    Credentials for remote computer access. Not required for local execution.

    .PARAMETER IncludeFileList
    When specified, includes an array of all cached file paths in the output.
    May impact performance on systems with a large cache.

    .EXAMPLE
    Get-VBSyncCenterStatus

    Returns Sync Center status for the local computer.

    .EXAMPLE
    Get-VBSyncCenterStatus -ComputerName 'WS001','WS002' -Credential (Get-Credential)

    Returns Sync Center status for two remote workstations using alternate credentials.

    .EXAMPLE
    'WS001' | Get-VBSyncCenterStatus -IncludeFileList

    Queries a remote computer via pipeline and includes the full list of cached files.

    .EXAMPLE
    Get-VBSyncCenterStatus -ComputerName $computers -Credential $cred |
        Where-Object { $_.CSCStatus -ne 'Disabled' } |
        Select-Object ComputerName, CSCStatus, CacheEnabled, FileCount

    Finds all machines where CSC is not disabled and reports key fields.

    .OUTPUTS
    PSCustomObject
    Returns an object with:
    - ComputerName     : Target computer
    - NetCachePolicy   : Group Policy setting ('Not Set', 'Enabled', 'Disabled', 'Unknown')
    - CSCServiceStatus : CscService startup type ('Automatic', 'Disabled', 'Unknown')
    - CSCStatus        : CSC driver startup type ('System', 'Disabled', 'Unknown')
    - CacheActive      : Boolean -- whether the cache is currently active
    - CacheEnabled     : Boolean -- whether offline files is enabled
    - CacheLocation    : Filesystem path to the CSC cache directory
    - FileCount        : Number of files found in the cache directory
    - GroupPolicyName  : GPO ID applying the NetCache policy (if any)
    - FileList         : Array of cached file paths (empty unless -IncludeFileList used)
    - CollectionTime   : Timestamp of data collection
    - Status           : 'Success' or 'Failed'
    - Error            : Error message (only present on failure)

    .NOTES
    Version : 1.1.1
    Author  : Vibhu Bhatnagar
    Category: Windows Workstation Administration

    Requirements:
    - PowerShell 5.1 or higher
    - Administrative privileges for full registry and CIM access
    - PowerShell Remoting enabled for remote targets
    #>

    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Name', 'Server', 'Host')]
        [string[]]$ComputerName = $env:COMPUTERNAME,

        [PSCredential]$Credential,

        [switch]$IncludeFileList
    )

    process {
        # Core logic defined once -- executed locally or via Invoke-Command
        $scriptBlock = {
            param([bool]$IncludeFiles)

            $cscStart       = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\CSC' `
                -Name 'Start' -ErrorAction SilentlyContinue
            $cscServiceStart = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\CscService' `
                -Name 'Start' -ErrorAction SilentlyContinue
            $netCacheEnabled = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\NetCache' `
                -Name 'Enabled' -ErrorAction SilentlyContinue
            $cacheLocationReg = Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Services\CSC' `
                -Name 'CacheLocation' -ErrorAction SilentlyContinue
            $cacheInfo       = Get-CimInstance -ClassName Win32_OfflineFilesCache -ErrorAction SilentlyContinue

            # Group Policy identification
            $gpoName = 'Not Set'
            try {
                $gpResult = Get-CimInstance -ClassName RSOP_RegistryPolicySetting `
                    -Namespace 'root\rsop\computer' `
                    -Filter "registryKey like '%NetCache%'" -ErrorAction SilentlyContinue
                if ($gpResult) { $gpoName = $gpResult.GPOID }
            }
            catch {
                $gpoName = 'Unable to determine'
            }

            $cacheLocation = if ($cacheLocationReg.CacheLocation) {
                $cacheLocationReg.CacheLocation
            }
            else { 'C:\Windows\CSC' }

            # Cache file enumeration
            $fileCount = 0
            $fileList  = @()
            if (Test-Path $cacheLocation) {
                try {
                    $files     = Get-ChildItem -Path $cacheLocation -Recurse -File -ErrorAction SilentlyContinue
                    $fileCount = ($files | Measure-Object).Count
                    if ($IncludeFiles -and $files) {
                        $fileList = $files | Select-Object -ExpandProperty FullName
                    }
                }
                catch {
                    $fileCount = 0
                }
            }

            return @{
                CSCStart        = $cscStart
                CSCServiceStart = $cscServiceStart
                NetCacheEnabled = $netCacheEnabled
                CacheInfo       = $cacheInfo
                GPOName         = $gpoName
                CacheLocation   = $cacheLocation
                FileCount       = $fileCount
                FileList        = $fileList
            }
        }

        foreach ($computer in $ComputerName) {
            try {
                Write-Verbose "Querying Sync Center status on: $computer"

                $includeFiles = $IncludeFileList.IsPresent

                if ($computer -eq $env:COMPUTERNAME) {
                    $data = & $scriptBlock $includeFiles
                }
                else {
                    $invokeParams = @{
                        ComputerName = $computer
                        ScriptBlock  = $scriptBlock
                        ArgumentList = $includeFiles
                        ErrorAction  = 'Stop'
                    }
                    if ($Credential) { $invokeParams['Credential'] = $Credential }
                    $data = Invoke-Command @invokeParams
                }

                # Translate registry values to human-readable strings
                $netCachePolicyStatus = switch ($data.NetCacheEnabled.Enabled) {
                    $null { 'Not Set' }
                    0     { 'Disabled' }
                    1     { 'Enabled' }
                    default { 'Unknown' }
                }

                $cscServiceStatus = switch ($data.CSCServiceStart.Start) {
                    2       { 'Automatic' }
                    4       { 'Disabled' }
                    default { 'Unknown' }
                }

                $cscStatus = switch ($data.CSCStart.Start) {
                    1       { 'System' }
                    4       { 'Disabled' }
                    default { 'Unknown' }
                }

                Write-Debug "CSC=$cscStatus | CscService=$cscServiceStatus | Policy=$netCachePolicyStatus | Files=$($data.FileCount)"

                [PSCustomObject]@{
                    ComputerName     = $computer
                    NetCachePolicy   = $netCachePolicyStatus
                    CSCServiceStatus = $cscServiceStatus
                    CSCStatus        = $cscStatus
                    CacheActive      = if ($data.CacheInfo) { $data.CacheInfo.Active }   else { $false }
                    CacheEnabled     = if ($data.CacheInfo) { $data.CacheInfo.Enabled }  else { $false }
                    CacheLocation    = $data.CacheLocation
                    FileCount        = $data.FileCount
                    GroupPolicyName  = $data.GPOName
                    FileList         = if ($IncludeFileList) { $data.FileList } else { @() }
                    CollectionTime   = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    Status           = 'Success'
                }
            }
            catch {
                Write-Error -Message "Failed to query '$computer': $($_.Exception.Message)" -ErrorAction Continue
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
