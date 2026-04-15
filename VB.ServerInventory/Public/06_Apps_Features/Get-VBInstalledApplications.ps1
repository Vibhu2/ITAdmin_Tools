# ============================================================
# FUNCTION : Get-VBInstalledApplications
# VERSION  : 1.0.0
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Query installed applications from registry (32-bit and 64-bit hives)
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Retrieves installed applications from local or remote computers.

.DESCRIPTION
    Queries the Windows registry (both 32-bit and 64-bit hives) to enumerate
    installed applications. Returns one PSCustomObject per application with
    name, version, publisher, and install date. Supports pipeline input and
    remote execution via WinRM.

.PARAMETER ComputerName
    Target computer(s) to query. Accepts pipeline input. Defaults to local
    machine ($env:COMPUTERNAME). Aliases: Name, Server, Host.

.PARAMETER Credential
    PSCredential object for remote execution. Required for cross-domain or
    credential-based remote queries.

.EXAMPLE
    Get-VBInstalledApplications

.EXAMPLE
    Get-VBInstalledApplications -ComputerName SERVER01

.EXAMPLE
    'SRV01','SRV02' | Get-VBInstalledApplications -Credential (Get-Credential)

.OUTPUTS
    [PSCustomObject]: ComputerName, DisplayName, DisplayVersion, Publisher,
    InstallDate, Status, CollectionTime

.NOTES
    Version  : 1.0.0
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : Applications

#>

function Get-VBInstalledApplications {
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
                # Step 1 -- Define script block for registry query
                $scriptBlock = {
                    $Applications = @()

                    # Step 2 -- Query 64-bit registry hive
                    $Applications += Get-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' `
                        -ErrorAction SilentlyContinue |
                        Where-Object { $_.DisplayName -ne $null }

                    # Step 3 -- Query 32-bit registry hive (WOW6432Node)
                    $Applications += Get-ItemProperty -Path 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' `
                        -ErrorAction SilentlyContinue |
                        Where-Object { $_.DisplayName -ne $null }

                    # Step 4 -- Return deduplicated and sorted results
                    $Applications | Sort-Object DisplayName -Unique
                }

                # Step 5 -- Execute locally or remotely
                if ($computer -eq $env:COMPUTERNAME) {
                    $results = & $scriptBlock
                } else {
                    $invokeParams = @{
                        ComputerName = $computer
                        ScriptBlock  = $scriptBlock
                    }
                    if ($PSBoundParameters.ContainsKey('Credential')) {
                        $invokeParams['Credential'] = $Credential
                    }
                    $results = Invoke-Command @invokeParams
                }

                # Step 6 -- Emit one PSCustomObject per application
                foreach ($item in $results) {
                    [PSCustomObject]@{
                        ComputerName   = $computer
                        DisplayName    = $item.DisplayName
                        DisplayVersion = $item.DisplayVersion
                        Publisher      = $item.Publisher
                        InstallDate    = $item.InstallDate
                        Status         = 'Success'
                        CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    }
                }
            }
            catch {
                # Step 7 -- Return error object on failure
                [PSCustomObject]@{
                    ComputerName   = $computer
                    DisplayName    = $null
                    DisplayVersion = $null
                    Publisher      = $null
                    InstallDate    = $null
                    Status         = 'Failed'
                    Error          = $_.Exception.Message
                }
            }
        }
    }
}
