# ============================================================
# FUNCTION : Get-VBWindowsStoreApps
# VERSION  : 1.0.2
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Query Windows Store (AppX) packages from local or remote systems
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Retrieves Windows Store application packages from local or remote computers.

.DESCRIPTION
    Enumerates AppX packages installed on Windows systems. Returns one
    PSCustomObject per package with name, version, publisher, and architecture.
    Supports pipeline input and remote execution via WinRM.

.PARAMETER ComputerName
    Target computer(s) to query. Accepts pipeline input. Defaults to local
    machine ($env:COMPUTERNAME). Aliases: Name, Server, Host.

.PARAMETER Credential
    PSCredential object for remote execution. Required for cross-domain or
    credential-based remote queries.

.EXAMPLE
    Get-VBWindowsStoreApps

.EXAMPLE
    Get-VBWindowsStoreApps -ComputerName SERVER01

.EXAMPLE
    'SRV01','SRV02' | Get-VBWindowsStoreApps -Credential (Get-Credential)

.OUTPUTS
    [PSCustomObject]: ComputerName, Name, Version, Publisher, Architecture,
    Status, CollectionTime

.NOTES
    Version  : 1.0.2
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : Applications

#>

function Get-VBWindowsStoreApps {
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
                # Step 1 -- Define script block to enumerate AppX packages
                $scriptBlock = {
                    Get-AppxPackage -ErrorAction SilentlyContinue | Sort-Object Name
                }

                # Step 2 -- Execute locally or remotely
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

                # Step 3 -- Emit one PSCustomObject per package
                foreach ($item in $results) {
                    [PSCustomObject]@{
                        ComputerName   = $computer
                        Name           = $item.Name
                        Version        = $item.Version
                        Publisher      = $item.Publisher
                        Architecture   = $item.Architecture
                        Status         = 'Success'
                        CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    }
                }
            }
            catch {
                # Step 4 -- Return error object on failure
                [PSCustomObject]@{
                    ComputerName   = $computer
                    Name           = $null
                    Version        = $null
                    Publisher      = $null
                    Architecture   = $null
                    Status         = 'Failed'
                    Error          = $_.Exception.Message
                }
            }
        }
    }
}
