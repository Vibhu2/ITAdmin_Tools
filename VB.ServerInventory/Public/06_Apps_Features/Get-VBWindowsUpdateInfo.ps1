# ============================================================
# FUNCTION : Get-VBWindowsUpdateInfo
# VERSION  : 1.0.2
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Query installed Windows Updates (hotfixes) from local or remote systems
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Retrieves installed Windows Updates (hotfixes) from local or remote computers.

.DESCRIPTION
    Enumerates all installed Windows hotfixes and KB articles. Returns one
    PSCustomObject per update with description, KB ID, and installation date.
    Supports pipeline input and remote execution via WinRM.

.PARAMETER ComputerName
    Target computer(s) to query. Accepts pipeline input. Defaults to local
    machine ($env:COMPUTERNAME). Aliases: Name, Server, Host.

.PARAMETER Credential
    PSCredential object for remote execution. Required for cross-domain or
    credential-based remote queries.

.EXAMPLE
    Get-VBWindowsUpdateInfo

.EXAMPLE
    Get-VBWindowsUpdateInfo -ComputerName SERVER01

.EXAMPLE
    'SRV01','SRV02' | Get-VBWindowsUpdateInfo -Credential (Get-Credential)

.OUTPUTS
    [PSCustomObject]: ComputerName, Description, HotFixID, InstalledOn,
    Status, CollectionTime

.NOTES
    Version  : 1.0.2
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : Updates

#>

function Get-VBWindowsUpdateInfo {
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
                # Step 1 -- Define script block to query hotfixes
                $scriptBlock = {
                    Get-HotFix -ErrorAction SilentlyContinue | Sort-Object InstalledOn -Descending
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

                # Step 3 -- Emit one PSCustomObject per update
                foreach ($item in $results) {
                    [PSCustomObject]@{
                        ComputerName   = $computer
                        Description    = $item.Description
                        HotFixID       = $item.HotFixID
                        InstalledOn    = $item.InstalledOn
                        Status         = 'Success'
                        CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    }
                }
            }
            catch {
                # Step 4 -- Return error object on failure
                [PSCustomObject]@{
                    ComputerName   = $computer
                    Description    = $null
                    HotFixID       = $null
                    InstalledOn    = $null
                    Status         = 'Failed'
                    Error          = $_.Exception.Message
                }
            }
        }
    }
}
