# ============================================================
# FUNCTION : Get-VBAzureADJoinStatusSimple
# VERSION  : 1.0.0
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu
# PURPOSE  : Simple yes/no Azure AD join status indicator
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Simple Azure AD join status check from target computer(s).

.DESCRIPTION
    Quick check for Azure AD join status. Returns simple Joined or Not Joined
    status via PSCustomObject. Does not return full dsregcmd output details.

.PARAMETER ComputerName
    Target computer(s). Defaults to local machine. Accepts pipeline input.
    Aliases: Name, Server, Host

.PARAMETER Credential
    Alternate credentials for remote execution.

.EXAMPLE
    Get-VBAzureADJoinStatusSimple
    Checks local computer's Azure AD join status.

.EXAMPLE
    Get-VBAzureADJoinStatusSimple -ComputerName SERVER01
    Checks SERVER01's Azure AD join status.

.EXAMPLE
    'SERVER01', 'SERVER02' | Get-VBAzureADJoinStatusSimple
    Checks multiple computers via pipeline.

.OUTPUTS
    [PSCustomObject]: ComputerName, Status, CollectionTime

.NOTES
    Version  : 1.0.0
    Author   : Vibhu
    Modified : 10-04-2026
    Category : Security
#>

function Get-VBAzureADJoinStatusSimple {
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
                # Step 1 -- Define remote script block
                $scriptBlock = {
                    $statusOutput = dsregcmd /status 2>&1
                    if ($statusOutput) {
                        return $true
                    } else {
                        return $false
                    }
                }

                # Step 2 -- Execute locally or remotely
                if ($computer -eq $env:COMPUTERNAME) {
                    $isJoined = & $scriptBlock
                } else {
                    $splat = @{
                        ComputerName = $computer
                        ScriptBlock = $scriptBlock
                    }
                    if ($Credential) {
                        $splat['Credential'] = $Credential
                    }
                    $isJoined = Invoke-Command @splat
                }

                # Step 3 -- Output result
                [PSCustomObject]@{
                    ComputerName = $computer
                    Status = if ($isJoined) { 'Joined' } else { 'Not Joined' }
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
            catch {
                [PSCustomObject]@{
                    ComputerName = $computer
                    Status = 'Failed'
                    Error = $_.Exception.Message
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
