# ============================================================
# FUNCTION : Get-VBAzureADJoinStatus
# VERSION  : 1.0.2
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Retrieve detailed Azure AD join status from target computer
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Retrieve detailed Azure AD join status from target computer(s).

.DESCRIPTION
    Executes dsregcmd /status command on target computer to retrieve comprehensive
    Azure AD join status information including device ID, tenant ID, and join state.
    Returns full status output or error message if not joined.

.PARAMETER ComputerName
    Target computer(s). Defaults to local machine. Accepts pipeline input.
    Aliases: Name, Server, Host

.PARAMETER Credential
    Alternate credentials for remote execution.

.EXAMPLE
    Get-VBAzureADJoinStatus
    Retrieves Azure AD join status from local computer.

.EXAMPLE
    Get-VBAzureADJoinStatus -ComputerName SERVER01
    Retrieves Azure AD join status from SERVER01.

.EXAMPLE
    'SERVER01', 'SERVER02' | Get-VBAzureADJoinStatus
    Retrieves Azure AD join status from multiple computers via pipeline.

.OUTPUTS
    [PSCustomObject]: ComputerName, Status, CollectionTime, Details (or Error)

.NOTES
    Version  : 1.0.2
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : Security
#>

function Get-VBAzureADJoinStatus {
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
                        return @{
                            Details = $statusOutput
                            IsJoined = $true
                        }
                    } else {
                        return @{
                            Details = 'Not joined'
                            IsJoined = $false
                        }
                    }
                }

                # Step 2 -- Execute locally or remotely
                if ($computer -eq $env:COMPUTERNAME) {
                    $result = & $scriptBlock
                } else {
                    $splat = @{
                        ComputerName = $computer
                        ScriptBlock = $scriptBlock
                    }
                    if ($Credential) {
                        $splat['Credential'] = $Credential
                    }
                    $result = Invoke-Command @splat
                }

                # Step 3 -- Output result
                [PSCustomObject]@{
                    ComputerName = $computer
                    Status = if ($result.IsJoined) { 'Joined' } else { 'Not Joined' }
                    Details = $result.Details
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
