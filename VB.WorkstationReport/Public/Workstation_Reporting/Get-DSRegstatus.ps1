# ============================================================
# FUNCTION : Get-DSRegstatus
# VERSION  : 2.0.0
# CHANGED  : 24-05-2026 -- Merged flat PSCustomObject output with remote/pipeline support
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Retrieve detailed Azure AD join status from target computer(s)
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Retrieve detailed Azure AD join status from target computer(s).

.DESCRIPTION
    Executes dsregcmd /status on target computer(s) and returns a flat PSCustomObject
    with all keys as named properties — spaces stripped for clean Select-Object usage.
    Supports local execution, remote via Invoke-Command, pipeline input, and alternate credentials.

.PARAMETER ComputerName
    Target computer(s). Defaults to local machine. Accepts pipeline input.
    Aliases: Name, Server, Host

.PARAMETER Credential
    Alternate credentials for remote execution.

.EXAMPLE
    Get-DSRegstatus
    Retrieves Azure AD join status from local computer.

.EXAMPLE
    Get-DSRegstatus -ComputerName SERVER01
    Retrieves Azure AD join status from SERVER01.

.EXAMPLE
    'SERVER01', 'SERVER02' | Get-DSRegstatus
    Pipeline input — retrieves status from multiple computers.

.EXAMPLE
    Get-DSRegstatus | Select-Object ComputerName, AzureAdJoined, TenantName, DeviceName
    Select specific properties from flat output.

.OUTPUTS
    [PSCustomObject]: ComputerName, Status, CollectionTime + all dsregcmd keys as properties

.NOTES
    Version  : 2.0.0
    Author   : Vibhu Bhatnagar
    Modified : 24-05-2026
    Category : Security
#>
# Note :- Get-VBAzureADJoinStatus is same copy of Get-DSRegstatus with different function names o that code will not break
function Get-DSRegstatus {
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
                # Step 1 -- Script block: parse dsregcmd into key/value pairs
                $scriptBlock = {
                    $raw = dsregcmd /status 2>&1 | Select-String ":" |
                        ForEach-Object {
                            $parts = $_.Line.Trim() -split "\s*:\s*", 2
                            if ($parts.Count -eq 2) {
                                [PSCustomObject]@{
                                    Key   = $parts[0].Trim() -replace '\s+', ''
                                    Value = $parts[1].Trim()
                                }
                            }
                        }
                    return $raw
                }

                # Step 2 -- Execute locally or remotely
                if ($computer -eq $env:COMPUTERNAME) {
                    $raw = & $scriptBlock
                } else {
                    $splat = @{
                        ComputerName = $computer
                        ScriptBlock  = $scriptBlock
                    }
                    if ($Credential) { $splat['Credential'] = $Credential }
                    $raw = Invoke-Command @splat
                }

                # Step 3 -- Build flat object, seed with metadata properties
                $obj = [PSCustomObject]@{
                    ComputerName   = $computer
                    Status         = 'Joined'
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }

                # Step 4 -- Add all dsregcmd keys as named properties
                foreach ($item in $raw) {
                    if ($item.Key -and -not ($obj.PSObject.Properties.Name -contains $item.Key)) {
                        $obj | Add-Member -NotePropertyName $item.Key -NotePropertyValue $item.Value
                    }
                }

                # Step 5 -- Set Status based on AzureAdJoined property
                if ($obj.PSObject.Properties.Name -contains 'AzureAdJoined') {
                    $obj.Status = if ($obj.AzureAdJoined -eq 'YES') { 'Joined' } else { 'Not Joined' }
                }

                $obj
            }
            catch {
                [PSCustomObject]@{
                    ComputerName   = $computer
                    Status         = 'Failed'
                    Error          = $_.Exception.Message
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
