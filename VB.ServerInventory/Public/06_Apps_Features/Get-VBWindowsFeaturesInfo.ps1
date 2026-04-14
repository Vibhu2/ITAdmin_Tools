# ============================================================
# FUNCTION : Get-VBWindowsFeaturesInfo
# VERSION  : 1.0.0
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu
# PURPOSE  : Query Windows Features, Roles, and Role Services status
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Retrieves Windows Features, Roles, and Role Services from local or remote computers.

.DESCRIPTION
    Enumerates installed Windows Features on Server SKUs (or Windows Optional
    Features on client SKUs). Returns one PSCustomObject per feature with name,
    display name, installation state, and feature type. Attempts Get-WindowsFeature
    (Server) first, then falls back to Get-WindowsOptionalFeature (client).
    Supports pipeline input and remote execution via WinRM.

.PARAMETER ComputerName
    Target computer(s) to query. Accepts pipeline input. Defaults to local
    machine ($env:COMPUTERNAME). Aliases: Name, Server, Host.

.PARAMETER Credential
    PSCredential object for remote execution. Required for cross-domain or
    credential-based remote queries.

.EXAMPLE
    Get-VBWindowsFeaturesInfo

.EXAMPLE
    Get-VBWindowsFeaturesInfo -ComputerName SERVER01

.EXAMPLE
    'SRV01','SRV02' | Get-VBWindowsFeaturesInfo -Credential (Get-Credential)

.OUTPUTS
    [PSCustomObject]: ComputerName, Name, DisplayName, InstallState, FeatureType,
    Status, CollectionTime

.NOTES
    Version  : 1.0.0
    Author   : Vibhu
    Modified : 10-04-2026
    Category : Features

#>

function Get-VBWindowsFeaturesInfo {
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
                # Step 1 -- Define script block to enumerate features
                $scriptBlock = {
                    $Features = $null

                    # Step 2 -- Try Get-WindowsFeature for Server SKUs
                    if (Get-Command -Name Get-WindowsFeature -ErrorAction SilentlyContinue) {
                        $Features = Get-WindowsFeature -ErrorAction SilentlyContinue |
                            Where-Object { $_.InstallState -eq 'Installed' }
                    }
                    # Step 3 -- Fall back to Get-WindowsOptionalFeature for client SKUs
                    elseif (Get-Command -Name Get-WindowsOptionalFeature -ErrorAction SilentlyContinue) {
                        $Features = Get-WindowsOptionalFeature -Online -ErrorAction SilentlyContinue |
                            Where-Object { $_.State -eq 'Enabled' } |
                            Select-Object -Property @{
                                Name       = 'Name'
                                Expression = { $_.FeatureName }
                            },
                            @{
                                Name       = 'DisplayName'
                                Expression = { $_.FeatureName }
                            },
                            @{
                                Name       = 'InstallState'
                                Expression = { 'Installed' }
                            },
                            @{
                                Name       = 'FeatureType'
                                Expression = { 'Feature' }
                            }
                    }

                    # Step 4 -- Return results if found
                    if ($Features) {
                        $Features | Sort-Object Name
                    }
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

                # Step 6 -- Emit one PSCustomObject per feature
                foreach ($item in $results) {
                    [PSCustomObject]@{
                        ComputerName   = $computer
                        Name           = $item.Name
                        DisplayName    = $item.DisplayName
                        InstallState   = $item.InstallState
                        FeatureType    = $item.FeatureType
                        Status         = 'Success'
                        CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    }
                }
            }
            catch {
                # Step 7 -- Return error object on failure
                [PSCustomObject]@{
                    ComputerName   = $computer
                    Name           = $null
                    DisplayName    = $null
                    InstallState   = $null
                    FeatureType    = $null
                    Status         = 'Failed'
                    Error          = $_.Exception.Message
                }
            }
        }
    }
}
