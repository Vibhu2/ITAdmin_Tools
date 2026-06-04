function Get-VBJoinType {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, Mandatory = $true)]
        [PSCustomObject]$DSRegStatus
    )

    process {
        $azureJoined = $DSRegStatus.AzureAdJoined -eq 'YES'
        $domainJoined = $DSRegStatus.DomainJoined  -eq 'YES'

        $DSRegStatus | Select-Object ComputerName, AzureAdJoined, DomainJoined, @{
            Name       = 'JoinType'
            Expression = {
                if ($azureJoined -and $domainJoined) { 'Hybrid (AD + Entra ID)' }
                elseif ($azureJoined)                { 'Entra ID Only (Cloud)'  }
                elseif ($domainJoined)               { 'On-Prem AD Only'        }
                else                                 { 'Workgroup'              }
            }
        }
    }
}
