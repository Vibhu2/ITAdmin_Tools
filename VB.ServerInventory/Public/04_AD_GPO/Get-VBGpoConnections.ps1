# ============================================================
# FUNCTION : Get-VBGpoConnections
# VERSION  : 1.0.0
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu
# PURPOSE  : Reports all GPO connections to organizational units including enforcement status
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Reports all GPO connections to organizational units.

.DESCRIPTION
    Retrieves all organizational units in the domain and identifies which GPOs are linked to each,
    including information about enforcement status and link order.

.PARAMETER ComputerName
    Domain Controller to query. Defaults to local machine. Accepts pipeline input.

.PARAMETER Credential
    Alternate credentials for the AD/GPO query.

.EXAMPLE
    Get-VBGpoConnections

.EXAMPLE
    Get-VBGpoConnections -ComputerName DC01

.EXAMPLE
    'DC01','DC02' | Get-VBGpoConnections -Credential (Get-Credential)

.OUTPUTS
    [PSCustomObject]: ComputerName, GPO, OU, Enforced, LinkOrder, Status, CollectionTime

.NOTES
    Version  : 1.0.0
    Author   : Vibhu
    Modified : 10-04-2026
    Category : AD / GPO
#>

function Get-VBGpoConnections {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'Server', 'Host')]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [PSCredential]$Credential
    )
    begin {
        Import-Module ActiveDirectory -ErrorAction Stop
        Import-Module GroupPolicy -ErrorAction Stop
    }



    process {
        foreach ($computer in $ComputerName) {
            try {
                # Step 1 -- Prepare AD/GPO query parameters
                $AdParams = @{}
                if ($computer -ne $env:COMPUTERNAME) {
                    $AdParams['Server'] = $computer
                }
                if ($Credential) {
                    $AdParams['Credential'] = $Credential
                }

                # Step 2 -- Retrieve all organizational units
                $Ous = Get-ADOrganizationalUnit -Filter * @AdParams

                # Step 3 -- Process each OU for GPO connections
                $Results = @()

                foreach ($ou in $Ous) {
                    # Step 4 -- Get GPO inheritance information for this OU
                    $Inheritance = Get-GPInheritance -Target $ou.DistinguishedName @AdParams

                    # Step 5 -- Process each GPO link
                    foreach ($link in $Inheritance.GpoLinks) {
                        # Step 6 -- Extract OU name from Distinguished Name
                        $OuName = $ou.DistinguishedName
                        if ($ou.DistinguishedName -match '^OU=([^,]+)') {
                            $OuName = $matches[1]
                        }

                        # Step 7 -- Create result object for each link
                        $Results += [PSCustomObject]@{
                            ComputerName   = $computer
                            GPO            = $link.DisplayName
                            OU             = $OuName
                            Enforced       = $link.Enforced
                            LinkOrder      = $link.Order
                            Status         = 'Success'
                            CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                        }
                    }
                }

                # Step 8 -- Return results
                $Results
            } catch {
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
