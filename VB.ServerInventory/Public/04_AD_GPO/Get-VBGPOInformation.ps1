# ============================================================
# FUNCTION : Get-VBGPOInformation
# VERSION  : 1.0.0
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Retrieves all Group Policy Objects with detailed properties and modification history
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Retrieves all Group Policy Objects with detailed properties.

.DESCRIPTION
    Fetches all GPOs from the domain and returns their properties including ID, display name, status,
    modification time, creation time, and description.

.PARAMETER ComputerName
    Domain Controller to query. Defaults to local machine. Accepts pipeline input.

.PARAMETER Credential
    Alternate credentials for the GPO query.

.EXAMPLE
    Get-VBGPOInformation

.EXAMPLE
    Get-VBGPOInformation -ComputerName DC01

.EXAMPLE
    'DC01','DC02' | Get-VBGPOInformation -Credential (Get-Credential)

.OUTPUTS
    [PSCustomObject]: ComputerName, Id, DisplayName, GpoStatus, ModificationTime, CreationTime,
    Description, Status, CollectionTime

.NOTES
    Version  : 1.0.0
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : AD / GPO
#>

function Get-VBGPOInformation {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'Server', 'Host')]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [PSCredential]$Credential
    )
    begin {
        Import-Module GroupPolicy -ErrorAction Stop
    }



    process {
        foreach ($computer in $ComputerName) {
            try {
                # Step 1 -- Prepare GPO query parameters
                $GpoParams = @{}
                if ($computer -ne $env:COMPUTERNAME) {
                    $GpoParams['Server'] = $computer
                }
                if ($Credential) {
                    $GpoParams['Credential'] = $Credential
                }

                # Step 2 -- Retrieve all GPOs
                $Gpos = Get-GPO -All @GpoParams

                # Step 3 -- Check if GPOs were found
                if (-not $Gpos) {
                    [PSCustomObject]@{
                        ComputerName   = $computer
                        Message        = "No Group Policy Objects found in the domain."
                        Status         = 'Failed'
                        CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    }
                } else {
                    # Step 4 -- Return each GPO as individual object
                    foreach ($gpo in $Gpos) {
                        [PSCustomObject]@{
                            ComputerName     = $computer
                            Id               = $gpo.Id
                            DisplayName      = $gpo.DisplayName
                            GpoStatus        = $gpo.GpoStatus
                            ModificationTime = $gpo.ModificationTime
                            CreationTime     = $gpo.CreationTime
                            Description      = $gpo.Description
                            Status           = 'Success'
                            CollectionTime   = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                        }
                    }
                }
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
