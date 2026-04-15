# ============================================================
# FUNCTION : Get-VBUnusedGPOs
# VERSION  : 1.0.2
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Identifies Group Policy Objects that are not linked to any domain or OU
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Identifies Group Policy Objects that are not linked to any domain or OU.

.DESCRIPTION
    Scans all GPOs in the domain and identifies those that have no links to any organizational units
    or domains. Returns detailed information including version numbers and creation/modification times.

.PARAMETER ComputerName
    Domain Controller to query. Defaults to local machine. Accepts pipeline input.

.PARAMETER Credential
    Alternate credentials for the GPO query.

.EXAMPLE
    Get-VBUnusedGPOs

.EXAMPLE
    Get-VBUnusedGPOs -ComputerName DC01

.EXAMPLE
    'DC01','DC02' | Get-VBUnusedGPOs -Credential (Get-Credential)

.OUTPUTS
    [PSCustomObject]: ComputerName, Name, UserVersion, ComputerVersion, Created, Modified, ID, Status, CollectionTime

.NOTES
    Version  : 1.0.2
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : AD / GPO
#>

function Get-VBUnusedGPOs {
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
                $AllGpos = Get-GPO -All @GpoParams

                # Step 3 -- Identify unused GPOs
                $UnusedGpos = @()

                foreach ($gpo in $AllGpos) {
                    # Step 4 -- Generate XML report for each GPO
                    $XmlReport = Get-GPOReport -Guid $gpo.Id -ReportType Xml @GpoParams

                    # Step 5 -- Load into XML object
                    [xml]$Doc = $XmlReport

                    # Step 6 -- Count link nodes under GPO/LinksTo
                    $LinkCount = $Doc.GPO.LinksTo.Link.Count

                    # Step 7 -- Add to unused list if no links found
                    if ($LinkCount -eq 0) {
                        $UnusedGpos += [PSCustomObject]@{
                            ComputerName    = $computer
                            Name            = $gpo.DisplayName
                            UserVersion     = [int]$Doc.GPO.UserVersion
                            ComputerVersion = [int]$Doc.GPO.ComputerVersion
                            Created         = $gpo.CreationTime
                            Modified        = $gpo.ModificationTime
                            ID              = $gpo.Id
                            Status          = 'Success'
                            CollectionTime  = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                        }
                    }
                }

                # Step 8 -- Return results
                if ($UnusedGpos.Count -eq 0) {
                    [PSCustomObject]@{
                        ComputerName   = $computer
                        Message        = "No unused GPOs found."
                        Status         = 'Success'
                        CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    }
                } else {
                    $UnusedGpos | Sort-Object Created
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
