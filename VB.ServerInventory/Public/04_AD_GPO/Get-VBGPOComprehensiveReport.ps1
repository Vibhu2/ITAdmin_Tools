# ============================================================
# FUNCTION : Get-VBGPOComprehensiveReport
# VERSION  : 1.0.2
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Generates comprehensive GPO report with all link scopes and configurations
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Generates comprehensive GPO report with all link scopes and configurations.

.DESCRIPTION
    Retrieves all GPOs and their link information, including scope paths, enabled status,
    enforcement settings, and GPO status details. Can filter to show only linked GPOs or all GPOs.

.PARAMETER ComputerName
    Domain Controller to query. Defaults to local machine. Accepts pipeline input.

.PARAMETER Credential
    Alternate credentials for the GPO query.

.PARAMETER ShowAll
    If specified, shows all GPOs including those with no links. If omitted, only shows GPOs with links.

.EXAMPLE
    Get-VBGPOComprehensiveReport

.EXAMPLE
    Get-VBGPOComprehensiveReport -ComputerName DC01

.EXAMPLE
    Get-VBGPOComprehensiveReport -ShowAll

.EXAMPLE
    'DC01','DC02' | Get-VBGPOComprehensiveReport -Credential (Get-Credential)

.OUTPUTS
    [PSCustomObject]: ComputerName, GPOName, GPOID, LinkScope, LinkEnabled, Enforced, GPOStatus,
    CreatedTime, ModifiedTime, Status, CollectionTime

.NOTES
    Version  : 1.0.2
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : AD / GPO
#>

function Get-VBGPOComprehensiveReport {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'Server', 'Host')]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [PSCredential]$Credential,
        [switch]$ShowAll
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

                # Step 3 -- Process each GPO
                foreach ($gpo in $Gpos) {
                    # Step 4 -- Get GPO report in XML format
                    $Report = Get-GPOReport -Guid $gpo.Id -ReportType Xml @GpoParams
                    [xml]$Xml = $Report

                    # Step 5 -- Process each link scope
                    $Links = @()

                    foreach ($scope in $Xml.GPO.LinksTo) {
                        $LinkObject = [PSCustomObject]@{
                            ComputerName   = $computer
                            GPOName        = $gpo.DisplayName
                            GPOID          = $gpo.Id
                            LinkScope      = $scope.SOMPath
                            LinkEnabled    = $scope.Enabled
                            Enforced       = $scope.NoOverride
                            GPOStatus      = $gpo.GpoStatus
                            CreatedTime    = $gpo.CreationTime
                            ModifiedTime   = $gpo.ModificationTime
                            Status         = 'Success'
                            CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                        }

                        $Links += $LinkObject
                    }

                    # Step 6 -- Return results based on ShowAll flag
                    if ($ShowAll -or $Links.Count -gt 0) {
                        if ($Links.Count -gt 0) {
                            $Links
                        } else {
                            # Return single object for unlinked GPO when ShowAll is specified
                            [PSCustomObject]@{
                                ComputerName   = $computer
                                GPOName        = $gpo.DisplayName
                                GPOID          = $gpo.Id
                                LinkScope      = "UNLINKED"
                                LinkEnabled    = $null
                                Enforced       = $null
                                GPOStatus      = $gpo.GpoStatus
                                CreatedTime    = $gpo.CreationTime
                                ModifiedTime   = $gpo.ModificationTime
                                Status         = 'Success'
                                CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                            }
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
