function Get-VBGPOInformation {
    <#
    .SYNOPSIS
    Retrieves comprehensive information about all Group Policy Objects in the domain.

    .DESCRIPTION
    This function queries the domain for all Group Policy Objects and returns detailed information including ID, display name, status, modification time, creation time, and description. The results are sorted by modification time and displayed in a formatted table for easy analysis.

    .PARAMETER None
    This function does not accept any parameters.

    .EXAMPLE
    Get-VBGPOInformation
    Get-VBGPOInformation | Format-Table -AutoSize
    Retrieves and displays all GPOs in the current domain with their key properties sorted by modification time.

    .EXAMPLE
    Get-VBGPOInformation | Export-Csv -Path "C:\Reports\GPOReport.csv" -NoTypeInformation
    Exports all GPO information to a CSV file for further analysis or documentation.

    .EXAMPLE
    $gpoData = Get-VBGPOInformation
    $gpoData | Where-Object {$_.GpoStatus -eq 'AllSettingsDisabled'}
    Retrieves GPO information and filters to show only disabled GPOs for cleanup purposes.

    .OUTPUTS
    System.Object[]
    Returns an array of objects containing GPO information with properties: Id, DisplayName, GpoStatus, ModificationTime, CreationTime, Description.

    .NOTES
    Version: 1.0
    Author: VB Admin Tools
    Category: Group Policy Management
    Requires: GroupPolicy PowerShell module and appropriate domain permissions
    Compatible: PowerShell 5.1+
    #>

    [CmdletBinding()]
    param()

    try {
        # Ensure the GroupPolicy module is loaded
        if (-not (Get-Module -Name GroupPolicy -ListAvailable)) {
            Import-Module GroupPolicy -ErrorAction Stop
        }

        # Retrieve all GPOs
        $gpos = Get-GPO -All

        # If no GPOs were found, return error object
        if (-not $gpos) {
            [PSCustomObject]@{
                Status = 'Failed'
                Error = 'No Group Policy Objects found in the domain'
                GPOCount = 0
            }
            return
        }

        # Select and return all desired properties
        $gpos | 
            Select-Object Id, DisplayName, GpoStatus, ModificationTime, CreationTime, Description |
            Sort-Object -Property ModificationTime
    }
    catch {
        [PSCustomObject]@{
            Status = 'Failed'
            Error = $_.Exception.Message
            GPOCount = 0
        }
    }
}