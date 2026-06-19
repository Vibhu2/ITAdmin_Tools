function Get-VBUserCustomProperties {
    <#
    .SYNOPSIS
        Gets user custom properties report
    
    .DESCRIPTION
        Retrieves users with custom properties like ProfilePath, ScriptPath, HomeDrive, HomeDirectory
    
    .PARAMETER Limit
        Maximum number of users to return (default: 100)
    #>
    
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [int]$Limit = 100
    )
    
    try {
        Write-Host "User Custom Properties" -ForegroundColor Cyan
        
        $users = Get-ADUser -Filter * -Properties SamAccountName, ProfilePath, ScriptPath, HomeDrive, HomeDirectory | 
                Where-Object { 
                    -not [string]::IsNullOrEmpty($_.ProfilePath) -or 
                    -not [string]::IsNullOrEmpty($_.ScriptPath) -or 
                    -not [string]::IsNullOrEmpty($_.HomeDrive) -or 
                    -not [string]::IsNullOrEmpty($_.HomeDirectory)
                } | 
                Select-Object -First $Limit

        if ($users.Count -gt 0) {
            $userPropertiesTable = foreach ($user in $users) {
                [PSCustomObject]@{
                    'SamAccountName' = $user.SamAccountName
                    'ProfilePath' = if ([string]::IsNullOrEmpty($user.ProfilePath)) { "N/A" } else { $user.ProfilePath }
                    'LogonScript' = if ([string]::IsNullOrEmpty($user.ScriptPath)) { "N/A" } else { $user.ScriptPath }
                    'HomeDrive' = if ([string]::IsNullOrEmpty($user.HomeDrive)) { "N/A" } else { $user.HomeDrive }
                    'HomeDirectory' = if ([string]::IsNullOrEmpty($user.HomeDirectory)) { "N/A" } else { $user.HomeDirectory }
                }
            }
            
            $userPropertiesTable | Format-Table -AutoSize
            
            if ($users.Count -eq $Limit) {
                Write-Host "Note: Showing first $Limit users with custom properties." -ForegroundColor Yellow
            }
        } else {
            Write-Host "No users found with custom properties." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "ERROR: Unable to retrieve user custom properties - $($_.Exception.Message)" -ForegroundColor Red
    }
}