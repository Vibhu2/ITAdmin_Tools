function Get-VBAzureADJoinStatus {
    <#
    .SYNOPSIS
    Retrieves Azure AD join status information from local or remote computers.

    .DESCRIPTION
    The Get-VBAzureADJoinStatus function uses dsregcmd.exe to determine Azure AD join status
    and returns structured information about device registration, workplace join status,
    and Azure AD connectivity. Supports both local and remote computer queries with 
    pipeline input for bulk operations. Use -RawOutput to see the complete dsregcmd output.

    .PARAMETER ComputerName
    Specifies the computer name(s) to query for Azure AD join status. Supports multiple 
    computers via pipeline input. Defaults to the local computer if not specified.
    Accepts pipeline input by value and by property name.

    .PARAMETER Credential
    Specifies credentials for authenticating to remote computers. Not required for 
    local computer queries. Use Get-Credential to create credential objects.

    .PARAMETER RawOutput
    Returns the raw dsregcmd output directly to the console instead of parsed objects.
    Useful for troubleshooting or when you need to see all available information.

    .EXAMPLE
    Get-VBAzureADJoinStatus
    
    ComputerName    : DESKTOP-ABC123
    AzureADJoined   : True
    WorkplaceJoined : False
    DomainJoined    : False
    DeviceId        : 12345678-1234-1234-1234-123456789012
    TenantId        : 87654321-4321-4321-4321-210987654321
    TenantName      : Contoso Corporation
    UserEmail       : user@contoso.com
    Status          : Success

    Returns Azure AD join status for the local computer with structured output.

    .EXAMPLE
    Get-VBAzureADJoinStatus -ComputerName "SERVER01" -Credential $cred
    
    Gets Azure AD join status from a remote server using specified credentials.

    .EXAMPLE
    "SERVER01", "SERVER02", "DESKTOP01" | Get-VBAzureADJoinStatus | Where-Object {$_.AzureADJoined -eq $true}
    
    Gets Azure AD join status from multiple computers via pipeline and filters for 
    devices that are Azure AD joined only.

    .EXAMPLE
    Get-VBAzureADJoinStatus -ComputerName "SERVER01" -RawOutput
    
    === SERVER01 ===
    
    | Device State              |
    +---------------------------+
    
             AzureADJoined : YES
          EnterpriseJoined : NO
              DomainJoined : YES
                   DeviceId : 12345678-1234-1234-1234-123456789012
    
    Shows the complete raw dsregcmd output for detailed analysis.

    .OUTPUTS
    System.Management.Automation.PSCustomObject (when -RawOutput is not used)
    Returns objects with the following properties:
    - ComputerName: Target computer name
    - AzureADJoined: Boolean indicating if device is Azure AD joined
    - WorkplaceJoined: Boolean indicating if device is workplace joined  
    - DomainJoined: Boolean indicating if device is domain joined
    - DeviceId: Azure AD device ID (if available)
    - TenantId: Azure AD tenant ID (if available)
    - TenantName: Azure AD tenant name (if available)
    - UserEmail: Current user's email address (if available)
    - Status: Operation status ('Success' or 'Failed')
    - Error: Error message (only present when Status is 'Failed')

    System.String (when -RawOutput is used)
    Returns the raw dsregcmd command output directly to the console.

    .NOTES
    Version: 1.0
    Author: System Administrator
    Category: Azure AD Management
    
    Requirements:
    - dsregcmd.exe must be available on target systems (Windows 10/Server 2016+)
    - For remote computers, WinRM must be enabled and properly configured
    - Function supports both Windows PowerShell 5.1 and PowerShell Core
    - Requires appropriate permissions to run dsregcmd on target systems
    
    Common Use Cases:
    - Verify Azure AD join status across multiple computers
    - Troubleshoot hybrid Azure AD join issues
    - Audit device registration status in bulk operations
    - Extract device and tenant information for reporting
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'Server')]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        
        [PSCredential]$Credential,
        
        [switch]$RawOutput
    )
    
    process {
        foreach ($computer in $ComputerName) {
            try {
                if ($computer -eq $env:COMPUTERNAME) {
                    $result = dsregcmd /status 2>&1
                } else {
                    $result = Invoke-Command -ComputerName $computer -Credential $Credential -ScriptBlock {
                        dsregcmd /status 2>&1
                    }
                }
                
                if ($RawOutput) {
                    Write-Output "=== $computer ==="
                    $result
                } else {
                    # Parse dsregcmd output
                    $statusText = $result -join "`n"
                    
                    $azureADJoined = $statusText -match "AzureAdJoined\s*:\s*YES"
                    $workplaceJoined = $statusText -match "WorkplaceJoined\s*:\s*YES"
                    $domainJoined = $statusText -match "DomainJoined\s*:\s*YES"
                    
                    $deviceId = if ($statusText -match "DeviceId\s*:\s*([a-f0-9\-]+)") { $matches[1] } else { $null }
                    $tenantId = if ($statusText -match "TenantId\s*:\s*([a-f0-9\-]+)") { $matches[1] } else { $null }
                    $tenantName = if ($statusText -match "TenantName\s*:\s*([^`r`n]+)") { 
                        $matches[1].Trim() -replace "TenantId\s*:\s*[a-f0-9\-]+", "" 
                    } else { $null }
                    $userEmail = if ($statusText -match "UserEmail\s*:\s*([^`r`n]+)") { 
                        $emailValue = $matches[1].Trim()
                        if ($emailValue -and $emailValue -ne "") { $emailValue } else { $null }
                    } else { $null }
                    
                    [PSCustomObject]@{
                        ComputerName = $computer
                        AzureADJoined = $azureADJoined
                        WorkplaceJoined = $workplaceJoined
                        DomainJoined = $domainJoined
                        DeviceId = $deviceId
                        TenantId = $tenantId
                        TenantName = $tenantName
                        UserEmail = $userEmail
                        Status = 'Success'
                    }
                }
            }
            catch {
                if ($RawOutput) {
                    Write-Output "=== $computer ==="
                    Write-Output "Error: $($_.Exception.Message)"
                } else {
                    [PSCustomObject]@{
                        ComputerName = $computer
                        Status = 'Failed'
                        Error = $_.Exception.Message
                    }
                }
            }
        }
    }
}